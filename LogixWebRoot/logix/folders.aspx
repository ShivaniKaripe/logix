<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" EnableSessionState="True"%>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>

<%
    ' *****************************************************************************
    ' * FILENAME: folders.aspx 
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
    Dim rst As System.Data.DataTable
    Dim AdminUserID As Long
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim dt As DataTable
    dim index as Integer
    Dim row As DataRow
    Dim ShowActionButton As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannersEnabled As Boolean = False
    Const MaxBigIntLength As Integer = 19
    Dim bProductionSystem As Boolean = True
    Dim bTestSystem As Boolean = False
    Dim bArchiveSystem As Boolean = False
    Dim bWorkflowActive As Boolean = False
    Dim bAllowTimeWithStartEndDates As Boolean = False

    bTestSystem = (MyCommon.Fetch_CM_SystemOption(77) = "1")
    bArchiveSystem = (MyCommon.Fetch_CM_SystemOption(77) = "2")
    bAllowTimeWithStartEndDates = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(200) = "1")

    If bTestSystem Or bArchiveSystem Then
        bProductionSystem = False
    Else
        bProductionSystem = True
    End If

    MyCommon.AppName = "folders.aspx"
    Response.Expires = 0
    MyCommon.Open_LogixRT()

    If MyCommon.IsEngineInstalled(0) Then
        bWorkflowActive = (MyCommon.Fetch_CM_SystemOption(74) = "1")
    End If

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    Send_HeadBegin("term.folders")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts(New String() {"datePicker.js", "popup.js", "jquery.min.js"})
  %>
<style type="text/css">
  body {
    overflow: visible;
  }
  .OfferNavigationDialog {
  display: none;
  height: 200px;
  margin: 0px auto;
  position: relative;
  top: 270px;
  width: 400px;
  z-index: 99;
}
  #foldersearch {
    top: 140px;
  }
  #fsearchResults {
    height:290px;
  }
  #execduplicateoffer
  {
   overflow-y: auto;   
  }
  #statusClose:hover {
    opacity: 1.0;
    filter: alpha(opacity=100); /* For IE8 and earlier */
}
#statusClose
{
  display: block;
  float:right;
  position:relative;
  height: 15px;
  opacity: 0.7;
    filter: alpha(opacity=70); /* For IE8 and earlier */
}
#Actionitems
{
    width:300px;
}
#Actionitems option 
{
    width:300px;
}
</style>
<%  
  Send_Scripts()
%>
<script type="text/javascript">

  <%Send_Calendar_Overrides(MyCommon) %>

  // constants
  var CREATE_FOLDER = 1;
  var RENAME_FOLDER = 2;
  var DELETE_FOLDER = 3;
  var VIEW_ITEMS = 4;
  var VIEW_AVAILABLE = 5;
  var REMOVE_ITEMS = 6;
  var ADD_ITEMS = 7;
  var ASSIGN_FOLDERS = 8;
  var FOLDER_SEARCH = 9;
  var GENERATE_DIV_FOLDER = 10;
  var MODIFY_FOLDER = 11;
  var DUP_OFFERS = 12;
  var MASSDEPLOY_OFFERS = 13;
  var NAVIGATETO_REPORTS  = 14;
  var SEND_OUTBOUND = 15;
  var WFSTAT_PREVALIDATE = 16;
  var WFSTAT_POSTVALIDATE = 17;
  var WFSTAT_READYTODEPLOY = 18;
  var FILTER_ITEMS = 19;
  var TRANSFER_OFFERS = 20;
  var FOLDERSTART_OFFER=21;
  var FOLDEREND_OFFER=22;
  var FOLDERSTARTEND_OFFER=23;
  var CHECK_DEFAULTFOLDER=24;
  var SUB_FOLDERS=25;
  var PARENT_FOLDERS=26;
  var DELETESELECTED_OFFERS=27;
  var MASSDEFERDEPLOY_OFFERS = 28;
  var SHOW_APPROVALACTIONITEMS = 29
  var MASSREQUESTAPPROVALFOROFFERS = 30;
  var DEPLOYANDAPPROVAL_REQUEST = 31;

  var selectedFolder = 0;
  var selectedItems = new Array();
  var selectedItemTypes = new Array();
  var selectedPromoEngines = new Array();
  var isCtrlDown = false;
  var isShiftDown = false;
  var lastCreatedFolderName = '';
  var folderstartdate = '';
  var folderenddate = '';
  var foldertheme = '';
  var datePickerDivID = "datepicker";
  var sourceFolder = 0; 
 var cancelClicked = false; 
 var IsNavtoReportsCancelled = false;

  function toggleDropdown() {
    if (document.getElementById("actionsmenu") != null) {
      bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
      if (bOpen) {
        document.getElementById("actionsmenu").style.visibility = 'visible';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
      } else {
        document.getElementById("actionsmenu").style.visibility = 'hidden';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
      }
    }
  }
  
  function xmlhttpPost(strURL, action, frmdata) {
    var xmlHttpReq = false;
    var self = this;
    var tokens = new Array();
    if(action == VIEW_ITEMS){
        showContents( "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>"+'<p><b>Loading</b></p>'+ '<\/div>');
     }
    if (window.XMLHttpRequest) { // Mozilla/Safari
      self.xmlHttpReq = new XMLHttpRequest();
    } else if (window.ActiveXObject) { // IE
      self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function() {
          if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {     
                    // if a web page is sent back in the response then replace the parent window with the returned paged.
                    
                   if (action == NAVIGATETO_REPORTS && IsNavtoReportsCancelled){
                    IsNavtoReportsCancelled = false;
                    return;
                    }
        if (needsRedirect(self.xmlHttpReq.responseText)) {
          writeRedirectPage(self.xmlHttpReq.responseText);
          return;
        }
        switch (action) {
          case CREATE_FOLDER:
            handleCreateFolder(self.xmlHttpReq.responseText);
            break;
          case RENAME_FOLDER:
            confirmSuccess(self.xmlHttpReq.responseText);
            break;
          case DELETE_FOLDER:
            confirmSuccess(self.xmlHttpReq.responseText);
            break;
          case VIEW_ITEMS:
        
          if(self.xmlHttpReq.responseText.trim().indexOf("~FIU~")>=0)
            {
              showContents("<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/><p><b>"+self.xmlHttpReq.responseText.trim().replace("~FIU~","")+"</b></p>");
              highlightSelected(null, selectedFolder);
            }
          else{
              showContents(self.xmlHttpReq.responseText);
              var selectedfolder = document.getElementById("hdnfolderid").value;
              xmlhttpPost('/logix/folder-feeds.aspx?Action=GetSubFolders', SUB_FOLDERS, 'ParentFolderID=' + selectedFolder)
           }
            break;
          case VIEW_AVAILABLE:
            showSearchContents(self.xmlHttpReq.responseText);
            document.body.style.cursor = 'default';
            break; 
          case REMOVE_ITEMS:
            tokens = self.xmlHttpReq.responseText.split('|');
            // set the folder to accurately show whether the folder has content.
            if (tokens.length > 0) {
              setFolderIcon(selectedFolder, (parseInt(tokens[0]) > 0));
            }
            // show the remaining items assigned to the folder.
            if (tokens.length > 1) {
              showContents(tokens[1]);
            }
            break;
          case ADD_ITEMS:
            showContents(self.xmlHttpReq.responseText);
            pageIndex=1;
            GetRecords();
            setFolderIcon(selectedFolder, true);
            toggleDialog('folderpopulate', false);
            resetSearch();
            break;
          case FOLDER_SEARCH:
            showfSearchContents(self.xmlHttpReq.responseText);
            break;
          case GENERATE_DIV_FOLDER:
            generatedivfmodify(self.xmlHttpReq.responseText);
            break;
          case MODIFY_FOLDER:
            cofirmsuccessmodfolder(self.xmlHttpReq.responseText);
            break; 
          case DUP_OFFERS:
            cofirmsuccessdupoffer(self.xmlHttpReq.responseText);
            break;
          case MASSDEPLOY_OFFERS:
            ClosePopUp("");
            break;
          case MASSREQUESTAPPROVALFOROFFERS:
            ClosePopUp("");
            break; 
          case DEPLOYANDAPPROVAL_REQUEST:
            ClosePopUp("");
            break;
          case MASSDEFERDEPLOY_OFFERS:
            ClosePopUp("");
            break; 
          case  FOLDERSTART_OFFER:
                folderdateschange(self.xmlHttpReq.responseText);
                break; 
        case  FOLDEREND_OFFER:
                folderdateschange(self.xmlHttpReq.responseText);
                break; 
            case  FOLDERSTARTEND_OFFER:
            folderdateschange(self.xmlHttpReq.responseText);
            break; 
		  case  DELETESELECTED_OFFERS:
            deleteoffers(self.xmlHttpReq.responseText);
            break; 		
          case NAVIGATETO_REPORTS:
            handlenavtoreports(self.xmlHttpReq.responseText);
            break;
          case SEND_OUTBOUND:
            handlesendoutbound(self.xmlHttpReq.responseText);
            break; 
          case WFSTAT_PREVALIDATE:
            handlewfstatprevalidate(self.xmlHttpReq.responseText);
            break;
          case WFSTAT_POSTVALIDATE:
            handlewfstatpostvalidate(self.xmlHttpReq.responseText);
            break;       
          case WFSTAT_READYTODEPLOY:
            handlewfstatreadytodeploy(self.xmlHttpReq.responseText);
            break;
          case FILTER_ITEMS:
            showContents(self.xmlHttpReq.responseText);
            setFolderIcon(selectedFolder, true);
            toggleDialog('performfilter', false);
            break;
          case TRANSFER_OFFERS:
            confirmsuccesstransferoffer(self.xmlHttpReq.responseText);
            break;	
          case CHECK_DEFAULTFOLDER:           
           SetDefaultFolder(self.xmlHttpReq.responseText);
           break;
          case SUB_FOLDERS:
            var i = 0;
            var subfoldersstr = (self.xmlHttpReq.responseText).replace("\r\n","");
            var subfolderids = new Array();
            subfolderids = subfoldersstr.split(',');
            var lowestsubfolderid = subfolderids[subfolderids.length - 1];
            if(subfolderids.length - 1 > 0){
              for(i = 0; i < subfolderids.length - 1; i++){
                //xmlhttpPost('/logix/folder-feeds.aspx?Action=GetSubFolders', SUB_FOLDERS, 'ParentFolderID=' + subfolderids[i])
              }
            }
            else{
                  var selectedfolder = 0;
                  selectedfolder = document.getElementById("hdnfolderid").value;
                  var lowestsubfolderstr = "foldername" + selectedfolder;
                  var foldernamecolor = (document.getElementById(lowestsubfolderstr)).style.color;
                  if (foldernamecolor == 'red'){
                    var folderstatus = document.getElementById("FolderStatus");
                    if(folderstatus != null){
                      var folderstatuscolor = folderstatus.style.backgroundColor;
                      if (folderstatuscolor == 'green'){
                        if (lowestsubfolderstr != null && lowestsubfolderstr != 'foldername0'){
                          (document.getElementById(lowestsubfolderstr)).style.color = 'black';
                        }
                      }
                    }
                    xmlhttpPost('/logix/folder-feeds.aspx?Action=GetParentFolders', PARENT_FOLDERS, 'FolderID=' + selectedfolder)
                  }
            }
            break;
          case PARENT_FOLDERS:  
              var selectedfolder = 0;
              selectedfolder = document.getElementById("hdnfolderid").value;
              var str = "foldername" + selectedFolder;
              var folder = document.getElementById(str);
              var parentfoldersstr = 0;
              parentfoldersstr = (self.xmlHttpReq.responseText).replace("\r\n","");
              var parentfolder = "foldername" + parentfoldersstr;
                if(parentfolder != null && parentfolder != 'foldername0' && folder.style.color == 'red'){
                  (document.getElementById(parentfolder)).style.color = 'red';
                }
                else if(parentfolder != null && parentfolder != 'foldername0')
                {
                  (document.getElementById(parentfolder)).style.color = 'black';
                }     
                if(parentfoldersstr != "0"){
                   xmlhttpPost('/logix/folder-feeds.aspx?Action=GetParentFolders', PARENT_FOLDERS, 'FolderID=' + parentfoldersstr)
                }       
            break;
        case SHOW_APPROVALACTIONITEMS:
            OnSuccessResponse((self.xmlHttpReq.responseText).replace("\r\n",""));
            break; 
        }
        updateStatusBar();
      }
    }
    self.xmlHttpReq.send(frmdata);
    cancelClicked=false;
  }

 function OnSuccessResponse(response) {
        if(response == "-1")
        {
            RemoveOptions('0, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27');
        }
        else
        {
            AddOption('12', '<%= Copient.PhraseLib.Lookup("perm.offers-delete", LanguageID) %>');
            AddOption('8', '<%= Copient.PhraseLib.Lookup("folders.transferoffers", LanguageID) %>');
            AddOption('11', '<%= Copient.PhraseLib.Lookup("folder.startenddatestooffer", LanguageID) %>');
            AddOption('10', '<%= Copient.PhraseLib.Lookup("folder.enddatettooffer", LanguageID) %>');
            AddOption('9', '<%= Copient.PhraseLib.Lookup("folder.startdatetooffer", LanguageID) %>');
            if(response == "0")
            {
                AddOption('13', '<%= Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) %>');
                AddOption('0', '<%= Copient.PhraseLib.Lookup("term.deploy", LanguageID) %>');

                RemoveOptions('14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27');
            }
            else if(response == "1")
            {
                RemoveOptions('0, 13, 16, 17, 18, 19, 20, 22, 23, 24, 25, 26, 27');
            
                AddOption('21', '<%= Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) %>');
                AddOption('15', '<%= Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('14', '<%= Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
            
            
            }
            else if(response == "2")
            {
                RemoveOptions('0, 13, 14, 15, 21');
                AddOption('24', '<%= Copient.PhraseLib.Lookup("term.deferdeployoffers", LanguageID) %>');
                AddOption('16', '<%= Copient.PhraseLib.Lookup("term.deployoffers", LanguageID) %>');
                AddOption('27', '<%= Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) %>');
                AddOption('26', '<%= Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('25', '<%= Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
                AddOption('23', '<%= Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) %>');
                AddOption('20', '<%= Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('19', '<%= Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
                AddOption('22', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) %>');
                AddOption('18', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('17', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
                
            
            }
            else if(response == "3")
            {
                RemoveOptions('0, 13, 14, 15, 17, 18, 19, 20, 21, 22, 23, 25, 26, 27');
                
                AddOption('24', '<%= Copient.PhraseLib.Lookup("term.deferdeployoffers", LanguageID) %>');
                AddOption('16', '<%= Copient.PhraseLib.Lookup("term.deployoffers", LanguageID) %>');
            
            }
            else if(response == "4")
            {
                RemoveOptions('0, 13, 14, 15, 16, 19, 20, 21, 23, 24, 25, 26, 27');
                
                AddOption('22', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) %>');
                AddOption('18', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('17', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
            
            }
            
        }
    }
    function isInArray(value, array) {
      return array.indexOf(value) > -1;
    }
    function RemoveOptions(optionValues){
        var values = optionValues.split(', ');
        var actionmenu = document.getElementById('Actionitems');
        var i;
        for(i=0; i<actionmenu.length; i++)
        {
            if(isInArray(actionmenu.options[i].value, values))
            {
                actionmenu.removeChild( actionmenu.options[i] ); 
                i--;
            }
        }
    }
    function AddOption(optionValue, optionText){
        var actionmenu = document.getElementById('Actionitems');
        var i, exists = 0;
        for(i=0; i<actionmenu.length; i++)
        {
            if(actionmenu.options[i].value == optionValue)
            {
                exists = 1;
                break;
            }
        }
        if(exists == 0)
        {
            actionmenu.options.add(new Option(optionText, optionValue), actionmenu.options[1])
        }
    }
  function needsRedirect(responseText) {
    return (responseText.indexOf('<html') > -1);
  }

  function writeRedirectPage(responseText) {
    document.open();
    document.write(responseText);
    document.close();
  }

  function confirmSuccess(responseText) {
    var response = trimString('<%Sendb(Copient.PhraseLib.Lookup("folders.CannotDelete", LanguageID))%>');
    if (responseText.substring(0, 2) != 'OK' && trimString(responseText) != response) {
      alert(responseText.substring(3));
      document.location = 'folders.aspx'; 
      }
    else if (trimString(responseText) == response){
      alert(responseText);  
      document.location = 'folders.aspx';
      }
  }
  function cofirmsuccessmodfolder(responseText) {
  var modifyfolderElem = document.getElementById("modifyfoldererror");
    if (responseText.substring(0, 2) == 'NO') {
      modifyfolderElem.style.display = 'block';
      modifyfolderElem.innerHTML = responseText.substring(3,responseText.length);
       toggleDialog('modifyfolder', true);
    }
    else if (responseText.substring(0, 2) != 'NO' && !cancelClicked) {
     modifyfolderElem.style.display = 'none';
     toggleDialog('modifyfolder', false);
     document.location = 'folders.aspx';
    }
  }
  
  function cofirmsuccessdupoffer(responseText) {
    
    //var performactionserrorElem = document.getElementById("performactionserror");
    //if (responseText.substring(0, 2) != 'OK') {
     // performactionserrorElem.style.display = 'block';
     // performactionserrorElem.innerHTML = responseText;
    //}
    //else if (responseText.substring(0, 2) == 'OK') {
     //performactionserrorElem.style.display = 'none';
     //toggleDialog('performactions', false);
     //document.location = 'folders.aspx';
    //}
	ClosePopUp("");
  }

  //Varma
  function ClosePopUp(responseText){
    var performactionserrorElem = document.getElementById("performactionserror");
    performactionserrorElem.style.display = 'none';
    toggleDialog('performactions', false);
    document.location = 'folders.aspx';
    
//    var failedoffersdesc = responseText;
//    var failedoffers = responseText;
//    failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
//    failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));
//    
//      if ((responseText.substring(0, 2) != 'OK') && (responseText.substring(0, 2) != 'NO')){
//        performactionserrorElem.style.display = 'block';
//        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.deployfail", LanguageID))%>';
//        failedoffersdesc = failedoffersdesc.split('|');
//        failedoffers = failedoffers.split(',');
//        populatedeploymenterrorrows(failedoffers,failedoffersdesc);                
//      }

//      else if (responseText.substring(0, 2) == 'NO') {
//              
//       var offerswithoutcon = responseText.substring(3, responseText.indexOf(','));        
//       handlemassDeployConditional(offerswithoutcon);
//      }

//      else if (responseText.substring(0, 2) == 'OK') {
//       performactionserrorElem.style.display = 'none';
//       toggleDialog('performactions', false);
//       document.location = 'folders.aspx';
//      }
    }
      function folderdateschange(responseText){
       var performactionserrorElem = document.getElementById("performactionserror"); 
        if (responseText.substring(0, 2) == 'NO') {            
            performactionserrorElem.style.display = 'block';  
            performactionserrorElem.style.color = 'red';     
            performactionserrorElem.innerHTML = responseText.substring(3,responseText.length);
            toggleDialog("performactions",true);
      }
      else{
       performactionserrorElem.style.display = 'none'; 
       toggleDialog("performactions",false); 
       document.location = 'folders.aspx';
      }
      
    }
	
	function deleteoffers(responseText){
       var performactionserrorElem = document.getElementById("performactionserror"); 
        if (responseText.substring(0, 2) == 'NO') {            
            performactionserrorElem.style.display = 'block';  
            performactionserrorElem.style.color = 'red';     
            performactionserrorElem.innerHTML = responseText.substring(3,responseText.length);
            toggleDialog("performactions",true);
      }
      else{
       performactionserrorElem.style.display = 'none'; 
       toggleDialog("performactions",false); 
       document.location = 'folders.aspx';
      }
      
    }
	
  function UpdateParentColor(selectedFolder) {
    var rootFolderId = "foldertree";
    var parents = $("#folderrow" + selectedFolder).parents();
    var parent;
    var i=0;
    while(i<parents.length && parents[i] != null && parents[i].id != rootFolderId)
    {    
        {
            var spans = $(parents[i]).find("span");       
            if(spans)
            {
             $(spans[0]).css("color","red");
            }
        }
        i=i+1;
      }
  }
  function confirmsuccesstransferoffer(responseText) {
    ClosePopUp('')
//    var performactionserrorElem = document.getElementById("performactionserror");
//    if (responseText.substring(0, 2) != 'OK') {
//      performactionserrorElem.style.display = 'block';
//      performactionserrorElem.innerHTML = responseText;
//    }
//    else if (responseText.substring(0, 2) == 'OK') {
//     performactionserrorElem.style.display = 'none';
//     toggleDialog('performactions', false);
//     document.location = 'folders.aspx';
//    }
  }	
  function populatedeploymenterrorrows(failedoffers, failedoffersdesc){     
     UpdateParentColor(selectedFolder);
     var failedItemsString = failedoffers.toString();
     for (var i = 0; i < selectedItems.length; i++) {
        var rowOfferRecord = document.getElementById("OfferRecord" + selectedItems[i]);
        if(failedItemsString.indexOf(selectedItems[i]) != -1)
        {
            var failedItemIndex = failedoffers.indexOf(selectedItems[i].toString());
            //alert(failedItemIndex + ":FailedOffers:" + failedoffers[0] + ":SelectedItem:" + selectedItems[i]);
            if(failedItemIndex != -1)
                rowOfferRecord.setAttribute("title", failedoffersdesc[failedItemIndex]);
            rowOfferRecord.setAttribute("style", "color: Red");
            //Set color for the anchor tag
            if($(rowOfferRecord).find("a").length>0)
                $($(rowOfferRecord).find("a")[0]).css("color","Red");
        }
        else
        {
            rowOfferRecord.setAttribute("style", "color: Black");
            if($(rowOfferRecord).find("a").length>0)
                $($(rowOfferRecord).find("a")[0]).css("color","Blue");
        }      
      }
    }
    function populaterow(failedoffers, failedoffersdesc){
     
     for (var i = 0; i < failedoffers.length; i++) {
        var row = document.getElementById("errdesc" + failedoffers[i]);        
        createCell(row.insertCell(0), failedoffersdesc[i], 'row');
      }      
    }

    // create DIV element and append to the table cell
    function createCell(cell, text, style) {      
      var div = document.createElement('div'), // create DIV element
      txt = document.createTextNode(text); // create text node
      div.appendChild(txt);                    // append text node to the DIV
      div.setAttribute('class', style);        // set DIV class attribute
      div.setAttribute('className', style);    // set DIV class attribute for IE (?!)      
      cell.appendChild(div);                   // append DIV to the table cell      
      cell.style.whiteSpace = "nowrap";
      cell.style.color = "#C82536";
      cell.colSpan = 6; 
                       
    }

    function CancelNavtoReports(){
        IsNavtoReportsCancelled = true;
        toggleDialog('NavigatetoReports', false);
        self.xmlHttpReq.abort();
        <%If Session("OFFERIDS") IsNot Nothing Then
        Session.Remove("OFFERIDS")
    End If %>
//        document.location = 'folders.aspx';
    }

    function handlenavtoreports(responseText){
      
        window.location = 'reports-custom.aspx';
    }

    function handlesendoutbound(responseText){    
      var performactionserrorElem = document.getElementById("performactionserror");
//      var failedoffersdesc = responseText;
//      var failedoffers = responseText;
//      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
//      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));     

//      if (responseText.substring(0, 2) != 'OK') {
//        performactionserrorElem.style.display = 'block';
//        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.sendoutboundfail", LanguageID))%>';           
//        failedoffersdesc = failedoffersdesc.split('|');
//        failedoffers = failedoffers.split(',');
//        populaterow(failedoffers,failedoffersdesc);     
//      }
//      else if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'folders.aspx';
//      }
    }

    function handlewfstatprevalidate(responseText){     
     
      var performactionserrorElem = document.getElementById("performactionserror");
      var failedoffersdesc = responseText;
      var failedoffers = responseText;
      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));  

      if ((responseText.substring(0, 2) != 'OK') && (responseText.indexOf('||') >= 0)) {
        //performactionserrorElem.style.display = 'block';
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.requirerevalidaion", LanguageID))%>');       
        performactionserrorElem.style.display = 'none';
        toggleDialog('performactions', false);
        document.location = 'folders.aspx';         
      }
      else if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'folders.aspx';
      }
      else{
        performactionserrorElem.style.display = 'block';
        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.prevalidatefail", LanguageID))%>'; 
        failedoffersdesc = failedoffersdesc.split('|');
        failedoffers = failedoffers.split(',');
        populaterow(failedoffers,failedoffersdesc);   
      }
    }

    function handlewfstatpostvalidate(responseText){     
      var performactionserrorElem = document.getElementById("performactionserror");
      var failedoffersdesc = responseText;
      var failedoffers = responseText;
      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));  

      if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'folders.aspx';
      }
      else{
        performactionserrorElem.style.display = 'block';
        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.postvalidatefail", LanguageID))%>';
        failedoffersdesc = failedoffersdesc.split('|');
        failedoffers = failedoffers.split(',');
        populaterow(failedoffers,failedoffersdesc);       
      }
    }

    function handlewfstatreadytodeploy(responseText){     
      var performactionserrorElem = document.getElementById("performactionserror");
      var failedoffersdesc = responseText;
      var failedoffers = responseText;
      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));  

      if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'folders.aspx';
      }
      else{
        performactionserrorElem.style.display = 'block';
        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.readytodeployfail", LanguageID))%>';
        failedoffersdesc = failedoffersdesc.split('|');
        failedoffers = failedoffers.split(',');
        populaterow(failedoffers,failedoffersdesc);                 
      }
    }

    function updateitemchecks(failedoffers){
     
     for (var i = 0; i < failedoffers.length; i++) {
         elem = document.getElementById("itemID" + failedoffers[i]);         
         var tr = elem.parentNode.parentNode;         
         tr.style.backgroundColor = 'red';               
    }     
   }

   function resetitemscolor(offers){
     
     for (var i = 0; i < (offers.length) - 1; i++) {     
         elem = document.getElementById("itemID" + offers[i]);         
         var tr = elem.parentNode.parentNode;
         tr.style.backgroundColor = 'transparent';         
    }     
   }

  function handleCreateFolder(responseText) {
    var folderid = 0;
    var folderDiv = null;
    var folderRow = null;
    var folderImg = null;
    var folderNbsp = null;
    var folderName = null;
    var hdnFolderElement = null;
    var className = '';
    var parentFolder = null;
    var descendants = null;
    var sibling = null;
    
    if (responseText.substring(0, 2) == 'OK') {
      folderid = parseInt(responseText.substring(3));
      if (selectedFolder == 0) {
        parentFolder = document.getElementById('foldertree');
      } else {
        parentFolder = document.getElementById('folder' + selectedFolder);
      }
      className = (parentFolder.getAttribute('className') != null) ? 'className' : 'class';
      
      // Find the sibling
      if (parentFolder != null) {
        descendants = parentFolder.getElementsByTagName('span');
        for (var i = 0; i < descendants.length; i++) {
          if (descendants[i].parentNode.parentNode.parentNode == parentFolder) {
            if (lastCreatedFolderName.toLowerCase() < descendants[i].innerHTML.toLowerCase()) {
              if (descendants[i].innerHTML.toLowerCase().length > 0 && descendants[i].innerHTML.toLowerCase().substring(0, 1) == "<") {
                 // skip these nodes as they are special folders which must float to the top
              } else {
                sibling = descendants[i].parentNode.parentNode;
                break;
              }
            }
          }
        }
      }
      //AMS-1369 The function highlightSelected() called in navigateToFolder() called in here expects the element: <input type=""hidden"" id=""hdnfolderid"" name=""hdnfolderid"" value=""" & FolderID & """/>
      hdnFolderElement = document.createElement('input');
      hdnFolderElement.setAttribute('type', 'hidden');
      hdnFolderElement.setAttribute('id', 'hdnfolderid');
      hdnFolderElement.setAttribute('name', 'hdnfolderid');
      hdnFolderElement.setAttribute('value', folderid);
      parentFolder.appendChild(hdnFolderElement);
      //Folder div
      folderDiv = document.createElement('div');
      folderDiv.setAttribute('id', 'folder' + folderid);
      folderDiv.setAttribute(className, 'folder');
      if (sibling != null) {
        try {
          parentFolder.insertBefore(folderDiv, sibling);
        } catch (ex) {
          parentFolder.appendChild(folderDiv);
        }
      } else {
        parentFolder.appendChild(folderDiv);
      }
      
      //FolderRow div
      folderRow = document.createElement('div');
      folderRow.setAttribute('id', 'folderrow' + folderid);
      folderRow.setAttribute(className, 'folderrow');
      folderRow.setAttribute('onclick', 'javascript:highlightSelected(event,' + folderid + ');');
      if (isIE()) {
        folderRow.onclick = function() {
          highlightSelected(event, folderid);
        }
      }
      folderDiv.appendChild(folderRow);
      
      //FolderImg image
      folderImg = document.createElement('img');
      folderImg.setAttribute('id', 'expander' + folderid);
      folderImg.src = '../images/plus.png';
      folderImg.setAttribute(className, 'expander');
      folderImg.onclick = function() {
        toggleFolder(folderid);
      }
      folderRow.appendChild(folderImg);

      folderImg = document.createElement('img');
      folderImg.setAttribute('id', 'folderimg' + folderid);
      folderImg.src = '../images/folder.png';
      folderImg.setAttribute(className, 'folderimg');
      folderRow.appendChild(folderImg);
      
      //FolderName span
      folderName = document.createElement('span');
      folderName.setAttribute('id', 'foldername' + folderid);
      folderName.setAttribute(className, 'foldername');
      folderName.innerHTML = lastCreatedFolderName;
      folderRow.appendChild(folderName);
      // move to the newly created folder
      navigateToFolder(folderid);
    
      // hide the create folderdialog box
      toggleDialog('foldercreate', false);
      
      var folderNameElem = document.getElementById('newFolderName');
      var startElem = document.getElementById('folderstart');
      var endElem = document.getElementById('folderend');
      folderNameElem.value='';
      startElem.value='';
      endElem.value='';


    }
    else if (responseText.substring(0, 2) != 'NO') {
             showfolderdateerror(responseText);
     }
     else if(responseText.substring(0, 2) == 'NO'){
        var res = responseText.split('|')
        if(res[1] !=''){
             showfolderdateerror(res[1]);
        }
     }
  }
  
  function navigateToFolder(FolderID) {
    var folderDiv = document.getElementById('folder' + FolderID);
    var prt = null;
    var fid = 0;
        
    if (folderDiv != null) {
      prt = folderDiv;
      while (prt != null) {
        if (prt.parentNode.id.length >= 6 && prt.parentNode.id != 'foldertree' && prt.parentNode.id.substring(0,6) == 'folder') {
          fid = parseInt(prt.parentNode.id.substring(6));
          expandFolder(fid);
          prt = prt.parentNode;
        } else {
          prt = null;
        }
      }
      highlightSelected(null, FolderID);
    }
  }
  
  function highlightSelected(e, FolderID) {
	sourceFolder = FolderID;    
    var folderDiv = document.getElementById('folder' + FolderID);
    var folderRow = document.getElementById('folderrow' + FolderID);
    var selectedDiv = document.getElementById('folder' + selectedFolder);
    var selectedRow = document.getElementById('folderrow' + selectedFolder);
    var source = null;
    if(document.getElementById('hdnfolderid') != null)
        document.getElementById('hdnfolderid').value = FolderID;
    if (e != null) {
      source = (document.all) ? window.event.srcElement : e.target;
    }
    
    if (source != null && FolderID==selectedFolder && source.id.length >= 8 && source.id.substring(0,8) == 'expander') {
      // clicking the expander for the selected folder doesn't change its selected state.    
    } else {
      if (FolderID != selectedFolder && selectedRow != null) {
        selectedRow.style.backgroundColor = '#ffffff';
        selectedRow.style.border = '1px solid #ffffff';
      }
      if (folderDiv != null && folderRow != null) {
        if (FolderID != selectedFolder) {
          selectedFolder = FolderID;
        } else {
          selectedFolder = 0;
        }
        
        if (folderRow.style.backgroundColor == '' || folderRow.style.backgroundColor == '#ffffff' || folderRow.style.backgroundColor == 'rgb(255, 255, 255)') {
          folderRow.style.border = '1px solid #7788cc';
          folderRow.style.backgroundColor = '#ccccff';
          var content = loadFolderItems(FolderID, 2);
          setSearchFocus();
        } else {
          folderRow.style.backgroundColor = '#ffffff';
          folderRow.style.border = '1px solid #ffffff';
        }
      }
    }
    
  }
  

  function loadFolderItems(folderID, SortDirection, SortText) {
    var FilterElem = document.getElementById("Filteritems");
    var selectedfilter = FilterElem.options[FilterElem.selectedIndex].value;
  var   urlStr='/logix/folder-feeds.aspx?Action=LoadFolderItems';
    //Reset variables
    pageIndex=1;
    loaded=true;
    if (selectedfilter != 0) {      
      urlStr= urlStr + '&WFStatus=' + selectedfilter;
    }
    if (SortDirection != '' && SortText != ''){
        m_SortText=SortText;
        m_SortDirection=SortDirection;
       urlStr= urlStr + '&SortText=' + SortText + '&SortDirection=' + SortDirection;
    }
    return xmlhttpPost(urlStr, VIEW_ITEMS, 'FolderID=' + selectedFolder);    
  }

  function loadOfferNavigationDialog(OfferID, CollisionReportNavigationEnabled, OfferValidationMessage) {
    if (CollisionReportNavigationEnabled.trim() == "True") {
      $("#lnkViewCollisionReport").show();
      $("#lnkViewCollisionReport").prop("href", "CollidingOffers-Report.aspx?ID=" + OfferID);
      $("#spnViewCollisionReport").hide();
    }
    else {
      $("#lnkViewCollisionReport").hide();
      $("#spnViewCollisionReport").show();
    }
    $("#spnOfferValidationMsg").text(OfferValidationMessage);
    $("#lnkViewOffer").prop("href", "offer-redirect.aspx?OfferID=" + OfferID);
    toggleDialogNoFade('OfferNavigationDialog', true);
    return false;
  }

  //AL-7000
  //Load offers iteratively in a folder
  var pageIndex = 1;
  var pageCount;
  var loaded=true;
  var m_SortDirection="";
  var m_SortText="";
  jQuery(document).ready(function($) {
      $('#foldercontents').on('scroll',function () {
      if(loaded)
      {
          if($('#foldercontents').scrollTop()+$('#foldercontents').innerHeight() > $('#foldercontents')[0].scrollHeight){
            loaded=false;
             GetRecords();
          }
      }
      });
    });
    
     //Attach Folder status close function
      $(document).on('click', '#statusClose', function() {
      var obj=this;
            var strURL="/logix/folder-feeds.aspx?Action=UpdateFolderStatus";
                    $.post(strURL,{FolderID: selectedFolder,
                                   FromOfferList:false},
                                   function (data) { 
                                     $(obj).parent().fadeTo(300,0,function(){
                                     $(obj).remove();
                                     $("#FolderStatus").empty();
                                });
                      },false);
              });
	//Feches records from Database when scroll bar reaches end
    function GetRecords() {
    var FilterElem = document.getElementById("Filteritems");
    var selectedfilter = FilterElem.options[FilterElem.selectedIndex].value;
        pageIndex++;
        //Due to some timing issue, its skipping 2nd-3rd iteration, to fix this added new check pageindex==3
        if (pageIndex == 2 || pageIndex <= pageCount) {
            var isChecked=false;
            if ($('#allitemIDs').is(":checked"))
            {
                isChecked=true;
            }
            $("#loader").show();
            var strURL="/logix/folder-feeds.aspx?Action=GetOffersIteratively";
            $.post(strURL,
            {pageIndex: pageIndex,
            FolderID: selectedFolder,
            checked:isChecked,
            WFStatus:selectedfilter,
            SortText:m_SortText,
            SortDirection:m_SortDirection},
            function (data) { 
            OnSuccess(data);
           },false);
        }
    }
    function ShowRequestApprovalActionItems(itemIds) {
        xmlhttpPost('/logix/folder-feeds.aspx?Action=ShowRequestApprovalActionItem', SHOW_APPROVALACTIONITEMS, 'ItemIds=' + itemIds + '&FromOfferList=' + false)
    }
     
	//If data fetched successfully, append new records to offer table
    function OnSuccess(response) {
        var Offertable = response.split("AMS_SPLITTER_AMS").pop()
        pageCount=parseInt(response.substring(0,response.indexOf('AMS_SPLITTER_AMS')));
        $('#tb1 tr:last').after(Offertable);
        $("#loader").hide();
        //$("#folderstatusbar")[0].innerText=($("#tb1 > tbody > tr").length -1)/2 +" item(s) loaded out of total Offers";
        loaded=true;
    } 
  
  function loadAllItems(folderID) {
    return xmlhttpPost('/logix/folder-feeds.aspx?Action=SendFoundOffers', VIEW_AVAILABLE, 'FolderID=' + selectedFolder);
  }
  
  function showContents(content) {
    selectedItems.length = 0; // clear any selected items in previous folder
    var folderContentsElem = document.getElementById("foldercontents");
    if (folderContentsElem != null) {
      folderContentsElem.innerHTML = content;
    }
  }
  //
 function showwarning(content){
  var AddItemWarningElem = document.getElementById("AddItemWarninng");
  var AddItemBtnElem = document.getElementById("btnAddItems");
  var datematced = trimString('<b><%Sendb(Copient.PhraseLib.Lookup("folders.DateMatched", LanguageID))%></b>');
  var datenotavailable = trimString('<b><%Sendb(Copient.PhraseLib.Lookup("folders.ExpiryDateUnavailable", LanguageID))%></b>');
    if (trimString(content) == datematced || trimString(content) == '' || trimString(content) == datenotavailable) {
      AddItemWarningElem.style.display = 'none';
      AddItemBtnElem.disabled = false;
      }
    else {
      AddItemWarningElem.style.display = 'block';
      AddItemWarningElem.innerHTML = content;
      AddItemBtnElem.disabled = true;
      } 
    }
 function showfolderdateerror(content){
  var createfolderElem = document.getElementById("createfoldererror");
      
      createfolderElem.style.display = '';
      createfolderElem.innerHTML = content;
    }
  function showSearchContents(content) {
    var addElem = document.getElementById("btnAddItems");
    var resultsElem = document.getElementById("searchResults");
    var countElem = null;
    var showAdd = true;
        
    if (resultsElem != null) {
      resultsElem.innerHTML = content;
    }
    
    // if no results are returned then don't show the add button
    countElem = document.getElementById("offerCount");
    if (countElem != null && parseInt(countElem.value) == 0) {
      showAdd = false;
    }
    
    if (addElem != null && showAdd) {
      addElem.style.visibility= 'visible';
    }
  }
  
  function showfSearchContents(content) {
    var resultsElem = document.getElementById("fsearchResults");
  
    if (resultsElem != null) {
      resultsElem.innerHTML = content;
    }
    document.body.style.cursor = 'default';
  }
  
  function toggleDialog(elemName, shown) {
    var elem = document.getElementById(elemName);
    var fadeElem = document.getElementById('fadeDiv');
    switch (elemName) {        
        case 'foldersearch':            
            var searchElem = document.getElementById("fsearchterms");
            var resultsElem = document.getElementById("fsearchResults"); 
            var searchType = document.getElementById("fsearchType");
            if (searchElem != null) {
                searchElem.value = ""
            }
            if (resultsElem != null) {
                resultsElem.innerHTML = "";
            }
            if (searchType != null && searchType.options.length > 0) {
                searchType.selectedIndex = "0";
            }
            break;
        case 'foldercreate':
            var folderstart = document.getElementById("folderstart");
            var folderend = document.getElementById("folderend"); 
            var newFolderName = document.getElementById("newFolderName");
            var defaultUEFolder = document.getElementById("defaultUEFolder");
            if (folderstart != null) {
                folderstart.value = ""
            }
            if (folderend != null) {
                folderend.value = "";
            }
            if (newFolderName != null) {
                newFolderName.value = ""
            }
            if (defaultUEFolder != null) {
                $('.defaultUEFolder').attr('checked', false);
            }
            break;  
        case 'performactions':
            if(!shown && document.getElementById("reloadpage").innerHTML == "1")
            {
                document.location = 'folders.aspx';                
            }
            break;  
    }
    
    if (elem != null) {
      elem.style.display = (shown) ? 'block' : 'none';
    }
    
    if (fadeElem != null) {
      fadeElem.style.display = (shown) ? 'block' : 'none';
    }
    
  }

  function toggleDialogNoFade(elemName, shown) {
    var elem = document.getElementById(elemName);
    
    if (elem != null) {
      elem.style.display = (shown) ? 'block' : 'none';
    }
    
  }
  
  function resetSearch() {
    var searchElem = document.getElementById('searchResults');
    var searchBox = document.getElementById('');
        
    if (searchElem != null) {
      searchElem.innerHTML = '';
      setSearchBoxText('');
    }
  }
  
  function setFolderIcon(folderID, hasContent) {
    var elemImg = document.getElementById('folderimg' + folderID);
    
    if (elemImg != null) {
      elemImg.src = (hasContent) ? '../images/folder-full.png' : '../images/folder.png';
    }    
    
  }
  
  function toggleDialogOfferDuplicate(elemName, shown) {
	var elem = document.getElementById(elemName);
    var offerfadeElem = document.getElementById('OfferfadeDiv');
	
    if (elem != null) {
      elem.style.display = (shown) ? 'block' : 'none';
    }
    
    if (offerfadeElem != null) {
      offerfadeElem.style.display = (shown) ? 'block' : 'none';
    }	
	
    if (shown)  {
	  toggleDialog('performactions', false);
	} else {
	  toggleDialog('performactions', true);
	}
 }
    
  function createFolder() {
    var dialogElem = document.getElementById('foldercreate');
    var newNameElem = document.getElementById('newFolderName');
    var folderstartdateele=document.getElementById('folderstart');
    var folderenddateele=document.getElementById('folderend');
    var folderthemeele=document.getElementById('Theme');
    var defaultUEFolderele=document.getElementById('defaultUEFolder');
     var newAccessLevelElem = document.getElementById('accessLevel');
    var folderElem = document.getElementById('folder' + selectedFolder);
    var accessLevel = 3;
    var isdefaultUEFolder=false;

    if (dialogElem != null && newNameElem != null) {
      dialogElem.style.display = 'block';

      if(newNameElem.value.trim() == "")
      {
         alert(' <%Sendb(Copient.PhraseLib.Lookup("folders.EnterFolderName", LanguageID))%> ');
      }
      else
      {
        lastCreatedFolderName = newNameElem.value
        if (lastCreatedFolderName != null) {
          lastCreatedFolderName = trimString(lastCreatedFolderName);
         
             if (newAccessLevelElem != null) {
                      accessLevel = newAccessLevelElem.value;
          }
          
           if (folderstartdateele != null){
              folderstartdate=trimString(folderstartdateele.value);
              }
              if (folderenddateele != null){
                  folderenddate=trimString(folderenddateele.value);
                  }
                if (folderthemeele != null){
                  foldertheme= folderthemeele.options[folderthemeele.selectedIndex].text;
                    }
    if(defaultUEFolderele!=null){
              isdefaultUEFolder=defaultUEFolderele.checked;
            }

            if (document.getElementById('form_FolderStartHr') != null && document.getElementById('form_FolderStartHr').value != "") {
                folderstartHr = document.getElementById('form_FolderStartHr').value;
            }
            if (document.getElementById('form_FolderStartMin') != null && document.getElementById('form_FolderStartMin').value != "") {
                folderstartMin = document.getElementById('form_FolderStartMin').value;
            }
            if (document.getElementById('form_FolderEndHr') != null && document.getElementById('form_FolderEndHr').value != "") {
                folderendHr = document.getElementById('form_FolderEndHr').value;
            }
            if (document.getElementById('form_FolderEndMin') != null && document.getElementById('form_FolderEndMin').value != "") {
                folderendMin = document.getElementById('form_FolderEndMin').value;
            }

            if (folderstartdate != "") {
                folderstartdate = folderstartdate + " " + folderstartHr + ":" + folderstartMin + ":00";
            }
            if (folderenddate != "") {
                folderenddate = folderenddate + " " + folderendHr + ":" + folderendMin + ":00";
            }

          xmlhttpPost('/logix/folder-feeds.aspx?Action=CreateFolder', CREATE_FOLDER, 'FolderID=' + selectedFolder + '&FolderName=' + encodeURIComponent(lastCreatedFolderName) + '&AccessLevel=' + accessLevel + '&FolderStartDate=' + folderstartdate + '&FolderEndDate=' + folderenddate + '&FolderTheme=' + foldertheme+'&IsdefaultUEFolder='+isdefaultUEFolder);
          lastCreatedFolderName=lastCreatedFolderName.replace(/</g, "&lt;");
          }
        }
      } 
    } 
    var modfoldername ="";
    var folderstartdate="";
    var folderstartHr="00";
    var folderstartMin="00";
    var folderenddate="";
    var folderendHr="00";
    var folderendMin="00";
    var foldertheme="";
    var isdefaultUEFolder=false;

   function savemodifiedFolder() {   
     
     var modfoldialogElem = document.getElementById('modifyfolder');
     var modfolderstartdateele=document.getElementById('modifyfolderstart');
     var modfolderenddateele=document.getElementById('modifyfolderend');
     var modfolderthemeele=document.getElementById('ModifyTheme');
     var modfoldialogName= document.getElementById('editFolderName');
     var moderror = document.getElementById('modifyfoldererror');  
	 moderror.innerHTML='';
     var statustext = document.getElementById('statustext');  
     var ismassupdateenable= <% Sendb(MyCommon.NZ(MyCommon.Fetch_SystemOption(133),0)) %>;
     var saveButton = document.getElementById('btnModFolder');
     var divMessg = document.getElementById('divMessage');
     var defaultUEFolderele=document.getElementById('defaultUEFolder-Mod');
     var dialogClose = document.getElementById('dialogClose');
     dialogclose.style.visibility = "visible";
     dialogclose.style.display = ""; 

      if (modfoldialogElem != null) 
      {
        if(modfoldialogName != null)
        {
            modfoldername = trimString(modfoldialogName.value);
            if(modfoldername == ""){
                 moderror.innerHTML = ' <%Sendb(Copient.PhraseLib.Lookup("folders.EnterFolderName", LanguageID))%>';
                toggleDialog('modifyfolder', true);
                 return;
             }
        }
         if (modfolderstartdateele != null) {
          folderstartdate= trimString(modfolderstartdateele.value);
         }     
         if (modfolderenddateele != null) {
          folderenddate=trimString(modfolderenddateele.value);
         } 

            if (document.getElementById('modifyfolderstartHr') != null && document.getElementById('modifyfolderstartHr').value != "") {
                folderstartHr = document.getElementById('modifyfolderstartHr').value;
            }
            if (document.getElementById('modifyfolderstartMin') != null && document.getElementById('modifyfolderstartMin').value != "") {
                folderstartMin = document.getElementById('modifyfolderstartMin').value;
            }
            if (document.getElementById('modifyfolderEndHr') != null && document.getElementById('modifyfolderEndHr').value != "") {
                folderendHr = document.getElementById('modifyfolderEndHr').value;
            }
            if (document.getElementById('modifyfolderEndMin') != null && document.getElementById('modifyfolderEndMin').value != "") {
                folderendMin = document.getElementById('modifyfolderEndMin').value;
            }

            if (folderstartdate != "") {
                folderstartdate = folderstartdate + " " + folderstartHr + ":" + folderstartMin + ":00";
            }
            if (folderenddate != "") {
                folderenddate = folderenddate + " " + folderendHr + ":" + folderendMin + ":00";
            }
          
         if (modfolderthemeele != null) {
          foldertheme=modfolderthemeele.options[modfolderthemeele.selectedIndex].text;
         }
         if(defaultUEFolderele!=null){
            isdefaultUEFolder=defaultUEFolderele.checked;
         }
        
          if (ismassupdateenable && statustext.value.indexOf("0 item(s)")==-1){
            saveButton.style.visibility = "hidden";
             divMessg.style.visibility = "visible";
            divMessg.style.display= "";
        }
        else{
           xmlhttpPost('/logix/folder-feeds.aspx?Action=ModifyFolder', MODIFY_FOLDER, 'FolderID=' + selectedFolder + '&ModFolderName=' + encodeURIComponent(modfoldername) + '&ModFolderStartDate=' + folderstartdate + '&ModFolderEndDate=' + folderenddate + '&ModFolderTheme=' + foldertheme +'&isMassupdateEnabled=false'+'&IsdefaultUEFolder='+isdefaultUEFolder);        
        } 
           }
   }
   function OKClicked()
   {
        var saveButton = document.getElementById('btnModFolder');
        var divMessg = document.getElementById('divMessage');
        xmlhttpPost('/logix/folder-feeds.aspx?Action=ModifyFolder', MODIFY_FOLDER, 'FolderID=' + selectedFolder + '&ModFolderName=' + encodeURIComponent(modfoldername) + '&ModFolderStartDate=' + folderstartdate + '&ModFolderEndDate=' + folderenddate + '&ModFolderTheme=' + foldertheme +'&isMassupdateEnabled=true'+'&IsdefaultUEFolder='+isdefaultUEFolder);
         saveButton.style.visibility = "visible";
        divMessg.style.visibility = "hidden";
         toggleDialog('modifyfolder', true);
   }
   function CancelClicked()
   {
        var isRestrictOperationEnabled= <% Sendb(MyCommon.NZ(MyCommon.Fetch_SystemOption(272),0)) %>;
        var saveButton = document.getElementById('btnModFolder');
        var divMessg = document.getElementById('divMessage');
		if(isRestrictOperationEnabled == false)
		{
		    xmlhttpPost('/logix/folder-feeds.aspx?Action=ModifyFolder', MODIFY_FOLDER, 'FolderID=' + selectedFolder + '&ModFolderName=' + encodeURIComponent(modfoldername) + '&ModFolderStartDate=' + folderstartdate + '&ModFolderEndDate=' + folderenddate + '&ModFolderTheme=' + foldertheme +'&isMassupdateEnabled=false'+'&IsdefaultUEFolder='+isdefaultUEFolder);             
        }		
        saveButton.style.visibility = "visible";
        divMessg.style.visibility = "hidden";
        cancelClicked=true;
   }
   function handleAction(){
     
     var elemExecAction = document.getElementById('ExecAction');
     var elemdupoffer = document.getElementById('dupoffer');
     var elemexecdupoffer = document.getElementById('execduplicateoffer');
     var actionele=document.getElementById('Actionitems');
     var folderList=document.getElementById('folderList');
     
     if(folderList != null){
        folderList.value = '';
     }


     if (actionele != null && ((actionele.value != '1') && (actionele.value != '8'))){
       elemExecAction.style.visibility = 'visible';
       elemdupoffer.style.visibility = 'hidden';
       elemexecdupoffer.style.visibility = 'hidden';
       }
     else if(actionele != null && ((actionele.value=='1')||(actionele.value=='8'))){
        elemdupoffer.style.visibility = 'visible';
        elemexecdupoffer.style.visibility = 'hidden';
        elemExecAction.style.visibility = 'hidden';
     }  
     
     if(actionele != null && actionele.value<0){
        elemExecAction.style.visibility = 'hidden';
     }         
   }

   function ExecAction(){
    var actionele=document.getElementById('Actionitems');
    var selectedaction="";         
       if (actionele != null) {
         selectedaction=actionele.options[actionele.selectedIndex].text; 
       }
       if (selectedaction != "") {
         switch (selectedaction) {
        
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%>':
            handlemassDeploy();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID))%>':
            handlemassDeferDeploy();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%>':
            handleReports();
            break;      
          case '<%Sendb(Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID))%>':
            SendOutBound();
            break;         
          case '<%Sendb(Copient.PhraseLib.Lookup("term.prevalidate", LanguageID))%>':
            ChangeWFStatustoPreValidate();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.postvalidate", LanguageID))%>':
            ChangeWFStatustoPostValidate();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID))%>':
            ChangeWFStatustoReadytoDeploy();
            break;
		  case '<%Sendb(Copient.PhraseLib.Lookup("folders.transferoffers", LanguageID))%>':
            TransferOffers();
            break;	
          case '<%Sendb(Copient.PhraseLib.Lookup("folder.startdatetooffer", LanguageID))%>':
          case '<%Sendb(Copient.PhraseLib.Lookup("folder.startdatetimetooffer", LanguageID))%>':
            ApplyFolderStartDatetoOffer();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("folder.enddatettooffer", LanguageID))%>':
          case '<%Sendb(Copient.PhraseLib.Lookup("folder.enddatetimettooffer", LanguageID))%>':
            ApplyFolderEndDatetoOffer();
            break;
		  case '<%Sendb(Copient.PhraseLib.Lookup("folder.startenddatestooffer", LanguageID))%>':
          case '<%Sendb(Copient.PhraseLib.Lookup("folder.startenddatewithtimetooffer", LanguageID))%>':
            ApplyFolderStartEndDatetoOffer();
            break;
		  case '<%Sendb(Copient.PhraseLib.Lookup("perm.offers-delete", LanguageID))%>':
            DeleteSelectedOffers();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%>':
            HandleMassRequestApproval_Async(13, '<%Sendb(Copient.PhraseLib.Lookup("folders.submit", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.submitapprovalconfirm", LanguageID))%>');
            break;	
          case '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%>':
            HandleMassRequestApproval_Async(14, '<%Sendb(Copient.PhraseLib.Lookup("folders.submit", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.approvewithdeployconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%>':
            HandleMassRequestApproval_Async(15, '<%Sendb(Copient.PhraseLib.Lookup("folders.submit", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.approvewithdeferdeployconfirm", LanguageID))%>');
            break;	
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deployoffers", LanguageID))%>':
            HandleRequestForMixOffers_Async(1, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.deploycomfirm", LanguageID))%> ');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%>':
            HandleRequestForMixOffers_Async(2, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.submitapprovalconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%>':
            HandleRequestForMixOffers_Async(3, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.approvewithdeployconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%>':
            HandleRequestForMixOffers_Async(4, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.approvewithdeferdeployconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%>':
            HandleRequestForMixOffers_Async(5, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deploy.submitapprovalconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%>':
            HandleRequestForMixOffers_Async(6, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deploy.approvewithdeployconfirm", LanguageID))%> ');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%>':
            HandleRequestForMixOffers_Async(7, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deploy.approvewithdeferdeployconfirm", LanguageID))%> ');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deferdeployoffers", LanguageID))%>':
            HandleRequestForMixOffers_Async(8, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.deferdeployconfirm", LanguageID))%> ');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%>':
            HandleRequestForMixOffers_Async(9, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deferdeploy.submitapprovalconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%>':
            HandleRequestForMixOffers_Async(10, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deferdeploy.approvewithdeployconfirm", LanguageID))%> ');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%>':
            HandleRequestForMixOffers_Async(11, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deferdeploy.approvewithdeferdeployconfirm", LanguageID))%> ');
            break;
         }
         actionele.selectedIndex=0;
       }          
   }
    var Wstatus=0;
   function ExecFilter(){
     var FilterElem = document.getElementById("Filteritems"); 
     Wstatus = FilterElem.options[FilterElem.selectedIndex].value;         
     frmdata = 'FolderID=' + selectedFolder + '&WFStatus=' + Wstatus;
     xmlhttpPost('folder-feeds.aspx?Action=LoadFilteredFolderItems', FILTER_ITEMS, frmdata);      
   
   }


  function getElementsByIdStartsWith(container, selectorTag, prefix) {
    var items = [];
    var myPosts = document.getElementById(container).getElementsByTagName(selectorTag);
    for (var i = 0; i < myPosts.length; i++) {
        //omitting undefined null check for brevity
        if (myPosts[i].id.lastIndexOf(prefix, 0) === 0) {
            items.push(myPosts[i]);
        }
    }
    return items;
  }
  
  function deletecells(rows){
     
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];        
      var mycell = row.cells[0];
		
      if (mycell != null){		           		  
        mycell.innerText = ""
      } 		
    }
  }
  
  function showactions() {
    var offerrows = [];
    //offerrows = document.getElementsByName("errdesc");
 
       offerrows = getElementsByIdStartsWith("tb1","tr","errdesc");
     
    if (selectedItems.length > 0) {
	  
        deletecells(offerrows);       
        ShowRequestApprovalActionItems(selectedItems);
        toggleDialog('performactions', true);
        
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }    
  }

  

  function renameFolder() {
    var folderElem = document.getElementById('folder' + selectedFolder);
    var folderNameElem = document.getElementById('foldername' + selectedFolder);
    
    if (folderElem != null && folderNameElem != null) {
      var rsp = prompt('<%Sendb(Copient.PhraseLib.Lookup("folders.EnterNewName", LanguageID))%>', '');
      if (rsp != null) {
        rsp = trimString(rsp);
        if (rsp != '') {
          if (rsp.length > 50) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NameTooLong", LanguageID))%>');
          } else {
              xmlhttpPost('/logix/folder-feeds.aspx?Action=RenameFolder', RENAME_FOLDER, 'FolderID=' + selectedFolder + '&FolderName=' + encodeURIComponent(rsp));
            folderNameElem.innerHTML = rsp.replace(/</g,'&lt;');
          }
        }
      } 
    }
  }
  
  function deleteFolder() {
    var folderElem = document.getElementById('folder' + selectedFolder);
    var deleteemptyfolder = <% Sendb(MyCommon.NZ(MyCommon.Fetch_SystemOption(134), 0))%>;
    var StatusValue= document.getElementById('statustext');
    if (folderElem != null) {
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.DeleteConfirm", LanguageID))%>')) {
            xmlhttpPost('/logix/folder-feeds.aspx?Action=DeleteFolder', DELETE_FOLDER, 'FolderID=' + selectedFolder);
            folderElem.parentNode.removeChild(folderElem);
            selectedFolder = 0; 
            if( deleteemptyfolder==0 || StatusValue.defaultValue=='0 item(s)' )
            {
                showContents( "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>"+'<p><b>Folder has been deleted</b></p>'+ '<\/div>');	
            }
        }
    }
  }
  function modifyFolder() {

    var folderElem = document.getElementById('folder' + selectedFolder);
    
    if (selectedFolder > 0) {
      if (folderElem != null) {
      xmlhttpPost('/logix/folder-feeds.aspx?Action=GenDivModFolder', GENERATE_DIV_FOLDER, 'FolderID=' + selectedFolder);
      } 
   } else {
      showContents('<p><%Sendb(Copient.PhraseLib.Lookup("folders.PleaseSelect", LanguageID))%><\/p>');
      }    
  }

  function generatedivfmodify(divHTML) {
  var elem = document.getElementById("modifyfolder");
   if (elem != null) {
     elem.innerHTML = divHTML;
     toggleDialog('modifyfolder', true);
   }
  } 
    
  function assignNoofDuplicateOffers(shown) {
	 var elem = document.getElementById('DuplicateNoofOffer');
     var OfferfadeElem = document.getElementById('OfferfadeDiv');
	 var fadeElem = document.getElementById('fadeDiv');
	 
	 if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
      }

      if (OfferfadeElem != null) {
        OfferfadeElem.style.display = (shown) ? 'block' : 'none';
      }

	  if (fadeElem != null) {
        fadeElem.style.display = (shown) ?  'none' : 'block';
      }

	  if (shown)  {
	   document.getElementById('txtDuplicateOffersCnt').value='1';
	   document.getElementById('txtDuplicateOffersCnt').focus();
	   ClearNoOfDuplicateOfferserror();
	   return false;
	  } else {
	  return true;
	  }
   } 
   
   
  	function showNoOfDuplicateOfferserror(content){
     var duplicateofferElem = document.getElementById("DuplicateOffererror");
      
      duplicateofferElem.style.display = 'block';
      duplicateofferElem.innerHTML = content;
    }
	
	function ClearNoOfDuplicateOfferserror(){
     var duplicateofferElem = document.getElementById("DuplicateOffererror");
      if (duplicateofferElem != null) {
        duplicateofferElem.style.display = 'none';
       }  
    }
	
	function addDuplicateOfferscount() {
	  var maxOffersperfolderduplicate = <% Sendb(MyCommon.NZ(MyCommon.Fetch_SystemOption(184),0)) %>;
	  if (maxOffersperfolderduplicate == 0 ) {
	    maxOffersperfolderduplicate = 99;
	  }
	  var dupOffersCntvalue =  document.getElementById("txtDuplicateOffersCnt").value;
	   if (dupOffersCntvalue != null && dupOffersCntvalue.trim() != "") {
	    if (!isNaN(dupOffersCntvalue)) {
		   ClearNoOfDuplicateOfferserror();
		   if (dupOffersCntvalue <= 0) {
		     showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.invalidDuplicateOfferCount", LanguageID))%>');
		   }
		   else if (dupOffersCntvalue > maxOffersperfolderduplicate) {
		     showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxOffersperfolderduplicate + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
		   }
		   else {
		       var actionele = document.getElementById('Actionitems');
               var folderlist = document.getElementById("folderList").value;
			   toggleDialogOfferDuplicate('DuplicateNoofOffer', false);
			   var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';
			   if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' ' + offerphrasetext + '. ' + ' <%Sendb(Copient.PhraseLib.Lookup("term.enteredDuplicateOffersCount", LanguageID))%> ' + dupOffersCntvalue + '.  <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
               frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=' + dupOffersCntvalue + '&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false + '&ActionItem=' + actionele.value + '&Fid=' + selectedFolder;
			   //ClosePopUp("");
			   xmlhttpPost('folder-feeds.aspx?Action=DuplicateOffers', DUP_OFFERS, frmdata);
               selectedItems = new Array();
               selectedItemTypes = new Array();
               selectedPromoEngines= new Array();
			   }	
            }
		}
		else {
		  showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.invalidDuplicateOfferCount", LanguageID))%>');
		}
	   }
	   else {
	     showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.enterDuplicateOfferCount", LanguageID))%>');
	   }
	}
	
  function DuplicateOfferstofolder() {
     if (document.getElementById('Actionitems').value == '8') {
	   TransferOffers();
	}
	else { 
    var actionele = document.getElementById('Actionitems');
    var folderlist = document.getElementById("folderList").value;
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(153)) %>;     
    if (selectedItems.length > 0) {
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';            
        if (maxoffers == 0 || selectedItems.length < maxoffers + 1) {
		   <%     if MyCommon.IsEngineInstalled(9) or MyCommon.IsEngineInstalled(0) then %>
		       assignNoofDuplicateOffers(true);
            <% Else %>
            if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + ' <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
               frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=' + 1 + '&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false + '&ActionItem=' + actionele.value + '&Fid=' + selectedFolder;
               //ClosePopUp("");
               xmlhttpPost('folder-feeds.aspx?Action=DuplicateOffers', DUP_OFFERS, frmdata);
               selectedItems = new Array();
               selectedItemTypes = new Array(); 
               selectedPromoEngines = new Array();
		     }
		   <% End If %>
        }
        else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
      }   
	}
  }

  function handlemassDeployConditional(offerswithoutcon) {
    //value of max offers should be fetched from a CM_System option
    
    if (selectedItems.length > 0) {
      
        if (confirm(offerswithoutcon + ' Offers do not have any condition. Are you sure you want to continue?')) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false +'&OffersWithoutConditions=' + true;
        xmlhttpPost('folder-feeds.aspx?Action=DeployOffers', MASSDEPLOY_OFFERS, frmdata);
        }
     }
  }
  function ApplyFolderStartDatetoOffer() {
  var frmdata="";  
   var performactionserrorElem = document.getElementById("performactionserror");
    if(selectedPromoEngines.length >0)
    {
        frmdata = 'PromoEngines='+selectedPromoEngines.join(',')+ "&";
    }
    if (selectedItems.length > 0) {
        frmdata += 'ItemIDs=' + selectedItems.join(',')+"&";
        frmdata += "FID="+selectedFolder;

        xmlhttpPost('folder-feeds.aspx?Action=ApplyFolderStartDatesToOffer', FOLDERSTART_OFFER, frmdata);
     }
  }
  
  function DeleteSelectedOffers() {
  var frmdata="";  
    var performactionserrorElem = document.getElementById("performactionserror");
     if(selectedPromoEngines.length >0)
    {
        frmdata = 'PromoEngines='+selectedPromoEngines.join(',')+ "&";
    }
    if (selectedItems.length > 0) {
        frmdata += 'ItemIDs=' + selectedItems.join(',')+"&";
        frmdata += "FID="+selectedFolder;

        xmlhttpPost('folder-feeds.aspx?Action=DeleteSelectedOffers', DELETESELECTED_OFFERS, frmdata);
     }
  }
  
  function ApplyFolderEndDatetoOffer() { 
  var frmdata="";  
   var performactionserrorElem = document.getElementById("performactionserror");
   if(selectedPromoEngines.length >0)
    {
        frmdata = 'PromoEngines='+selectedPromoEngines.join(',')+ "&";
    }
    if (selectedItems.length > 0) {
        frmdata += 'ItemIDs=' + selectedItems.join(',')+"&";
        frmdata += "FID="+selectedFolder;

        xmlhttpPost('folder-feeds.aspx?Action=ApplyFolderEndDatesToOffer', FOLDEREND_OFFER, frmdata);
     }
  }
  function ApplyFolderStartEndDatetoOffer() {
  var frmdata="";  
    var performactionserrorElem = document.getElementById("performactionserror");
     if(selectedPromoEngines.length >0)
    {
        frmdata = 'PromoEngines='+selectedPromoEngines.join(',')+ "&";
    }
    if (selectedItems.length > 0) {
        frmdata += 'ItemIDs=' + selectedItems.join(',')+"&";
        frmdata += "FID="+selectedFolder;

        xmlhttpPost('folder-feeds.aspx?Action=ApplyFolderStartEndDatesToOffer', FOLDERSTARTEND_OFFER, frmdata);
     }
  }

  function handlemassDeploy() {
    //value of max offers should be fetched from a CM_System option
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;   
    var deploynontemplates= '<%Sendb(Logix.UserRoles.DeployNonTemplateOffers) %>';
    var performactionserrorElem = document.getElementById("performactionserror"); 
    if (selectedItems.length > 0) {
      var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if(deploynontemplates == "False") {         
                performactionserrorElem.style.display = 'block';
                performactionserrorElem.style.color = 'red';
                performactionserrorElem.innerHTML ='<%Sendb(Copient.PhraseLib.Lookup("folder.Nodeploypermission", LanguageID))%> ';
                return;
        }
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.deployofferwarning", LanguageID))%> ' + selectedItems.length +' '+offerphrasetext+'. '+' <%Sendb(Copient.PhraseLib.Lookup("folder-deployBackground", LanguageID))%> \n' +  ' <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false + '&OffersWithoutConditions=' + false + '&FolderID=' +selectedFolder;
        xmlhttpPost('folder-feeds.aspx?Action=DeployOffers', MASSDEPLOY_OFFERS, frmdata);
        }
      }      
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.deployofferwarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        } 
      }    
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 

  }

  function handlemassDeferDeploy() {
    //value of max offers should be fetched from a CM_System option
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;   
    var performactionserrorElem = document.getElementById("performactionserror"); 
    if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.deferdeployofferwarning", LanguageID))%> ' + selectedItems.length +' '+offerphrasetext+'. '+' <%Sendb(Copient.PhraseLib.Lookup("folder-deferdeployBackground", LanguageID))%> \n' +  ' <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false + '&OffersWithoutConditions=' + false + '&FolderID=' +selectedFolder;
        xmlhttpPost('folder-feeds.aspx?Action=DeferDeployOffers', MASSDEFERDEPLOY_OFFERS, frmdata);
        }
      } 
      }   
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 

  }
function HandleMassRequestApproval_Async(approvalType, messageText) {
    //value of max offers should be fetched from a System option 152
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;   
    if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
         if (confirm(messageText)) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&ApprovalType=' + approvalType + '&FromOfferList=' + false + '&OffersWithoutConditions=' + false + '&FolderID=' +selectedFolder;
        xmlhttpPost('folder-feeds.aspx?Action=RequestApproval', MASSREQUESTAPPROVALFOROFFERS, frmdata);
        }
      } 
    else{
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
        alert('You are going to perform mass action on ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        } 
      }   
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 

  }
function HandleRequestForMixOffers_Async(deploymentType, messageText) {
    //value of max offers should be fetched from a System option 152
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;   
    if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
         if (confirm(messageText)) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&DeploymentType=' + deploymentType + '&OffersWithoutConditions=' + false + '&FolderID=' +selectedFolder + '&FromOfferList=' + false;
        xmlhttpPost('folder-feeds.aspx?Action=HandleDeployAndApprovalRequest', DEPLOYANDAPPROVAL_REQUEST, frmdata);
        }
      } 
    else{
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
        alert('You are going to perform mass action on ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        } 
      }   
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 

  }
  function handleReports(){     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(103)) %>;
     if (selectedItems.length > 0) {
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
        if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.navigatetoreportswarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false +'&FolderID='+selectedFolder;
          toggleDialog("performactions",false);
          toggleDialog("NavigatetoReports",true);
          xmlhttpPost('folder-feeds.aspx?Action=NavigatetoReports', NAVIGATETO_REPORTS, frmdata);         
          
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.navigatetoreportswarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 
  }

  function SendOutBound(){     
       
     var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(154)) %>;
     var str = '<%Sendb(Copient.PhraseLib.Lookup("folders.sendoutboundwarning", LanguageID))%> '
     str= str.replace("&#39;","''");
     if (selectedItems.length > 0) {
      var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm(str + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false+'&FolderID='+selectedFolder;
          xmlhttpPost('folder-feeds.aspx?Action=SendOutbound', SEND_OUTBOUND, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.sendoutboundwarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 
  }

  function ChangeWFStatustoPreValidate(){     
     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(100)) %>;
     if (selectedItems.length > 0) {
      var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.prevalidatewarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false;
          xmlhttpPost('folder-feeds.aspx?Action=WFStatustoPreValidate', WFSTAT_PREVALIDATE, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.prevalidatewarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 
  }

  function ChangeWFStatustoPostValidate(){     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(101)) %>;
     if (selectedItems.length > 0) {
      var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.postvalidatewarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false;
          xmlhttpPost('folder-feeds.aspx?Action=WFStatustoPostValidate', WFSTAT_POSTVALIDATE, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.postvalidatewarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 
  }

  function ChangeWFStatustoReadytoDeploy(){     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(102)) %>;
     if (selectedItems.length > 0) {
      var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';                    
        if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.readytodeploywarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + false;
          xmlhttpPost('folder-feeds.aspx?Action=WFStatustoReadytoDeploy', WFSTAT_READYTODEPLOY, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.readytodeploywarning", LanguageID))%> ' + selectedItems.length + ' '+offerphrasetext+'. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 
  }

  function TransferOffers(){    
    var actionele = document.getElementById('Actionitems');
    var destinationFolder = document.getElementById("folderList").value;
	if (sourceFolder != destinationFolder) {
      if (selectedItems.length > 0) {      
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';
	    if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.transferofferswarning", LanguageID))%> ' + selectedItems.length + ' ' + offerphrasetext + '. ' + ' <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          
         frmdata = 'ItemIDs=' + selectedItems.join(',')+'&sFolder=' + sourceFolder + '&dFolder=' + destinationFolder + '&FromOfferList=' + false +  '&ActionItem=' + actionele.value;
          xmlhttpPost('folder-feeds.aspx?Action=TransferOffers', TRANSFER_OFFERS, frmdata);
          selectedItems = new Array();
          selectedPromoEngines= new Array();
          
		}  
      }
      else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
      }
	}
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.samesourcedest", LanguageID))%>');
		document.getElementById('btnDupOffer').disabled = true; 
    }	
  }
  
   function handleAllItems() {
    var ids = new Array();
    var strURL="/logix/folder-feeds.aspx?Action=GetFolderItemIDs";
    $.post(strURL,{FolderID: selectedFolder,WFStatus:Wstatus},function (data) { 
    updatepage(data.trim())
    });
   
  }
  function updatepage(data){
     var ids = new Array();
     ids = data.split(',')
    var elem = null;
    var elemAll = document.getElementById("allitemIDs");    
    if (elemAll != null) {
      selectedItems = new Array();
      selectedPromoEngines=new Array();
       for (var i = 0; i < ids.length; i++) {
          if(ids[i]!= ""){
          //Update allItemIDs
          if (elemAll.checked) { 
            updateidlist(ids[i], elemAll.checked); 
          }
          //Update the Checkboxes on the page
         elem = document.getElementById("itemID" + ids[i]);
         if(elem!=null){
            elem.checked = elemAll.checked;
         }
     }
    }                 
    }
  }

  function removeFromFolder() {
    var frmdata = '';
    var response = '';
    var msg = '';
    
    if (selectedItems.length > 0) {
      frmdata = 'FolderID=' + selectedFolder + '&ItemIDs=' + selectedItems.join(',');
      
      msg = '<%Sendb(Copient.PhraseLib.Lookup("folders.RemoveConfirm", LanguageID))%>';
      response = confirm(msg);
      if (response != '') {
        xmlhttpPost('folder-feeds.aspx?Action=RemoveItemsFromFolder', REMOVE_ITEMS, frmdata);
        selectedItems = new Array();
        selectedPromoEngines = new Array();
      }
    } else {
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedToRemove", LanguageID))%>');
    }
  }
  
  function addToFolder(FolderID, Items) {
    var frmdata = '';
    var FilterElem = document.getElementById("Filteritems"); 
    var selectedfilter = FilterElem.options[FilterElem.selectedIndex].value;
    if (selectedfilter != 0) {      
      frmdata = 'FolderID=' + selectedFolder + '&LinkIDs=' + selectedItems.join(',') + '&LinkTypeIDs=' + selectedItemTypes.join(',') + '&searchterms=' + document.getElementById('searchterms').value + '&WFStatus=' + selectedfilter;
    }else{
      frmdata = 'FolderID=' + selectedFolder + '&LinkIDs=' + selectedItems.join(',') + '&LinkTypeIDs=' + selectedItemTypes.join(',') + '&searchterms=' + document.getElementById('searchterms').value;
    }
    
    xmlhttpPost('folder-feeds.aspx?Action=AddItemsToFolder', ADD_ITEMS, frmdata);
    selectedItems = new Array();
    selectedItemTypes = new Array();
    selectedPromoEngines = new Array();
  }
  
  function submitToAddItems(linkid, linktypeid,promoengine, bChecked, warningmessage) {
    var index = -1;
    var i = 0;
    var iID = 0;
    var iTypeID = 0;
    var iPromoEngine = "";
    
    if (typeof warningmessage == 'undefined') {
    warningmessage = '';
    }
    if (warningmessage == trimString('<b><%Sendb(Copient.PhraseLib.Lookup("folders.OfferNotInDateRange", LanguageID))%>')){
     showwarning(warningmessage);
    }
    else { 
    iID = parseInt(linkid);
    iTypeID = parseInt(linktypeid);
    iPromoEngine = promoengine.toString();
    if (!bChecked) {
      index = search(selectedItems, iID, false)
      if (index > -1) {
        selectedItems.splice(index, 1);
        selectedItemTypes.splice(index, 1);
        selectedPromoEngines.splice(index,1);
      }
    } else {
      index = search(selectedItems, iID, true)
      selectedItems.splice(index, 0, iID);
      selectedItemTypes.splice(index, 0, iTypeID);
      selectedPromoEngines.splice(index,0,iPromoEngine);
    }
   } 
     showwarning(warningmessage);  
  }

  
  function submitToRemoveItems(itemid,promoengine, bChecked) {
    var index = -1;
    var i = 0;
    var iID = 0;   
     var iPromoEngine ="";
    var elemAll = document.getElementById("allitemIDs");
    iID = parseInt(itemid);
    iPromoEngine = promoengine.toString();
    if (!bChecked) {
      index = search(selectedItems.sort(), iID, false)
      if (index > -1) {
       selectedItems.splice(index, 1);
       selectedPromoEngines.splice(index,1);
      }
    } else {
      index = search(selectedItems, iID, true)
      selectedItems.splice(index, 0, iID);
       selectedPromoEngines.splice(index,0,iPromoEngine);
    }    
    elemAll.checked = iID.checked; 
  }
  
  function updateidlist(itemid, bChecked) {    
    var index = -1;
    var i = 0;
    var iID = 0;        
    iID = parseInt(itemid);
    if (!bChecked) {      
      
      index = selectedItems.indexOf(iID)
      if (index > -1) {
       selectedItems.splice(index, 1);
      }
      
    } else {    
     
      selectedItems.splice(index, 0, iID);
      
    }        
  }

  function search(o, v, i) {
    /*
    vector (o):  array that will be looked up
    value (v):   object that will be searched
    insert (i):  if true, the function will return the index where the value should be inserted
    to keep the array ordered, otherwise returns the index where the value was found
    or -1 if it wasn't found
    */
    var h = o.length, l = -1, m;
    while (h - l > 1)
      if (o[m = h + l >> 1] < v) l = m;
    else h = m;
    return o[h] != v ? i ? h : -1 : h;
  };
  
  function submitSearch() {
    var frmdata = '';
    var searchElem = document.getElementById('searchterms');
    var searchTerms = '';
    
    if (searchElem != null) {
        if(document.getElementById('searchterms').value.trim().length > 0){
            searchTerms = document.getElementById('searchterms').value;
            frmdata = 'FolderID=' + selectedFolder + '&searchterms=' + encodeURIComponent(searchTerms);

            document.body.style.cursor = 'wait';
      
            var elemfolderbox = document.getElementById("folderpopulate");
            var elemsresults = document.getElementById("searchResults");
            elemsresults.style.height = elemfolderbox.offsetHeight - elemsresults.offsetTop - 5 + "px";

            xmlhttpPost('folder-feeds.aspx?Action=SendFoundOffers', VIEW_AVAILABLE, frmdata); 
        } else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("folders.EnterFolderName", LanguageID))%>');
        }
    }
  }
  
  function submitfSearch() {
    var frmdata = '';
    var searchElem = document.getElementById('fsearchterms');
    var typeElem = document.getElementById('fsearchType');
    var searchTerms = '';
    var searchType = '';
    var action = '';
    
    if (searchElem != null && typeElem != null) {
      searchTerms = searchElem.value;
      searchType = typeElem.value;
      
      frmdata = '&searchterms=' + encodeURIComponent(searchTerms)
      document.body.style.cursor = 'wait';
      
      var elemfolderbox = document.getElementById("foldersearch");
      var elemsresults = document.getElementById("fsearchResults");
      elemsresults.style.height = elemfolderbox.offsetHeight - elemsresults.offsetTop - 5 + "px";

      action = (searchType == 2) ? 'SendOfferSearch' : 'SendFolderSearch';
      xmlhttpPost('folder-feeds.aspx?Action=' + action, FOLDER_SEARCH, frmdata);
    }
  }
  
  function updateStatusBar() {
    var statusTextElem = document.getElementById('statustext');
    var statusBarElem = document.getElementById('folderstatusbar');
    var statusText = '';
    
    if (statusBarElem != null) {
      if (statusTextElem != null) {
        statusText = statusTextElem.value;
      } else {
        statusText = '<%Sendb(Copient.PhraseLib.Lookup("term.ready", LanguageID))%>.';
      }
      statusBarElem.innerHTML = statusText;
    }
  }

  function setSearchFocus() {
    var searchTextElem = document.getElementById('searchterms');
    
    if (searchTextElem != null && searchTextElem.style.display != 'none' &&  searchTextElem.style.display != '') {
      searchTextElem.focus();
      searchTextElem.select();
    }
  }

  function setSearchBoxText(text) {
    var searchTextElem = document.getElementById('searchterms');
    
    if (searchTextElem != null) {
      searchTextElem.value = text;
    }
  }

  function handleSearchKeyDown(e) {
    var key = e.which ? e.which : e.keyCode;

    if (key == 13) {
      submitSearch();
    }
  }
  
  function handlefSearchKeyDown(e) {
    var key = e.which ? e.which : e.keyCode;

    if (key == 13) {
      submitfSearch();
    }
  }
  
  function toggleSearch(shown) {
    var searchTextElem = document.getElementById('searchterms');
    var searchSubmitElem = document.getElementById('searchsubmit');
    
    if (searchTextElem != null && searchSubmitElem != null) {
      searchTextElem.style.display = (shown) ? 'inline' : 'none';
      searchSubmitElem.style.display = (shown) ? 'inline' : 'none';
    }
    
    if (!shown) {
      setSearchBoxText('');
    }
  }
  
  function toggleFolder(folderid) {
    var elem = document.getElementById('folder' + folderid);
    var expanderElem = document.getElementById('expander' + folderid);
    var isOpen = false;
    var className = '';

    if (elem != null && expanderElem != null) {
      isOpen = (expanderElem.src.indexOf('minus.png') > -1)
      expanderElem.src = (isOpen) ? '../images/plus.png' : '../images/minus.png';
      className = (elem.getAttribute('className') != null) ? 'className' : 'class';
      if (isOpen) {
        elem.setAttribute(className, "folder closed");
      } else {
        elem.setAttribute(className, "folder");
      }
    }
  }
  
  function expandFolder(folderid) {
    var elem = document.getElementById('folder' + folderid);
    var expanderElem = document.getElementById('expander' + folderid);
    var isOpen = false;
    var className = '';

    if (elem != null && expanderElem != null) {
      expanderElem.src = '../images/minus.png';

      className = (elem.getAttribute('className') != null) ? 'className' : 'class';
      elem.setAttribute(className, "folder");
    }
  }

  function handleAddItemClick() {
    var searchElem = document.getElementById("searchterms");

    if (selectedFolder > 0) {
      clearAddItem();      
      toggleDialog('folderpopulate', true);
      if (searchElem != null) {
        searchElem.focus();
      }  
    } else {
      showContents('<p><%Sendb(Copient.PhraseLib.Lookup("folders.PleaseSelect", LanguageID))%><\/p>');
    }
  }
  
  function clearAddItem() {
    var searchElem = document.getElementById("searchterms");
    var resultsElem = document.getElementById("searchResults");
    var btnAddElem = document.getElementById("btnAddItems");
//
    var resultsWarningElem = document.getElementById("AddItemWarninng");
    
    if (searchElem != null) {
      searchElem.value = ""
    }
    if (resultsElem != null) {
      resultsElem.innerHTML = "";
    }
    if (btnAddElem != null) {
      btnAddElem.style.visibility = 'hidden';
      btnAddElem.disabled = false;  
    }
//
    if (resultsWarningElem != null) {
      resultsWarningElem.style.display = 'none';
    }
  }
  
  function clearErrorContents() {
    var modifyfolderElem = document.getElementById("modifyfoldererror");
    var createfolderElem = document.getElementById("createfoldererror");
    var performActionsElem = document.getElementById("performactionserror");
    var duplicateofferdiv = document.getElementById("execduplicateoffer");

     if (createfolderElem != null) {
      createfolderElem.style.display = 'none';
      }     
     if (modifyfolderElem != null) {
      modifyfolderElem.style.display = 'none';
      } 
     if (performActionsElem != null) {
      performActionsElem.style.display = 'none';
     } 
     if (duplicateofferdiv != null) {
      duplicateofferdiv.style.visibility = 'hidden';
     }
  }

  function folderLinkClicked(folderID) {
    navigateToFolder(folderID);
    ensureFolderSelected(folderID);
    toggleDialog('foldersearch', false);
    showfSearchContents('');
    var searchElem = document.getElementById('fsearchterms');
    searchElem.value = '';
  }
  
  function ensureFolderSelected(FolderID) {
    if (FolderID != selectedFolder) {
      highlightSelected(null, FolderID)
    }
  }

  function isIE() {
    return /msie/i.test(navigator.userAgent) && !/opera/i.test(navigator.userAgent);
  } 
       

       function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el=(typeof event!=='undefined')? event.srcElement : e.target        
      
      if (el != null) {
        var pickerDiv = document.getElementById(datePickerDivID);
        if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
          if (el.id!="folder-start-picker" && el.id!="folder-end-picker" 
             && el.id!="foldermodified-start-picker" && el.id!="foldermodified-end-picker") {
            if (!isDatePickerControl(el.className)) {
              pickerDiv.style.visibility = "hidden";
              pickerDiv.style.display = "none";  
              if (calFrame != null) {
                calFrame.style.visibility = 'hidden';
                calFrame.style.display = 'none';
              }
            }
          } else  {
              pickerDiv.style.visibility = "visible";            
              pickerDiv.style.display = "block";            
              if (calFrame != null) {
                calFrame.style.visibility = 'visible';
                calFrame.style.display = 'block';
              }
          }
        }
      }
    }

function isDatePickerControl(ctrlClass) {
      var retVal = false;
      
      if (ctrlClass != null && ctrlClass.length >= 2) {
        if (ctrlClass.substring(0,2) == "dp") {
          retVal = true;
        }
      }

      return retVal;
    }   

   function SetDefaultFolder(responseText){ 
        var defaultUEFolderele=document.getElementById('defaultUEFolder');
        var defaultUEFolderLabel=document.getElementById('defaultUEFolderLabel');
        
        if(responseText !='' && defaultUEFolderele != null && defaultUEFolderLabel !=null){
            defaultUEFolderele.checked = false; 
                    
            if (responseText == 1 ){
                defaultUEFolderele.style.display='none';
                defaultUEFolderLabel.style.display='none';           
            }
            else{
                defaultUEFolderele.style.display='';
                defaultUEFolderLabel.style.display='';
            }      
        }

  }

  function CheckDefaultFolder(){
    var defaultUEFolderele = document.getElementById("defaultUEFolder");  
    if(defaultUEFolderele !=null){
        xmlhttpPost('/logix/folder-feeds.aspx?Action=CheckDefaultFolder', CHECK_DEFAULTFOLDER,'');
    }
  }

       
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 2)
  Send_Subtabs(Logix, 20, 6)
  
  If (Logix.UserRoles.AccessFolders = False) Then
    Send_Denied(1, "perm.folders-access")
    GoTo done
  End If
%>
<div id="intro">
  <div id="foldertoolbar" class="foldertoolbar">
    <%
        Sendb("<h1 id=""title"">")
        Sendb(Copient.PhraseLib.Lookup("term.folders", LanguageID))
        Send("</h1>")
        If (Logix.UserRoles.CreateFolders OrElse Logix.UserRoles.DeleteFolders OrElse Logix.UserRoles.EditFolders) Then
            Send("<div id=""foldertools"">")
            Send("  <img src=""../images/folders/vr.png"" />")
            If (Logix.UserRoles.CreateFolders) Then
                Send("  <img src=""../images/folders/folder-create.png"" alt=""" & Copient.PhraseLib.Lookup("folders.CreateFolder", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.CreateFolder", LanguageID) & """ onclick=""javascript:clearErrorContents();toggleDialog('foldercreate', true);CheckDefaultFolder();"" />")
            End If
            If (Logix.UserRoles.DeleteFolders) Then
                Send("  <img src=""../images/folders/folder-delete.png"" alt=""" & Copient.PhraseLib.Lookup("folders.DeleteFolder", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.DeleteFolder", LanguageID) & """ onclick=""javascript:deleteFolder();"" />")
            End If
            If (Logix.UserRoles.EditFolders) Then
                Send("  <img src=""../images/folders/foldergear.jpg"" alt=""Folder Settings"" title=""" & Copient.PhraseLib.Lookup("folders.RenameFolder", LanguageID) & """ onclick=""javascript:clearErrorContents();modifyFolder();"" />")
            End If
            Send("</div>")
        End If
        If (Logix.UserRoles.EditFolders) OrElse (Logix.UserRoles.AssignOfferstoFolderOnly) Then
            Send("<div id=""itemtools"">")
            Send("  <img src=""../images/folders/vr.png"" />")
            Send("  <img src=""../images/folders/item-add.png"" alt=""" & Copient.PhraseLib.Lookup("folders.AddItems", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.AddItems", LanguageID) & """ onclick=""javascript:handleAddItemClick();"" />")
            Send("  <img src=""../images/folders/item-remove.png"" alt=""" & Copient.PhraseLib.Lookup("folders.RemoveItems", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.RemoveItems", LanguageID) & """ onclick=""javascript:removeFromFolder();"" />")
            Send("</div>")
        End If
        If (Logix.UserRoles.EditFolders) Then
            Send("<div id=""itemtools"">")
            Send("  <img src=""../images/folders/deploy.png"" alt=""" & Copient.PhraseLib.Lookup("folders.performaction", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.performaction", LanguageID) & """ onclick=""javascript:clearErrorContents(); javascript:showactions();"" />")
            Send("  <img src=""../images/folders/vr.png"" />")
            Send("  <img src=""../images/folders/item-search.png"" alt=""" & Copient.PhraseLib.Lookup("folders.Search", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.Search", LanguageID) & """ onclick=""javascript:toggleDialog('foldersearch', true);"" />")
            If bWorkflowActive Then
                Send("  <img src=""../images/folders/filter.png"" alt=""" & Copient.PhraseLib.Lookup("folders.performfilter", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.performfilter", LanguageID) & """ onclick=""javascript:toggleDialog('performfilter',true);"" />")
            End If
            'Send("  <input type=""text"" id=""globalterms"" name=""globalterms"" value="""" onkeydown="""" />")
            'Send("  <input type=""button"" id=""globalsubmit"" name=""globalsubmit"" value=""Search"" onclick="""" />")                        
            Send("</div>")
        End If
        If (Logix.UserRoles.AssignFolders And Not Logix.UserRoles.EditFolders) OrElse (Logix.UserRoles.AssignOfferstoFolderOnly And Not Logix.UserRoles.EditFolders) Then
            Send("<div id=""itemtools"">")
            Send("  <img src=""../images/folders/item-search.png"" alt=""" & Copient.PhraseLib.Lookup("folders.Search", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.Search", LanguageID) & """ onclick=""javascript:toggleDialog('foldersearch', true);"" />")
            Send("</div>")
        End If
    %>
  </div>
</div>
<div id="fadeDiv"></div>
<span id="reloadpage" style="display:none;">0</span>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="foldertree" class="foldertree">
    <%
      Send("")      
      LoadFolderStats(MyCommon)
      WriteFolders(MyCommon, 0)
      Send("")
    %>
    <hr class="hidden" />
  </div>
  <div id="foldercontents" class="foldercontents">
    <%
    %>
    <hr class="hidden" />
  </div>
  <div id="folderstatusbar" class="folderstatusbar">
    <%Sendb(Copient.PhraseLib.Lookup("term.ready", LanguageID))%>
  </div>
</div>

<script runat="server">
    Dim Counter As Integer = 0
    Dim htFolders As Hashtable = Nothing
    Dim folderWithDeploymentErrors() As Integer

    Public Sub LoadFolderStats(ByRef MyCommon As Copient.CommonInc)
        Dim FolderData As FolderMetaData
        Dim dt As DataTable
        Dim row As DataRow
        Dim count As Int16 = 0

        MyCommon.QueryStr = "dbo.pa_FolderSummary_Select"
        MyCommon.Open_LRTsp()
        dt = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        htFolders = New Hashtable(dt.Rows.Count + 1)

        For Each row In dt.Rows
            If Not htFolders.ContainsKey(MyCommon.NZ(row.Item("FolderID"), "0").ToString) Then
                FolderData = New FolderMetaData(MyCommon.NZ(row.Item("HasContent"), False), MyCommon.NZ(row.Item("HasChildren"), False))
                htFolders.Add(MyCommon.NZ(row.Item("FolderID"), "0").ToString, FolderData)
            End If
        Next

        MyCommon.QueryStr = "dbo.pa_GetFolders_WithDeploymentErrors"
        MyCommon.Open_LRTsp()
        dt = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            ReDim folderWithDeploymentErrors(dt.Rows.Count)
            For Each row In dt.Rows
                folderWithDeploymentErrors(count) = Integer.Parse(MyCommon.NZ(row.Item("FolderID"), "0"))
                count = count + 1
            Next
        End If
    End Sub

    Public Sub WriteFolders(ByRef MyCommon As Copient.CommonInc, ByVal ParentFolderID As Integer)
        Dim dt As DataTable
        Dim row As DataRow
        Dim level As Integer = 0
        Dim FolderID As Integer = 0
        Dim FolderData As FolderMetaData
        Dim ExpanderImg As String = ""
        Dim FolderImg As String = ""
        Dim ExpanderClickable As Boolean = False

        MyCommon.QueryStr = "select FolderID, ParentFolderID, FolderName " &
                            "from Folders with (NoLock) where ParentFolderID=" & ParentFolderID & " order by FolderName;"
        dt = MyCommon.LRT_Select
        For Each row In dt.Rows
            FolderID = MyCommon.NZ(row.Item("FolderID"), 0)

            FolderData = Nothing
            ExpanderImg = ""
            FolderImg = ""
            If htFolders.ContainsKey(FolderID.ToString) Then
                FolderData = htFolders.Item(FolderID.ToString)
                ExpanderImg = IIf(FolderData.HasSubFolders, "plus.png", "blank-clear.png")
                ExpanderClickable = IIf(FolderData.HasSubFolders, True, False)
                FolderImg = IIf(FolderData.HasItems, "folder-full.png", "folder.png")
            End If

            Dim folderName As String = MyCommon.NZ(row.Item("folderName"), "Unnamed")
            Send("<input type=""hidden"" id=""hdnfolderid"" name=""hdnfolderid"" value=""" & FolderID & """/>")
            Send("<div class=""folder closed"" id=""folder" & FolderID & """" & IIf(ParentFolderID = 0, " style=""display:block !important;""", "") & ">")
            Sendb("<div class=""folderrow"" id=""folderrow" & FolderID & """ onclick=""javascript:highlightSelected(event," & FolderID & ");"">")
            Sendb("<img src=""../images/" & ExpanderImg & """ class=""expander"" id=""expander" & FolderID & """" & IIf(ExpanderClickable, " onclick=""javascript:toggleFolder(" & FolderID & ");""", "") & " />")
            Sendb("<img src=""../images/" & FolderImg & """ class=""folderimg"" id=""folderimg" & FolderID & """ />")
            If folderWithDeploymentErrors IsNot Nothing AndAlso folderWithDeploymentErrors.Contains(FolderID) Then
                Sendb("<span class=""foldername"" style=""color:red;"" id=""foldername" & FolderID & """>" & MyCommon.NZ(row.Item("FolderName"), "Unnamed") & "</span>")
            Else
                Sendb("<span class=""foldername"" id=""foldername" & FolderID & """>" & MyCommon.NZ(row.Item("FolderName"), "Unnamed") & "</span>")
            End If
            Send("</div>")
            WriteFolders(MyCommon, FolderID)
            Send("</div>")
        Next
    End Sub

    Public Class FolderMetaData
        Public HasItems As Boolean
        Public HasSubFolders As Boolean
        Public Sub New(ByVal Items As Boolean, ByVal SubFolders As Boolean)
            HasItems = Items
            HasSubFolders = SubFolders
        End Sub
    End Class

</script>

<script type="text/javascript">
  if (window.captureEvents) {
    window.captureEvents(Event.CLICK);
    window.onclick = handlePageClick;
  } else {
    document.onclick = handlePageClick;
    }

    /// Function to allow only numeric values to be input ///
function NumberChecking(evt, src, allowDecimal) {
    var nkeycode = (window.event) ? window.event.keyCode : evt.which;
    var exceptionKeycodes = new Array(8, 9, 13, 16);

    if ((nkeycode >= 48 && nkeycode <= 57)) {
        return true;
    } else {
        for (var i = 0; i < exceptionKeycodes.length; i++) {
            if (nkeycode == exceptionKeycodes[i]) {
                return true;
            }
        }
        if (allowDecimal == true && (nkeycode == 110 || nkeycode == 190)) {
            if (src != null && src.value.indexOf(".") < 0) {
                return true;
            }
        }
        return false;
    }
}
</script>
<%
done:
  Send_FocusScript("mainform", "searchterms")
  Send_WrapEnd()
%>

<div id="foldercreate" class="folderdialog" style="height:auto; padding-bottom:10px">
<div class="foldertitlebar">
<span class="dialogtitle"><% Sendb(Copient.PhraseLib.Lookup("folders.CreateFolder", LanguageID))%></span>
<span class="dialogclose" onclick="toggleDialog('foldercreate', false);">X</span>
</div>
<div id="createfoldererror" style="display:none;color:red;">
</div>
<div class="dialogcontents">

<br class="half"/>
<label for="newFolderName"><% Sendb(Copient.PhraseLib.Lookup("folders.EnterFolderName", LanguageID))%>:</label><br />
<input type="text" id="newFolderName" name="newFolderName" maxlength="50" value="" class="mediumlong" />
<br />
<label for="folderstart"><% Sendb(Copient.PhraseLib.Lookup("folders.FolderDate", LanguageID))%>:</label><br />
<input type="text" class="short" id="folderstart" name="folderstart" maxlength="10" value="" />
<img src="../images/calendar.png" class="calendar" id="folder-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('folderstart', event);" />

    <%If bAllowTimeWithStartEndDates Then%>
              <input class="shortest" id="form_FolderStartHr" maxlength="2" onkeypress="return NumberChecking(event,false,false)" name="form_FolderStartHr"
               type="text" value="00" />:<input
              class="shortest" id="form_FolderStartMin" maxlength="2" name="form_FolderStartMin" type="text" onkeypress="return NumberChecking(event,false,false)"
               value="00" />
    <% End If%>

<% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
<input type="text" class="short" id="folderend" name="folderend" maxlength="10" value="" />      
<img src="../images/calendar.png" class="calendar" id="folder-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('folderend', event);" />
<%If bAllowTimeWithStartEndDates Then%>
              <input class="shortest" id="form_FolderEndHr" maxlength="2" name="form_FolderEndHr" onkeypress="return NumberChecking(event,false,false)"
               type="text" value="00" />:<input
              class="shortest" id="form_FolderEndMin" maxlength="2" name="form_FolderEndMin" type="text" onkeypress="return NumberChecking(event,false,false)"
               value="00" />
            <% Else%>
            <% Sendb("(" & MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern & ")")%>
            <% End If%>

<%--<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern)%>--%>
   

<%

    Dim hasDeferDeployPermission As Boolean = False
    If (MyCommon.Fetch_SystemOption(262)) Then
        hasDeferDeployPermission = Logix.UserRoles.DeferDeployNonTemplateOffers OrElse Logix.UserRoles.DeferDeployTemplateOffers
    Else
        hasDeferDeployPermission = Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers
    End If
    
If MyCommon.Fetch_SystemOption(123) = "1" Then
MyCommon.QueryStr = "select ThemeId,ThemeDescription from Themes;"
rst = MyCommon.LRT_Select
Send(" <label for=""Theme"">Theme:</label><br />")
Send("    <select name=""Theme"" id=""Theme"" >")
Send("      <option value=""-1""></option>")
index = 0
For Each row In rst.Rows
  Send(" <option value=""" & index & """>" & MyCommon.NZ(row.Item("ThemeDescription"), "") & "</option>")
  index = index + 1
Next
Send("    </select>")
End If
    %>
          
<!-- Remove hidden field when access levels are implemented -->
<input type="hidden" id="accessLevel" name="accessLevel" value="3" />
<%
  'MyCommon.QueryStr = "select FolderAccessLevelID, Description, PhraseID " & _
  '                    "from FolderAccessLevels with (NoLock) order by FolderAccessLevelID desc;"
  'dt = MyCommon.LRT_Select
  'If dt.Rows.Count > 0 Then
  '  Send("<label for=""accessLevel"">" & Copient.PhraseLib.Lookup("folders.SelectAccessLevel", LanguageID) & ":</label><br />")
  '  Send("<select id=""accessLevel"" name=""accessLevel"" class=""mediumlong"">")
  '  For Each row In dt.Rows
  '    Send("<option value=""" & MyCommon.NZ(row.Item("FolderAccessLevelID"), 3) & """>" & _
  '         Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("Description"), "")) & "</option>")
  '  Next
  '  Send("</select>")
  'End If
%>
<br />  
    <%
      'If Ue engine is installed and no folders are selected as default
        If (MyCommon.IsEngineInstalled(9)) Then
            
        Send("<input type=""checkbox"" name=""defaultUEFolder"" id=""defaultUEFolder""  /><label for=""defaultUEFolder"" id =""defaultUEFolderLabel"" checked="""">" & Copient.PhraseLib.Lookup("term.defaultuefolder", LanguageID) & "</label><br/><br/>")
           
    End If
    %>
    <br />
<input type="button" name="btnNewFolder" id="btnNewFolder" value="<%Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>" onclick="javascript:createFolder();" />
    <br />
</div>
</div>



<div id="OfferNavigationDialog" class="folderdialog OfferNavigationDialog" style="height: auto; width: 250px;">
  <div class="foldertitlebar">
    <span class="dialogtitle">
      <% Sendb(Copient.PhraseLib.Lookup("term.offer-navigation", LanguageID))%>
    </span>
    <span class="dialogclose" onclick="toggleDialogNoFade('OfferNavigationDialog', false);">X</span>
  </div>
  <div class="dialogcontents">
    <div id="Div2" style="display: none; color: red;">
    </div>
    <table style="width: 90%">
      <tr><td>
        <span id="spnOfferValidationMsg"></span><br /><br />
      </td>
      </tr>
      <tr align="left">
        <td>
        <a href="#" id="lnkViewOffer">
            <% Sendb(Copient.PhraseLib.Lookup("term.viewoffer", LanguageID))%>
        </a><br /><br />
        <a href="#" id="lnkViewCollisionReport">
            <% Sendb(Copient.PhraseLib.Lookup("term.view", LanguageID))%>&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.collisionreport", LanguageID))%>
        </a>
        <span id="spnViewCollisionReport" style="color:Gray">
            <% Sendb(Copient.PhraseLib.Lookup("term.view", LanguageID))%>&nbsp;<% Sendb(Copient.PhraseLib.Lookup("term.collisionreport", LanguageID))%>
        </span><br /><br />
        </td>
      </tr>
    </table>
  </div>
</div>

<div id="modifyfolder" class="folderdialog" style="height:auto; padding-bottom:10px"> 
</div>
<div id="performfilter" class="folderdialog"> 

<div class="foldertitlebar">
<span class="dialogtitle">WorkFlow Status Filter</span>
<span class="dialogclose" onclick="toggleDialog('performfilter', false);">X</span>
</div>
<div class="dialogcontents">
<br />
<br class="half"/>
<label for="Theme">Perform Filter:</label>&nbsp;&nbsp;
<select name="Filteritems" id="Filteritems">
<option value="0">No Filter</option>
<option value="4"><%Sendb(Copient.PhraseLib.Lookup("list.showonlydraft", LanguageID))%></option>
<option value="1"><%Sendb(Copient.PhraseLib.Lookup("list.showonly-prevalidate", LanguageID))%></option>
<option value="2"><%Sendb(Copient.PhraseLib.Lookup("list.showonly-postvalidate", LanguageID))%></option>
<option value="3"><%Sendb(Copient.PhraseLib.Lookup("list.showonly-readytodeploy", LanguageID))%></option>
</select>
<input type="button" name="ExecAction" id="Button1" value="Execute" onclick="javascript:ExecFilter();"/>
</div>
</div>
<div id="performactions" class="folderdialog"> 
<div class="foldertitlebar">
<span class="dialogtitle"><%Sendb(Copient.PhraseLib.Lookup("folders.performaction", LanguageID))%></span>
<span class="dialogclose" onclick="toggleDialog('performactions', false);">X</span>
</div>
<div id="performactionserror" style="display:none;">
</div>
<div class="dialogcontents">
<br />
<br class="half"/>
<label for="Theme"><%Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>:</label>&nbsp;&nbsp;
<select name="Actionitems" id="Actionitems" width=300 STYLE="width: 300px"  onchange="javascript:handleAction();">
<option value="-1" ><%Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%></option>
<option id="deploy" value="0"><%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%></option>
    <%If Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers Then%>
    <option id="reqapproval" value="14"><%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%></option>
    <option id="reqapprovedeploy" value="15"><%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%></option>
<% End If %>
<%If (hasDeferDeployPermission) Then%>
    <option id="reqapprovedeferdeploy" value="21"><%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%></option>
<% End If %>
<%If Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers Then%>
    <option id="onlyreqapproval" value="17"><%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%></option>
    <option id="onlyreqapprovedeploy" value="18"><%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%></option>
<% End If %>
<%If (hasDeferDeployPermission) Then%>
    <option id="onlyreqapprovedeferdeploy" value="22"><%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%></option>
<% End If %>
<%If Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers Then%>
    <option id="deployreqapproval" value="19"><%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%></option>
    <option id="deployreqapprovedeploy" value="20"><%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%></option>
<% End If %>
<%If (hasDeferDeployPermission) Then%>
    <option id="deployreqapprovedeferdeploy" value="23"><%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%></option>
<% End If %>
<%If (hasDeferDeployPermission) Then%>
    <option id="deferdeploy" value="13"><%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID))%></option>
    <option id="deferdeployreqapproval" value="25"><%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%></option>
    <option id="deferdeployreqapprovedeploy" value="26"><%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%></option>
    <option id="deferdeployreqapprovedeferdeploy" value="27"><%Sendb(Copient.PhraseLib.Lookup("term.deferdeploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID))%></option>
<%End If%>
<%If Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers Then%>
    <option id="deploymixoffers" value="16"><%Sendb(Copient.PhraseLib.Lookup("term.deployoffers", LanguageID))%></option>
<% End If %>
<%If (hasDeferDeployPermission) Then%>
    <option id="deferdeploymixoffers" value="24"><%Sendb(Copient.PhraseLib.Lookup("term.deferdeployoffers", LanguageID))%></option>
<% End If %>
    <option value="9"><% If bAllowTimeWithStartEndDates Then
                              Sendb(Copient.PhraseLib.Lookup("folder.startdatetimetooffer", LanguageID))
                          Else
                              Sendb(Copient.PhraseLib.Lookup("folder.startdatetooffer", LanguageID))
                          End If
                          %>
                        
    </option>
    <option value="10"><%
                           If bAllowTimeWithStartEndDates Then
                               Sendb(Copient.PhraseLib.Lookup("folder.enddatetimettooffer", LanguageID))
                           Else
                               Sendb(Copient.PhraseLib.Lookup("folder.enddatettooffer", LanguageID))
                           End If

                           %></option>
    <option value="11"><%
                           If bAllowTimeWithStartEndDates Then
                               Sendb(Copient.PhraseLib.Lookup("folder.startenddatewithtimetooffer", LanguageID))
                           Else
                               Sendb(Copient.PhraseLib.Lookup("folder.startenddatestooffer", LanguageID))
                           End If

                           %></option>
<%  If MyCommon.Fetch_SystemOption(286) = "1" Then%>
<%  If ((Logix.UserRoles.CreateOfferFromBlank OrElse Logix.UserRoles.CopyOfferCreatedFromBlank OrElse Logix.UserRoles.CopyOfferCreatedFromTemplate) And Not bTestSystem) Then%>
    <option value="1"><%Sendb(Copient.PhraseLib.Lookup("folders.duplicateoffer", LanguageID))%></option>
<%End If %>
<% Else%>
<%  If (Logix.UserRoles.CreateOfferFromBlank And Not bTestSystem) Then%>
    <option value="1"><%Sendb(Copient.PhraseLib.Lookup("folders.duplicateoffer", LanguageID))%></option>
<%End If %>
<% End If%>
<%If (Logix.UserRoles.AssignPreValidate) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
<option value="2"><%Sendb(Copient.PhraseLib.Lookup("term.prevalidate", LanguageID))%></option>
<%End If%>
<%If (Logix.UserRoles.AssignPostValidate) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
<option value="3"><%Sendb(Copient.PhraseLib.Lookup("term.postvalidate", LanguageID))%></option>
<%End If%>
<%If (Logix.UserRoles.AssignReadyToDeploy) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
<option value="4"><%Sendb(Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID))%></option>
<%End If%>
<%If (Logix.UserRoles.SendOffersToCRM) Then%>
<option value="5"><%Sendb(Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID))%></option>
<%End If%>
<option value="8"><%Sendb(Copient.PhraseLib.Lookup("folders.transferoffers", LanguageID))%></option>
<option value="6"><%Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%></option>
<%If (Logix.UserRoles.DeleteOffersFromFolder) Then%>
<option value="12"><%Sendb(Copient.PhraseLib.Lookup("perm.offers-delete", LanguageID))%></option>
<%End If%>
</select>
<input type="button" name="ExecAction" id="ExecAction" value="Execute" onclick="javascript:ExecAction();" style="visibility:hidden;"/>
<div id="dupoffer"" style="visibility:hidden;">
<label for="lblselectfolder"><%Sendb(Copient.PhraseLib.Lookup("folders.PleaseSelect", LanguageID))%>:</label>&nbsp;&nbsp;
<input type="hidden" id="folderList" name="folderList" value="" />&nbsp;
<input type="button" name="folderbrowse" id="folderbrowse" value="<%Sendb(Copient.PhraseLib.Lookup("term.browse", LanguageID))%>" onclick="javascript:openPopup('folder-browse.aspx');" />
<br />
</div>
<div id="execduplicateoffer" style="visibility:hidden;">
<table summary="">
<tr>
<td valign="top" id="folderNames"><%Sendb(Copient.PhraseLib.Lookup("term.noneselected", LanguageID))%></td>
</tr>
<tr>
<td>
<input type="button" name="btnDupOffer" id="btnDupOffer" value="Execute" onclick="javascript:DuplicateOfferstofolder();" />
</td>
</tr>
</table>
</div>
</div>
</div>


<div id="NavigatetoReports" class="folderdialog"> 
<div class="foldertitlebar">
<span class="dialogtitle"><% Sendb(Copient.PhraseLib.Lookup("term.customreports", LanguageID) & " " & Copient.PhraseLib.Lookup("term.page", LanguageID) & " " & Copient.PhraseLib.Lookup("message.loading", LanguageID))%></span>
<span class="dialogclose" onclick="toggleDialog('NavigatetoReports', false);">X</span>
</div>
<div class="dialogcontents">
<br class="half"/>
<br />
 <div class="loading">
    <img id="loader" alt="Loading" src="..\images\loadingAnimation.gif" />
</div>
<input type="button" class="cancel" name="btnCancel" id="btnCancelNvaigation" value="<%Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>" onclick="javascript:CancelNavtoReports();" />
<br />
</div>
</div>




<div id="datepicker" class= "dpDiv">
</div>
<%
If Request.Browser.Type = "IE6" Then
  Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
End If
%>
<div id="folderpopulate" class="folderdialog">
<div class="foldertitlebar">
  <span class="dialogtitle"><%Sendb(Copient.PhraseLib.Lookup("hierarchy.additems", LanguageID))%></span>
  <span class="dialogclose" onclick="toggleDialog('folderpopulate', false);">X</span>
</div>
<div class="dialogcontents">
  <br class="half"/>
  <div id="AddItemWarninng">
  </div>
  <label for="searchterms"><%Sendb(Copient.PhraseLib.Lookup("folders.SearchTerm", LanguageID))%>:</label><br />
  <input type="text" id="searchterms" name="searchterms" value="" maxlength="200" onkeydown="javascript:handleSearchKeyDown(event);" class="long" />
  <input type="button" id="searchsubmit" name="searchsubmit" value="<%Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" onclick="javascript:submitSearch();" /><br />
  <br />
  <input type="button" class="regular" id="btnAddItems" name="btnAddItems" style="visibility:hidden;" onclick="javascript:addToFolder();" value="<%Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID))%>" /><br />
  <br />
  <div id="searchResults"></div> 
</div>
</div>
<div id="foldersearch" class="folderdialog">
<div class="foldertitlebar">
  <span class="dialogtitle"><%Sendb(Copient.PhraseLib.Lookup("folders.SearchFoldersOffers", LanguageID))%></span>
  <span class="dialogclose" onclick="toggleDialog('foldersearch', false);">X</span>
</div>
<div class="dialogcontents">
  <br class="half"/>
  <label for="fsearchterms"><%Sendb(Copient.PhraseLib.Lookup("folders.SearchTerm", LanguageID))%>:</label><br />
  <input type="text" id="fsearchterms" name="fsearchterms" value="" maxlength="200" onkeydown="javascript:handlefSearchKeyDown(event);" class="long" />
  <select name="fsearchType" id="fsearchType">
    <option value="1" selected="selected"><%Sendb(Copient.PhraseLib.Lookup("term.folders", LanguageID))%></option>
    <option value="2"><%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%></option>
  </select>
  <input type="button" id="fsearchsubmit" name="fsearchsubmit" value="<%Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" onclick="javascript:submitfSearch();" /><br />
  <br />
  <div id="fsearchResults"></div> 
</div>
</div>

<div id="OfferfadeDiv"></div>
<div id="DuplicateNoofOffer" class="folderdialog" style="position:relative; top: 100px; WIDTH: 400px; HEIGHT: 150px">
  <div class="foldertitlebar">
    <span class="dialogtitle"><% Sendb(Copient.PhraseLib.Lookup("folders.copyofferstofolder", LanguageID)) %></span> 
	<span class="dialogclose" onclick="toggleDialogOfferDuplicate('DuplicateNoofOffer', false);">X</span>
  </div>
  <div class="dialogcontents">
    <div id="DuplicateOffererror" style="display: none; color: red;">
    </div>
    <table style="width:90%">
		<tr><td>&nbsp;</td></tr>
      <tr>
        <td>
		  <label for="lbldupOffers"><% Sendb(Copient.PhraseLib.Lookup("term.duplicateOfferstoCreate", LanguageID).Replace("99", MyCommon.NZ(MyCommon.Fetch_SystemOption(184), 0).ToString()))%></label>
		  <input type="text" style="width:20px" id="txtDuplicateOffersCnt" name="txtDuplicateOffersCnt" maxlength="2" value="" /> 
        </td>
      </tr>
		<tr><td>&nbsp;</td></tr>
	  <tr align="right">
        <td>
          <input type="button" name="btnOk" id="btnOk" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>" onclick="addDuplicateOfferscount();" />
		  <input type="button" name="btnCancel" id="btnCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>" onclick="toggleDialogOfferDuplicate('DuplicateNoofOffer', false);" />
                  </td>
       </tr>	  
    </table>
  </div>
</div>
  
<%  
  Send_PageEnd()
  If (Request.QueryString("new") <> "") Then
    Send("<script type=""text/javascript"">")
    Send("  toggleDialog('foldercreate', true);")
    Send("</script>")
  End If
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
