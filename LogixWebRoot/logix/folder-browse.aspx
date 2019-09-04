<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%
  ' *****************************************************************************
  ' * FILENAME: folder-browse.aspx
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
  Dim dt As DataTable
  Dim row As DataRow
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim OfferID As Long = 0
  Dim FindFolderID As Integer
  Dim FolderList As String = ""
  Dim rst As System.Data.DataTable
  Dim index As Integer
  Dim UpdateLevel As Long = 0
  Dim buyerid As String = ""
  Dim id As Integer = 0
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String=String.Empty
  Dim fromCopyAction As Integer = 0
  Dim EngineID As String = -1
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  buyerid = Request.QueryString("buyerid")

  Response.Expires = 0
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  MyCommon.AppName = "folder-browse.aspx"
  CMS.AMS.CurrentRequest.Resolver.AppName = "folder-browse.aspx"

    OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    fromCopyAction = MyCommon.Extract_Val(Request.QueryString("fromCopyAction"))
    FindFolderID = MyCommon.Extract_Val(Request.QueryString("FolderID"))
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    
  If Not String.IsNullOrEmpty(buyerid) Then
    id = GetBuyerID(buyerid)
    MyCommon.QueryStr = "select distinct FolderID from FolderItems as FI (NoLock)" & _
                          "where LinkID=" & id & " and LinkTypeID=2;"
    dt = MyCommon.LRT_Select
    For Each row In dt.Rows
      If FolderList <> String.Empty Then FolderList &= ","
      FolderList &= MyCommon.NZ(row.Item("FolderID"), "")
    Next
  End If

  If OfferID > 0 Then
    MyCommon.QueryStr = "select distinct FolderID from FolderItems as FI with (NoLock) " & _
                        "where LinkID=" & OfferID & " and LinkTypeID=1;"
    dt = MyCommon.LRT_Select
    For Each row In dt.Rows
      If FolderList <> String.Empty Then FolderList &= ","
      FolderList &= MyCommon.NZ(row.Item("FolderID"), "")
    Next

    Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
    If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
        StatusText = Copient.PhraseLib.Lookup("term.active", LanguageID)
    ElseIf (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED) Then
            StatusText = Copient.PhraseLib.Lookup("term.expired", LanguageID)
        ElseIf (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_SCHEDULED) Then
            StatusText = Copient.PhraseLib.Lookup("term.scheduled", LanguageID)
    End If

    'find if the offer has been deployed
    MyCommon.QueryStr = "select updatelevel from cpe_incentives with (nolock) where incentiveid=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      UpdateLevel = MyCommon.NZ(rst.Rows(0).Item("UpdateLevel"), 0)
    End If
  End If

  Send_HeadBegin("term.offer", "term.folder")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
    Send_Scripts(New String() {"datePicker.js", "popup.js"})
    ' Check User Permission for Folder access.
    If (Logix.UserRoles.AccessFolders = False) Then
        Send_Denied(1, "perm.folders-access")
        GoTo done
    End If
%>
<style type="text/css">
  body
  {
    overflow: visible;
  }
  #foldercreate
  {
    left: 0px;
    top: 100px;
  }
  #foldersearch
  {
    top: 50px;
  }
  #folderstatusbar
  {
    left: 16px;
    top: 478px;
    width: 650px;
  }
  * html #folderstatusbar
  {
    left: 16px;
    top: 474px;
    width: 650px;
  }
  #folderinfostatusbar
  {
    left: 16px;
    top: 498px;
    width: 650px;
  }
  * html #folderinfostatusbar
  {
    left: 16px;
    top: 494px;
    width: 650px;
  }
  #searchResults
  {
    width: 95%;
  }
</style>
<script type="text/javascript" src="../javascript/jquery.min.js"></script>
<script type="text/javascript">
  // constants
  var CREATE_FOLDER = 1;
  var MODIFY_FOLDER = 2;
  var DELETE_FOLDER = 3;
  var VIEW_ITEMS = 4;
  var VIEW_AVAILABLE = 5;
  var REMOVE_ITEMS = 6;
  var ADD_ITEMS = 7;
  var ASSIGN_FOLDERS = 8;
  var FOLDER_SEARCH = 9;
  var FOLDER_INFO = 10;
  var FOLDER_FUTURE_DATE = 11;
  var GENERATE_DIV_FOLDER = 12;
  var CHECK_DEFAULTFOLDER=24;
  var COPY_EXPIREDOFFER=25;

  var selectedFolder = 0;
  var checkedFolders = new Array();
  var lastCreatedFolderName = '';
  var folderstartdate = '';
  var folderenddate = '';
  var foldertheme = '';
  var datePickerDivID = "datepicker";
  var offerId = <%= OfferID %>

<%  Send_Calendar_Overrides(MyCommon)%>
  function toggleDropdown() {
    if (document.getElementById("actionsmenu") != null) {
      bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible');
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
    if (window.XMLHttpRequest) { // Mozilla/Safari
      self.xmlHttpReq = new XMLHttpRequest();
    } else if (window.ActiveXObject) { // IE
      self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
        switch (action) {
        case COPY_EXPIREDOFFER:
            assignFolderUpdateForCopy(self.xmlHttpReq.responseText);
            break;
            case CREATE_FOLDER:
            handleCreateFolder(self.xmlHttpReq.responseText);
            break;
          case DELETE_FOLDER:
            if(confirmSuccess(self.xmlHttpReq.responseText)){
                saveFolders();
                if(opener.document.getElementById("reloadpage") != null)
                {
                    opener.document.getElementById("reloadpage").innerHTML = "1";
                }
            }
            break;
          case ASSIGN_FOLDERS:
            assignFolderUpdate(self.xmlHttpReq.responseText);
            ChangeParentDocument();
            break;
            case FOLDER_SEARCH:
            showSearchContents(self.xmlHttpReq.responseText);
            break;
             case GENERATE_DIV_FOLDER:
            generatedivfmodify(self.xmlHttpReq.responseText);
            break;
          case FOLDER_INFO:
            UpdateFolderInfo(self.xmlHttpReq.responseText);
            break;
          case FOLDER_FUTURE_DATE:
            HandleFolderFutureDate(self.xmlHttpReq.responseText);
            break;
             case MODIFY_FOLDER:
            cofirmsuccessmodfolder(self.xmlHttpReq.responseText);
            break;
          case CHECK_DEFAULTFOLDER:
           SetDefaultFolder(self.xmlHttpReq.responseText);
           break;

        }
      }
    }
    self.xmlHttpReq.send(frmdata);
  }

  function clearContents(elemName)
  {
      switch (elemName) {        
          case 'foldersearch':            
              var searchElem = document.getElementById("searchterms");
              var resultsElem = document.getElementById("searchResults"); 
              var searchType = document.getElementById("searchType");
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
              var elem=document.getElementById("foldercreate");
              if(elem!=null)
              {
                  var Fname=document.getElementById("newFolderName");
                  if(Fname!=null)
                  {
                      if(Fname.value!="")
                      {
                          Fname.value="";
                      }
                  }
                  var Fstart=document.getElementById("folderstart");
                  if(Fstart!=null)
                  {
                      if(Fstart.value!="")
                      {
                          Fstart.value="";
                      }
                  }
                  var Fend=document.getElementById("folderend");
                  if(Fend!=null)
                  {
                      if(Fend.value!="")
                      {
                          Fend.value="";
                      }
                  }
              }
              break;  
      }    
  }
  function assignFolderUpdate(resp) {
    var fadeElem = document.getElementById('fadeDiv');

    if (confirmSuccess(resp)) {
      updateStatusBar('<%Sendb(Copient.PhraseLib.Lookup("uom-set.saved", LanguageID))%>');
      ChangeParentDocument();
    }

    if (fadeElem != null) {
      fadeElem.style.display = 'none';
      document.body.style.cursor = 'default';
    }
  }

  function assignFolderUpdateForCopy(resp) {
  var fadeElem = document.getElementById('fadeDiv');
   var folderList = resp.split('|');
    if (folderList[0] == 'OK')  {
      var offerId=folderList[1];
      if(offerId > 0)
      {
       if (confirmSuccess(resp)) {
      updateStatusBar('<%Sendb(Copient.PhraseLib.Lookup("uom-set.saved", LanguageID))%>');
      RedirectParentDocument(offerId);
       }
        
        }
       else
        {
            alert("Could not copy offer");
            return false;
        }
     }
     else if(folderList[0] == 'NO'){
         alert(folderList[1]);
              if (fadeElem != null) {
           fadeElem.style.display = 'none';
           document.body.style.cursor = 'default';
         }
         return false;
     }
     if (fadeElem != null) {
      fadeElem.style.display = 'none';
      document.body.style.cursor = 'default';
    }
  }

  function RedirectParentDocument(offerid) {
 
      /*
     * queryParameters -> handles the query string parameters
     * queryString -> the query string without the fist '?' character
     * re -> the regular expression
     * m -> holds the string matching the regular expression
     */
    var queryParameters = {}, queryString = opener.location.search.substring(1),
        re = /([^&=]+)=([^&]*)/g, m;

    // Creates a map with the query string parameters
    while (m = re.exec(queryString)) {
        queryParameters[decodeURIComponent(m[1])] = decodeURIComponent(m[2]);
    }

    // Add new parameters or update existing ones
    queryParameters['OfferID'] = offerid;

    /*
     * Replace the query portion of the URL.
     * jQuery.param() -> create a serialized representation of an array or
     *     object, suitable for use in a URL query string or Ajax request.
     */
    opener.location.search = $.param(queryParameters); // Causes page to reload
    window.close();
         
  }

 
  function HandleFolderFutureDate(resp) {
    var folderList = trimString(resp.substring(3));
    if (resp.substring(0, 2) != 'OK' && resp.substring(0, 2) != 'NO') {
     if( opener.document.getElementById("execduplicateoffer") != null)
        opener.document.getElementById("execduplicateoffer").style.visibility = 'hidden';
     alert(resp);
     }
    else if (resp.substring(0, 2) != 'NO') {
       if (opener != null && opener.document.getElementById('folderList') != null) {
        opener.document.getElementById('folderList').value = folderList;
      if( opener.document.getElementById("execduplicateoffer") != null)
        opener.document.getElementById("execduplicateoffer").style.visibility = 'visible';
        writeFolderNames(folderList);
      }
      updateStatusBar('<%Sendb(Copient.PhraseLib.Lookup("uom-set.saved", LanguageID))%>');
    }
   }

  function confirmSuccess(responseText) {
    var folderid=0;
    var response = trimString('<%Sendb(Copient.PhraseLib.Lookup("folders.CannotDelete", LanguageID))%>');

    if (responseText.substring(0, 2) != 'OK' && trimString(responseText) != response) {
      alert(responseText);
      return false;
    }
    else if (trimString(responseText) == response){
      alert(responseText);
      document.location = 'folder-browse.aspx';
      return false;
      }
    else if (responseText.substring(0, 2) != 'NO')  {
    return true;
    }
  }

    function cofirmsuccessmodfolder(responseText) {
      var errorstring = '<%Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))%>' + " > " + '<%Sendb(Copient.PhraseLib.Lookup("term.error", LanguageID))%>';
        
    var modifyfolderElem = document.getElementById("modifyfoldererror");
    var location = "";
    var div = document.createElement('div'); 
    div.innerHTML = responseText;       
    if(responseText.toUpperCase().search("</HTML>") != -1 && $(div).find('title').text() == errorstring){
        document.getElementById("modifyfoldererror").innerHTML =$(responseText).find('#main').text();
    }
    else {
        if (responseText.substring(0, 2) == 'NO') {
            modifyfolderElem.style.display = 'block';
            modifyfolderElem.innerHTML = responseText.substring(3,responseText.length);
            toggleDialog('modifyfolder', true);
        }
        else if (responseText.substring(0, 2) != 'NO') {
            modifyfolderElem.style.display = 'none';
            toggleDialog('modifyfolder', false);

            if(offerId)
                location = 'folder-browse.aspx?OfferID='+offerId;
            else
                location = 'folder-browse.aspx';
            document.location = location;
        }
        folderid = parseInt(responseText.substring(3));
        if(folderid > 0){
            highlightSelected(null,selectedFolder);
            var folderChk = document.getElementById('chk' + selectedFolder);
            folderChk.setAttribute("checked",true);
        }
    }

  }
  
  var modfoldername ="";
  var folderstartdate="";
  var folderenddate="";
  var foldertheme="";

   function savemodifiedFolder() {
     var modfoldialogElem = document.getElementById('modifyfolder');
     var modfolderstartdateele=document.getElementById('modifyfolderstart');
     var modfolderenddateele=document.getElementById('modifyfolderend');
     var modfolderthemeele=document.getElementById('ModifyTheme');
     var modfoldialogName= document.getElementById('editFolderName');
     var moderror = document.getElementById('modifyfoldererror');
     var ismassupdateenable= <% Sendb(MyCommon.NZ(MyCommon.Fetch_SystemOption(133),0)) %>;
     var saveButton = document.getElementById('btnModFolder');
     var divMessg = document.getElementById('divMessage');
     var defaultUEFolderele=document.getElementById('defaultUEFolder-Mod');
     var isdefaultUEFolder=false;
      if (modfoldialogElem != null)
      {
        if(modfoldialogName != null)
        {
            modfoldername = trimString(modfoldialogName.value);
            if(modfoldername == ""){
                 moderror.innerHTML = "Please enter folder name";
                  toggleDialog('modifyfolder', true);
                 return;
             }
               if(modfoldername.length >99)
             {
                moderror.innerHTML ="Folder name should be less than 100 characters";
                toggleDialog('modifyfolder', true);
                return;
             }
        }
         if (modfolderstartdateele != null) {
          folderstartdate=trimString(modfolderstartdateele.value);
         }
         if (modfolderenddateele != null) {
          folderenddate=trimString(modfolderenddateele.value);
         }
         if (modfolderthemeele != null) {
          foldertheme=modfolderthemeele.options[modfolderthemeele.selectedIndex].text;
         }
         if(defaultUEFolderele!=null){
            isdefaultUEFolder=defaultUEFolderele.checked;
         }

          if(folderstartdate != "" && folderenddate != "")
         {
                var fstartdate=Date.parse(ConvertToISODate(folderstartdate, '<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>'));
                var fenddate=Date.parse(ConvertToISODate(folderenddate, '<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>'));
                if(isNaN(fstartdate) || isNaN(fenddate))
                {
                    moderror.innerHTML= "folder dates should be in correct date format";
                      toggleDialog('modifyfolder', true);
                    return;
                }
                if(fstartdate > fenddate){
                     moderror.innerHTML = "folder start date should be less than folder end date";
                       toggleDialog('modifyfolder', true);
                     return;
                 }

            }
            else
            {
              xmlhttpPost('/logix/folder-feeds.aspx?Action=ModifyFolder', MODIFY_FOLDER, 'FolderID=' + selectedFolder + '&ModFolderName=' + encodeURIComponent(modfoldername) + '&ModFolderStartDate=' + folderstartdate + '&ModFolderEndDate=' + folderenddate + '&ModFolderTheme=' + foldertheme +'&isMassupdateEnabled=false'+'&IsdefaultUEFolder='+isdefaultUEFolder);
                 return;
          }
         if (ismassupdateenable)
         {
           saveButton.style.visibility = "hidden";
            divMessg.style.display= "block";
        }
         else
             xmlhttpPost('/logix/folder-feeds.aspx?Action=ModifyFolder', MODIFY_FOLDER, 'FolderID=' + selectedFolder + '&ModFolderName=' + encodeURIComponent(modfoldername) + '&ModFolderStartDate=' + folderstartdate + '&ModFolderEndDate=' + folderenddate + '&ModFolderTheme=' + foldertheme +'&isMassupdateEnabled=false'+'&IsdefaultUEFolder='+isdefaultUEFolder);
          }
   }
  function OKClicked()
   {
        var saveButton = document.getElementById('btnModFolder');
        var divMessg = document.getElementById('divMessage');
        xmlhttpPost('/logix/folder-feeds.aspx?Action=ModifyFolder', MODIFY_FOLDER, 'FolderID=' + selectedFolder + '&ModFolderName=' + encodeURIComponent(modfoldername) + '&ModFolderStartDate=' + folderstartdate + '&ModFolderEndDate=' + folderenddate + '&ModFolderTheme=' + foldertheme +'&isMassupdateEnabled=true');
        saveButton.setAttribute("display","block");
        divMessg.setAttribute("display","none");
   }
   function CancelClicked()
   {
        var saveButton = document.getElementById('btnModFolder');
        var divMessg = document.getElementById('divMessage');
        xmlhttpPost('/logix/folder-feeds.aspx?Action=ModifyFolder', MODIFY_FOLDER, 'FolderID=' + selectedFolder + '&ModFolderName=' + encodeURIComponent(modfoldername) + '&ModFolderStartDate=' + folderstartdate + '&ModFolderEndDate=' + folderenddate + '&ModFolderTheme=' + foldertheme +'&isMassupdateEnabled=false');
        saveButton.setAttribute("display","block");
        divMessg.setAttribute("display","none");
   }
  function modifyFolder() {
    var folderElem = document.getElementById('folder' + selectedFolder);

    if (selectedFolder > 0) {
      if (folderElem != null) {
      xmlhttpPost('/logix/folder-feeds.aspx?Action=GenDivModFolder', GENERATE_DIV_FOLDER, 'FolderID=' + selectedFolder);
      }
   } else {
//      showContents('<p>Please select a folder.<\/p>');
      }
  }
  
  function generatedivfmodify(divHTML) {
  var elem = document.getElementById("modifyfolder");
   if (elem != null) {
     elem.innerHTML = divHTML;
     toggleDialog('modifyfolder', true);
   }
  }
  
  function showSearchContents(content) {
    var resultsElem = document.getElementById("searchResults");
    if (resultsElem != null) {
      resultsElem.innerHTML = content;
    }
  }

  function UpdateFolderInfo(content) {
    var statusBarElem = document.getElementById('folderinfostatusbar');
    if (trimString(content) != '') {
      if (statusBarElem != null) {
      statusBarElem.style.display = 'block';
      statusBarElem.innerHTML = content;
      }
    }
  }

  function showfolderdateerror(content){
    var createfolderElem = document.getElementById("createfoldererror");

      createfolderElem.style.display = '';
      createfolderElem.innerHTML = content;
    }

  function handleCreateFolder(responseText) {
      var errorstring = '<%Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))%>' + " > " + '<%Sendb(Copient.PhraseLib.Lookup("term.error", LanguageID))%>';
      var div = document.createElement('div'); 
      div.innerHTML = responseText;       
      if(responseText.toUpperCase().search("</HTML>") != -1 && $(div).find('title').text() == errorstring){
          document.getElementById("createfoldererror").innerHTML =$(responseText).find('#main').text();
          document.getElementById("createfoldererror").style.display = 'block';
      }
      else {
          document.getElementById("createfoldererror").innerHTML = "";
          document.getElementById("createfoldererror").style.display = 'none';
          var folderid = 0;
          var folderDiv = null;
          var folderRow = null;
          var folderImg = null;
          var folderNbsp = null;
          var folderName = null;
          var folderChk = null;
          var className = '';
          var parentFolder = null;
          var descendants = null;
          var sibling = null;

          if (responseText.substring(0, 2) == 'OK') {
              folderid = parseInt(responseText.substring(3));

              if (selectedFolder == 0) {
                  parentFolder = document.getElementById('foldertreebrowse');
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
              folderRow.onclick = function() {
                  highlightSelected(null, folderid);
              }
              folderDiv.appendChild(folderRow);

              //FolderImg image
              folderImg = document.createElement('img');
              folderImg.setAttribute('id', 'expander' + folderid);
              folderImg.src = '../images/blank-clear.png';
              folderImg.setAttribute(className, 'expander');
              folderImg.onclick = function() {
                  toggleFolder(folderid);
              }
              folderRow.appendChild(folderImg);

              //FolderChk checkbox
              folderChk = document.createElement('input');
              folderChk.setAttribute('id', 'chk' + folderid);
              folderChk.setAttribute('type', 'checkbox');
              folderChk.setAttribute('style', 'position:relative');
              folderChk.setAttribute('value', folderid);
              folderChk.onclick = function() {
                  handleFolderClick(folderid, this.checked);
              }
              folderRow.appendChild(folderChk);

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
  }

  function navigateToFolder(FolderID) {
    var folderDiv = document.getElementById('folder' + FolderID);
    var prt = null;
    var fid = 0;

    if (folderDiv != null) {
      prt = folderDiv;
      while (prt != null) {
        if (prt.parentNode.id.length >= 6 && prt.parentNode.id != 'foldertreebrowse' && prt.parentNode.id.substring(0,6) == 'folder') {
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
    var folderDiv = document.getElementById('folder' + FolderID);
    var folderRow = document.getElementById('folderrow' + FolderID);
    var selectedDiv = document.getElementById('folder' + selectedFolder);
    var selectedRow = document.getElementById('folderrow' + selectedFolder);
    var source = null;

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
          var content = ShowFolderInfo(FolderID);
        } else {
          folderRow.style.backgroundColor = '#ffffff';
          folderRow.style.border = '1px solid #ffffff';
        }
      }
    }
  }

  function ShowFolderInfo(folderID) {
   var urlStr='/logix/folder-feeds.aspx?Action=ShowFolderInfo';
    return xmlhttpPost(urlStr, FOLDER_INFO, 'FolderID=' + selectedFolder);
  }

  function ensureFolderSelected(FolderID) {
    if (FolderID != selectedFolder) {
      highlightSelected(null, FolderID)
    }
  }

  function showSearchContents(content) {
      var errorstring = '<%Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))%>' + " > " + '<%Sendb(Copient.PhraseLib.Lookup("term.error", LanguageID))%>';
      var div = document.createElement('div'); 
      div.innerHTML = content;       
      if(content.toUpperCase().search("</HTML>") != -1 && $(div).find('title').text() == errorstring){
          content = $(content).find('#main').text().fontcolor("red");
      }
    var resultsElem = document.getElementById("searchResults");
    if (resultsElem != null) {
      resultsElem.innerHTML = content;
    }
    document.body.style.cursor = 'default';
  }

  function toggleDialog(elemName, shown) {
    var elem = document.getElementById(elemName);
    var fadeElem = document.getElementById('fadeDiv');

    if (elem != null) {
      elem.style.display = (shown) ? 'block' : 'none';
    }
    if (fadeElem != null) {
      fadeElem.style.display = (shown) ? 'block' : 'none';
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

  function createFolder() {
    var dialogElem = document.getElementById('foldercreate')
    var newNameElem = document.getElementById('newFolderName');
    var newAccessLevelElem = document.getElementById('accessLevel');
     //
    var folderstartdateele=document.getElementById('folderstart');
    var folderenddateele=document.getElementById('folderend');
    var folderthemeele=document.getElementById('Theme');
    var defaultUEFolderele=document.getElementById('defaultUEFolder');
    //
    var folderElem = document.getElementById('folder' + selectedFolder);
    var accessLevel = 3;
    var isdefaultUEFolder=false;

    if (dialogElem != null && newNameElem != null) {
      dialogElem.style.display = 'block';

      if (trimString(newNameElem.value) == "") {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.EnterFolderName", LanguageID))%>');
      } else {
        if (newNameElem.value.length > 50) {
          alert('    <%Sendb(Copient.PhraseLib.Lookup("folders.NameTooLong", LanguageID))%>');
        } else {
          lastCreatedFolderName = newNameElem.value;
          if (lastCreatedFolderName != null) {
            lastCreatedFolderName = trimString(lastCreatedFolderName);
            if (newAccessLevelElem != null) {
              accessLevel = newAccessLevelElem.value;
            }
            if (folderstartdateele != null) {
              folderstartdate = trimString(folderstartdateele.value);
            }
            if (folderenddateele != null) {
              folderenddate = trimString(folderenddateele.value);
            }
            if (folderthemeele != null) {
              foldertheme = folderthemeele.options[folderthemeele.selectedIndex].text;
            }
            if(defaultUEFolderele!=null){
              isdefaultUEFolder=defaultUEFolderele.checked;
            }
            xmlhttpPost('/logix/folder-feeds.aspx?Action=CreateFolder', CREATE_FOLDER, 'FolderID=' + selectedFolder + '&FolderName=' + encodeURIComponent(lastCreatedFolderName) + '&AccessLevel=' + accessLevel + '&FolderStartDate=' + folderstartdate + '&FolderEndDate=' + folderenddate + '&FolderTheme=' + foldertheme+'&IsdefaultUEFolder='+isdefaultUEFolder);
            lastCreatedFolderName = lastCreatedFolderName.replace("<", "&lt;");
          }
        }
      }
    }
  }

//  function renameFolder() {
//    var folderElem = document.getElementById('folder' + selectedFolder);
//    var folderNameElem = document.getElementById('foldername' + selectedFolder);
//
//    if (folderElem != null && folderNameElem != null) {
//      var rsp = prompt('<%Sendb(Copient.PhraseLib.Lookup("folders.EnterNewName", LanguageID))%>', '');
//      if (rsp != null) {
//        rsp = trimString(rsp);
//        if (rsp != '') {
//          if (rsp.length > 50) {
//            alert('  <%Sendb(Copient.PhraseLib.Lookup("folders.NameTooLong", LanguageID))%>');
//          } else {
//            xmlhttpPost('/logix/folder-feeds.aspx?Action=RenameFolder', RENAME_FOLDER, 'FolderID=' + selectedFolder + '&FolderName=' + rsp);
//            folderNameElem.innerHTML = rsp;
//          }
//        }
//      }
//    }
//  }

  function deleteFolder() {
    var folderElem = document.getElementById('folder' + selectedFolder);

    if (folderElem != null) {
      if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.DeleteConfirm", LanguageID))%>')) {
        xmlhttpPost('/logix/folder-feeds.aspx?Action=DeleteFolder', DELETE_FOLDER, 'FolderID=' + selectedFolder);
        folderElem.parentNode.removeChild(folderElem);
        selectedFolder = 0;
      }
    }
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
  function handleFolderClick(folderid, bChecked) {
    var index = -1;
    var fID = 0;

    fID = parseInt(folderid);
    if (!bChecked) {
      index = search(checkedFolders, fID, false);
      if (index > -1) {
       checkedFolders.splice(index, 1);
      }
    } else {
      index = search(checkedFolders, fID, true);
      checkedFolders.splice(index, 0, fID);
    }

    updateStatusBar('<%Sendb(Copient.PhraseLib.Lookup("term.UnsavedChanges", LanguageID))%>');
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
    var typeElem = document.getElementById('searchType');
    var searchTerms = '';
    var searchType = '';
    var action = '';

    if (searchElem != null && typeElem != null) {
      searchTerms = searchElem.value;
      searchType = typeElem.value;
      frmdata = '&searchterms=' + encodeURIComponent(searchTerms);
      document.body.style.cursor = 'wait';

      var elemfolderbox = document.getElementById("foldersearch");
      var elemsresults = document.getElementById("searchResults");
      elemsresults.style.height = elemfolderbox.offsetHeight - elemsresults.offsetTop - 5 + "px";

      action = (searchType == 2) ? 'SendOfferSearch' : 'SendFolderSearch';
      xmlhttpPost('folder-feeds.aspx?Action=' + action, FOLDER_SEARCH, frmdata);
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
      isOpen = (expanderElem.src.indexOf('minus.png') > -1);
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
    var className = '';

    if (elem != null && expanderElem != null) {
      expanderElem.src = '../images/minus.png';
      className = (elem.getAttribute('className') != null) ? 'className' : 'class';
      elem.setAttribute(className, "folder");
    }
  }

  function ChangeParentDocument() {
    if (opener && !opener.closed) {
      opener.location.reload();
    }
  }
<% If (OfferID > 0) Then  %>
  function saveFolders() {
    var folderList = '';
    var fadeElem = document.getElementById('fadeDiv');

    if (fadeElem != null) {
      fadeElem.style.display = 'block';
      document.body.style.cursor = 'wait';
    }

    folderList = checkedFolders.join(",")
    <% If MyCommon.Fetch_SystemOption(132) = "1" Then %>
        if (folderList == '') {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folder-selectfolder", LanguageID))%>');
            if (fadeElem != null) {
            fadeElem.style.display = 'none';
            document.body.style.cursor = 'default';
            }
        return false;
        }
    <%End If %>
    <% If MyCommon.Fetch_SystemOption(191) = "1" Then %>
        if (folderList.indexOf(",") > -1) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folder.offerinmultiplefolder", LanguageID))%>');
            if (fadeElem != null) {
            fadeElem.style.display = 'none';
            document.body.style.cursor = 'default';
            }
        return false;
        }
    <%End If %>
  
 
// CLOUDSOL-2159 
 var forCopy ='<%Sendb(fromCopyAction)%>';
  <% If (MyCommon.Fetch_SystemOption(133) = "0" Or MyCommon.Fetch_SystemOption(272) = "1") Then%>
  var statusText = '<%Sendb(StatusText)%>';
    if(statusText !="" && statusText.length >0)
    {
       if (forCopy != 1){
         alert("Cannot change folder. The offer has been " + statusText );
         if (fadeElem != null) {
           fadeElem.style.display = 'none';
          document.body.style.cursor = 'default';
          }
         return false;
       }
     }
   <%End If %>
       <% If (fromCopyAction = 1 ) Then %>
            xmlhttpPost('/logix/folder-feeds.aspx?Action=CopyExpiredOffer', COPY_EXPIREDOFFER, 'OfferID=<%Sendb(OfferID)%>&FolderList=' + folderList);
       <%Else %>
        xmlhttpPost('/logix/folder-feeds.aspx?Action=SaveOfferFolders', ASSIGN_FOLDERS, 'OfferID=<%Sendb(OfferID)%>&FolderList=' + folderList);
        ChangeParentDocument();
       <%End If %>
  }
<% Else if (id > 0) Then  %>
    function saveFolders(){
    //buyerpart
    var folderList = '';
    var elemFolders = null;
    var elemTree = document.getElementById('foldertreebrowse');
    var fadeElem = document.getElementById('fadeDiv');
     if (fadeElem != null) {
      fadeElem.style.display = 'block';
      document.body.style.cursor = 'wait';
    }

    //else part
    if (elemTree != null) {
      elemFolders = elemTree.getElementsByTagName('input');
      for (var i=0; i < elemFolders.length; i++) {
        if (elemFolders[i].type == "checkbox" && elemFolders[i].checked) {
          if (folderList != '') {
            folderList += ',';
          }
          folderList += elemFolders[i].value;
        }
      }
    }

    folderList = checkedFolders.join(",")
     if (folderList.indexOf(",") > -1) {
         alert('<%Sendb(Copient.PhraseLib.Lookup("folder.buyerinmultiplefolder", LanguageID))%>');
          if (fadeElem != null) {
            fadeElem.style.display = 'none';
            document.body.style.cursor = 'default';
            }
       return false;
       }

    <% If MyCommon.Fetch_SystemOption(132) = "1" Then %>
     if (folderList == '') {
        alert("Please select at least one folder");
            if (fadeElem != null) {
            fadeElem.style.display = 'none';
            document.body.style.cursor = 'default';
            }
        return false;
        }
    <%End If %>
    <% If MyCommon.Fetch_SystemOption(191) = "1" Then %>
        if (folderList.indexOf(",") > -1) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folder.offerinmultiplefolder", LanguageID))%>');
            if (fadeElem != null) {
            fadeElem.style.display = 'none';
            document.body.style.cursor = 'default';
            }
        return false;
        }
    <%End If %>

    xmlhttpPost('/logix/folder-feeds.aspx?Action=SaveBuyerFolders', ASSIGN_FOLDERS, 'id=<%Sendb(id)%>&FolderList=' + folderList);
}
<% Else %>
  function saveFolders() {
    var elemTree = document.getElementById('foldertreebrowse');
    var elemFolders = null;
    var folderList = '';
    //else part
    if (elemTree != null) {
      elemFolders = elemTree.getElementsByTagName('input');
      for (var i=0; i < elemFolders.length; i++) {
        if (elemFolders[i].type == "checkbox" && elemFolders[i].checked) {
          if (folderList != '') {
            folderList += ',';
          }
          folderList += elemFolders[i].value;
        }
      }
   }
<% If buyerid IsNot Nothing Then%>
      folderList = checkedFolders.join(",")
      if (folderList.indexOf(",") > -1) {
          alert('<%Sendb(Copient.PhraseLib.Lookup("folder.buyerinmultiplefolder", LanguageID))%>');
          return false;
      }
   <% End If%>
    <% If MyCommon.Fetch_SystemOption(132) = "1" Then %>
        if (folderList == '') {
        alert("Please select at least one folder");
            if (fadeElem != null) {
            fadeElem.style.display = 'none';
            document.body.style.cursor = 'default';
            }
        return false;
        }
    <%End If %>
    <% If MyCommon.Fetch_SystemOption(191) = "1" Then %>
        if (folderList.indexOf(",") > -1) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folder.offerinmultiplefolder", LanguageID))%>');
            if (fadeElem != null) {
            fadeElem.style.display = 'none';
            document.body.style.cursor = 'default';
            }
        return false;
        }
    <%End If %>
    <% If MyCommon.Fetch_SystemOption(143) = "1" Then %>
       xmlhttpPost('/logix/folder-feeds.aspx?Action=SendOfferFoldersWithFutureDate', FOLDER_FUTURE_DATE, '&FolderList=' + folderList);
    <% Else %>

       if (opener != null && opener.document.getElementById('folderList') != null) {
        opener.document.getElementById('folderList').value = folderList;
        if (opener.document.getElementById("execduplicateoffer") != null) {
          //else part
          opener.document.getElementById("execduplicateoffer").style.visibility = 'visible';
          opener.document.getElementById("execduplicateoffer").style.height = opener.document.getElementById("performactions").offsetHeight - opener.document.getElementById("execduplicateoffer").offsetTop - 5 + "px";

        }
        writeFolderNames(folderList);
      }
	  if ((window.opener.location.pathname == '/logix/offer-list.aspx')||(window.opener.location.pathname == '/logix/folders.aspx')||(window.opener.location.pathname == '/logix/Enhanced-extoffer-list.aspx')){
        if (opener.document.getElementById('btnDupOffer').disabled == true){
	      updateStatusBar('<%Sendb(Copient.PhraseLib.Lookup("term.UnsavedChanges", LanguageID))%>');
	    }
	    else {
          updateStatusBar('<%Sendb(Copient.PhraseLib.Lookup("uom-set.saved", LanguageID))%>');
	    }
	  }
      else {
          updateStatusBar('<%Sendb(Copient.PhraseLib.Lookup("uom-set.saved", LanguageID))%>');
	   }

    <% End If %>
    }
<% End If %>

  function writeFolderNames(folderList) {
    opener.document.getElementById('folderNames').innerHTML =""
    var elemName = null;
    var folderIDs;
    var folderNames = '';
    var folderPath = '';
    var selectedfolders = 0;

    if (folderList != '') {
      folderIDs = folderList.split(',');
      if (folderIDs.length > 0) {
        folderNames += '<ul>';
        for (var i=0; i < folderIDs.length; i++) {
          elemName = document.getElementById('foldername' + folderIDs[i]);
          if (elemName != null) {
            folderNames += '<li>';
            folderPath = getFolderPath(elemName);
            folderNames += folderPath;
            if (folderPath.length > 0) {
              folderNames += ' / ';
            }
            folderNames += elemName.innerHTML;
            folderNames += '<\/li>';
			selectedfolders = 1 + selectedfolders;
          }
        }
        folderNames += '<ul>';
      }
    } else {
   folderNames = '<%Sendb(Copient.PhraseLib.Lookup("term.noneselected", LanguageID))%>';

	  if ((window.opener.location.pathname == '/logix/offer-list.aspx')||(window.opener.location.pathname == '/logix/folders.aspx')||(window.opener.location.pathname == '/logix/Enhanced-extoffer-list.aspx')){
	    opener.document.getElementById('btnDupOffer').disabled = true;
      }
    }

    if ((opener != null && opener.document.getElementById('folderNames') != null) && (folderNames != "None selected")){
	  if ((window.opener.location.pathname == '/logix/offer-list.aspx')||(window.opener.location.pathname == '/logix/folders.aspx')||(window.opener.location.pathname == '/logix/Enhanced-extoffer-list.aspx')){
	   if ((opener.document.getElementById("Actionitems").value == '8') && (selectedfolders >1)){
		alert('<%Sendb(Copient.PhraseLib.Lookup("folders.selectonlyonefolder", LanguageID))%>');
	    opener.document.getElementById('btnDupOffer').disabled = true;
	   }
	   else {
		opener.document.getElementById('folderNames').innerHTML = folderNames;
		opener.document.getElementById('btnDupOffer').disabled = false;
	  }
     }
	   else {
		opener.document.getElementById('folderNames').innerHTML = folderNames;
	  }
	}
  }

  function getFolderPath(elemName) {
    var folderPath = new Array();
    var elemFolderName = null;
    var prt = null;
    var folderID = '';
    var i = 0;
    var fullPath = '';

    if (elemName != null && elemName.parentNode != null && elemName.parentNode.parentNode != null && elemName.parentNode.parentNode.parentNode != null && elemName.parentNode.parentNode.parentNode.id != 'foldertreebrowse') {
      prt = elemName.parentNode.parentNode.parentNode;
      while (prt != null) {
        if (prt.id.length >= 7 && prt.id.substring(0,6)=='folder' && prt.id != 'foldertreebrowse') {
          folderID = prt.id.substring(6);
          elemFolderName = document.getElementById('foldername' + folderID);
          if (elemFolderName != null) {
            folderPath[i] = elemFolderName.innerHTML;
            i++;
          }
          prt = prt.parentNode;
        } else {
          prt = null;
        }
      }
    }
    folderPath.reverse();

    return folderPath.join(' / ');
  }

  function markSelectedFolders(folderList) {
    var folderIDs = new Array();
    var elem = null;

    // use the folder list if supplied, otherwise check if the opener is supplying the folder list (e.g. from Offer New)
    if (folderList != null && folderList != '') {
      folderIDs = folderList.split(',');
    } else if (opener != null && opener.document.getElementById('folderList') != null && opener.document.getElementById('folderList').value != '') {
      folderIDs = opener.document.getElementById('folderList').value.split(',');
    } else {
      folderIDs = null;
    }

    if (folderIDs != null && folderIDs.length > 0) {
      for (var i=0; i < folderIDs.length; i++) {
        elem = document.getElementById('chk' + folderIDs[i]);
        if (elem != null) {
          elem.checked = true;
          checkedFolders[checkedFolders.length] = folderIDs[i];
        }
        navigateToFolder(folderIDs[i]);
      }
    }
  }

  function updateStatusBar(statusText) {
    var statusBarElem = document.getElementById('folderstatusbar');

    if (statusText ==null || statusText=='') {
      statusText = '<%Sendb(Copient.PhraseLib.Lookup("term.ready", LanguageID))%>.';
    }

    if (statusBarElem != null) {
      statusBarElem.innerHTML = statusText;
    }
  }

  function searchFolders() {
    var elem = document.getElementById('searchterms');

    toggleDialog('foldersearch', true);

    if (elem != null) {
      elem.focus();
      elem.select();
    }
  }

  function folderLinkClicked(folderID) {
    navigateToFolder(folderID);
    ensureFolderSelected(folderID);
    toggleDialog('foldersearch', false);
  }
 function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el=(typeof event!=='undefined')? event.srcElement : e.target

      if (el != null) {
        var pickerDiv = document.getElementById(datePickerDivID);
        if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
          if (el.id!="folder-start-picker" && el.id!="folder-end-picker" && el.id!="foldermodified-start-picker" && el.id!="foldermodified-end-picker") {
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

   function clearErrorContent() {
      var createfolderElem = document.getElementById("createfoldererror");
       var modifyfolderElem = document.getElementById("modifyfoldererror");
      if (createfolderElem != null) {
      createfolderElem.style.display = 'none';
      }
        if (modifyfolderElem != null) {
      modifyfolderElem.style.display = 'none';
      }
   }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(3)
%>
<div id="intro">
  <div id="foldertoolbarbrowse" class="foldertoolbar">
    <%
        Sendb("<h1 id=""title"">")
        Sendb(Copient.PhraseLib.Lookup("term.folders", LanguageID))
        Send("</h1>")
        If (Logix.UserRoles.CreateFolders OrElse Logix.UserRoles.DeleteFolders OrElse Logix.UserRoles.EditFolders) Then
            Send("<div id=""foldertools"">")
            Send("  <img src=""../images/folders/vr.png"" />")
            If (Logix.UserRoles.CreateFolders) Then
                Send("  <img src=""../images/folders/folder-create.png"" alt=""" & Copient.PhraseLib.Lookup("folders.CreateFolder", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.CreateFolder", LanguageID) & """ onclick=""javascript:clearContents('foldercreate');javascript:clearErrorContent(); javascript:toggleDialog('foldercreate', true);CheckDefaultFolder();"" />")
            End If
            If (Logix.UserRoles.DeleteFolders) Then
                Send("  <img src=""../images/folders/folder-delete.png"" alt=""" & Copient.PhraseLib.Lookup("folders.DeleteFolder", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.DeleteFolder", LanguageID) & """ onclick=""javascript:deleteFolder();"" />")
            End If
            'If (Logix.UserRoles.EditFolders) Then
            '    Send("  <img src=""../images/folders/folder-rename.png"" alt=""" & Copient.PhraseLib.Lookup("folders.RenameFolder", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.RenameFolder", LanguageID) & """ onclick=""javascript:renameFolder();"" />")
            'End If
            If (Logix.UserRoles.EditFolders) Then
                Send("  <img src=""../images/folders/foldergear.jpg"" alt=""Folder Settings"" title=""" & Copient.PhraseLib.Lookup("folders.RenameFolder", LanguageID) & """ onclick=""javascript:clearErrorContent(); modifyFolder();"" />")
            End If
            Send("  <img src=""../images/folders/vr.png"" />")
            Send("</div>")
        End If
        If (Logix.UserRoles.EditFolders OrElse Logix.UserRoles.AssignFolders) Then
            Send("<div id=""itemtools"">")
            Send("&nbsp;")
            Send("  <img src=""../images/folders/item-search.png"" alt=""" & Copient.PhraseLib.Lookup("folders.SearchForFolder", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.SearchForFolder", LanguageID) & """ onclick=""javascript:searchFolders();"" />")
            Send("  <img src=""../images/save.png"" alt=""" & Copient.PhraseLib.Lookup("folders.SaveFolderSelection", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("folders.SaveFolderSelection", LanguageID) & """ onclick=""javascript:saveFolders();"" />")
            Send("</div>")
        End If
    %>
  </div>
</div>
<div id="fadeDiv">
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="foldertreebrowse" class="foldertree">
    <%
      Send("")
      LoadFolderStats(MyCommon)
      WriteFolders(MyCommon, 0)
      Send("")
    %>
    <hr class="hidden" />
  </div>
  <div id="folderstatusbar" class="folderstatusbar">
    <%Sendb(Copient.PhraseLib.Lookup("term.ready", LanguageID))%>.
  </div>
  <div id="folderinfostatusbar" style="display: none;" class="folderinfostatusbar">
  </div>
</div>
<script runat="server">
  Dim htFolders As Hashtable = Nothing

  Public Sub LoadFolderStats(ByRef MyCommon As Copient.CommonInc)
    Dim FolderData As FolderMetaData
    Dim dt As DataTable
    Dim row As DataRow

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

  End Sub
  Private Function GetBuyerID(ByVal buyerid As String) As Integer
    Dim buyerroledataservice As IBuyerRoleData = CMS.AMS.CurrentRequest.Resolver.Resolve(Of IBuyerRoleData)()
    Dim buyer As Buyer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of Buyer)()
    If buyerid <> "0" Then
      buyer = buyerroledataservice.LookupBuyerRoleByExternalId(buyerid).Result
    Else
      buyer.ID = 0
    End If
    Return buyer.ID
  End Function

  Public Sub WriteFolders(ByRef MyCommon As Copient.CommonInc, ByVal ParentFolderID As Integer)
    Dim dt As DataTable
    Dim row As DataRow
    Dim FolderID As Integer = 0
    Dim FolderData As FolderMetaData
    Dim ExpanderImg As String = ""
    Dim FolderImg As String = ""
    Dim ExpanderClickable As Boolean = False

    MyCommon.QueryStr = "select FolderID, ParentFolderID, FolderName " & _
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
      Send("<div class=""folder closed"" id=""folder" & FolderID & """" & IIf(ParentFolderID = 0, " style=""display:block !important;""", "") & ">")
      Sendb("<div class=""folderrow"" id=""folderrow" & FolderID & """ onclick=""javascript:highlightSelected(event," & FolderID & ");"">")
      Sendb("<img src=""../images/" & ExpanderImg & """ class=""expander"" id=""expander" & FolderID & """" & IIf(ExpanderClickable, " onclick=""javascript:toggleFolder(" & FolderID & ");""", "") & " />")
      Sendb("<input type=""checkbox"" id=""chk" & FolderID & """" & " value=""" & FolderID & """ style=""position:relative;"" onclick=""javascript:handleFolderClick(" & FolderID & ", this.checked);"" />")
      Sendb("<img src=""../images/" & FolderImg & """ class=""folderimg"" id=""folderimg" & FolderID & """ />")
      Sendb("<span class=""foldername"" id=""foldername" & FolderID & """>" & MyCommon.NZ(row.Item("FolderName"), "Unnamed") & "</span>")
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

  // mark the folders that were already selected from the offer-new page.
  markSelectedFolders('<%Sendb(FolderList)%>');
  <%
    If (FindFolderID > 0) Then
      Send("ensureFolderSelected(" & FindFolderID & ");")
    End If
  %>
</script>
<%
done:
  Send_FocusScript("mainform", "searchterms")
  Send_WrapEnd()
%>
<div id="foldercreate" class="folderdialog">
  <div class="foldertitlebar">
    <span class="dialogtitle">
      <% Sendb(Copient.PhraseLib.Lookup("folders.CreateFolder", LanguageID))%></span>
    <span class="dialogclose" onclick="toggleDialog('foldercreate', false);">X</span>
  </div>
  <div class="dialogcontents">
    <div id="createfoldererror" style="display: none;color: red;">
    </div>
    <br class="half" />
    <label for="newFolderName">
      <% Sendb(Copient.PhraseLib.Lookup("folders.EnterFolderName", LanguageID))%>:</label><br />
    <input type="text" id="newFolderName" name="newFolderName" value="" class="mediumlong" />
    <br />
    <label for="folderstart">
      <% Sendb(Copient.PhraseLib.Lookup("folders.FolderDate", LanguageID))%>:</label><br />
    <input type="text" class="short" id="folderstart" name="folderstart" maxlength="10"
      value="" />
    <img src="../images/calendar.png" class="calendar" id="folder-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
      title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('folderstart', event);" />
    <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
    <input type="text" class="short" id="folderend" name="folderend" maxlength="10" value="" />
    <img src="../images/calendar.png" class="calendar" id="folder-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
      title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('folderend', event);" />
    <% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>
    <br />
    <%

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

        Send("<input type=""checkbox"" name=""defaultUEFolder"" id=""defaultUEFolder"" /><label for=""defaultUEFolder"" id =""defaultUEFolderLabel"" checked="""">" & Copient.PhraseLib.Lookup("term.defaultuefolder", LanguageID) & "</label><br /><br />")

      End If
    %>
    <input type="button" name="btnNewFolder" id="btnNewFolder" value="<%Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>"
      onclick="javascript:createFolder();" />
  </div>
</div>
<div id="modifyfolder" class="folderdialog" style="height: 225px">
</div>
<div id="datepicker" class="dpDiv">
</div>
<%
  If Request.Browser.Type = "IE6" Then
    Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
  End If
%>
<div id="foldersearch" class="folderdialog">
  <div class="foldertitlebar">
    <span class="dialogtitle">
      <%Sendb(Copient.PhraseLib.Lookup("folders.SearchFoldersOffers", LanguageID))%></span>
    <span class="dialogclose" onclick="javascript:clearContents('foldersearch');toggleDialog('foldersearch', false);">X</span>
  </div>
  <div class="dialogcontents">
    <br class="half" />
    <label for="searchterms">
      <%Sendb(Copient.PhraseLib.Lookup("folders.SearchTerm", LanguageID))%>:</label><br />
    <input type="text" id="searchterms" name="searchterms" value="" onkeydown="javascript:handleSearchKeyDown(event);"
      class="long" />
    <select name="searchType" id="searchType">
      <option value="1" selected="selected">
        <%Sendb(Copient.PhraseLib.Lookup("term.folders", LanguageID))%></option>
      <option value="2">
        <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%></option>
    </select>
    <input type="button" id="searchsubmit" name="searchsubmit" value="<%Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>"
      onclick="javascript:submitSearch();" /><br />
    <br />
    <div id="searchResults">
    </div>
  </div>
</div>
<%
  Send_PageEnd()

  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>