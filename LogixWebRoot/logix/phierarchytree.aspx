﻿﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.IO" %>
<%
    ' *****************************************************************************
    ' * FILENAME: phierarchytree.aspx 
    ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' * Copyright © 2002 - 2009.  All rights reserved by:script
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
    Dim dr As DataRow
    Dim dt As DataTable = Nothing
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim nodeID As Long
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim AssignedCount As Integer = 0

    Dim SelectedCount As Integer = 0
    Dim DispSelectedCount As Integer = 0
    Dim MaxNumberOfItemsInProdGroup As Integer = 0

    Dim pgID As Integer = 0
    Dim searchPathIDs As String() = Nothing
    Dim searchHierarchyPath As String = ""
    Dim searchSelNodeID As String = ""
    Dim itemSelectedPK As String = ""
    Dim SelectedOption As String = ""
    Dim levelPos, idPos As Integer
    Dim ParentNodeIdList As String = ""
    Dim SelectedNodeId As String = ""
    Dim qryStr As String = ""
    Dim Level As Integer
    Dim ID As Integer
    Dim CancelRefresh As String = "false"
    Dim ActionDisabled As String = ""
    Dim SearchString As String = ""
    Dim LevelDisplay As String = ""
    Dim Name As String = ""
    Dim ExternalID As String = ""
    Dim ProductID As String = ""
    Dim i As Integer = 0
    Dim Shaded As String = "shaded"
    Dim ShowNoItemMsg As Boolean = True
    Dim BannersEnabled As Boolean = False
    Dim BannerHierarchyIDs As String = ""
    Dim IdType As String = ""
    Dim OfferID As Long
    Dim EngineID As Integer
    Dim CreatedFromOffer As Boolean = False
    Dim InLinkMode As Boolean = False
    Dim rstItemAttrDetails As DataTable
    Dim rowItemAttrDetails As DataRow
    Dim BuyerID As Long = -1
    Dim ConditionID As String = ""
    Dim Disqualifier As String = ""
    Dim count As Integer = 0
    Dim PAB As String = ""
    Dim popupFlag As String
    Dim SelectedNodeIDs As String = ""
    'Dim AttributeProductGroupID As String = ""
	
    'AMSPS-1570: there are two separate user permissions for creating product groups and editing product groups, but unless Access Configuration is assigned as well the permissions dont work and they get access denied. Access Configuration has other undesired permissions for this user level. 
    'They would like the two permissions to work without Access Configuration

    Dim AutoAccessProductGroup As Boolean = IIf(MyCommon.Fetch_SystemOption(261) = "1", True, False)
    
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    MyCommon.AppName = "phierarchytree.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    DispSelectedCount = MyCommon.Fetch_SystemOption(258)
    MaxNumberOfItemsInProdGroup = MyCommon.Fetch_SystemOption(259)

    OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    CreatedFromOffer = (OfferID > 0 AndAlso EngineID > 0)
    
    If Request.QueryString("BuyerID") IsNot Nothing Then
        BuyerID = Convert.ToInt64(Request.QueryString("BuyerID"))
    End If
    ConditionID = GetCgiValue("ConditionID")
    Disqualifier = GetCgiValue("Disqualifier")
    'AttributeProductGroupID = GetCgiValue("AttributeProductGroupID")
    
    SearchString = Request.QueryString("searchString")
    SelectedOption = Request.QueryString("selected")
    SelectedNodeIDs = IIf(Request.QueryString("SelectedNodeIDs") Is Nothing, "", Request.QueryString("SelectedNodeIDs"))
    pgID = Request.QueryString("ProductGroupID")
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    PAB = MyCommon.Extract_Val(Request.QueryString("PAB"))
    popupFlag = Request.QueryString("PopupFlag")
    If (BannersEnabled) Then
        MyCommon.QueryStr = "select BPH.HierarchyID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join BannerProdHierarchies BPH with (NoLock) on BPH.BannerID = AUB.BannerID " & _
                            "where AdminUserID=" & AdminUserID
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
            For Each dr In dt.Rows
                If BannerHierarchyIDs <> "" Then BannerHierarchyIDs &= ","
                BannerHierarchyIDs &= MyCommon.NZ(dr.Item("HierarchyID"), -1)
            Next
        End If
    End If
  
    CancelRefresh = IIf(pgID > 0, "false", "true")
    ActionDisabled = IIf(pgID > 0, "", " disabled=""disabled""")
  
    If (SelectedOption <> "") Then
        levelPos = SelectedOption.IndexOf("L")
        idPos = SelectedOption.IndexOf("ID")
        If (levelPos > -1 AndAlso idPos > -1) Then
            Integer.TryParse(SelectedOption.Substring(levelPos + 1, idPos - 1), Level)
            Integer.TryParse(SelectedOption.Substring(idPos + 2), ID)
            ' Find parents based on the level
            Select Case Level
                Case 0 ' Root Level
                    nodeID = ID
                    searchHierarchyPath = nodeID
                    searchPathIDs = searchHierarchyPath.Split(",")
                    searchSelNodeID = nodeID
                  
                Case 1  ' Node Level
                    nodeID = ID
                    searchHierarchyPath = GetParentNodeList(nodeID)
                    If (searchHierarchyPath <> "") Then
                        searchPathIDs = searchHierarchyPath.Split(",")
                    End If
                    searchSelNodeID = nodeID
                Case 2 ' Item Level
                    nodeID = GetItemNodeID(ID)
                    searchHierarchyPath = GetParentNodeList(nodeID)
                    If (searchHierarchyPath <> "") Then
                        searchPathIDs = searchHierarchyPath.Split(",")
                    End If
                    searchSelNodeID = nodeID
                    itemSelectedPK = ID
                Case Else
            End Select
        End If
    End If
    
    InLinkMode = (Request.QueryString("Linking") = "1")
  
    Send_HeadBegin("term.hierarchy", "term.product")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
%>
<% 
    If popupFlag <> "0" Then%>
<style type="text/css">
    * html body {
        margin: 10px 10px 57px 10px !important;
    }

    #wrap {
        width: auto !important;
    }
</style>
<% Else%>
<style type="text/css">
    #custom1 {
        display: none;
    }

    #custom2 {
        display: none;
    }
</style>
<%  End If%>
<%  Send_Scripts({"jquery.min.js"})%>
<script type="text/javascript">
    var bLoading = false;
    var hierSel = -1;
    var treeNodeSel = -1;
    var listItemSel = -1;
    var levelSel = -1;
    var NODE_INDENT_SIZE = 17;
    var cancelRefresh =  <%Sendb(CancelRefresh)%>;
    var idList = "";
    var nodeidlist="";
    var selectednodenamesList ="";
    var StartIndex = 1;
    var LastNodeID = 0;
    var masterSortCode = 0;
    var ValuesNotSelectedCnt = 0;
    var selectedAttribute = '';
    var LoadLastResult = false;
    var loadlastItem = '';
    var nodeid= -1;
    var productcount =0;
    var PAB=<%Sendb(PAB)%>;
  var PABPath = "<%= Request.Url.AbsolutePath %>";
    var PrevPageCount=0;

  

    function toggleNode(id, level) {
        if (isNodeOpen(id, level)) {
            collapseNode(id, level);
        } else {
            expandNode(id, level, false);
        }
    }
  
    function handleItemDblClick(id, level, parent) {
        if ( ($('#disableHierarchyTree').val()==undefined )|| ($('#disableHierarchyTree').val()!=undefined &&  $('#disableHierarchyTree').val() != "true") ){
            var tblID = (level <= 1) ? "H" : "";
            var elem = document.getElementById("hId" + tblID + id);
            var reload = false;
            var pg = document.getElementById("productgroup").value;
            var lblProductcount= document.getElementById("lblProductsCount");
            nodeidlist="";
            selectednodenamesList = "";
            idList="";
            $('#pcount').hide();
   
            nodeid= id;
            xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=drill&node=' + id + '&level=' + level + '&parent=' + parent + '&reload=' + reload + '&pg=' + pg, 'drill', id );

        }
    }
    function expandNode(id, level, bReload) {
        var tblID = (level <= 1) ? "H" : "";
        var elem = document.getElementById("hId" + tblID + id);
        var imgElem = document.getElementById("img"+ tblID + id);
        var reload = (bReload) ? 1 : 0;
        var pg = document.getElementById("productgroup").value;
    
        if (elem != null) {
            xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=expand&node=' + id + '&level=' + level + "&reload=" + reload + "&pg=" + pg, 'expand', id );
        } else if (elem != null) {
            elem.style.display = "";   
        }
    
        if (imgElem != null) {
            imgElem.src = "/images/minus.png";
        }
    }
  
    function isNodeOpen(id, level) {
        var retVal = false;
        var tblID = (level <= 1) ? "H" : "";
        var elem = document.getElementById("hId" + tblID + id);
    
        if (elem != null) {
            retVal = (elem.style.display != 'none');
        }
    
        return retVal;    
    }
  
    function xmlhttpPost(strURL, qryStr, action, id) {
        var xmlHttpReq = false;
        var self = this;
        var level = -1;
        var parentID = 0
    
        bLoading = true;
        if (action != 'findMatchesItemAttrb' && action != 'CheckItemAttrb') {
            if (action != 'LastItemAttrbSearchCriteria' && action != 'BindlastItemAttributevalues') 
                setTimeout('showWait()', 2000);
        }
        else if (action == 'findMatchesItemAttrb' || action == 'BindlastItemAttributevalues' || action == 'CheckItemAttrb' ) {
            setTimeout('showWaitProdItemAttrb()', 1000);
        }
    
        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
            // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }
    
        self.xmlHttpReq.open('POST', strURL + '?' + qryStr, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.onreadystatechange = function() {
            if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
                if (action == 'drill') {
                    level = parseToken(qryStr, "level");
                    parentID = parseToken(qryStr, "parent");
                    reload = parseToken(qryStr, "reload");
                    //         productcount=self.xmlHttpReq.responseText.substr(self.xmlHttpReq.responseText.indexOf("productcount:")+13,self.xmlHttpReq.responseText.indexOf("</label>"));
                    //         if(document.getElementById("lblProductsCount") != null)
                    //         {
                    //        document.getElementById("lblProductsCount").innerHTML="Contains "+ productcount+"number of products";
                    //        }
                    self.xmlHttpReq.responseText.replace(self.xmlHttpReq.responseText.substr(self.xmlHttpReq.responseText.indexOf("productcount:"),self.xmlHttpReq.responseText.length),"");
                    updateNodes(id, parentID, self.xmlHttpReq.responseText, parseInt(level));
                    if (reload) {
                        highlightNode(id, parseInt(level), false);
                    }
       
                } else if (action == 'expand') {
                    level = parseToken(qryStr, "level");
                    reload = parseToken(qryStr, "reload");
                    updateNodes(id, parentID, self.xmlHttpReq.responseText, parseInt(level));
                    if (reload) {
                        highlightNode(id, parseInt(level), false);
                    }
                } else if (action == 'openFolder') {
                    level = parseToken(qryStr, "level");
                    updateNodes(id, parentID, self.xmlHttpReq.responseText, parseInt(level));
                    if (level != '' && !isNaN(level)) {
                        highlightNode(id, parseInt(level), false);
                    }
                } else if (action == 'showItems') {
                    //       productcount=self.xmlHttpReq.responseText.substr(self.xmlHttpReq.responseText.indexOf("productcount:")+13,self.xmlHttpReq.responseText.indexOf("</label>"));
                    //           if(document.getElementById("lblProductsCount") != null)
                    //         {
                    //        document.getElementById("lblProductsCount").innerHTML="Contains "+ productcount+"number of products";
                    //        }
                    self.xmlHttpReq.responseText.replace(self.xmlHttpReq.responseText.substr(self.xmlHttpReq.responseText.indexOf("productcount:"),self.xmlHttpReq.responseText.length),"");
                    updateItems(id, self.xmlHttpReq.responseText);
                } else if (action == 'addNode') {
                    updateAddNode(id, self.xmlHttpReq.responseText);
                } else if (action == 'deleteNode') {
                    updateDeleteNode(id, self.xmlHttpReq.responseText);
                } else if (action == 'deleteItems') {
                    updateItems(id, self.xmlHttpReq.responseText);
                } else if (action == 'showAvailItems') {
                    updateShowAvailItems(id, self.xmlHttpReq.responseText);
                } else if (action == 'removeAll') {
                    updateForRemoveAll(id, self.xmlHttpReq.responseText);
                } else if (action == 'findMatches') {
                    updateSearch(id, self.xmlHttpReq.responseText);
                } else if (action == 'findMatchesItemAttrb') {
                    updateSearchItemAttrib(id, self.xmlHttpReq.responseText);
                } else if (action == 'CheckItemAttrb') {
                    bindItemAttrbRangeValues(id, self.xmlHttpReq.responseText);
                } else if (action == 'BindlastItemAttributevalues') {
                    LoadBindlastItemAttributevalues(id, self.xmlHttpReq.responseText);
                } else if (action == 'LastItemAttrbSearchCriteria') {
                    displayLastItemAttrbSearchCriteria(id, self.xmlHttpReq.responseText);
                } else if (action == 'LinkMatchestoHierarchy' || action == 'RemoveMatchesFromHierarchy') {
                    displayLinkMatchestoHierarchy(id, self.xmlHttpReq.responseText);
                } else if (action == 'handleSearchItemAdjust') {
                    updateSearchItem(id, self.xmlHttpReq.responseText);
                } else if (action == 'handleSearchNodeAdjust') {
                    updateSearchNode(id, self.xmlHttpReq.responseText);
                } else if (action == 'delFromHierarchy') {
                    updateDelFromHierarchy(self.xmlHttpReq.responseText);
                } else if (action == 'linkToGroup') {
                    updateLinkToGroup(id, self.xmlHttpReq.responseText);
                } else if (action == 'removeLinkToGroup') {
                    updateRemoveLinkToGroup(id, self.xmlHttpReq.responseText);
                } else if (action == 'excludeFromGroup') {
                    updateExcludeFromGroup(id, self.xmlHttpReq.responseText);
                } else if (action == 'removeExclusion') {
                    updateRemoveExclusion(id, self.xmlHttpReq.responseText);
                } else if (action == 'AssignSigns') {
                    confirmSuccess(self.xmlHttpReq.responseText);
                } else if (action == 'GenerateDivForSigns') {
                    generatedivforsigns(self.xmlHttpReq.responseText);
                }
      
                bLoading = false;
                if (action != 'findMatchesItemAttrb' && action != 'CheckItemAttrb' ) {
                    if (action != 'LastItemAttrbSearchCriteria'  && action != 'BindlastItemAttributevalues') {
                        hideWait();
                    }  
                    else {
                        hideWaitProdItemAttrb();
                    }
                }	
                else if (action == 'findMatchesItemAttrb' || action == 'CheckItemAttrb') {
                    hideWaitProdItemAttrb();
                }
            }
        }
        self.xmlHttpReq.send("?" + qryStr);    
    }
  
    function parseToken(qryStr, tokenName) {
        var tokenValue = '';
        var startPos, endPos;
    
        if (qryStr != null && tokenName != null) {
            startPos = qryStr.indexOf(tokenName);
            if (startPos > -1) {
                endPos = qryStr.indexOf("&", startPos);
                // adjust startPos to account for the token name and equal sign
                startPos = startPos + tokenName.length + 1;
                if (endPos > -1) {
                    tokenValue = qryStr.substring(startPos, endPos);
                } else {
                    tokenValue = qryStr.substring(startPos);
                }
            }
        }
        return tokenValue;
    }
  
    function transmitGroups(strURL, qryStr, params) {
        var xmlHttpReq = false;
        var self = this;
    
        bLoading = true;
        setTimeout('showWait()', 1000);
    
        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
            // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }
    
        self.xmlHttpReq.open('POST', strURL + '?' + qryStr, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.setRequestHeader('Content-length', params.length);
        self.xmlHttpReq.setRequestHeader('Connection', 'close');
        self.xmlHttpReq.onreadystatechange = function() {
            if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
                showAddedCount(self.xmlHttpReq.responseText);
                bLoading = false;
                hideWait();
                idList = "";
                nodeidlist="";
                selectednodenamesList ="";
                window.close();
            }
        }
        self.xmlHttpReq.send(params);    
    }
  
    function showAddedCount(response) {
        var htmlBuf = '';
        var msgbuf = '';
        var trailerPos = -1;
        var msgStart, msgEnd;
        var ctStart, ctEnd;
    
        if (response != null && response.length >0) {
            ChangeParentDocument();            
            trailerPos = response.indexOf('<trailer>');
            if ( trailerPos > -1) {
                htmlBuf = response.substring(0, trailerPos-1);
                msgStart = response.indexOf('<message>', trailerPos);
                if (msgStart > -1) {
                    msgEnd = response.indexOf('<\/message>', msgStart);
                    if (msgEnd > -1) {
                        msgBuf = response.substring(msgStart + 9, msgEnd); 
                    }
                }
                ctStart = response.indexOf('<count>', trailerPos);
                if (ctStart > -1) {
                    ctEnd = response.indexOf('<\/count>', ctStart);
                    if (ctEnd > -1) {
                        var elemAssigned = document.getElementById('assignedCt');
                        if (elemAssigned != null) {
                            elemAssigned.innerHTML = response.substring(ctStart + 7, ctEnd);
                        }
                    }
                }
            } else {
                htmlBuf = response;
            }
      
            if (msgBuf != '') {
                alert(msgBuf);
            }

            highlightNode(treeNodeSel, levelSel, false);
        }
    }
  
    function showWait() {
        var elem = document.getElementById("itemList");
        var elemWait = document.getElementById("WaitDiv");
        var elemSearch = document.getElementById("SearchDiv");
        var elemToolbar = document.getElementById("toolset");
        var elemSearchType = document.getElementById("searchType");
    
        if (bLoading && elem != null) {
            if (elemToolbar != null) {
                elemToolbar.style.display = "none";
            }
            if (elemSearch != null && elemSearch.style.display != "none") {
                // disable background so the user can't make changes during wait.
                if (elemWait != null) {
                    elemWait.style.display = "block";
                    elemWait.innerHTML = '<div class=\"loading\"><br \/><img id=\"clock\" src=\"/images/clock22.png\" \/><br \/><% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';
                }
                if (elemSearchType != null) {
                    elemSearchType.style.visibility = 'hidden';
                }
            } else {
                elem.innerHTML = '<div class=\"loading\"><br \/><img id=\"clock\" src=\"/images/clock22.png\" \/><br \/><% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';
                // disable background so the user can't make changes during wait.
                if (elemWait != null) {
                    elemWait.style.display = "block";
                    elemWait.innerHTML = '<div class=\"loading\"><br \/><img id=\"clock\" src=\"/images/clock22.png\" \/><br \/><% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';
          }
      }
  }
}

function showWaitProdItemAttrb() {
    var elem = document.getElementById("itemList");
    var elemWait = document.getElementById("WaitDiv");
    var elemSearch = document.getElementById("SearchProdItemAttrbDiv");
    var elemToolbar = document.getElementById("toolset");
    
    if (bLoading && elem != null) {
        if (elemToolbar != null) {
            elemToolbar.style.display = "none";
        }
        if (elemSearch != null && elemSearch.style.display != "none") {
            // disable background so the user can't make changes during wait.
            if (elemWait != null) {
                elemWait.style.display = "block";
                elemWait.innerHTML = '<div class=\"loading\"><br \/><br \/><br \/><br \/><br \/><br \/><br \/><br \/><img id=\"clock\" src=\"/images/clock22.png\" \/><br \/><% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';
            }
        } else {
            elem.innerHTML = '<div class=\"loading\"><br \/><br \/><br \/><br \/><br \/><br \/><br \/><br \/><img id=\"clock\" src=\"/images/clock22.png\" \/><br \/><% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';
            // disable background so the user can't make changes during wait.
            if (elemWait != null) {
                elemWait.style.display = "block";
                elemWait.innerHTML = '<div class=\"loading\"><br \/><br \/><br \/><br \/><br \/><br \/><br \/><br \/><img id=\"clock\" src=\"/images/clock22.png\" \/><br \/><% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID))%><\/div>';
          }
      }
  }
}
  
function hideWaitProdItemAttrb() {
    var elemWait = document.getElementById("WaitDiv");
    var elem = document.getElementById("itemList");
    var elemToolbar = document.getElementById("toolset");
    elemToolbar.style.display = "none";
    if (elemWait != null) {
        elemWait.style.display = "none";
    }
        
    if (elem != null && elem.innerHTML.indexOf("clock22.png") > 0) {
        elem.innerHTML = '';
    }
}
  
function hideWait() {
    var elemWait = document.getElementById("WaitDiv");
    var elem = document.getElementById("itemList");
    var elemToolbar = document.getElementById("toolset");
    var elemSearch = document.getElementById("SearchDiv");
    var elemSearchItemAttrbSearch = document.getElementById("SearchProdItemAttrbDiv");
    var elemSearchType = document.getElementById("searchType");
    
    if (elemToolbar != null) {
        if (elemSearch == null || elemSearch.style.display != 'block') {
            if (elemSearchItemAttrbSearch.style.display != 'block')
                elemToolbar.style.display = "block";
        }
    }
    
    if (elemWait != null) {
        elemWait.style.display = "none";
    }
    
    if (elemSearchType != null) {
        elemSearchType.style.visibility = 'visible';
    }
    
    if (elem != null && elem.innerHTML.indexOf("clock22.png") > 0) {
        elem.innerHTML = '';
    }
}
  
function dynamicWidth(elem,evt,orglength) {
    var selection = document.getElementById(elem.id);
    selection.style.width = (orglength)+'px';
}

function updateNodes(id, parentID, divHTML, level) {
    var tblID = (level <= 2) ? "H" : "";
    var elem = document.getElementById("hId" + tblID + parentID);
    var elemImg = document.getElementById("img" + tblID + parentID);
    
    if (elem != null && divHTML != "") {
        if (parentID == 0 && level >= 2) {
        } else {
            elem.style.display = "block";
            elem.innerHTML = divHTML;
        }
    } else if (elemImg != null) {
        elemImg.src = "/images/blank.png";
        elem.style.display = "none";
        elem.innerHTML = "";
    }
}
  
function updateItems(id, divHTML) {
    var trailerStart = -1;
    var trailerEnd = -1;
    var htmlBuf = divHTML;
    var hlIndex = -1;
    var strArray = null;
    var statusMsg = '';
    var pg = document.getElementById("productgroup").value;
    var totalItemElem  = document.getElementById("totalItemCt");
    var elem = document.getElementById("itemList");
    
    if (elem != null) {
        if (divHTML != null && divHTML.length >0) {
            trailerStart = divHTML.indexOf('<trailer>');
            if ( trailerStart > -1) {
                htmlBuf = divHTML.substring(0, trailerStart-1);
                trailerEnd = divHTML.indexOf('<\/trailer>', trailerStart);
                if (trailerEnd > -1) {
                    strArray = divHTML.substring(trailerStart + 9, trailerEnd).split(",", 3);
                    //hlIndex = strArray[0];
                    if (totalItemElem != null && strArray[1] > 0) {
                        if (pg != "" && parseInt(pg) > 0) {
                            statusMsg = " | ";
                        }
                        statusMsg += " Items: " + strArray[1];
                        if (strArray[2] > 0) {
                            statusMsg += " (" + strArray[2] + " selected)";
                        }
                        totalItemElem.innerHTML = statusMsg;
                    } else if (totalItemElem != null) {
                        totalItemElem.innerHTML = "";
                    }
                    //hlIndex = parseInt(divHTML.substring(trailerStart + 9, trailerEnd));
                }
            }
            elem.innerHTML = htmlBuf;
            elem.style.display = "block";
        
            if (strArray != null && strArray[0]!= null && strArray[0] > -1) {
                highlightItem(strArray[0]);
                listItemSel = strArray[0];
                scrollToItem(strArray[0]);
            }
        }
    }
}

function updateForRemoveAll(id, response) {
    var pg = document.getElementById("productgroup").value;
    var msg = response;
    var ct = 0;
    var preCt, postCt;
    var commaPos = -1;
    var tokenValues = [];
    
    if (response != null && response.length > 0) {
        ChangeParentDocument();            
        if (response.substring(0,1) == "|") {
            commaPos = response.indexOf(",");
            preCt = parseInt(response.substring(1, commaPos));
            postCt = parseInt(response.substring(commaPos+1));
            if (postCt == 0) {
                tokenValues = [preCt, pg]
                msg = detokenizeString('<% Sendb(Copient.PhraseLib.Lookup("phierarchy.RemovedItems", LanguageID))%>', tokenValues);
            } else {
                tokenValues = [prect, pg, postCt];
                msg = detokenizeString('<% Sendb(Copient.PhraseLib.Lookup("phierarchy.PartialItemRemoval", LanguageID))%>', tokenValues);
      }
      var elemAssigned = document.getElementById('assignedCt');
      if (elemAssigned != null) {
          elemAssigned.innerHTML = postCt;
      }
      alert(msg);
  } else {
            // a database error was returned in the response.
      alert(response);
  }
    if (treeNodeSel > -1) {
        showNodeItems(id, levelSel, false);
    }
}
}
  
function  updateSearch(id, response) {
    var resultsElem = document.getElementById("linerresults");
    var itemElem = document.getElementById("searchItemCt");
    var trailerPos = -1, trailerEnd = -1;
    var resultCt = 0;      
    var MAX_RESULTS = 500;
    var msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.MaxResults", LanguageID))%>';
    var tokenValues = [MAX_RESULTS];
    
    if (response != null && response.length >0) {
        trailerPos = response.indexOf('<trailer>');
        if ( trailerPos > -1) {
            htmlBuf = response.substring(0, trailerPos-1);
            if (itemElem != null) {
                trailerEnd = response.indexOf("<\/trailer>", trailerPos);
                if (trailerEnd > -1) {
                    resultCt = response.substring(trailerPos+9, trailerEnd);
                } else {
                    resultCt = response.substring(trailerPos+9); 
                }
                itemElem.innerHTML = '| <%Sendb(Copient.PhraseLib.Lookup("term.results", LanguageID))%>: ' + resultCt
                if (resultCt >= MAX_RESULTS) {
                    alert(detokenizeString(msg, tokenValues));
                }
            }
        } else {
            htmlBuf = response;
        }
    }
    if (resultsElem != null) {
        resultsElem.innerHTML = htmlBuf;
    }
}
  
function LoadBindlastItemAttributevalues(id, response) {
    // alert(response);
    var str='';
    var Stindx=-1;
    var endindx=-1;
    var pg = document.getElementById("productgroup").value; 
    Stindx = response.indexOf('<ItemAttrb1>');  
    endindx=response.indexOf("<\/ItemAttrb1>", Stindx);
    str = response.substring(Stindx+12, endindx);
    document.getElementById("ItemAttrib1").value = str;
    Stindx = response.indexOf('<ItemAttrb2>');  
    endindx=response.indexOf("<\/ItemAttrb2>", Stindx);
    str = response.substring(Stindx+12, endindx);
    document.getElementById("ItemAttrib2").value = str;
    Stindx = response.indexOf('<ItemAttrb3>');  
    endindx=response.indexOf("<\/ItemAttrb3>", Stindx);
    str = response.substring(Stindx+12, endindx);
    document.getElementById("ItemAttrib3").value = str;
    loadlastItem = "ItemAttrib1";
    xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=LastItemAttrbSearchCriteria&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel + '&AttributeId=ItemAttrib1', 'LastItemAttrbSearchCriteria', treeNodeSel );		  
}
 
function displayLinkMatchestoHierarchy(id, response) {
    var str='';
    var strMessage='';
    var Stindx=-1;
    var endindx=-1;
    //alert(response);
    Stindx = response.indexOf('<Status>');  
    endindx=response.indexOf("<\/Status>", Stindx);
    str = response.substring(Stindx+8, endindx);
    if (str == 'Success') {
        Stindx = response.indexOf('<message>');
        if ( Stindx > -1) {
            endindx=response.indexOf("<\/message>", Stindx);
            strMessage= response.substring(Stindx+9, endindx);
            alert(strMessage);
            closeSearchItemAttrb();
        }
        else {
            closeSearchItemAttrb();
            if (handleLinkingRefresh(id, response)) {
                updateFolderIcons(id, true);
            }
        }   
    }
    else
        alert(str);
}
 
function displayLastItemAttrbSearchCriteria(id, response) {
    var resultsElem = document.getElementById("ItemAttrbDiv1StartEndValues");
    var pg = document.getElementById("productgroup").value; 
    if (loadlastItem == "ItemAttrib1") {
        resultsElem = document.getElementById("ItemAttrbDiv1StartEndValues");
        resultsElem.innerHTML = response;
        resultsElem.style.visibility = 'visible';
        loadlastItem = "ItemAttrib2";
        xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=LastItemAttrbSearchCriteria&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel + '&AttributeId=ItemAttrib2', 'LastItemAttrbSearchCriteria', treeNodeSel );		  
    }
    else if (loadlastItem == "ItemAttrib2") {
        resultsElem = document.getElementById("ItemAttrbDiv2StartEndValues");
        resultsElem.innerHTML = response;
        resultsElem.style.visibility = 'visible';
        loadlastItem = "ItemAttrib3";
        xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=LastItemAttrbSearchCriteria&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel + '&AttributeId=ItemAttrib3', 'LastItemAttrbSearchCriteria', treeNodeSel );		  
    }
    else if (loadlastItem == "ItemAttrib3") {
        resultsElem = document.getElementById("ItemAttrbDiv3StartEndValues");     
        resultsElem.innerHTML = response;
        resultsElem.style.visibility = 'visible';
        displayAttributeSearchDiv();
    }
    var itemToolbarDiv = document.getElementById("toolset");
    itemToolbarDiv.style.display = 'none';
} 
   
function displayAttributeSearchDiv() {
    var searchDiv = document.getElementById("SearchDiv");
    var SearchProdItemAttrbDiv = document.getElementById("SearchProdItemAttrbDiv");
    var searchBoxDiv = document.getElementById("searchbox");
    var itemToolbarDiv = document.getElementById("toolset");
    var totalItemElem  = document.getElementById("totalItemCt");
    var resultsElem = document.getElementById("searchItemCt");
    var searchFolder = document.getElementById("ItemAttrbsearchFolder");
    var ItemAttribLineResults = document.getElementById("linerresultsItemAttrib");
    var elemToolbar = document.getElementById("toolset");
	
    ItemAttribLineResults.style.visibility='hidden';
    //if (searchDiv != null) {
    //  searchDiv.style.display = "none";
    //}  
    if (searchBoxDiv != null) {
        searchBoxDiv.style.display = 'none';
    }
    elemToolbar.style.display = 'none';
    var leftPaneDiv = document.getElementById("leftpane");
    if (leftPaneDiv != null) {
        leftPaneDiv.style.display = 'none';
    }
    var rightPaneDiv = document.getElementById("rightpane");
    if (rightPaneDiv != null) {
        rightPaneDiv.style.display = 'none';
    }
    // disable main pane's controls 
    if (itemToolbarDiv != null) { 
        itemToolbarDiv.style.display = 'none';
    }
    if (totalItemElem != null) {
        totalItemElem.style.display = 'none';
    }
    if (resultsElem != null) {
        resultsElem.style.display = 'inline';
    }            
    if (searchFolder != null) {
        searchFolder.innerHTML = getSelectedFolderFullName();
    }  
    if (SearchProdItemAttrbDiv != null) {
        SearchProdItemAttrbDiv.style.display = "block";
    }  
    hideWaitProdItemAttrb();
}

function bindItemAttrbRangeValues(id, response) {
    var resultsElem = document.getElementById("ItemAttrbDiv1StartEndValues");
    if (selectedAttribute == "ItemAttrib1")
        var resultsElem = document.getElementById("ItemAttrbDiv1StartEndValues");
    if (selectedAttribute == "ItemAttrib2")
        var resultsElem = document.getElementById("ItemAttrbDiv2StartEndValues");
    if (selectedAttribute == "ItemAttrib3")
        var resultsElem = document.getElementById("ItemAttrbDiv3StartEndValues");     
    //alert(response);
    if (resultsElem != null) {
        resultsElem.innerHTML = response;
    }
    resultsElem.style.visibility = 'visible';
}
  
function  updateSearchItemAttrib(id, response) {
    var resultsElem = document.getElementById("linerresultsItemAttrib");
    var resultsCount = document.getElementById("lblAttrSearchResultCnt");
	
    var linkgroup = document.getElementById("btnAttrSearchLinktoHierarchy");
    var excludegroup = document.getElementById("btnAttrSearchExcludefromHierarchy");
    var pg = document.getElementById("productgroup").value;
    resultsElem.style.visibility = 'visible';
    var trailerPos = -1, trailerEnd = -1;
    var resultCt = 0;      
    var MAX_RESULTS = 500;
    var msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.MaxResults", LanguageID))%>';
    var tokenValues = [MAX_RESULTS];
    
    if (response != null && response.length >0) {
        trailerPos = response.indexOf('<trailer>');
        if ( trailerPos > -1) {
            trailerEnd = response.indexOf("<\/trailer>", trailerPos);
            if (trailerEnd > -1) {
                resultCt = response.substring(trailerPos+9, trailerEnd);
            } else {
                resultCt = response.substring(trailerPos+9); 
            }
            if (resultCt >= MAX_RESULTS) {
                alert(detokenizeString(msg, tokenValues));
            }
        } 
    }
    resultsCount.innerHTML = resultCt;
    if (resultCt > 0 ) {
        if ( pg > 0) {
            linkgroup.disabled="";
            excludegroup.disabled="";
        }
        else {
            linkgroup.disabled="disabled";
            excludegroup.disabled="disabled";
        }
    }
    else {
        linkgroup.disabled="disabled";
        excludegroup.disabled="disabled";	
    }
}

function updateSearchItem(id, response) {
    if (response == 'ADDED' || response == 'REMOVED') {
        var elemAssigned = document.getElementById('assignedCt');
        var elemLink = document.getElementById("link" + id);
        var elemImg = document.getElementById("img" + id);
        var newAction = (response == 'ADDED') ? 'Remove' : 'Add';
        var newImg = (response== 'ADDED') ? '/images/upc-on.png' : '/images/upc.png';
        var adjVal = (response == 'ADDED') ? 1 : -1;
      
        if (elemLink != null ) {
            elemLink.innerHTML = newAction;
        }
        if (elemImg != null) {
            elemImg.src = newImg;
        }
      
        if (elemAssigned != null) {
            if (!isNaN(elemAssigned.innerHTML)) {
                elemAssigned.innerHTML = parseInt(elemAssigned.innerHTML) + adjVal
            }
        }
      
    } else {
        alert(response);
    }
}
  
function updateSearchNode(id, response) {
    var htmlBuf = '';
    var msgbuf = '';
    var trailerPos = -1;
    var msgStart, msgEnd;
    var ctStart, ctEnd;
    
    if (response != null && response.length >0) {
        trailerPos = response.indexOf('<trailer>');
        if ( trailerPos > -1) {
            htmlBuf = response.substring(0, trailerPos-1);
            msgStart = response.indexOf('<message>', trailerPos);
            if (msgStart > -1) {
                msgEnd = response.indexOf('<\/message>', msgStart);
                if (msgEnd > -1) {
                    msgBuf = response.substring(msgStart + 9, msgEnd); 
                }
            }
            ctStart = response.indexOf('<count>', trailerPos);
            if (ctStart > -1) {
                ctEnd = response.indexOf('<\/count>', ctStart);
                if (ctEnd > -1) {
                    var elemAssigned = document.getElementById('assignedCt');
                    if (elemAssigned != null) {
                        elemAssigned.innerHTML = response.substring(ctStart + 7, ctEnd);
                    }
                }
            }
        } else {
            htmlBuf = response;
        }
        if (msgBuf != '') {
            alert(msgBuf);
        }
    }
}
  
function updateAddNode(id, resp) {
    var level = 0;
    if (treeNodeSel > - 1) {
        var elemIndent = document.getElementById("indent" + treeNodeSel);
        level = (parseInt(elemIndent.style.paddingLeft) / NODE_INDENT_SIZE) + 1;
    }
    var elemTopLevel = document.getElementById("topLevel");
    if (elemTopLevel != null && elemTopLevel.checked) {
        level = 0;
    }
    if (level == 0) {
        //window.location = 'phierarchytree.aspx';
        displayTopNode(resp);
        showAddNode(false);
    } else {
        expandNode(id, level, true);
        showAddNode(false)
    }
}
  
function displayTopNode(divHTML) {
    var elem = document.getElementById("topNodesDiv");
    elem.innerHTML = elem.innerHTML + divHTML;
}
  
function collapseNode(id, level) {
    var tblID = (level <= 1) ? "H" : "";
    var elem = document.getElementById("hId" + tblID + id);
    var imgElem = document.getElementById("img" + tblID + id);
    
    if (elem != null) {
        elem.innerHTML = "";
        elem.style.display = 'none';
    }
    if (imgElem != null) {
        imgElem.src = "/images/plus.png";
    }
}
  
function showNodeItems(id, level, passItemSelected,selectCheckboxes) {
    var name = "";    
    
    var pg = document.getElementById("productgroup").value;
    var itemPK = document.getElementById("itemSelectedPK").value;
    var tblID = (level <= 1) ? "H" : "";
    var pab2=<%Sendb(PAB)%>;
    var PAB1=(pab2 == "0" || pab2 == null)?"0":"1";
    if (id > -1) {
        elem = document.getElementById('name' + tblID + id);
        if (elem != null) {
            name = elem.innerHTML;
        }

        var qryStr ='';
        qryStr =  'action=showItems&node=' + id + '&level=' + level + "&nodeName=" + encodeURIComponent(name) + "&pg=" + pg + "&itemPK=" + itemPK +"&buyerid=<% Sendb(BuyerID)%>&StartIndex=" + StartIndex + ((masterSortCode > 0) ? "&sort=" + masterSortCode : '') +"&PAB="+PAB1;
        if(typeof selectCheckboxes != undefined && selectCheckboxes == true){
            qryStr = qryStr + '&SelectedNodeIDs=<%Sendb(SelectedNodeIDs)%>';
      }
      else if (nodeidlist!=""){
          qryStr = qryStr + '&SelectedNodeIDs=' + nodeidlist;
      }
      else
      {
          idList = '';
          nodeidlist="";
          selectednodenamesList ="";
      }

      if (passItemSelected) {
          qryStr = qryStr + "&itemSelectedPK=" + <%Sendb(IIf(itemSelectedPK <> "", itemSelectedPK, "0"))%>;       
      }
      xmlhttpPost('/logix/HierarchyFeeds.aspx',qryStr, 'showItems', id  );      
  }
}
  
function highlightNode(id, level, passItemSelected, selectCheckboxes) {
    var elem = null;
    var hIdelem = null;
    var tblID = (level <= 1) ? "H" : "";

    
    StartIndex = 1;
    showNodeItems(id, level, passItemSelected,selectCheckboxes);
    
    // highlight the selected node
    elem = document.getElementById('name' + tblID + id);
    if (elem != null) {
        elem.style.backgroundColor = '#000080';
        elem.style.color = '#ffffff';
    }
    
    // unselect the previously selected node
    if (treeNodeSel > -1 && (id != treeNodeSel || level != levelSel)) {
        if (levelSel <= 1) {
            elem = document.getElementById('nameH' + treeNodeSel);
        } else {
            elem = document.getElementById('name' + treeNodeSel);
        }
        if (elem != null) {
            elem.style.backgroundColor = '';
            elem.style.color = '';
        }            
    }
    
    // empty the contents of the selected node's hID span (which contains its subnodes)
    hIdelem = document.getElementById("hId" + tblID + id);
    if (hIdelem != null) {
        hIdelem.style.display = "none";
        hIdelem.innerHTML = "";
    }
    
    if (hierSel > - 1 && (id != hierSel || level != levelSel)) {
        elem = document.getElementById('nameH' + hierSel);
        if (elem != null) {
            elem.style.backgroundColor = '';
            elem.style.color = '';
        }
    }
    
    treeNodeSel = (level == 1) ? -1 : id;
    hierSel = (level == 1) ? id : -1;
    listItemSel = -1;
    levelSel = level;
}
  
function highlightItem(id) {
    if ( ($('#disableHierarchyTree').val()==undefined )|| ($('#disableHierarchyTree').val()!=undefined &&  $('#disableHierarchyTree').val() != "true") ){
        var elem = null;
        var lblProductcount= document.getElementById("lblProductsCount");
    
        //      if(lblProductcount!= null)
        //    {
        //        lblProductcount.innerHTML="Contains " + productcount +" Products";
        //    }
        elem = document.getElementById('itemRow' + id);
        if (elem != null) {
            elem.style.backgroundColor = '#000080';
            elem.style.color = '#ffffff';
        }
        if (listItemSel > -1) {
            elem = document.getElementById('itemRow' + listItemSel);
            if (elem != null) {
                elem.style.backgroundColor = '#ffffff';
                elem.style.color = '#000000';
            }            
        }
        if (id == listItemSel) {
            listItemSel = -1;
        } else {
            listItemSel = id;
        }
    }
}
function showAddNode(bShow) {
    var elem = document.getElementById("AddNodeDiv");
    var elemAddItem = document.getElementById("AddItemDiv");
    var elemName = document.getElementById("nodeName")    
    var elemTopLevel = document.getElementById("topLevel");
    
    if (elemAddItem != null && elemAddItem.style.display!="none") {
        showAddItem(false);
    }
    
    if (elem != null) {
        elem.style.display = (bShow) ?  'block' : 'none';
        if (bShow && elemName != null) {
            elemName.focus();
            elemName.value = "";
        }
        if (bShow && elemTopLevel!=null) {
            elemTopLevel.checked = (treeNodeSel==-1) ? true : false;
        }
    }
}
  
function addNode() {
    var elemTopLevel = document.getElementById("topLevel");
    var elemNodeName = document.getElementById("nodeName");
    var level = 0, parentId = 0, hierId =0;
    var nodeName = '', qryStr = '';
    
    if (treeNodeSel > - 1) {
        var elemHier = document.getElementById("hierarchy" + treeNodeSel);
        if (elemHier != null) {
            hierId = elemHier.value;
        }
        var elemIndent = document.getElementById("indent" + treeNodeSel);
        level = (parseInt(elemIndent.style.paddingLeft) / NODE_INDENT_SIZE) + 1;
    } else {
        hierId = 0;
        level = 0;
    }
    // override the selected node
    if (elemTopLevel.checked) {
        level = 0;
        hierId = 0;
    }                        
    parentId = (treeNodeSel==-1) ? 0 : treeNodeSel;
    nodeName = elemNodeName.value;
    qryStr = 'action=addNode&node=' + parentId + '&level=' + level;
    qryStr += '&nodeName=' + nodeName + '&hierId=' + hierId;
    xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'addNode', parentId  );
}
  
function showDeleteNode() {
    var elemName = null;
    var name = "";
    var response;
    var msg = '';
    var tokenValues = [];
    
    if (treeNodeSel == -1) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.SelectNodeToDelete", LanguageID))%>')
    } else {
        elemName = document.getElementById("name" + treeNodeSel);
        if (elemName != null) {
            name = elemName.innerHTML;            
        }
        msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmDeleteNode", LanguageID))%>';
        tokenValues[0] = name;
        response = confirm(detokenizeString(msg, tokenValues));
        if (response) {
            deleteNode(treeNodeSel);
        }
    }
}
  
function deleteNode(id) {
    var elemHier = document.getElementById("hierarchy" + treeNodeSel);
    
    if (elemHier != null) {
        hierId = elemHier.value;
    }
    var elemIndent = document.getElementById("indent" + treeNodeSel);
    var level = (parseInt(elemIndent.style.paddingLeft) / NODE_INDENT_SIZE);
    var qryStr = ""
      
    qryStr = 'action=deleteNode&node=' + id + '&level=' + level + '&hierId=' + hierId;
    xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'deleteNode', id );
}
  
function updateDeleteNode(id, resp) {
    var elemNode = document.getElementById("node" + id);
    var pos = 0, level = 1;
    
    if (elemNode != null && resp.indexOf("OK") > -1) {
        elemNode.parentNode.removeChild(elemNode);
        pos = resp.indexOf("ParentID=");
        if (pos > -1) {
            var parentId = resp.substring(pos + 9);
            if (!isNaN(parentId)) {
                level = (parseInt(parentId) == -1) ? 0 : 1;
                // TO DO: Collapse Node if this is the last node for the parent
                //collapseNode(parseInt(parentId), level);
            }
        }
    } else {
        alert(resp);
    }
}
  
function showDeleteItems() {
    var elem = null;
    var i = 0;
    var delItems = "";
    var response;
    
    elem = document.getElementById("chk" + i);
    
    while (elem != null) {
        if (elem.checked) {
            delItems += elem.value;
            delItems += ",";
        }
        i++;
        elem = document.getElementById("chk" + i);
    }
    if (delItems == "") {
        alert(Copient.PhraseLib.Lookup("hierarchy.selectdelete", LanguageID));
    } else {
        response = confirm(Copient.PhraseLib.Lookup("confirm.delete", LanguageID));
        if (response) {
            qryStr = 'action=deleteItems&items=' + delItems + "&node=" + treeNodeSel;
            xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'deleteItems', treeNodeSel);
        }
    }
}
  
function showAddItem(bShow) {
    var elem = document.getElementById("AddItemDiv");
    var elemAddNode = document.getElementById("AddNodeDiv");
    var elemNodeName = null;
    var qryStr = "", nodeName = "";
    
    if (treeNodeSel > -1) {
        if (elemAddNode != null && elemAddNode.style.display!="none") {
            showAddNode(false);
        }
        if (elem != null) {
            elem.style.display = (bShow) ?  'block' : 'none';
            if (bShow) {
                elemNodeName = document.getElementById('name' + treeNodeSel);
                if (elemNodeName != null) {
                    nodeName = elemNodeName.innerHTML;
                }
                elem.innerHTML = "<br \/><br \/><center><h2>" & Copient.PhraseLib.Lookup("message.loading", LanguageID) & "<\/h2><\/center>";
                qryStr = 'action=showAvailItems&node=' + treeNodeSel + '&nodeName=' + nodeName;
                xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'showAvailItems', treeNodeSel);
            }
        }
    } else {
        alert(Copient.PhraseLib.Lookup("hierarchy.selectnode", LanguageID));
    }        
}
  
function updateShowAvailItems(id, divHTML) {
    var elem = document.getElementById("AddItemDiv");
    
    if (elem != null) {
        elem.innerHTML = divHTML;
    }
}
  
function addToGroup() {
    var pg = document.getElementById("productgroup").value;
    var elemChk = null;
    var i = 0;
    var nodeIDs = "";

    var assignedCt = document.getElementById("assignedCt").innerHTML.toString();
    var selectedNotAssignedCt = document.getElementById("selectedNotAssignedCt").value.toString();
    var assignedPlusSelected = parseInt(assignedCt) + parseInt(selectedNotAssignedCt);
    var maxAllowedItems = '<% Sendb(MaxNumberOfItemsInProdGroup)%>';

    if (assignedPlusSelected > maxAllowedItems){
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NonItemsAdded", LanguageID))%>');
        return;
    }
 
    nodeIDs = idList;
    if (nodeIDs > "") {
        transmitGroups('/logix/HierarchyFeeds.aspx', "action=assignNodes", "pg=" + pg + "&sel=" + treeNodeSel + "&ids=" + nodeIDs);       
    } else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NoItemsSelected", LanguageID))%>');
    }
}
  
 
function linkToGroup() {
    var pg = document.getElementById("productgroup").value;
    var qryStr = '';
    var nodeIDs = "";

    var assignedCt = document.getElementById("assignedCt").innerHTML;
    var selectedNotAssignedCt = document.getElementById("selectedNotAssignedCt").value;
    var assignedPlusSelected = parseInt(assignedCt) + parseInt(selectedNotAssignedCt);
    var maxAllowedItems = '<% Sendb(MaxNumberOfItemsInProdGroup)%>';

    if (assignedPlusSelected > maxAllowedItems){
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NonItemsLinked", LanguageID))%>');
        return;
    }
    else
    {
 
        nodeIDs = idList;
        if (nodeIDs > "") {
            if (hasItemSelected(nodeIDs)) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.InvalidItemSelection", LanguageID))%>');            
            } else {
                qryStr = "action=linkToGroup&pg=" + pg + "&hid=" + hierSel + "&sel=" + treeNodeSel + "&ids=" + nodeIDs
                xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'linkToGroup', treeNodeSel);
                setTimeout(function (){window.close()},1000);
            }
        } else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NoItemsSelected", LanguageID))%>');
        }
    }
}
  
  
  
function removeLinkToGroup() {
    var pg = document.getElementById("productgroup").value;
    var qryStr = '';
    var nodeIDs = "";
    
    nodeIDs = idList;
    if (nodeIDs > "") {
        if (hasItemSelected(nodeIDs)) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.InvalidItemSelection", LanguageID))%>');            
        } else {
            qryStr = "action=removeLinkToGroup&pg=" + pg + "&sel=" + treeNodeSel + "&ids=" + nodeIDs
            xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'removeLinkToGroup', treeNodeSel);
            setTimeout(function (){window.close()},1000);
        }
    } else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NoItemsSelectedToRemove", LanguageID))%>');            
    }
   
}
  
function excludeFromGroup() {
    var pg = document.getElementById("productgroup").value;
    var qryStr = '';
    var nodeIDs = "";
    
    nodeIDs = idList;
    if (nodeIDs > "") {
        //alert(nodeIDs);
        qryStr = "action=excludeFromGroup&pg=" + pg + "&hid=" + hierSel + "&sel=" + treeNodeSel + "&ids=" + nodeIDs
        xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'excludeFromGroup', treeNodeSel);
        setTimeout(function (){window.close()},1000);
    } else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NoItemsSelectedToExclude", LanguageID))%>');            
    }
    
}
  
function removeExclusion() {
    var pg = document.getElementById("productgroup").value;
    var qryStr = '';
    var nodeIDs = "";
    
    nodeIDs = idList;
    if (nodeIDs > "") {
        //alert(nodeIDs);
        qryStr = "action=removeExclusion&pg=" + pg + "&sel=" + treeNodeSel + "&ids=" + nodeIDs
        xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'removeExclusion', treeNodeSel);
        setTimeout(function (){window.close()},1000);
    } else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NoItemsSelectedRemoveExclusion", LanguageID))%>');            
    }
  
}
  
function removeFromGroup() {
    var pg = document.getElementById("productgroup").value;
    var elemChk = null;
    var i = 0;
    var nodeIDs = "";
    
    nodeIDs = idList;
    if (nodeIDs > "") {
        transmitGroups('/logix/HierarchyFeeds.aspx', "action=unassignNodes", "pg=" + pg + "&sel=" + treeNodeSel + "&ids=" + nodeIDs);       
    } else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NoItemsSelectedForRemovePG", LanguageID))%>');
    }
}
  
function AssignSigns() {
    var signele=document.getElementById('Signs');
    var nodeIDs = "";
    var hiesign = '';
    nodeIDs = idList;
    
    if (signele != null) {
        hiesign = signele.options[signele.selectedIndex].text;
    }
    qryStr = "action=AssignSigns&sel=" + treeNodeSel + "&ids=" + nodeIDs + "&hID=" + hierSel + "&hiesign=" + hiesign
    xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'AssignSigns');
}
  
function divsign() {
    if (treeNodeSel > 0 || hierSel > -1)  {
        //alert(treeNodeSel);
        qryStr = "action=GenerateDivForSigns&sel=" + treeNodeSel + "&hID=" + hierSel
        xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'GenerateDivForSigns');      
    } else {
        alert("No Item are currently selected to assign signs");
    }
}

function search() {
    var pg = document.getElementById("productgroup").value;                
    var searchDiv = document.getElementById("SearchDiv");
    var SearchProdItemAttrbDiv = document.getElementById("SearchProdItemAttrbDiv");
    var searchBoxDiv = document.getElementById("searchbox");
    var itemToolbarDiv = document.getElementById("toolset");
    var totalItemElem  = document.getElementById("totalItemCt");
    var resultsElem = document.getElementById("searchItemCt");
    var searchFolder = document.getElementById("searchFolder");
    var lblProductcount= document.getElementById("lblProductsCount");
    if(lblProductcount != null)
    {
        // document.getElementById("lblProductsCount").innerHTML="";
    }
    if (searchDiv != null) {
        searchDiv.style.display = "block";
      
        if (searchBoxDiv != null) {
            searchBoxDiv.style.display = 'block';
        }
        if (SearchProdItemAttrbDiv != null) {
            SearchProdItemAttrbDiv.style.display = 'none';
        }
        var leftPaneDiv = document.getElementById("leftpane");
        if (leftPaneDiv != null) {
            leftPaneDiv.style.display = 'none';
        }
        var rightPaneDiv = document.getElementById("rightpane");
        if (rightPaneDiv != null) {
            rightPaneDiv.style.display = 'none';
        }
        // disable main pane's controls 
        if (itemToolbarDiv != null) { 
            itemToolbarDiv.style.display = 'none';
        }
        if (totalItemElem != null) {
            totalItemElem.style.display = 'none';
        }
        if (resultsElem != null) {
            resultsElem.style.display = 'inline';
        }
      
        if (searchFolder != null) {
            searchFolder.innerHTML = getSelectedFolderName();
        }
    }
}

  

function searchItemAttrb(lowerhirlevel,upperhirlevel) {
    // alert("lowerhirlevel" + lowerhirlevel + " :: upperhirlevel" + upperhirlevel);
    //alert('nodeid=' + treeNodeSel + ' :: hierid=' + hierSel);
    var allowfunction = false;
    if (lowerhirlevel > 0 && upperhirlevel > 0){
        if (levelSel >= lowerhirlevel && levelSel <= upperhirlevel)
            allowfunction = true;
        else {
            allowfunction = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.notallowProdHierLevel", LanguageID))%>');
            //alert("Product hierarchy level should allow between (" + lowerhirlevel + " - " + upperhirlevel + ")" );
        }
    } 
    else {
        if (lowerhirlevel == 0 && upperhirlevel == 0) {
            allowfunction = true;
        }
    }
    if (allowfunction == true) {
        if (LoadLastResult == false) {
            //LoadLastResult = true;
            var pg = document.getElementById("productgroup").value;  
            xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=BindlastItemAttributevalues&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel, 'BindlastItemAttributevalues', treeNodeSel );		  			
        }
        else
            displayAttributeSearchDiv(); 
    }  
}
  
function closeSearch() {
    var searchDiv = document.getElementById("SearchDiv");
    var searchBoxDiv = document.getElementById("searchbox");
    var itemToolbarDiv = document.getElementById("toolset");
    var totalItemElem  = document.getElementById("totalItemCt");
    var resultsElem = document.getElementById("searchItemCt");
    var itemAttrbsearchBoxDiv = document.getElementById("searchboxitemattrb");
    var searchtext = document.getElementById("searchString");
    var resultmsg = document.getElementById("noresultmsg");
    if (searchDiv != null) {
        searchDiv.style.display = "none";
      
        if(resultmsg != null)
            resultmsg.innerText = "";
        if (searchtext != null) 
            searchtext.value = "";

        if (searchBoxDiv != null) {
            searchBoxDiv.style.display = 'none';
        }
        if (itemAttrbsearchBoxDiv != null) {
            itemAttrbsearchBoxDiv.style.display = 'none';
        }      
        var leftPaneDiv = document.getElementById("leftpane");
        if (leftPaneDiv != null) {
            leftPaneDiv.style.display = 'block';
        }
      
        var rightPaneDiv = document.getElementById("rightpane");
        if (rightPaneDiv != null) {
            rightPaneDiv.style.display = 'block';
        }
      
        // disable main pane's controls 
        if (itemToolbarDiv != null) { 
            itemToolbarDiv.style.display = 'block';
        }
      
        if (totalItemElem != null) {
            totalItemElem.style.display = 'inline';
        }
      
        if (resultsElem != null) {
            resultsElem.style.display = 'none';
        }
    }
}

function closeSearchItemAttrb() {
    var searchDiv = document.getElementById("SearchDiv");
    var SearchProdItemAttrbDiv = document.getElementById("SearchProdItemAttrbDiv");
    var searchBoxDiv = document.getElementById("searchbox");
    var itemToolbarDiv = document.getElementById("toolset");
    var totalItemElem  = document.getElementById("totalItemCt");
    var resultsElem = document.getElementById("searchItemCt");
	   		
    if (searchDiv != null) {
        searchDiv.style.display = "none";
          
        if (searchBoxDiv != null) {
            searchBoxDiv.style.display = 'none';
        }
        if (SearchProdItemAttrbDiv != null) {
            SearchProdItemAttrbDiv.style.display = 'none';
        }
        var leftPaneDiv = document.getElementById("leftpane");
        if (leftPaneDiv != null) {
            leftPaneDiv.style.display = 'block';
        }
         
        var rightPaneDiv = document.getElementById("rightpane");
        if (rightPaneDiv != null) {
            rightPaneDiv.style.display = 'block';
        }
         
        // disable main pane's controls 
        if (itemToolbarDiv != null) { 
            itemToolbarDiv.style.display = 'block';
        }
        
        if (totalItemElem != null) {
            totalItemElem.style.display = 'inline';
        }
         
        if (resultsElem != null) {
            resultsElem.style.display = 'none';
        }
    }
}
  
function scrollToItem(itemIndex) {
    var elem = document.getElementById("itemRow" + itemIndex);
    var column2Div = document.getElementById("rightpane");
    
    if (elem != null) {
        ScrollToElement(elem, column2Div);
    }
}
  
function ScrollToElement(theElement, div){
    var selectedPosX = 0;
    var selectedPosY = 0;
    var divHeight = 0;
    
    while(theElement != null){
        selectedPosX += theElement.offsetLeft;
        selectedPosY += theElement.offsetTop;
        theElement = theElement.offsetParent;
    }
    if (div != null) {
        divHeight = parseInt(div.style.height);
        div.scrollTop = selectedPosY - (divHeight / 2);
    }
}
  
function removeAll() {
    var pg = document.getElementById("productgroup").value;
    var msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmRemoveAll", LanguageID))%>';
    var tokenValues = [pg];
    var confirmResponse = confirm(detokenizeString(msg, tokenValues));
    
    if (confirmResponse) {
        xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=removeAll&node=' + treeNodeSel + "&pg=" + pg, 'removeAll', treeNodeSel );
    }
}
  
function findMatches() {
    var searchElem = document.getElementById("searchString");
    var searchTypeElem = document.getElementById("searchType");
    var pg = document.getElementById("productgroup").value;
    var sType = 0;
    var pab2=<%Sendb(PAB)%>;
    var PAB1=(pab2 == "0" || pab2 == null)?"0":"1";
    if (searchElem != null) {
        if (searchElem.value == '') {
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.searchblank", LanguageID))%>');
            searchElem.focus();
        } else {
            if (searchTypeElem != null) { sType = searchTypeElem.value; }
            xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=findMatches&pg=' + pg + '&search=' + encodeURIComponent(htmlEntities(searchElem.value)) + '&stype=' + sType + '&nodeid=' + treeNodeSel + '&hierid=' + hierSel+'&PAB='+PAB1, 'findMatches', treeNodeSel );
        }
    }
}

function ItemAttributeSearchCriteria(ItemAttributeValue, exactORrange, StartValue, EndValue) {
    var newSearchString = '';
    if (ItemAttributeValue != '-1' && ItemAttributeValue != '') {
        if (StartValue == '' && EndValue == '') {
            //alert("Invalid start/End value.");
            //newSearchString = "NotAllow";
            //alert("ValuesNotSelectedCnt :: " + ValuesNotSelectedCnt);
            ValuesNotSelectedCnt = ValuesNotSelectedCnt + 1;
        }		
        else {
            if (exactORrange == 'exact') {
                newSearchString = "(HierAttribID = " + ItemAttributeValue + " and (HierAttribValue = " + StartValue + " or HierAttribValue = " + EndValue + "))";
            } 
            else if (exactORrange == 'range') {
                if (parseInt(StartValue) > parseInt(EndValue)) {
                    //alert("Invalid start/ end range value");
                    alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.invaliditemattrbrangeselection", LanguageID))%>');
                    newSearchString = "NotAllow";
                }
                else {
                    newSearchString = "(HierAttribID = " + ItemAttributeValue + " and HierAttribValue >= " + StartValue + " and HierAttribValue <= " + EndValue + ")";
                }
            }
            else {
                alert("Please select Exact or Range");
                newSearchString = "NotAllow";
            }		
    }
}
    return newSearchString;
}
  
function CheckItemAttrb(AttributeId) {
    selectedAttribute = AttributeId.id;
    var pg = document.getElementById("productgroup").value;
    var ItemAttribVal1 = document.getElementById(selectedAttribute);
    var attrbValue = ItemAttribVal1.options[ItemAttribVal1.selectedIndex];
    xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=CheckItemAttrb&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel+ '&ItemAttrbValue=' + attrbValue.value + '&AttributeId=' + selectedAttribute, 'CheckItemAttrb', treeNodeSel );
}    

function findMatchesItemAttrb() {
    ValuesNotSelectedCnt = 0
    var ItemAttribLineResults = document.getElementById("linerresultsItemAttrib");
    ItemAttribLineResults.style.visibility='hidden';
    var pg = document.getElementById("productgroup").value; 
    var sType = 0;
    var ExactSelectCnt = 0;
    var allowfunction = true;
    var ItemAttrib1SearchCriteria = '';
    var ItemAttrib2SearchCriteria = '';
    var ItemAttrib3SearchCriteria = '';
    var Attr1Selection = '0';
    var Attr2Selection = '0';
    var Attr3Selection = '0';
    var ItemAttribVal1 =  document.getElementById("ItemAttrib1");
    if (ItemAttribVal1 != null) {
        var ItemAttribVal2 =  document.getElementById("ItemAttrib2");
        var ItemAttribVal3 =  document.getElementById("ItemAttrib3");
        var startTextRange1 =  document.getElementById("StartRange1");
        var startTextRange2 =  document.getElementById("StartRange2");
        var startTextRange3 =  document.getElementById("StartRange3");
        var endTextRange1 =  document.getElementById("EndRange1");
        var endTextRange2 =  document.getElementById("EndRange2");
        var endTextRange3 =  document.getElementById("EndRange3");
        var Attr1exact = document.getElementById("Rb1");
        var Attr2exact = document.getElementById("Rb2");
        var Attr3exact = document.getElementById("Rb3");
        var Attr1range = document.getElementById("RbRange1");
        var Attr2range = document.getElementById("RbRange2");
        var Attr3range = document.getElementById("RbRange3");

        if ((ItemAttribVal1.value == '-1' || ItemAttribVal1.value == '') && (ItemAttribVal2.value == '-1' || ItemAttribVal2.value == '') && (ItemAttribVal3.value == '-1' || ItemAttribVal3.value == '')) {
            alert("Please select any one of the item attributes");
        }
        else {
            if (ItemAttribVal1.value != '-1' && ItemAttribVal1.value != '' && ItemAttribVal2.value != '-1' && ItemAttribVal2.value != '' && ItemAttribVal1.value == ItemAttribVal2.value) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.Sameitemattrbsearch", LanguageID))%>');
                allowfunction = false;
            }
            if (allowfunction == true && ItemAttribVal1.value != '-1' && ItemAttribVal1.value != '' && ItemAttribVal3.value != '-1' && ItemAttribVal3.value != '' && ItemAttribVal1.value == ItemAttribVal3.value) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.Sameitemattrbsearch", LanguageID))%>');
            allowfunction = false;
        }  
        if (allowfunction == true && ItemAttribVal2.value != '-1' && ItemAttribVal2.value != '' && ItemAttribVal3.value != '-1' && ItemAttribVal3.value != '' && ItemAttribVal2.value == ItemAttribVal3.value) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.Sameitemattrbsearch", LanguageID))%>');
            allowfunction = false;
        }  
        if (allowfunction == true) {
            if (Attr1exact.checked == true) {
                Attr1Selection = '1';
                if (ItemAttribVal1.value != '' && ItemAttribVal1.value != '-1' )
                    ExactSelectCnt = ExactSelectCnt + 1;
                ItemAttrib1SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal1.value, 'exact', startTextRange1.value, endTextRange1.value)
            } 
            else if (Attr1range.checked == true)	
                ItemAttrib1SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal1.value, 'range', startTextRange1.value, endTextRange1.value)
            else
                ItemAttrib1SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal1.value, '', startTextRange1.value, endTextRange1.value)

            if (ItemAttrib1SearchCriteria !=  'NotAllow') {
                if (Attr2exact.checked == true) {  
                    Attr2Selection = '1';
                    if (ItemAttribVal2.value != '' && ItemAttribVal2.value != '-1' )
                        ExactSelectCnt = ExactSelectCnt + 1;
                    ItemAttrib2SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal2.value, 'exact', startTextRange2.value, endTextRange2.value)
                }
                else if (Attr2range.checked == true)	
                    ItemAttrib2SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal2.value, 'range', startTextRange2.value, endTextRange2.value)
                else
                    ItemAttrib2SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal2.value, '', startTextRange2.value, endTextRange2.value)

                if (ItemAttrib2SearchCriteria !=  'NotAllow') {
                    if (Attr3exact.checked == true) {
                        Attr3Selection = '1';
                        if (ItemAttribVal3.value != '' && ItemAttribVal3.value != '-1' )
                            ExactSelectCnt = ExactSelectCnt + 1;
                        ItemAttrib3SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal3.value, 'exact', startTextRange3.value, endTextRange3.value)
                    }
                    else if (Attr3range.checked == true)	
                        ItemAttrib3SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal3.value, 'range', startTextRange3.value, endTextRange3.value)
                    else
                        ItemAttrib3SearchCriteria = ItemAttributeSearchCriteria(ItemAttribVal3.value, '', startTextRange3.value, endTextRange3.value)
                } 				
            } 			
            if (ItemAttrib1SearchCriteria != 'NotAllow' && ItemAttrib2SearchCriteria != 'NotAllow' && ItemAttrib3SearchCriteria != 'NotAllow') {
                if (ItemAttrib1SearchCriteria == '' && ItemAttrib2SearchCriteria == '' && ItemAttrib3SearchCriteria == '') {  
                    alert("Please select any one of the item attributes");
                } 
                else {   
                    if (ExactSelectCnt > 2) {
                        alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.invaliditemattrbexactselection", LanguageID))%>');
                    }
                    else { 
                        //alert(ItemAttrib1SearchCriteria);  
                        if (ValuesNotSelectedCnt >= 3) {
                            //alert("ValuesNotSelectedCnt :: " + ValuesNotSelectedCnt); 
                            ValuesNotSelectedCnt = 0;
                            //alert("Invalid start/End value.");
                            alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.invaliditemattrbrangeselection", LanguageID))%>');
                  }
                  else {
                      ValuesNotSelectedCnt = 0;
                      //alert("ItemAttrb1Search:: " + ItemAttrib1SearchCriteria);
                      //alert("ItemAttrb2Search:: " + ItemAttrib2SearchCriteria);
                      //alert("ItemAttrb3Search:: " + ItemAttrib3SearchCriteria);
                      //xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=findMatchesItemAttrb&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel+ '&ItemAttrb1Search=' + ItemAttrib1SearchCriteria + '&ItemAttrb2Search=' + ItemAttrib2SearchCriteria + '&ItemAttrb3Search=' + ItemAttrib3SearchCriteria, 'findMatchesItemAttrb', treeNodeSel );
                      // alert('action=findMatchesItemAttrb&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel+ '&ItemAttribVal1='+ ItemAttribVal1.value + '&Attr1Selection='+ Attr1Selection + '&low1Value='+ startTextRange1.value + '&high1Value='+ endTextRange1.value + '&ItemAttribVal2='+ ItemAttribVal2.value + '&Attr2Selection='+ Attr2Selection + '&low2Value='+ startTextRange2.value + '&high2Value='+ endTextRange2.value + '&ItemAttribVal3='+ ItemAttribVal3.value + '&Attr3Selection='+ Attr3Selection + '&low3Value='+ startTextRange3.value + '&high3Value='+ endTextRange3.value );
                      xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=findMatchesItemAttrb&pg=' + pg + '&search=1&stype=1&nodeid=' + treeNodeSel + '&hierid=' + hierSel+ '&ItemAttribVal1='+ ItemAttribVal1.value + '&Attr1Selection='+ Attr1Selection + '&low1Value='+ startTextRange1.value + '&high1Value='+ endTextRange1.value + '&ItemAttribVal2='+ ItemAttribVal2.value + '&Attr2Selection='+ Attr2Selection + '&low2Value='+ startTextRange2.value + '&high2Value='+ endTextRange2.value + '&ItemAttribVal3='+ ItemAttribVal3.value + '&Attr3Selection='+ Attr3Selection + '&low3Value='+ startTextRange3.value + '&high3Value='+ endTextRange3.value , 'findMatchesItemAttrb', treeNodeSel );
                  }       
              }
          }
      } 			
  }	
}
}
}

function LinkMatchestoHierarchy(){
    xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=LinkMatchestoHierarchy','LinkMatchestoHierarchy', treeNodeSel);
}

function RemoveMatchesFromHierarchy(){
    xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=RemoveMatchesFromHierarchy','RemoveMatchesFromHierarchy', treeNodeSel);
}
function htmlEntities(str){
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
  
function handleFindKeyDown(e) {
    var key = e.which ? e.which : e.keyCode;
    
    if (key == 13) {
        findMatches();
    }
}
  
function locateItem(selValue) {
    var pg = document.getElementById("productgroup").value;
    var locateHierarchyQueryString = "&selected=" + selValue + "&Linking=" + '<% Sendb(IIf(InLinkMode, "1", "0"))%>';
    var PABQueryString = "<%= Request.Url.Query %>";
    if (selValue != '') {
        cancelRefresh = true;
        if(PAB == 1)
        {         
            //      debugger; 
            $("#hdnLocateHierarchyURL").val(locateHierarchyQueryString);
            $("#btnLocateHierarchyTree").click();
            //          PABQueryString = removeExistingHierarchyParams(PABQueryString);
            //          PABQueryString = VerifyQueryStringSyntax(PABQueryString);
            //          document.location = PABPath + PABQueryString + "PABStage=1" + "&LocateHierarchyURL=" + encodeURIComponent(locateHierarchyQueryString);
        }
        else
        {
            document.location = '/logix/phierarchytree.aspx?selected=' + selValue + '&ProductGroupID=' + pg + '&Linking=<%Sendb(IIf(InLinkMode, "1", "0"))%>';
          }
      }
  }
  function VerifyQueryStringSyntax(QueryString)
  {
      //var tempQueryString = removeExistingHierarchyParams(QueryString);
      //alert(PABPath + PABQueryString + "PABStage=1" + "&LocateHierarchyURL=" + encodeURIComponent(locateHierarchyQueryString));
      if(QueryString != "")
          QueryString = QueryString + "&";
      else
          QueryString = "?";

      return QueryString;
  }
  function removeExistingHierarchyParams(PABQueryString)
  {
      var splitArray = PABQueryString.split("&");
      var indexPABStage = splitArray.indexOf("PABStage=1");
      if(indexPABStage > -1)
          splitArray.splice(indexPABStage,1);
    
      indexPABStage = splitArray.indexOf("PABStage=2");
      if(indexPABStage > -1)
          splitArray.splice(indexPABStage,1);

      var indexLocHier = -1;
      $.each(splitArray, function(index, value){
          if(value.indexOf("LocateHierarchyURL") > -1)
              indexLocHier = index;
      })
      if(indexLocHier > -1)
          splitArray.splice(indexLocHier);

      return splitArray.join("&");
  }



  function handleSearchItemAdjust(id, linkID, hID, extID) {
      var pg = document.getElementById("productgroup").value;
      var linkID = document.getElementById(linkID);
      var type = '';
      var pab2=<%Sendb(PAB)%>;
    var PAB1=(pab2 == "0" || pab2 == null)?"0":"1";
    if (linkID != null) {
        type = (linkID.innerHTML=='Add') ? 'add' : 'remove';
      
        var qryStr = 'action=handleSearchItemAdjust&pg=' + pg + '&productID=' + id + '&type=' + type + '&hID=' + hID + '&extID=' + extID+'&PAB='+PAB1;
      
        if (id != null) {
            xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'handleSearchItemAdjust', id + 'H' + hID );
        }
    }
}
  
function handleSearchNodeAdjust(id, linkID, hID, type) {
    var pg = document.getElementById("productgroup").value;
    var linkID = document.getElementById(linkID);
    
    if (id != null) {
        var qryStr = 'action=handleSearchNodeAdjust&pg=' + pg + '&nodeID=' + id + '&type=' + type + '&hID=' + hID + '&Linking=<%Sendb(IIf(InLinkMode, "1", "0"))%>';
        xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'handleSearchNodeAdjust', id + 'H' + hID );
    }
}
  
function select_Click() {
    var selValue = getSelectedValue();
    var pg = document.getElementById("productgroup").value;
    
    if (selValue != '') {
        cancelRefresh = true;
        document.location = 'phierarchytree.aspx?selected=' + selValue + '&ProductGroupID=' + pg + '&Linking=<%Sendb(IIf(InLinkMode, "1", "0"))%>';
    }
}
  
function getSelectedValue() {
    var elemOpt = document.getElementsByName("selected");
    var selectedValue = '';
    
    if (elemOpt != null) {
        if (isArray(elemOpt)) {
            for (var i=0; i < elemOpt.length; i++) {
                if (elemOpt[i].checked) {
                    selectedValue = elemOpt[i].value;
                    break;
                }
            }
        } else {
            selectedValue = elemOpt.value;
        }
        if (selectedValue == '') {
            alert('<% Sendb(Copient.PhraseLib.Lookup("lhierarchy.selectfromlist", LanguageID))%>');
        } 
    } else {
        alert('<% Sendb(Copient.PhraseLib.Lookup("phierarchy.NoItems", LanguageID))%>');
    }
    return selectedValue;            
}    
  
function isArray(obj) {
    return(typeof(obj.length)=="undefined") ? false:true;
}
  
function toggleDropdown() {
    if (document.getElementById("actionsmenu") != null) {
        bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
        if (bOpen) {
            document.getElementById("actionsmenu").style.visibility = 'visible';
            if(typeof document.mainform != "undefined"){
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▲';
            }
        } else {
            document.getElementById("actionsmenu").style.visibility = 'hidden';
            if(typeof document.mainform != "undefined"){
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▼';
        }
    }
}
}
  
function ChangeParentDocument() {
    var pab2=<%Sendb(PAB)%>;
    var PAB1=(pab2 == "0" || pab2 == null)?"0":"1";
    var IncentiveProductGroup = document.getElementById('IncentiveProductGroupID');
    var PABQueryString = "<%= Request.Url.Query %>";
      if (opener != null && !cancelRefresh && PAB1 != "1") {
          opener.location = '/logix/pgroup-edit.aspx?ProductGroupID=<%Sendb(pgID)%>&OfferID=<%Sendb(OfferID)%>&EngineID=<%Sendb(EngineID)%>' 
    }
    if(PAB1 == "1" )
    {
        if(IncentiveProductGroup != null){
            opener.location="/logix/UE/UEoffer-con.aspx?OfferID=<%Sendb(OfferID)%>";
        }
        else 
        { 
            opener.location="/logix/UE/UEoffer-rew.aspx?OfferID=<%Sendb(OfferID)%>";
        }
    }    
}
var progress=null;
  
function updateIdList(elemChkBox,nodeid,nodename) {
    var id = "";
    var isChecked = 0;
    $("#save").removeAttr("disabled");
    var lblProductcount= document.getElementById('lblProductsCount');
    if (elemChkBox != null) {
        id = elemChkBox.value;
        if (elemChkBox.checked) {
            isChecked=1;
            idList += id + ","
            nodeidlist += nodeid + ",";
            selectednodenamesList += nodename + ",";

        } else {
            var elemAll = document.getElementById("chkAll");
            elemAll.checked=elemChkBox.checked;
            idList = idList.replace(id + ",", "");
            nodeidlist = nodeidlist.replace(nodeid + ",", "");
            selectednodenamesList = selectednodenamesList.replace(nodename + ",", "");
        }
    }
    if(idList !="")
    {
        $('#pcount').show();
    }
    else{
        $('#pcount').hide();
    }
    var strurl = "/logix/HierarchyFeeds.aspx?action=getCount"
    if(progress)
    {
        progress.abort();
    }
    progress= $.ajax({
        type: 'POST',
        url: strurl,
        data:{nodeidl : nodeidlist},
        beforeSend : function(){
            $('#lblProductsCount').css("color","");
            $('#lblProductsCount').hide();
            $('#Img1').show();
            $('#warning').hide();
            $('#warning').css("color","");
            $('#contains').css("color","");
            $('#products').css("color","");
              
        }
    })
    .always(function(){
        $('#Img1').hide();
    })
   .done(function (data) { 
       if(data == "0\r\n")
       {
           $('#warning').css("color","red");
           $('#warning').show();
           $('#contains').css("color","red");
           $('#products').css("color","red");
           $('#lblProductsCount').css("color","red");
           $('#lblProductsCount').show();
           $('#lblProductsCount').html(data);
       }
       else
       {
           $('#contains').css("color","");
           $('#products').css("color","");
           $('#warning').hide();
           $('#lblProductsCount').show();
           $('#lblProductsCount').html(data);
       }
       progress=null;
   });

    if (<%Sendb(DispSelectedCount)%> != 0){
        if(id.indexOf("I") == 0){
            update_statusfooter(nodeidlist.toString(), '<%Sendb(pgID)%>', isChecked.toString(), '1', '0');
               } else{
                   update_statusfooter(nodeidlist.toString(), '<%Sendb(pgID)%>', isChecked.toString(), '0', '0');
               }
           }

       }
  
       function toggleHierarchyBox(id) {
           var elem = document.getElementById(id);
    
           if (elem != null) {
               if (elem.style.display != 'block') {
                   bLoading = true;
                   showWait();
                   elem.style.display = 'block';
               } else {
                   elem.style.display = 'none';
                   bLoading = false;
                   hideWait();
               }
           }
       }
  
       function handleAllItems(level) {
           var elem = null;
           var i = 0;
           var elemAll = document.getElementById("chkAll");
           var pgID = document.getElementById("prodgroupID").value;
           var isChecked = 0;
           if (elemAll != null) {
               if(PrevPageCount==0){
                   idList = "";
                   nodeidlist = "";
                   selectednodenamesList = '';
               }
               elem = document.getElementById("chk" + i);
               while (elem != null) {
                   if(elem.disabled != true)
                   {
                       elem.checked = elemAll.checked;
                   }
                   if (elem.checked) { 
                       isChecked = 1;
                       var id = "";
                       id = elem.value;
                       nodeid=id.substring(1);
                       if(idList.indexOf(id + ",")<0)
                           idList += id + ","
                       if(nodeidlist.indexOf(nodeid + ",")<0)
                           nodeidlist += nodeid + ",";
                       if(selectednodenamesList.indexOf(elem.name.substring(3) + ",")<0)
                           selectednodenamesList += elem.name.substring(3) + ",";
                       //updateIdList(elem,elem.value.substring(1)); 
                   }
                   else
                   {
                       var id = "";
                       id = elem.value;
                       nodeid=id.substring(1);
                       idList = idList.replace(id + ",", "");
                       nodeidlist = nodeidlist.replace(nodeid + ",", "");
                       selectednodenamesList = selectednodenamesList.replace(elem.name.substring(3) + ",", "");
                   }
                   i++;
                   elem = document.getElementById("chk" + i);
               }
               if(idList !="")
               {
                   $('#pcount').show();
               }
               else{
                   $('#pcount').hide();
               }
           
               var strurl = "/logix/HierarchyFeeds.aspx?action=getCount"
               if(progress)
               {
                   progress.abort();
               }
               progress= $.ajax({
                   type: 'POST',
                   url: strurl,
                   data:{nodeidl : nodeidlist},
                   beforeSend : function(){
                       $('#lblProductsCount').css("color","");
                       $('#lblProductsCount').hide();
                       $('#Img1').show();
                       $('#warning').hide();
                       $('#warning').css("color","");
                       $('#contains').css("color","");
                       $('#products').css("color","");
                   }
               })
               .always(function(){
                   $('#Img1').hide();
               })
               .done(function (data) { 
                   if(data == "0\r\n")
                   {
                       $('#lblProductsCount').css("color","red");
                       $('#warning').css("color","red");
                       $('#contains').css("color","red");
                       $('#products').css("color","red");
                       $('#warning').show();
                       $('#lblProductsCount').show();
                       $('#lblProductsCount').html(data);
                   }
                   else
                   {
                       $('#warning').hide();
                       $('#lblProductsCount').show();
                       $('#lblProductsCount').html(data);
                   }
                   progress=null;
               });
       

               if (<%Sendb(DispSelectedCount)%> != 0){
            //alert("DispSelectedCount="+ '<%Sendb(DispSelectedCount)%>');
            update_statusfooter(nodeidlist.toString(), '<%Sendb(pgID)%>', isChecked.toString(), '0', level.toString());
        }

    }
}
  
function deleteFromHierarchy() {
    var okToDelete = false;
    var msg = '';
    var nodeID = '';
    var itemID = '';
    var isHier = 0;
    var tdElems = null;
    var tokenValues = [];

    if (listItemSel > -1 && treeNodeSel != 0) {
        // delete the item selected in the right pane
        var elem = document.getElementById("PKID" + listItemSel);
        if (elem != null) {
            nodeID = treeNodeSel;
            itemID = elem.value;
            var elemListItem = document.getElementById("itemRow" + listItemSel);
            if (elemListItem != null && elemListItem.cells != null && elemListItem.cells.length >= 3) {
                msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmDeleteID", LanguageID))%>: ' + elemListItem.cells[1].innerHTML + " " + elemListItem.cells[2].innerHTML;         
            } else {
                msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmItemDelete", LanguageID))%>';
            }
        } else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ErrorOnItemDelete", LanguageID))%>');
        }
    } else if (treeNodeSel > 0) {
        // delete the folder selected in the left pane 
        nodeID = treeNodeSel;
        itemID = ''
        var elemTree = document.getElementById("name" + treeNodeSel);
        if (elemTree == null) {
            elemTree = document.getElementById("nameH" + treeNodeSel);
            isHier = 1;
        }
        if (elemTree != null) {
            msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmDelete", LanguageID))%>';
          tokenValues[0] = cleanCellText(elemTree.innerHTML);
          msg = detokenizeString(msg, tokenValues);
      } else {
          msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmDeleteItem", LanguageID))%>';
      }
  } else {
      alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.NoDeletableItem", LanguageID))%>');
  }
    
    if (msg != '') {
        okToDelete = confirm(msg);
        if (okToDelete) {
            var qryStr = 'action=delFromHierarchy&nodeID=' + nodeID + '&itemID=' + itemID + "&isHier=" + isHier;
            xmlhttpPost('/logix/HierarchyFeeds.aspx', qryStr, 'delFromHierarchy', itemID );
        }
    }
}
  
function cleanCellText(cellText) {
    var newCellText = cellText;
    
    newCellText = cellText.replace('&nbsp;', '');
    newCellText = newCellText.replace('&amp;', '&');
    
    return newCellText;
}
  
function updateDelFromHierarchy(response) {
    var id = '', parent = '';
    var refresh = '';
    
    refresh = parseTagValue(response, "refresh");
    id = parseTagValue(response, "node");
    
    if (id != null) {
        if (refresh == "2") {
            // remove the top-level from the tree
            var elemTop = document.getElementById("nodeH" + id);
            if (elemTop != null) {
                elemTop.style.display = "none";
                var elemList = document.getElementById("itemList");
                if (elemList != null) elemList.innerHTML = "";
            }
        } else if (refresh == "3") {
            var elemLeft = document.getElementById("node" + id);
            if (elemLeft != null) {
                elemLeft.style.display = "none";
                var elemList = document.getElementById("itemList");
                if (elemList != null) elemList.innerHTML = "";
                treeNodeSel = -1;
                listItemSel = -1;
            }
        } else if (refresh == "1") {
            // item refresh only
            highlightNode(id, levelSel, false);
        } else {
            // full refresh
            parent = parseTagValue(response, "parent");
            if (parent == 0 && treeNodeSel > -1 ) {
                parent = treeNodeSel;
            }
            expandNode(parent, levelSel, true);
        }
    }            
}
  
function handleLinkingRefresh(id, response) {
    var idToSend = id;
    var msg = '';
    var ct = 0;
    
    msg = parseTagValue(response, 'message');
    ct = parseTagValue(response, 'count');
    
    // assume success if no message is returned
    if (msg == '') {
        // update the status bar panel for assigned items
        updateAssignedItemCt(ct);
        ChangeParentDocument();
      
        // refresh both the folder and items panes
        if (id==-1) {
            idToSend = hierSel;
        }
      
        expandNode(idToSend, levelSel, true);
        return true;
    } else {
        // notify illegal operation
        alert(msg);
        return false;
    }
}
  
function updateLinkToGroup(id, response) {
    if (handleLinkingRefresh(id, response)) {
        updateFolderIcons(id, true);
    }
}
  
function generatedivforsigns(divHTML) {
    var elem = document.getElementById("divAssignSigns");
    
    if (elem != null) {
        elem.innerHTML = divHTML;
        toggleDialog('divAssignSigns', true);
    }
}
  
function updateFolderIcons(id, isLinked) {
    var elemFldr = document.getElementById('imgfldr' + id);
    var parentId = '';
    var elemParent = null;
    var prntIdStr = '';
    
    // traverse the parent nodes to change the icons are a link or delink of a node
    while (elemFldr != null) {
        elemFldr.src = (isLinked) ? '/images/folder-down.png' : '/images/folder.png';
        elemParent = document.getElementById('node' + id);
        
        if (elemParent != null && elemParent.parentNode != null) {
            prntIdStr = elemParent.parentNode.id;
            parentId = prntIdStr.replace('hId','')
        } else {
            parentId = '-1'
        }
        elemFldr = document.getElementById('imgfldr' + parentId);
        id = parentId;
    }
}
  
function updateAssignedItemCt(assignedCount) {
    var elemAssigned = document.getElementById('assignedCt');
    
    if (elemAssigned != null) {
        elemAssigned.innerHTML = assignedCount
    }
}
  
function updateRemoveLinkToGroup(id, response) {
    var linkChildCt = 0;
    
    if (handleLinkingRefresh(id, response)) {
        linkChildCt = parseTagValue(response, 'linked');
      
        // only update the parent folder icons if this is the last linked folder
        if (linkChildCt == 0) {
            updateFolderIcons(id, false);
        }
    }
}
  
function updateExcludeFromGroup(id, response) {
    handleLinkingRefresh(id, response);
}
  
function updateRemoveExclusion(id, response) {
    handleLinkingRefresh(id, response);
}
  
function confirmSuccess(response) {
    if (response.substring(0, 2) == 'OK') {
        toggleDialog('divAssignSigns', false);
    }
}
  
function parseTagValue(doc, tag) {
    var startPos = -1, endPos = -1;
    var value = '';
    
    if (doc != null && tag != null) {
        startPos = doc.indexOf("<" + tag + ">");
        if (startPos > -1 ) {
            endPos = doc.indexOf("</" + tag + ">");
            if (endPos > -1) {
                value = doc.substring(startPos + tag.length + 2, endPos);
            }
        }
    }
    return value;
}
  
function sortByColumn(colNbr, curSortCode) {
    var name = "";
    var pg = document.getElementById("productgroup").value;
    var itemPK = document.getElementById("itemSelectedPK").value;
    var PAB1= (PAB == "0")?"":"1"; 
    if (colNbr == 1 && curSortCode == 1) {
        curSortCode = 2;
    } else if (colNbr == 1 && curSortCode == 2) {
        curSortCode = 1;
    } else if (colNbr == 2 && curSortCode == 3) {
        curSortCode = 4;
    } else if (colNbr == 2 && curSortCode == 4) {
        curSortCode = 3;
    } else if (colNbr == 1) {
        curSortCode = 1;
    } else if (colNbr == 2) {
        curSortCode = 3;
    }
    
    masterSortCode = curSortCode;
    if (levelSel <= 1) {
        xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=showItems&node=' + hierSel + '&level=' + levelSel + '&nodeName=' + name + '&pg=' + pg + '&itemPK=' + itemPK +"&buyerid="+ <% Sendb(BuyerID)%> + '&sort=' + masterSortCode + '&StartIndex=1'+'&PAB='+PAB1+'&SelectedNodeIDs=<%Sendb(SelectedNodeIDs)%>', 'showItems', treeNodeSel);
    } else {
        xmlhttpPost('/logix/HierarchyFeeds.aspx', 'action=showItems&node=' + treeNodeSel + '&level=' + levelSel + '&nodeName=' + name +'&pg=' + pg + '&itemPK=' + itemPK +"&buyerid="+ <% Sendb(BuyerID)%> + '&sort=' + curSortCode + '&StartIndex=1'+'&PAB='+PAB1+'&SelectedNodeIDs=<%Sendb(SelectedNodeIDs)%>', 'showItems', treeNodeSel);
    }
}

function getSelectedFolderFullName() {
    var name = '<%Sendb(Copient.PhraseLib.Lookup("term.searching", LanguageID))%> '
    var elem = null
    
    if (treeNodeSel > 0) {
        elem = document.getElementById('name' + treeNodeSel);
        if (elem != null) {
            name += '<%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%>: ' + elem.innerHTML;
          } else {
              name += '<%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%>: [<%Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>]';
          }
      } else if (hierSel > 0) {
          elem = document.getElementById('nameH' + hierSel);
          if (elem != null) {
              name += '<%Sendb(Copient.PhraseLib.Lookup("term.hierarchy", LanguageID))%>: ' + elem.innerHTML;
      } else {
          name += '<%Sendb(Copient.PhraseLib.Lookup("term.hierarchy", LanguageID))%>: [<%Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>]';
      }
  } else {
      name += ' <%Sendb(Copient.PhraseLib.Lookup("term.AllHierarchies", LanguageID))%>'
  }
    
    return name;
}

function getSelectedFolderName() {
    var name = ' | <%Sendb(Copient.PhraseLib.Lookup("term.searching", LanguageID))%> '
    var elem = null
    
    if (treeNodeSel > 0) {
        elem = document.getElementById('name' + treeNodeSel);
        if (elem != null) {
            name += '<%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%>: ' + elem.innerHTML;//substring(0,20);
          } else {
              name += '<%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%>: [<%Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>]';
          }
      } else if (hierSel > 0) {
          elem = document.getElementById('nameH' + hierSel);
          if (elem != null) {
              name += '<%Sendb(Copient.PhraseLib.Lookup("term.hierarchy", LanguageID))%>: ' + elem.innerHTML;//.substring(0,20);
      } else {
          name += '<%Sendb(Copient.PhraseLib.Lookup("term.hierarchy", LanguageID))%>: [<%Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>]';
      }
  } else {
      name += ' <%Sendb(Copient.PhraseLib.Lookup("term.AllHierarchies", LanguageID))%>'
  }
    
    return name;
}
  
function hasItemSelected(idList) {
    var re = new RegExp("I[0-9]*");
    
    // check if there is a match indicating that an item is selected.
    var m = re.exec(idList);
    return (m != null)
}
  
function toggleDialog(elemName, shown) {
    var elem = document.getElementById(elemName);
    var fadeElem = document.getElementById('phfadeDiv');
    
    if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
    }
    if (fadeElem != null) {
        fadeElem.style.display = (shown) ? 'block' : 'none';
    }
}
  
function pageResults(nodeid, level, offset) {
    if($('#lblProductsCount')!=null)
    {
        if(parseInt($('#lblProductsCount').innerHTML)!=0)
        {
            PrevPageCount=parseInt($('#lblProductsCount').innerHTML)
        }
    }
    if (nodeid != LastNodeID) {
        StartIndex = 1;
    }
    
    if (offset == -9999999) { //go to beginning
        StartIndex = 1;
    } else if (offset == 9999999) { //go to end
      
    } else { //go forward or back by the offset
        StartIndex = (StartIndex + offset);
        if (StartIndex < 1) {
            StartIndex = 1;
        }
    }
    showNodeItems(nodeid, level, false);
    LastNodeID = nodeid;
}
</script>
<%
    Send_HeadEnd()
    If popupFlag = "0" Then 'When ProductGroup page is executing this page inside it
        Send_BodyBegin(1)
    Else
        Send_BodyBegin(IIf(CreatedFromOffer, 3, 2))
    End If
  
    If (Not AutoAccessProductGroup AndAlso Logix.UserRoles.EditSystemConfiguration = False) Then
        Send_Denied(1, "perm.admin-configuration")
        GoTo done
    End If
%>
<div>
    <input type="hidden" id="productgroup" name="productgroup" value="<%Sendb(pgID)%>" />
    <input type="hidden" id="itemSelectedPK" name="itemSelectedPK" value="<%Sendb(itemSelectedPK)%>" />
    <input type="hidden" id="searchPathIDs" name="searchPathIDs" value="<%Sendb(searchHierarchyPath)%>" />
    <input type="hidden" id="selNode" name="selNode" value="<%Sendb(searchSelNodeID)%>" />
</div>

<%  If (PAB = 1) Then%>
<div class="greybox" <%Sendb(IIf(InLinkMode, " style=""background-image:url('/images/box-important.png');""", ""))%>
    id="toolbar">
    <div class="greyboxwrap panel_header">
        <h2>
            <%
                If InLinkMode Then
                    Sendb("<span style=""float: left;color: white;"">")
                    Sendb(Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID))
                    Sendb(" (" & Copient.PhraseLib.Lookup("term.linking", LanguageID).ToLower & ")")
                    Send("</span>")
                Else
                    Sendb("<span style=""float: left;"">")
                    Sendb(Copient.PhraseLib.Lookup("term.hierarchyselection", LanguageID))
                    Send("</span>")
                End If
            %>
        </h2>
        <br clear="left" />
    </div>
</div>
<% Else%>
<div class="box" <%Sendb(IIf(InLinkMode, " style=""background-image:url('/images/box-important.png');""", ""))%>
    id="toolbar">
    <h2>
        <%
            If InLinkMode Then
                Sendb("<span style=""float: left;color: white;"">")
                Sendb(Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID))
                Sendb(" (" & Copient.PhraseLib.Lookup("term.linking", LanguageID).ToLower & ")")
                Send("</span>")
            Else
                Sendb("<span style=""float: left;"">")
                Sendb(Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID))
                Sendb(" (" & Copient.PhraseLib.Lookup("term.selection", LanguageID).ToLower & ")")
                Send("</span>")
            End If
        %>
    </h2>
    <br clear="left" />
</div>
<% End If%>



<div id="hmain" style="background: url('');">
    <div class="column" id="leftpane" style="padding-top: 5px; border-bottom: 1px solid black; float: none; height: 30%; min-height: 30%; width: 100%;"
        enableviewstate="true">
        <div class="liner" id="topNodesDiv">
            <%
                'Create a static, top-level product hierarchies container node
                Send(vbTab & "<span class=""hrow"" id=""nodeH0"">")
                Send(vbTab & " <span id=""indentH0"" style=""line-height:18px;left:0px;"" onclick=""highlightNode(0,0,false)"">")
                Send(vbTab & "  <span><img src=""/images/hierarchy.png"" border=""0"" alt="""" /></span>")
                Send(vbTab & "  <span id=""nameH0"">&nbsp;" & Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID) & "</span>")
                Send(vbTab & "  <input type=""hidden"" id=""hierarchyH0"" name=""hierarchyH0"" value=""0"" /><br />")
                Send(vbTab & " </span>")
        
                'Check if this is a search result
                If (Not searchPathIDs Is Nothing AndAlso searchPathIDs.Length >= 1) Then
                    Send(vbTab & " <span id=""hIdH0"" style=""display:inline;"">")
                    Dim SearchID As Integer
                    If (searchPathIDs.Length > 1) Then
                        SearchID = searchPathIDs(1)
                    Else
                        SearchID = searchPathIDs(0)
                    End If
                    GenerateDrilledNodeDiv(SearchID, 1, searchPathIDs, pgID)
                    Send(vbTab & " </span>")
                Else
                    Send(vbTab & " <span id=""hIdH0"" style=""display:none;"">")
                    Send(vbTab & " </span>")
                End If
                Send(vbTab & "</span>")
            %>
        </div>
    </div>
    <div class="column" id="rightpane" style="float: none; height: 70%; min-height: 70%; width: 100%;">
        <div class="liner" id="itemList">
        </div>
    </div>
    <div class="searchresults" id="SearchDiv" style="width: 100%;">
        <div id="linerresults">
            <br clear="left" />
            <%
                Send("<table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """ style=""width:100%;"">")
                Send("  <thead>")
                Send("    <tr>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.select", LanguageID) & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.action", LanguageID) & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.level", LanguageID) & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.hierarchy", LanguageID) & "</th>")
                'Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.Product", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "&nbsp;&nbsp;" & Copient.PhraseLib.Lookup("term.Name", LanguageID) & ")" & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.Node", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "&nbsp;&nbsp;" & Copient.PhraseLib.Lookup("term.Name", LanguageID) & ")" & "</th>")
                Send("    </tr>")
                Send("  </thead>")
                Send("  <tbody>")
                Send("    <tr>")
                Send("      <td colspan=""6""></td>")
                Send("    </tr>")
                Send("  </tbody>")
                Send("</table>")
            %>
        </div>
    </div>
    <div class="searchresults" id="SearchProdItemAttrbDiv">
        <table class="list" style="width: 640px">
            <tr>
                <td>
                    <table>
                        <tr>
                            <td style="width: 50px"></td>
                            <td style="width: 130px">
                                <label id="lblSelAttrb" text="Label">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.SelectAttribute", LanguageID))%></label>
                            </td>
                            <td colspan="4" style="width: 340px">
                                <label id="lblExact" text="Label">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%></label>&nbsp;
                                <label id="lblRange" text="Label">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.range", LanguageID))%></label>&nbsp;&nbsp;&nbsp;
                                <label id="lblStart" text="Label">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.start", LanguageID))%></label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <label id="lblEnd" text="Label">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.end", LanguageID))%></label>
                            </td>
                            <td style="width: 25px"></td>
                            <td style="width: 25px"></td>
                        </tr>
                        <tr>
                            <td style="width: 50px;">
                                <label for="lblAttribute1">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.Attribute1", LanguageID))%></label>
                            </td>
                            <td style="width: 130px;">
                                <select id="ItemAttrib1" onchange="CheckItemAttrb(this)" style="width: 120px">
                                    <%
                                        MyCommon.QueryStr = "Select HierAttribID,HierAttribDescription from HierAttribDefinition order by HierAttribID"
                                        rstItemAttrDetails = MyCommon.LRT_Select
                                        Send("<option value=""-1""></option>")
                                        For Each rowItemAttrDetails In rstItemAttrDetails.Rows
                                            Send("<option value=""" & rowItemAttrDetails.Item("HierAttribID") & """>" & rowItemAttrDetails.Item("HierAttribDescription") & " </option>")
                                        Next
                                    %>
                                </select>
                            </td>
                            <td colspan="4" style="width: 330px">
                                <div id="ItemAttrbDiv1StartEndValues" style="width: 330px">
                                    <%
                                        Send("<table class=""list""><thead></thead><tbody><tr>")
                                        Send("<td style=""width: 30px;""><input id=""Rb1"" type=""radio"" name=""Rb1""/></td>")
                                        Send("<td style=""width: 40px;""><input id=""RbRange1"" type=""radio"" name=""Rb1"" checked=""CHECKED""/></td>")
                                        Send("<td style=""width: 120px;"">")
                                        Send("<select name=""StartRange1"" id=""StartRange1"" style=""width: 120px;"">")
                                        Send("<option value=""""></option>")
                                        Send("</select>")
                                        Send("</td>")
                                        Send("<td style=""width: 120px;"">")
                                        Send("<select name=""EndRange1"" id=""EndRange1"" style=""width: 120px;"">")
                                        Send("<option value=""""></option>")
                                        Send("</select></td></tr></tbody></table>")
                                    %>
                                </div>
                            </td>
                            <td style="width: 25px;"></td>
                            <td style="width: 25px;"></td>
                        </tr>
                        <tr>
                            <td style="width: 50px;">
                                <label for="lblAttribute2">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.Attribute2", LanguageID))%></label>
                            </td>
                            <td style="width: 130px;">
                                <select id="ItemAttrib2" onchange="CheckItemAttrb(this)" style="width: 120px">
                                    <%
                                        Send("<option value=""-1""></option>")
                                        For Each rowItemAttrDetails In rstItemAttrDetails.Rows
                                            Send("<option value=""" & rowItemAttrDetails.Item("HierAttribID") & """>" & rowItemAttrDetails.Item("HierAttribDescription") & " </option>")
                                        Next
                                    %>
                                </select>
                            </td>
                            <td colspan="4" style="width: 330px">
                                <div id="ItemAttrbDiv2StartEndValues" style="width: 330px">
                                    <%
                                        Send("<table class=""list""><thead></thead><tbody><tr>")
                                        Send("<td style=""width: 30px;""><input id=""Rb2"" type=""radio"" name=""Rb2""/></td>")
                                        Send("<td style=""width: 40px;""><input id=""RbRange2"" type=""radio"" name=""Rb2"" checked=""CHECKED""/></td>")
                                        Send("<td style=""width: 120px;"">")
                                        Send("<select name=""StartRange2"" id=""StartRange2"" style=""width: 120px;"">")
                                        Send("<option value=""""></option>")
                                        Send("</select>")
                                        Send("</td>")
                                        Send("<td style=""width: 120px;"">")
                                        Send("<select name=""EndRange2"" id=""EndRange2"" style=""width: 120px;"">")
                                        Send("<option value=""""></option>")
                                        Send("</select></td></tr></tbody></table>")
                                    %>
                                </div>
                            </td>
                            <td style="width: 25px;"></td>
                            <td style="width: 25px;"></td>
                        </tr>
                        <tr>
                            <td style="width: 50px;">
                                <label for="lblAttribute3">
                                    <% Sendb(Copient.PhraseLib.Lookup("term.Attribute3", LanguageID))%></label>
                            </td>
                            <td style="width: 130px;">
                                <select id="ItemAttrib3" onchange="CheckItemAttrb(this)" style="width: 120px">
                                    <%
                                        Send("<option value=""-1""></option>")
                                        For Each rowItemAttrDetails In rstItemAttrDetails.Rows
                                            Send("<option value=""" & rowItemAttrDetails.Item("HierAttribID") & """>" & rowItemAttrDetails.Item("HierAttribDescription") & " </option>")
                                        Next
                                    %>
                                </select>
                            </td>
                            <td colspan="4" style="width: 330px">
                                <div id="ItemAttrbDiv3StartEndValues" style="width: 330px">
                                    <%
                                        Send("<table class=""list""><thead></thead><tbody><tr>")
                                        Send("<td style=""width: 30px;""><input id=""Rb3"" type=""radio"" name=""Rb3""/></td>")
                                        Send("<td style=""width: 40px;""><input id=""RbRange3"" type=""radio"" name=""Rb3"" checked=""CHECKED""/></td>")
                                        Send("<td style=""width: 120px;"">")
                                        Send("<select name=""StartRange3"" id=""StartRange3"" style=""width: 120px;"">")
                                        Send("<option value=""""></option>")
                                        Send("</select>")
                                        Send("</td>")
                                        Send("<td style=""width: 120px;"">")
                                        Send("<select name=""EndRange3"" id=""EndRange3"" style=""width: 120px;"">")
                                        Send("<option value=""""></option>")
                                        Send("</select></td></tr></tbody></table>")
                                    %>
                                </div>
                            </td>
                            <td style="width: 25px;" align="right">
                                <input type="button" id="finditemattrb" name="finditemattrb1" value="<% Sendb(Copient.PhraseLib.Lookup("term.find", LanguageID))%>"
                                    onclick="findMatchesItemAttrb();" />
                            </td>
                            <td style="width: 25px;">
                                <input type="button" id="closeitemattrb" name="closeitemattrb1" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
                                    onclick="closeSearchItemAttrb();" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="width: 100%">
                    <span id="ItemAttrbsearchFolder" style="color: Gray;"></span>
                </td>
            </tr>
        </table>
        <div id="linerresultsItemAttrib" style="visibility: hidden">
            <br clear="left" />
            <table style="width: 98%; border-right: black 1pt solid; border-top: black 1pt solid; border-left: black 1pt solid; border-bottom: black 1pt solid;"
                cellpadding="0"
                cellspacing="0">
                <tr>
                    <td align="right">
                        <input id="btnAttrSearchLinktoHierarchy" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.LinkToGroup", LanguageID))%>"
                            style="width: 150px" onclick="LinkMatchestoHierarchy();" disabled="disabled" />
                        <input id="btnAttrSearchExcludefromHierarchy" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.RemoveLinkToGroup", LanguageID))%>"
                            style="width: 150px" onclick="RemoveMatchesFromHierarchy();" disabled="disabled" />
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <label id="lblmatches">
                            <% Sendb(Copient.PhraseLib.Lookup("term.AttributeSearchResults", LanguageID))%>:</label><b><span
                                id="lblAttrSearchResultCnt" style="color: RED;"></span>
                    </td>
                    </b> </td>
                </tr>
            </table>
        </div>
    </div>
</div>
<div id="searchbox">
    <input type="text" class="medium" id="searchString" name="searchString" maxlength="120"
        value="<% Sendb(SearchString)%>" onkeydown="handleFindKeyDown(event);" />
    <select id="searchType" name="searchType">
        <option value="0">
            <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
        </option>
        <option value="1">
            <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
        </option>
        <option value="2">
            <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
        </option>
    </select>
    <input type="button" id="find" name="find" value="<% Sendb(Copient.PhraseLib.Lookup("term.find", LanguageID))%>"
        onclick="findMatches();" />
    <input type="button" id="close" name="close" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID))%>"
        onclick="closeSearch();" />
    <span id="searchFolder" style="color: Gray; white-space: pre-line; table-layout: fixed; word-break: break-all;"></span>
</div>
<%  If PAB = 0 Then%>
<div id="statusfooter">

    <script>

        function update_statusfooter(nodeList, pgID, elemChkBoxChecked, notProdGroup, level)
        {
            var assignedCt = document.getElementById("assignedCt").innerHTML.toString();
            var lastSelectedCt = document.getElementById("selectedNotAssignedCt").value.toString();
            var strurl = "/logix/HierarchyStatusfooter.aspx?prodgroupid=" + pgID + "&assignedCt=" + assignedCt + "&nodelist=" + nodeList + "&lastSelectedCt=" + lastSelectedCt + "&notProdGroup=" + notProdGroup + "&chkboxChecked=" + elemChkBoxChecked.toString() + "&level=" + level.toString();
            //var strurl = "/logix/HierarchyStatusfooter.aspx?prodgroupid=" + pgID + "&assignedCt=" + assignedCt + "&nodelist=" + nodeList;
            //var strurl = "/logix/HierarchyStatusfooter.aspx?prodgroupid=" + pgID + "&assignedCt=" + assignedCt + "&nodelist=" + nodeList + "&lastSelectedCt=" + lastSelectedCt + "&chkboxChecked=" + elemChkBoxChecked.toString();
            //alert(strurl);
            if(nodeList.length >= 0){
                var xmlhttp = new XMLHttpRequest();
                xmlhttp.onreadystatechange = function() {
                    if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
                        document.getElementById("statusfooter").innerHTML = xmlhttp.responseText;
                    }
                }

                xmlhttp.open("GET", strurl, true);
                xmlhttp.send();
            }

        }

    </script>

    <%
        If (pgID > 0) Then
            MyCommon.QueryStr = "Select count(*) as AssignedCount from ProdGroupItems with (NoLock) where ProductGroupID=" & pgID & " and Deleted=0;"
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                AssignedCount = MyCommon.NZ(dt.Rows(0).Item("AssignedCount"), 0)
            End If

            'Send("<span>" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & ": <span id=""prodgroupID"">" & pgID & "</span></span>&nbsp;|&nbsp;<span>" & Copient.PhraseLib.Lookup("term.Assigned", LanguageID) & ": <span id=""assignedCt"">" & AssignedCount & "</span></span>&nbsp;|&nbsp;<span>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ": <span id=""selectedCt"">" & SelectedCount & "</span></span>")
            Send("<span>" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & ": <span id=""prodgroupID"">" & pgID & "</span></span>&nbsp;|&nbsp;<span>" & Copient.PhraseLib.Lookup("term.Assigned", LanguageID) & ": <span id=""assignedCt"">" & AssignedCount & "</span></span>")
            Send(" <input type=""hidden"" id=""selectedNotAssignedCt"" name=""selectedNotAssignedCt"" value=""0"" />")
 
        End If
    %>
    <!-- <span id="totalItemCt">&nbsp;</span><span id="searchItemCt"></span> -->
</div>
<%  End If%>
<div id="WaitDiv" style="">
</div>
<div id="deleteBox" class="modifyBox">
    <div style="width: 100%; background-color: Blue; color: White;">
        <b>Delete from Hierarchy</b>
    </div>
    <br />
    <center>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>">
            <tr>
                <td colspan="2">
                    <%Sendb(Copient.PhraseLib.Lookup("phierarchy.DeleteLevelMessage", LanguageID))%>
                </td>
            </tr>
            <tr style="height: 15px;">
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td style="padding-left: 20px;">
                    <%Sendb(Copient.PhraseLib.Lookup("term.level", LanguageID))%>:
                </td>
                <td>
                    <select id="delHierType" name="delHierType" class="mediumshort">
                        <option value="1">
                            <%Sendb(Copient.PhraseLib.Lookup("term.EntireHierarchy", LanguageID))%></option>
                        <option value="2">
                            <%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%></option>
                        <option value="3">
                            <%Sendb(Copient.PhraseLib.Lookup("term.product", LanguageID))%></option>
                    </select>
                </td>
            </tr>
            <tr>
                <td style="padding-left: 20px;">
                    ID:
                </td>
                <td>
                    <input type="text" class="mediumshort" id="delID" name="delID" value="" />
                </td>
            </tr>
            <tr style="height: 15px;">
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center;">
                    <input type="button" id="btnDelOK" name="btnDelOK" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>"
                        onclick="alert('ok clicked');" />
                    <input type="button" id="btnDelCancel" name="btnDelCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>"
                        onclick="toggleHierarchyBox('deleteBox');" />
                </td>
            </tr>
        </table>
    </center>
</div>
<div id="addBox" class="modifyBox">
    <div style="width: 100%; background-color: Blue; color: White;">
        <b>Add to Hierarchy</b>
    </div>
    <br />
    <center>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID))%>">
            <tr>
                <td colspan="2">
                    <%Sendb(Copient.PhraseLib.Lookup("phierarchy.SelectAddLevel", LanguageID))%>
                </td>
            </tr>
            <tr style="height: 15px;">
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td style="padding-left: 20px;">
                    Level:
                </td>
                <td>
                    <select id="addHierType" name="addHierType" class="mediumshort">
                        <option value="1">
                            <%Sendb(Copient.PhraseLib.Lookup("term.EntireHierarchy", LanguageID))%></option>
                        <option value="2">
                            <%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%></option>
                        <option value="3">
                            <%Sendb(Copient.PhraseLib.Lookup("term.product", LanguageID))%></option>
                    </select>
                </td>
            </tr>
            <tr>
                <td style="padding-left: 20px;">
                    Parent ID:
                </td>
                <td>
                    <input type="text" id="addParentID" name="addParentID" class="mediumshort" value="" />
                </td>
            </tr>
            <tr>
                <td style="padding-left: 20px;">
                    ID:
                </td>
                <td>
                    <input type="text" id="addID" name="addID" class="mediumshort" value="" />
                </td>
            </tr>
            <tr>
                <td style="padding-left: 20px;">
                    Desc:
                </td>
                <td>
                    <input type="text" id="addDesc" name="addDesc" class="mediumshort" value="" />
                </td>
            </tr>
            <tr style="height: 15px;">
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center;">
                    <input type="button" id="btnAddOK" name="btnAddOK" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>"
                        onclick="alert('ok clicked');" />
                    <input type="button" id="btnAddCancel" name="btnAddCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>"
                        onclick="toggleHierarchyBox('addBox');" />
                </td>
            </tr>
        </table>
    </center>
</div>
<div id="toolset" style="width: 300px">
    <%If PAB <> "1" And MyCommon.Fetch_SystemOption(136) = "1" Then%>
    <%
        Dim startHirRangeLevel As Integer = 0
        Dim EndHirRangeLevel As Integer = 0
        Try
            startHirRangeLevel = MyCommon.Fetch_SystemOption(137)
            EndHirRangeLevel = MyCommon.Fetch_SystemOption(139)
        Catch ex As Exception
        End Try
        If (pgID > 0 AndAlso InLinkMode) Then
            Send("<input type=""button"" class=""regular"" id=""btnItemAttrbSearch"" name=""btnItemAttrbSearch"" onclick=""searchItemAttrb(" & startHirRangeLevel & "," & EndHirRangeLevel & ");"" value=""" & Copient.PhraseLib.Lookup("term.AttributeSearch", LanguageID) & """ style=""width:110px"" />")
        Else
            Send("<input type=""button"" class=""regular"" id=""btnItemAttrbSearch"" name=""btnItemAttrbSearch"" onclick=""searchItemAttrb(" & startHirRangeLevel & "," & EndHirRangeLevel & ");"" value=""" & Copient.PhraseLib.Lookup("term.AttributeSearch", LanguageID) & """ style=""width:110px"" disabled=""disabled"" />")
        End If
    %>
    <%End If%>
    <input type="button" class="regular" id="btnSearch" name="btnSearch" onclick="search();"
        title="<% Sendb(Copient.PhraseLib.Lookup("hierarchy.search", LanguageID))%>"
        value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
    <% If PAB = "1" Then%>
    <input type="button" class="regular" id="btnContinueWithAttributes" onclick="ShoworHideDivs(true);RaiseSeverEvent();"
        title="<% Sendb(Copient.PhraseLib.Lookup("term.continueWithAttributes", LanguageID))%>"
        value="<% Sendb(Copient.PhraseLib.Lookup("term.continueWithAttributes", LanguageID))%>" />
    <% End If%>
    <% If (PAB <> "1" And (pgID > 0)) Then%>
    <input type="button" class="regular" id="actions" name="actions" value="<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>&#9660;"
        onclick="toggleDropdown();" style="width: 90px" />
    <div class="actionsmenu" id="actionsmenu">
        <% If (pgID > 0 AndAlso InLinkMode) Then%>
        <input type="button" class="regular" id="btnLinkToGroup" name="btnLinkToGroup" value="<%Sendb(Copient.PhraseLib.Lookup("term.LinkToGroup", LanguageID))%>"
            onclick="linkToGroup();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.LinkToGroupMessage", LanguageID))%>" /><br />
        <input type="button" class="regular" id="btnRemoveLink" name="btnRemoveLink" value="<%Sendb(Copient.PhraseLib.Lookup("term.RemoveLinkToGroup", LanguageID))%>"
            onclick="removeLinkToGroup();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.RemoveLinkMessage", LanguageID))%>" /><br />
        <input type="button" class="regular" id="btnExcludeLink" name="btnExcludeLink" value="<%Sendb(Copient.PhraseLib.Lookup("term.ExcludeFromGroup", LanguageID))%>"
            onclick="excludeFromGroup();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ExcludeFromGroupMessage", LanguageID))%>" />
        <input type="button" class="regular" id="btnRemoveExclude" name="btnRemoveExclude"
            value="<%Sendb(Copient.PhraseLib.Lookup("term.RemoveExclusion", LanguageID))%>"
            onclick="removeExclusion();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.RemoveExclusionMessage", LanguageID))%>" />
        <% If MyCommon.Fetch_CPE_SystemOption(140) = "1" Then%>
        <input type="button" class="regular" id="btnAssignSigns" name="btnAssignSigns" value="<%Sendb(Copient.PhraseLib.Lookup("term.AssignSigns", LanguageID))%>"
            onclick="divsign();" title="Assign Signs to Hierarchy" />
        <% End If%>
        <% ElseIf (pgID > 0 AndAlso Not InLinkMode) Then%>
        <input type="button" class="regular" id="btnAddToGroup" name="btnAddToGroup" value="<%Sendb(Copient.PhraseLib.Lookup("term.AddToGroup", LanguageID))%>"
            onclick="addToGroup();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.AddToGroupMessage", LanguageID))%>" /><br />
        <input type="button" class="regular" id="btnRemove" name="btnRemove" value="<%Sendb(Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID))%>"
            onclick="removeFromGroup();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.RemoveFromGroupMessage", LanguageID))%>" /><br />
        <input type="button" class="regular" id="btnRemoveAll" name="btnRemoveAll" value="<%Sendb(Copient.PhraseLib.Lookup("term.RemoveAll", LanguageID))%>"
            onclick="removeAll();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ClearProductGroup", LanguageID))%>" />
        <% Else%>
        <!--
        <input type="button" class="regular" id="btnAdd" name="btnAdd" value="Add to Hierarchy" onclick="toggleHierarchyBox('addBox');" title="Add a hierarchy, folder or item." /><br />
       
        <% If (Logix.UserRoles.DeleteFromHierarchy) Then%>
        <input type="button" class="regular" id="btnDelete" name="btnDelete" value="<% Sendb(Copient.PhraseLib.Lookup("hierarchy.delete", LanguageID))%>"
            onclick="deleteFromHierarchy();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.DeleteMessage", LanguageID))%>" /><br />
        <% End If%>
        -->
        <% End If%>
    </div>
    <% End If%>
</div>
<%--
<div id="toolset">
  <%If (pgID > 0) Then%>
    <input type="button" id="actions" name="actions" value="<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>&#9660;" onclick="toggleDropdown();" />
    <div class="actionsmenu" id="actionsmenu" style="display: none;">
      <input type="button" name="btnAddToGroup" id="btnAddToGroup" value="Add to Group" onclick="addToGroup();" title="Add all checked items in this folder to this product group." /><br />
      <input type="button" name="btnRemove" id="btnRemove" value="Remove from Group" onclick="removeFromGroup();" title="Remove all checked items in this folder from this product group." /><br />
      <input type="button" name="btnRemoveAll" id="btnRemoveAll" value="Remove All" onclick="removeAll();" title="Clear this product group of all items." />
    </div>
  <% End If %>
  <span id="searchmenu">
    <input type="button" name="btnSearch" id="btnSearch" value="Search" onclick="search();" title="Search for items within all product hierarchies." />
  </span>
</div>
--%>
<script runat="server">
    Public MyCommon As New Copient.CommonInc
  
    Sub GenerateDrilledNodeDiv(ByVal nodeId As Integer, ByVal level As Integer, Optional ByVal searchPathIDs As String() = Nothing, Optional ByVal ProductGroupID As Long = 0)
        Dim dt As DataTable
        Dim row As DataRow
        Dim newLevel As Integer = level
        Dim newLeft As Integer
        Dim newNodeId As Integer
        Dim hierId As Integer
        Dim name As String = ""
        Dim sQuery As String = ""
        Dim IdType As String = ""
        Dim iconType() As String = {"", "-green", "-red", "-down", "-purple"}
        Dim excludeIndex As Integer = 0
        Dim Resyncer As New Copient.HierarchyResync(MyCommon, "HierarchyFeeds", "Hierarchy.txt")
        Dim ExtHierarchyID As String = ""
        Dim ExtNodeID As String = ""
        Dim FolderAltText As String = ""
        Dim ParentID As Long = 0
    
        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "/logix/HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            If Not (String.IsNullOrEmpty(Request.QueryString("productgroupid"))) Then
                ProductGroupID = MyCommon.Extract_Val(Request.QueryString("productgroupid"))
            End If
            IdType = MyCommon.Fetch_SystemOption(62)
      
            sQuery = "select Name="
            Select Case IdType
                Case "0" ' none
                    sQuery &= "Name "
                Case "1" ' ExternalID
                    sQuery &= "   case  " & _
                              "       when ExternalID is NULL then Name " & _
                              "       when ExternalID = '' then Name " & _
                              "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                              "       else ExternalID " & _
                              "   end "
                Case "2" ' DisplayID
                    sQuery &= "   case  " & _
                              "       when DisplayID is NULL or DisplayID='' then Name " & _
                              "       else DisplayID + '-' + Name " & _
                              "   end "
                Case Else ' default to none
                    sQuery &= "Name "
            End Select
            sQuery &= " , isnull(HierarchyID, 0) as HierarchyID "
            If (level <= 1) Then
                sQuery &= " from ProdHierarchies with (NoLock) "
                sQuery &= " where HierarchyID=" & nodeId & ";"
            Else
                sQuery &= ", NodeID "
                sQuery &= " from PHNodes with (NoLock) "
                sQuery &= "where NodeID=" & nodeId & ";"
            End If
            MyCommon.QueryStr = sQuery
            dt = MyCommon.LRT_Select()
      
            For Each row In dt.Rows
                If level > 1 Then
                    newNodeId = row.Item("NodeID")
                Else
                    newNodeId = row.Item("HierarchyID")
                End If
                hierId = row.Item("HierarchyID")
                newLeft = level * 17
                name = MyCommon.NZ(row.Item("Name"), "[No Name]")
                name = name.Replace("<", "&lt;")
                FolderAltText = ""
        
                excludeIndex = 0
        
                ' determine if the folder is linked or excluded to the current product group
                If ProductGroupID > 0 Then
                    Resyncer.FindHierarchyExtIDs(newNodeId, ExtHierarchyID, ExtNodeID)
                    If Resyncer.IsNodeLinked(ProductGroupID, ExtHierarchyID, ExtNodeID, True) Then
                        excludeIndex = 1
                        FolderAltText = Copient.PhraseLib.Detokenize("hierarchy.LinkedToProductGroup", LanguageID, ProductGroupID)
                    ElseIf Resyncer.IsNodeLinkedWithAttribute(ProductGroupID, ExtHierarchyID, ExtNodeID) Then
                        excludeIndex = 4
                        FolderAltText = Copient.PhraseLib.Detokenize("hierarchy.LinkedToProductGroup", LanguageID, ProductGroupID)
                    ElseIf Resyncer.IsNodeExcluded(ProductGroupID, ExtHierarchyID, ExtNodeID) Then
                        excludeIndex = 2
                        FolderAltText = Copient.PhraseLib.Detokenize("hierarchy.ExcludedFromProductGroup", LanguageID, ProductGroupID)
                    ElseIf Resyncer.IsChildNodeLinked(ProductGroupID, ExtHierarchyID, ExtNodeID) Then
                        excludeIndex = 3
                        FolderAltText = Copient.PhraseLib.Lookup("hierarchy.ChildLinkedToProductGroup", LanguageID)
                    End If
                End If
        
                Send("<span class=""hrow"" id=""node" & IIf(level <= 1, "H", "") & newNodeId & """>")
                Send(" <span id=""indent" & IIf(level <= 1, "H", "") & newNodeId & """ style=""line-height:18px;"" onclick=""highlightNode(" & newNodeId & ", " & newLevel & ", false)"">")
                Send("  <img src=""/images/clear.png"" style=""height:1px;width:" & newLeft & "px;"" />")
                Send("  <a href=""#""><img id=""imgfldr" & IIf(level <= 1, "H", "") & newNodeId & """ src=""/images/" & IIf(level <= 1, "hierarchy", "folder" & iconType(excludeIndex)) & ".png"" alt=""" & FolderAltText & """ title=""" & FolderAltText & """ border=""0"" /></a>")
                Send("  <span id=""name" & IIf(level <= 1, "H", "") & newNodeId & """ style=""left:5px;"">" & name & "</span>")
                Send("  <input type=""hidden"" id=""hierarchy" & IIf(level <= 1, "H", "") & newNodeId & """ name=""hierarchy" & IIf(level <= 1, "H", "") & newNodeId & """ value=""" & hierId & """ />")
                Send("  <br class=""zero"" />")
                Send(" </span>")
        
                If (Not searchPathIDs Is Nothing) AndAlso (searchPathIDs.GetUpperBound(0) > level) AndAlso (newNodeId = searchPathIDs(level)) Then
                    Send(" <span id=""hId" & IIf(level <= 1, "H", "") & newNodeId & """ style=""display:inline;"">")
                    GenerateDrilledNodeDiv(searchPathIDs(level + 1), level + 1, searchPathIDs, ProductGroupID)
                    Send(" </span>")
                Else
                    Send(" <span id=""hId" & IIf(level <= 1, "H", "") & newNodeId & """ style=""display:none;"">")
                    Send(" </span>")
                End If
        
                Send("</span>")
            Next
        Catch ex As Exception
            Send(Copient.PhraseLib.Lookup("hierarchy.unabletoloadnodes", LanguageID))
            MyCommon.Error_Processor(, ex.ToString(), "/logix/HierarchyFeeds.aspx", , )
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub
  
    Function GetParentNodeList(ByVal NodeId As Long) As String
        Dim ParentNodeIdList As String = ""
        Dim ParentID As Long
    
        ParentID = GetParentNode(NodeId)
        ParentNodeIdList = NodeId
    
        'Get nodes
        While (ParentID > 0)
            ParentNodeIdList = ParentID & "," & ParentNodeIdList
            ParentID = GetParentNode(ParentID)
        End While
    
        'Get hierarchy
        ParentID = GetHierarchyID(NodeId)
        ParentNodeIdList = ParentID & "," & ParentNodeIdList
    
        'Attach the topmost (0) level
        ParentNodeIdList = "0," & ParentNodeIdList
    
        Return ParentNodeIdList
    End Function
  
    Function GetHierarchyID(ByVal NodeId As Integer) As Integer
        Dim HierarchyID As Integer = 0
        Dim MyCommon As New Copient.CommonInc
        Dim dt As DataTable
    
        Try
            MyCommon.AppName = "phierarchytree.aspx"
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select HierarchyID from PHNodes with (NoLock) where NodeId =" & NodeId
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                HierarchyID = MyCommon.NZ(dt.Rows(0).Item("HierarchyID"), 0)
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    
        Return HierarchyID
    End Function
  
    Function GetParentNode(ByVal NodeId As Long) As Integer
        Dim MyCommon As New Copient.CommonInc
        Dim ParentID As Long = 0
        Dim dt As DataTable
    
        Try
            MyCommon.AppName = "phierarchytree.aspx"
            If NodeId <= 0 Then
                ParentID = 0
            Else
                MyCommon.Open_LogixRT()
                MyCommon.QueryStr = "SELECT ParentID, HierarchyID FROM PHNodes WITH (NoLock) WHERE NodeID=" & NodeId & ";"
                dt = MyCommon.LRT_Select
                If (dt.Rows.Count > 0) Then
                    ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0)
                End If
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    
        Return ParentID
    End Function
  
    Function GetItemNodeID(ByVal PKID As Integer) As Integer
        Dim NodeID As Integer = 0
        Dim MyCommon As New Copient.CommonInc
        Dim dt As DataTable
    
        Try
            MyCommon.AppName = "phierarchytree.aspx"
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select NodeID from PHContainer with (NoLock) where PKID =" & PKID
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                NodeID = MyCommon.NZ(dt.Rows(0).Item("NodeID"), 0)
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
    
        Return NodeID
    End Function

</script>
<script type="text/javascript" language="javascript">
    <%If (searchSelNodeID <> "") Then%>
  <% If (Level > 0) Then%>
    highlightNode(<% Sendb(searchSelNodeID)%>, <%Sendb(IIf(searchPathIDs.Length = 1, 2, searchPathIDs.Length))%>, true,true);
    <%Else%>
    highlightNode(<% Sendb(searchSelNodeID)%>, 1, true,true);
    
    <% End If
End If%>
  <%If (pgID = -1 OrElse SelectedNodeIDs = "-1") AndAlso searchPathIDs Is Nothing Then%>
    highlightNode(0,0,false);
  <%End If%>
    if (window.captureEvents) {
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    } else {
        document.onclick=handlePageClick;
    }
    if( typeof CloseWindow != 'undefined' && CloseWindow==true)
    {
        window.close();
    }
</script>
<div id="phfadeDiv">
</div>
<div id="divAssignSigns" class="folderdialog">
</div>
<%
done:
    MyCommon.Close_LogixRT()
    Send_BodyEnd()
    Logix = Nothing
    MyCommon = Nothing
%>