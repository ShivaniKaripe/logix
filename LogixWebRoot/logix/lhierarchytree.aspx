<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%
  ' *****************************************************************************
  ' * FILENAME: lhierarchytree.aspx 
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
  Dim dr As DataRow
  Dim dt As DataTable = Nothing
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim nodeID As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim AssignedCount As Integer = 0
  Dim LocGroupID As Integer = 0
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
  Dim SearchString As String = ""
  Dim BannersEnabled As Boolean = False
  Dim BannerHierarchyID As Integer = -1
  Dim BannerHierarchyIDs As String = ""
  Dim BannerName As String = ""
  Dim IdType As String = ""
  Dim OfferID As Long
  Dim EngineID As Integer
  Dim CreatedFromOffer As Boolean = False
  
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
  EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  CreatedFromOffer = (OfferID > 0 AndAlso EngineID > 0)
  
  MyCommon.AppName = "lhierarchytree.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  LocGroupID = Request.QueryString("LocationGroupID")
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  If (BannersEnabled And LocGroupID > 0) Then
    MyCommon.QueryStr = "select distinct BLH.HierarchyID, BAN.Name as BannerName from LocationGroups LG with (NoLock) " & _
                        "inner join  BannerLocHierarchies BLH with (NoLock) on BLH.BannerID = LG.BannerID " & _
                        "inner join Banners BAN with (NoLock) on BAN.BannerID = BLH.BannerID " & _
                        "where LG.Deleted = 0 and LG.LocationGroupID=" & LocGroupID
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      For Each dr In dt.Rows
        If BannerHierarchyIDs <> "" Then BannerHierarchyIDs &= ","
        BannerHierarchyIDs &= MyCommon.NZ(dr.Item("HierarchyID"), -1)
        BannerName = " " & Copient.PhraseLib.Lookup("term.for", LanguageID).ToLower & " " & MyCommon.TruncateString(MyCommon.NZ(dt.Rows(0).Item("BannerName"), ""), 30)
      Next
    Else
      BannerHierarchyID = 0
    End If
  ElseIf (BannersEnabled) Then
    MyCommon.QueryStr = "select BLH.HierarchyID from AdminUserBanners AUB with (NoLock) " & _
                        "inner join BannerLocHierarchies BLH with (NoLock) on BLH.BannerID = AUB.BannerID " & _
                        "where AdminUserID=" & AdminUserID
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      For Each dr In dt.Rows
        If BannerHierarchyIDs <> "" Then BannerHierarchyIDs &= ","
        BannerHierarchyIDs &= MyCommon.NZ(dr.Item("HierarchyID"), -1)
      Next
    Else
      BannerHierarchyID = 0
    End If
  End If
  
  CancelRefresh = IIf(LocGroupID > 0, "false", "true")
  SearchString = Request.QueryString("searchString")
  SelectedOption = Request.QueryString("selected")
  
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
  
  Send_HeadBegin("term.hierarchy", "term.location")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
* html body {
  margin: 10px 10px 57px 10px !important;
  }
#wrap {
  width: auto !important;
  }
</style>
<%  
  Send_Scripts()
%>

<script type="text/javascript">
    var bLoading = false;
    var hierSel = -1;
    var treeNodeSel = -1;
    var listItemSel = -1;
    var levelSel = -1;
    var NODE_INDENT_SIZE = 17;
    var cancelRefresh = <%Sendb(CancelRefresh) %>;
   
    function toggleNode(id, level) {
        if (isNodeOpen(id, level)) {
            collapseNode(id, level);
        } else {
            expandNode(id, level, false);
        }
    }
    
    function handleItemDblClick(id, level) {
        var tblID = (level == 1) ? "H" : "";
        var elem = document.getElementById("hId" + tblID + id);
        var imgElem = document.getElementById("img" + tblID + id);
        
        if (elem == null && treeNodeSel > -1) {
            expandNode(treeNodeSel, level,false);
        }
        
        if (elem != null && (elem.innerHTML == "")) {
            xmlhttpPost('LocHierarchyFeeds.aspx', 'action=openFolder&node=' + id + '&level=' + level, 'openFolder', id );
        } else if (elem != null) {
            elem.style.display = "";   
        }
        
        if (imgElem != null) {
            imgElem.src = "../images/minus.PNG";
        }
    }
    
    function expandNode(id, level, bReload) {
        var tblID = (level == 1) ? "H" : "";
        var elem = document.getElementById("hId" + tblID + id);
        var imgElem = document.getElementById("img"+ tblID + id);
        var reload = (bReload) ? 1 : 0;
        
        if (elem != null && (elem.innerHTML == "" || bReload)) {
            xmlhttpPost('LocHierarchyFeeds.aspx', 'action=expand&node=' + id + '&level=' + level + "&reload=" + reload, 'expand', id );
            elem.style.display = "";   
        } else if (elem != null) {
            elem.style.display = "";   
        }
        
        if (imgElem != null) {
            imgElem.src = "../images/minus.PNG";
        }
    }
    
    function isNodeOpen(id, level) {
        var retVal = false;
        var tblID = (level == 1) ? "H" : "";
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
        var reload = 0;
        
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
        //alert(qryStr)
        self.xmlHttpReq.open('POST', strURL + '?' + qryStr, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
            if (action == 'expand') {
                level = parseToken(qryStr, "level");
                reload = parseToken(qryStr, "reload");
                updateNodes(id, self.xmlHttpReq.responseText, parseInt(level));
                if (reload) {
                    highlightNode(id, parseInt(level));
                }
            } else if (action == 'openFolder') {
                level = parseToken(qryStr, "level")
                updateNodes(id, self.xmlHttpReq.responseText, parseInt(level));
                if (level != '' && !isNaN(level)) {
                    highlightNode(id, parseInt(level));
                }
            } else if (action == 'showItems') {
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
            } else if (action == 'handleSearchItemAdjust') {
              updateSearchItem(id, self.xmlHttpReq.responseText);
            } else if (action == 'handleSearchNodeAdjust') {
              updateSearchNode(id, self.xmlHttpReq.responseText);
            } else if (action == 'addToHierarchy') {
              updateAddToHierarchy(self.xmlHttpReq.responseText);
            } else if (action == 'delFromHierarchy') {
              updateDelFromHierarchy(self.xmlHttpReq.responseText);
            }
            bLoading = false;
            hideWait();
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
        //alert(qryStr)
        self.xmlHttpReq.open('POST', strURL + '?' + qryStr, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.setRequestHeader('Content-length', params.length);
        self.xmlHttpReq.setRequestHeader('Connection', 'close');
        self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
            showAddedCount(self.xmlHttpReq.responseText);
            bLoading = false;
            hideWait();
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
            
            updateItems(0, response);
            //ChangeParentDocument();            
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
              }
              if (elemSearchType != null) {
                elemSearchType.style.visibility = 'hidden';
              }
            } else {
              elem.innerHTML = '<div class=\"loading\"><br \/><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/><% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
              
              // disable background so the user can't make changes during wait.
              if (elemWait != null) {
                  elemWait.style.display = "block";
              }
              
            }
        }
        
    }
    
    function hideWait() {
      var elemWait = document.getElementById("WaitDiv");
      var elem = document.getElementById("itemList");
      var elemToolbar = document.getElementById("toolset");
      var elemSearch = document.getElementById("SearchDiv");
      var elemSearchType = document.getElementById("searchType");
      
      if (elemToolbar != null) {
        if (elemSearch == null || elemSearch.style.display != 'block') {
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
    
    function updateNodes(id, divHTML, level) {
        var tblID = (level == 1) ? "H" : "";
        var elem = document.getElementById("hId" + tblID + id);
        var elemImg = document.getElementById("img" + tblID + id);
        
        if (elem != null && divHTML != "") {
            elem.style.display = "block";
            elem.innerHTML = divHTML;
        } else if (elemImg != null) {
            elemImg.src = "../images/blank.PNG";
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
        
        //alert(divHTML);
        var lg = document.frmLHier.locgroup.value;
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
                            if (lg != "" && parseInt(lg) > 0) {
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
        var lg = document.frmLHier.locgroup.value;
        var msg = response;
        var ct = 0;
        var preCt, postCt;
        var commaPos = -1;
        var tokenValues = [];
                       
        if (response != null && response.length > 0) {
            if (response.substring(0,1) == "|") {
                commaPos = response.indexOf(",");
                preCt = parseInt(response.substring(1, commaPos));
                postCt = parseInt(response.substring(commaPos+1));
                if (postCt == 0) {
                    tokenValues = [preCt, lg]
                    msg = detokenizeString('<% Sendb(Copient.PhraseLib.Lookup("lhierarchy.RemovedItems", LanguageID)) %>', tokenValues);
                } else {
                    tokenValues = [(preCt - postCt), lg, postCt];
                    msg = detokenizeString('<% Sendb(Copient.PhraseLib.Lookup("lhierarchy.PartialItemRemoval", LanguageID)) %>', tokenValues);
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
                showNodeItems(id, levelSel);   
            }
            //ChangeParentDocument();            
        }
    }
    
    function  updateSearch(id, response) {
      var resultsElem = document.getElementById("linerresults");
      var itemElem = document.getElementById("searchItemCt");
      var trailerPos = -1, trailerEnd = -1;
      var resultCt = 0;      
      var MAX_RESULTS = 500;
      var msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.MaxResults", LanguageID)) %>';
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
    
    function updateSearchItem(id, response) {
      if (response == 'ADDED' || response == 'REMOVED') {
        var elemAssigned = document.getElementById('assignedCt');
        var elemLink = document.getElementById("link" + id);
        var elemImg = document.getElementById("img" + id);
        var newAction = (response == 'ADDED') ? 'Remove' : 'Add';
        var newImg = (response== 'ADDED') ? '../images/store-on.png' : '../images/store.png';
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
        var tblID = (level == 1) ? "H" : "";
        var elem = document.getElementById("hId" + tblID + id);
        var imgElem = document.getElementById("img" + tblID + id);
        
        if (elem != null) {
            elem.style.display = 'none';
        }
        
        if (imgElem != null) {
            imgElem.src = "../images/plus.PNG";
        }
    }
    function showNodeItems(id, level) {
        var name = "";
        var lg = document.frmLHier.locgroup.value;
        var itemPK = document.frmLHier.itemSelectedPK.value;
        var tblID = (level == 1) ? "H" : "";        
        
        if (id > -1) {
            elem = document.getElementById('name' + tblID + id);
            if (elem != null) {
                name = elem.innerHTML;
            }
            xmlhttpPost('LocHierarchyFeeds.aspx', 'action=showItems&node=' + id + '&level=' + level + "&nodeName=" + name + "&lg=" + lg + "&itemPK=" + itemPK, 'showItems', id  );
        }
    }
    
    function highlightNode(id, level) {
        var elem = null;
        var tblID = (level == 1) ? "H" : "";        
        
        showNodeItems(id, level);
        
        // highlight the selected node
        elem = document.getElementById('name' + tblID +  id);
        if (elem != null) {
            elem.style.backgroundColor = '#000080';
            elem.style.color = '#ffffff';
        }
        
        // unselect the previously selected node
        if (treeNodeSel > -1 && (id != treeNodeSel || level != levelSel)) {
          if (levelSel == 1) {
            elem = document.getElementById('nameH' + treeNodeSel);
          } else {
            elem = document.getElementById('name' + treeNodeSel);
          }
          if (elem != null) {
              elem.style.backgroundColor = '#ffffff';
              elem.style.color = '#000000';
          }
        }
        
        if (hierSel > - 1 && (id != hierSel || level != levelSel)) {
          elem = document.getElementById('nameH' + hierSel);
          if (elem != null) {
              elem.style.backgroundColor = '#ffffff';
              elem.style.color = '#000000';
          }
        }
        
        treeNodeSel = (level == 1) ? -1 : id;
        hierSel = (level == 1) ? id : -1;
        listItemSel = -1;
        levelSel = level;
    }
    
    function highlightItem(id) {
        var elem = null;
        
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
        
        xmlhttpPost('LocHierarchyFeeds.aspx', qryStr, 'addNode', parentId  );
    }
    
    function showDeleteNode() {
        var elemName = null;
        var name = "";
        var response;
        var msg = '';
        var tokenValues = [];
        
        if (treeNodeSel == -1) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.SelectNodeToDelete", LanguageID)) %>')
        } else {
            elemName = document.getElementById("name" + treeNodeSel);
            if (elemName != null) {
                name = elemName.innerHTML;            
            }
            msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmDeleteNode", LanguageID)) %>';
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
        
        //alert('delete node ' + id + " level" + level + " hier: " + hierId);
        qryStr = 'action=deleteNode&node=' + id + '&level=' + level + '&hierId=' + hierId;
        xmlhttpPost('LocHierarchyFeeds.aspx', qryStr, 'deleteNode', id );
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
                xmlhttpPost('LocHierarchyFeeds.aspx', qryStr, 'deleteItems', treeNodeSel);
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
                    xmlhttpPost('LocHierarchyFeeds.aspx', qryStr, 'showAvailItems', treeNodeSel);
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
        var lg = document.frmLHier.locgroup.value;
        var elemChk = null;
        var i = 0;
        var nodeIDs = "";
        
        elemChk = document.getElementById("chk" + i);
        while (elemChk != null) {
            if (elemChk.checked) {
                if (nodeIDs != "") { nodeIDs = nodeIDs + ","; }
                nodeIDs += elemChk.value;
            }
            i++;
            elemChk = document.getElementById("chk" + i);
        }
        if (nodeIDs > "") {
            transmitGroups('LocHierarchyFeeds.aspx', "action=assignNodes", "lg=" + lg + "&sel=" + treeNodeSel + "&ids=" + nodeIDs);       
        } else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.NoItemsSelected", LanguageID)) %>');
        }
    }
    
    function removeFromGroup() {
        var lg = document.frmLHier.locgroup.value;
        var elemChk = null;
        var i = 0;
        var nodeIDs = "";
        
        elemChk = document.getElementById("chk" + i);
        while (elemChk != null) {
            if (elemChk.checked) {
                if (nodeIDs != "") { nodeIDs = nodeIDs + ","; }
                nodeIDs += elemChk.value;
            }
            i++;
            elemChk = document.getElementById("chk" + i);
        }
        if (nodeIDs > "") {
            //alert(nodeIDs);
            transmitGroups('LocHierarchyFeeds.aspx', "action=unassignNodes", "lg=" + lg + "&sel=" + treeNodeSel + "&ids=" + nodeIDs);       
        } else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.NoItemsSelected", LanguageID))%>');
        }
    }
    
    function search() {
      var lg = document.frmLHier.locgroup.value;
      var searchDiv = document.getElementById("SearchDiv");
      var searchBoxDiv = document.getElementById("searchbox");
      var itemToolbarDiv = document.getElementById("toolset");
      var totalItemElem  = document.getElementById("totalItemCt");
      var resultsElem = document.getElementById("searchItemCt");
      var searchFolder = document.getElementById("searchFolder");
      
      if (searchDiv != null) {
        searchDiv.style.display = "block";
        
        if (searchBoxDiv != null) {
          searchBoxDiv.style.display = 'block';
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
    
    function closeSearch() {
        var searchDiv = document.getElementById("SearchDiv");
        var searchBoxDiv = document.getElementById("searchbox");
        var itemToolbarDiv = document.getElementById("toolset");
        var totalItemElem  = document.getElementById("totalItemCt");
        var resultsElem = document.getElementById("searchItemCt");
         
        if (searchDiv != null) {
          searchDiv.style.display = "none";
          
          if (searchBoxDiv != null) {
            searchBoxDiv.style.display = 'none';
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
    
    function scrollToNode(nodeID) {
        var elem = document.getElementById("node" + nodeID);
        var column1Div = document.getElementById("leftpane");
        
        if (elem != null) {
            ScrollToElement(elem, column1Div);
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
        var addedstrcount = window.opener.$("#contents option").length;
        if(addedstrcount != 0)
        {
        var lg = document.frmLHier.locgroup.value;
        var msg = '<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.ConfirmRemoveAll", LanguageID)) %>';
        msg= msg.replace("&#39;","'");
        var tokenValues = [lg];

        var confirmResponse = confirm(detokenizeString(msg, tokenValues));
        
        if (confirmResponse) {
            xmlhttpPost('LocHierarchyFeeds.aspx', 'action=removeAll&node=' + treeNodeSel + "&lg=" + lg, 'removeAll', treeNodeSel );
            }
        }
        else
        {
            alert('<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.nostores", LanguageID)) %>');
        }
    }
    
    function findMatches() {
      var searchElem = document.getElementById("searchString");
      var searchTypeElem = document.getElementById("searchType");
      var lg = document.frmLHier.locgroup.value;
      var sType = 0;
      
      if (searchElem != null) {
        if (searchElem.value == '') {
          alert('<%Sendb(Copient.PhraseLib.Lookup("folders.SearchTerm", LanguageID))%>');
          searchElem.focus();
        } else{
            if (searchTypeElem != null) { sType = searchTypeElem.value; }
            xmlhttpPost('LocHierarchyFeeds.aspx', 'action=findMatches&lg=' + lg + '&search=' + encodeURIComponent(htmlEntities(searchElem.value)) + '&stype=' + sType + '&nodeid=' + treeNodeSel + '&hierid=' + hierSel, 'findMatches', treeNodeSel );
        }
      }
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
      var lg = document.frmLHier.locgroup.value;
      
      if (selValue != '') {
        cancelRefresh = true;
        document.location = 'lhierarchytree.aspx?selected=' + selValue + '&LocationGroupID=' + lg;
      }
    }
    
    function handleSearchItemAdjust(id, linkID, hID, extID) {
      var lg = document.frmLHier.locgroup.value;
      var linkID = document.getElementById(linkID);
      var type = '';
      
      if (linkID != null) {
        type = (linkID.innerHTML=='Add') ? 'add' : 'remove';  
      
        var qryStr = 'action=handleSearchItemAdjust&lg=' + lg + '&store=' + id + '&type=' + type + '&hID=' + hID + '&extID=' + extID;
        
        if (id != null) {
          xmlhttpPost('LocHierarchyFeeds.aspx', qryStr, 'handleSearchItemAdjust', id + 'H' + hID );
        }
      }
    }
    
    function handleSearchNodeAdjust(id, linkID, hID, type) {
      var lg = document.frmLHier.locgroup.value;
      var linkID = document.getElementById(linkID);
      
      if (id != null) {
        var qryStr = 'action=handleSearchNodeAdjust&lg=' + lg + '&nodeID=' + id + '&type=' + type + '&hID=' + hID;
        xmlhttpPost('LocHierarchyFeeds.aspx', qryStr, 'handleSearchNodeAdjust', id + 'H' + hID );
      }
    }
        
    function select_Click() {
      var selValue = getSelectedValue();
      var lg = document.frmLHier.locgroup.value;
      
      if (selValue != '') {
        cancelRefresh = true;
        document.location = 'lhierarchytree.aspx?selected=' + selValue + '&LocationGroupID=' + lg;
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
        
    function ChangeParentDocument() {
        if (opener != null && !cancelRefresh) {
            opener.document.location = 'lgroup-edit.aspx?LocationGroupID=<%Sendb(LocGroupID) %>'
        }
    }
    
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
    
    function deleteFromHierarchy() {
      var okToDelete = false;
      var msg = '';
      var nodeID = '';
      var itemID = '';
      var isHier = 0;
      var tdElems = null;
                  
      if (listItemSel  > -1) {
        // delete the item selected in the right pane
        var elem = document.getElementById("PKID" + listItemSel);
        if (elem != null) {
          nodeID = treeNodeSel;
          itemID = elem.value;
          var elemListItem = document.getElementById("itemRow" + listItemSel);
          if (elemListItem != null && elemListItem.cells != null && elemListItem.cells.length >= 3) {       
            msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmDeleteID", LanguageID))%>: ' + elemListItem.cells[1].innerHTML + " " + elemListItem.cells[2].innerHTML;         
          } else {
            msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmItemDelete", LanguageID)) %>';
          }
        } else {
          alert('<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ErrorOnItemDelete", LanguageID)) %>');
        }
      } else if (treeNodeSel > - 1) {
        // delete the folder selected in the left pane 
        nodeID = treeNodeSel;
        itemID = ''
        var elemTree = document.getElementById("name" + treeNodeSel);
        if (elemTree == null) {
          elemTree = document.getElementById("nameH" + treeNodeSel);
          isHier = 1;
        }
        if (elemTree != null) {  
          msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmDelete", LanguageID)) %>';
          tokenValues[0] = cleanCellText(elemTree.innerHTML);
          msg = detokenizeString(msg, tokenValues);
        } else {
          msg = '<%Sendb(Copient.PhraseLib.Lookup("phierarchy.ConfirmItemDelete", LanguageID)) %>';
        }
      } else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.NoDeletableItem", LanguageID)) %>');
      }
      
      if (msg != '') {
        okToDelete = confirm(msg);
        if (okToDelete) {
          var qryStr = 'action=delFromHierarchy&nodeID=' + nodeID + '&itemID=' + itemID + "&isHier=" + isHier;
          xmlhttpPost('LocHierarchyFeeds.aspx', qryStr, 'delFromHierarchy', itemID );
        }
      }
      
    }
    
    function cleanCellText(cellText) {
      var newCellText = cellText;
      
      newCellText = cellText.replace('&nbsp;', '');
      
      return newCellText;
    }
    
    function updateAddToHierarchy(response) {
      alert(response);  
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
          highlightNode(id, levelSel);
        } else {
          // full refresh
          parent = parseTagValue(response, "parent");
          if (parent == 0 && treeNodeSel > -1 ) { parent = treeNodeSel; }          
          expandNode(parent, levelSel, true);
        }
      }            
    }
   
    function handleAllItems() {
      var elem = null;
      var i = 0;
      
      var elemAll = document.getElementById("chkAll");
      
      if (elemAll != null) {
        elem = document.getElementById("chk" + i);
        while (elem != null) {
          elem.checked = elemAll.checked;
          i++;
          elem = document.getElementById("chk" + i);
        }
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
      var lg = document.frmLHier.locgroup.value;
      var itemPK = document.frmLHier.itemSelectedPK.value;
      
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
      
      //xmlhttpPost('LocHierarchyFeeds.aspx', 'action=showItems&node=' + treeNodeSel + '&level=' + levelSel + "&nodeName=" + name + "&lg=" + lg + "&itemPK=" + itemPK + "&sort=" + curSortCode, 'showItems', treeNodeSel  );

	  masterSortCode = curSortCode;
		if (levelSel <= 1) 
        {
		xmlhttpPost('LocHierarchyFeeds.aspx', 'action=showItems&node=' + hierSel + '&level=' + levelSel + '&nodeName=' + name + "&lg=" + lg + '&itemPK=' + itemPK + '&sort=' + masterSortCode + '&StartIndex=1', 'showItems', treeNodeSel);
        } 
        else 
        {
		xmlhttpPost('LocHierarchyFeeds.aspx', 'action=showItems&node=' + treeNodeSel + '&level=' + levelSel + '&nodeName=' + name + "&lg=" + lg +  '&itemPK=' + itemPK + '&sort=' + curSortCode + '&StartIndex=1', 'showItems', treeNodeSel);
        }
    }
    
    function getSelectedFolderName() {
      var name = ' | <%Sendb(Copient.PhraseLib.Lookup("term.searching", LanguageID)) %> '
      var elem = null
      
      if (treeNodeSel > 0) {
        elem = document.getElementById('name' + treeNodeSel);
        if (elem != null) {
          name += '<%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%>: ' + elem.innerHTML.substring(0,20);
        } else {
          name += '<%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%>: [<%Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>]';
        }
      } else if (hierSel > 0) {
        elem = document.getElementById('nameH' + hierSel);
        if (elem != null) {
          name += '<%Sendb(Copient.PhraseLib.Lookup("term.hierarchy", LanguageID))%>: ' + elem.innerHTML.substring(0,20);
        } else {
          name += '<%Sendb(Copient.PhraseLib.Lookup("term.hierarchy", LanguageID))%>: [<%Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>]';
        }
      } else {
        name += ' <%Sendb(Copient.PhraseLib.Lookup("term.AllHierarchies", LanguageID)) %>'
      }
      
      return name;
    }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(IIf(CreatedFromOffer, 3, 2))
    
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration")
    GoTo done
  End If
%>
<form id="frmLHier" name="frmLHier" action="#">
  <input type="hidden" id="locgroup" name="locgroup" value="<%Sendb(LocGroupID) %>" />
  <input type="hidden" id="itemSelectedPK" name="itemSelectedPK" value="<%Sendb(itemSelectedPK) %>" />
  <input type="hidden" id="searchPathIDs" name="searchPathIDs" value="<%Sendb(searchHierarchyPath) %>" />
  <input type="hidden" id="selNode" name="selNode" value="<%Sendb(searchSelNodeID) %>" />
</form>
<div class="box" id="toolbar">
  <h2>
    <span style="float: left;">
      <%Sendb(Copient.PhraseLib.Lookup("term.storehierarchies", LanguageID) & BannerName)%>
    </span>
  </h2>
  <br clear="left" />
</div>
<div id="hmain">
  <div class="column" id="leftpane">
    <div class="liner" id="topNodesDiv">
      <%
        'Find the top level product hierarchy levels
        IdType = MyCommon.Fetch_SystemOption(63)
        MyCommon.QueryStr = "select HierarchyID, ExternalID, Name = "
        
        Select Case IdType
          Case "0" ' None
            MyCommon.QueryStr &= "Name "
          Case "1" ' ExternalID
            MyCommon.QueryStr &= "   case  " & _
                                 "       when ExternalID is NULL then Name " & _
                                 "       when ExternalID = '' then Name " & _
                                 "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                                 "       else ExternalID " & _
                                 "   end "
          Case "2" ' DisplayID
            MyCommon.QueryStr &= "   case  " & _
                                 "       when DisplayID is NULL or DisplayID='' then Name " & _
                                 "       else DisplayID + '-' + Name " & _
                                 "   end "
          Case Else ' default to none
            MyCommon.QueryStr &= "Name "
        End Select
        
        MyCommon.QueryStr &= "from LocationHierarchies with (nolock) " & _
                             "" & IIf(BannerHierarchyID > -1, "where HierarchyID = " & BannerHierarchyID & " ", "") & _
                             "" & IIf(BannerHierarchyIDs <> "", "where HierarchyID in (" & BannerHierarchyIDs & ") ", "") & _
                             "order by ExternalID;"
        'Send(MyCommon.QueryStr)
        'GoTo done
        dt = MyCommon.LRT_Select()
        Send("")
        For Each dr In dt.Rows
          nodeID = MyCommon.NZ(dr.Item("HierarchyID"), 0)
          Send(vbTab & "<span id=""nodeH" & nodeID & """><span id=""indentH" & nodeID & """ style=""line-height:18px;left:0px;""><a href=""javascript:toggleNode(" & nodeID & ", 1);""><img id=""imgH" & nodeID & """ src=""../images/plus.png"" border=""0"" alt="""" /></a>")
          Send(vbTab & "<a href=""#"" ondblclick=""toggleNode(" & nodeID & ",1);""><span onclick=""javascript:highlightNode(" & nodeID & ",1)""><img src=""../images/hierarchy.png"" border=""0"" alt=""""/></span></a>")
          Send(vbTab & "<span id=""nameH" & nodeID & """ onclick=""highlightNode(" & nodeID & ",1)"" ondblclick=""toggleNode(" & nodeID & ",1);"">&nbsp;" & MyCommon.NZ(dr.Item("Name"), "[" & Copient.PhraseLib.Lookup("term.noname", LanguageID) & "]") & "</span>")
          Sendb(vbTab & "<input type=""hidden"" id=""hierarchy" & nodeID & """ name=""hierarchy" & nodeID & """ value=""" & nodeID & """ />")
          Send("<br /></span>")
          ' check if this is a search result find and the source hierarchy
          If (Not searchPathIDs Is Nothing AndAlso searchPathIDs.Length >= 1 AndAlso nodeID = searchPathIDs(0)) Then
            Sendb("<span id=""hIdH" & nodeID & """ style=""display:inline;"">")
            GenerateNodeDiv(searchPathIDs(0), 1, searchPathIDs)
          Else
            Sendb("<span id=""hIdH" & nodeID & """ style=""display:none;"">")
          End If
          Send("</span></span>")
        Next
      %>
    </div>
  </div>
  <div class="column" id="rightpane">
    <div class="liner" id="itemList">
    </div>
  </div>
  <div class="searchresults" id="SearchDiv">
    <div id="linerresults">
      <br clear="left" />
      <%
        Send("<table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """ style=""width:100%;"">")
        Send("  <thead>")
        Send("    <tr>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.select", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.action", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.level", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.hierarchy", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
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
</div>
<div id="searchbox">
  <input type="text" class="medium" id="searchString" name="searchString" maxlength="100" onkeydown="handleFindKeyDown(event);" value="<% Sendb(SearchString) %>" />
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
  <input type="button" id="find" name="find" value="<% Sendb(Copient.PhraseLib.Lookup("term.find", LanguageID))%>" onclick="findMatches();" />
  <input type="button" id="close" name="close" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID)) %>" onclick="closeSearch();" />
  <span id="searchFolder" style="color:Gray;"></span>
</div>
<div id="statusfooter">
  <%
    If (LocGroupID > 0) Then
      MyCommon.QueryStr = "Select count(*) as AssignedCount from LocGroupItems with (NoLock) where LocationGroupID=" & LocGroupID & " and Deleted=0;"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        AssignedCount = MyCommon.NZ(dt.Rows(0).Item("AssignedCount"), 0)
      End If
  %>
  <span><%Sendb(Copient.PhraseLib.Lookup("term.locationgroup", LanguageID))%>:<%Sendb(LocGroupID)%></span>&nbsp;|&nbsp;<span><%Sendb(Copient.PhraseLib.Lookup("term.Assigned", LanguageID))%>: <span id="assignedCt"><% Sendb(AssignedCount)%></span></span>
  <% End If%>
  <span id="totalItemCt"></span><span id="searchItemCt"></span>
</div>
<div id="WaitDiv" style="">
</div>
<div id="deleteBox" class="modifyBox">
  <div style="width:100%; background-color:Blue; color:White;">
    <b>Delete from Hierarchy</b>
  </div>
  <br />
  <center>
    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>">
      <tr>
        <td colspan="2">
          <%Sendb(Copient.PhraseLib.Lookup("phierarchy.DeleteLevelMessage", LanguageID)) %>
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
            <option value="1"><%Sendb(Copient.PhraseLib.Lookup("term.EntireHierarchy", LanguageID)) %></option>
            <option value="2"><%Sendb(Copient.PhraseLib.Lookup("term.folder", LanguageID))%></option>
            <option value="3"><%Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))%></option>
          </select>
        </td>
      </tr>
      <tr>
        <td style="padding-left: 20px;">
          <%Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>:
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
          <input type="button" id="btnDelOK" name="btnDelOK" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>" onclick="alert('ok clicked');" />
          <input type="button" id="btnDelCancel" name="btnDelCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID)) %>" onclick="toggleHierarchyBox('deleteBox');" />
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
          <%Sendb(Copient.PhraseLib.Lookup("phierarchy.SelectAddLevel", LanguageID)) %>
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
          <select id="addHierType" name="addHierType" class="mediumshort">
            <option value="1">Entire Hierarchy</option>
            <option value="2">Folder</option>
            <option value="3">Store</option>
          </select>
        </td>
      </tr>
      <tr>
        <td style="padding-left: 20px;">
          <%Sendb(Copient.PhraseLib.Lookup("term.ParentID", LanguageID))%>:
        </td>
        <td>
          <input type="text" class="mediumshort" id="addParentID" name="addParentID" value="" />
        </td>
      </tr>
      <tr>
        <td style="padding-left: 20px;">
          <%Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>:
        </td>
        <td>
          <input type="text" class="mediumshort" id="addID" name="addID" value="" />
        </td>
      </tr>
      <tr>
        <td style="padding-left: 20px;">
          <%Sendb(Left(Copient.PhraseLib.Lookup("term.description", LanguageID), 4))%>:
        </td>
        <td>
          <input type="text" class="mediumshort" id="addDesc" name="addDesc" value="" />
        </td>
      </tr>
      <tr style="height: 15px;">
        <td colspan="2">
        </td>
      </tr>
      <tr>
        <td colspan="2" style="text-align: center;">
          <input type="button" id="btnAddOK" name="btnAddOK" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>" onclick="alert('ok clicked');" />
          <input type="button" id="btnAddCancel" name="btnAddCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID)) %>" onclick="toggleHierarchyBox('addBox');" />
        </td>
      </tr>
    </table>
  </center>
</div>
<div id="toolset">
  <form id="mainform" name="mainform" action="">
    <input type="button" class="regular" id="btnSearch" name="btnSearch" onclick="search();" title="<% Sendb(Copient.PhraseLib.Lookup("hierarchy.search", LanguageID)) %>" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>" />
    <% If (LocGroupID > 0 OrElse Logix.UserRoles.DeleteFromHierarchy) Then%>
    <input type="button" class="regular" id="actions" name="actions" value="<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>&#9660;" onclick="toggleDropdown();" />
    <div class="actionsmenu" id="actionsmenu">
      <% If (LocGroupID > 0) Then%>
      <input type="button" class="regular" id="btnAddToGroup" name="btnAddToGroup" value="<%Sendb(Copient.PhraseLib.Lookup("term.AddToGroup", LanguageID)) %>" onclick="addToGroup();" title="<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.AddToGroupMessage", LanguageID)) %>" /><br />
      <input type="button" class="regular" id="btnRemove" name="btnRemove" value="<%Sendb(Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID)) %>" onclick="removeFromGroup();" title="<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.RemoveFromGroupMessage", LanguageID)) %>" /><br />
      <input type="button" class="regular" id="btnRemoveAll" name="btnRemoveAll" value="<%Sendb(Copient.PhraseLib.Lookup("term.RemoveAll", LanguageID)) %>" onclick="removeAll();" title="<%Sendb(Copient.PhraseLib.Lookup("lhierarchy.ClearGroup", LanguageID)) %>" />
      <% Else%>
      <!--
      <input type="button" class="regular" id="btnAdd" name="btnAdd" value="Add to Hierarchy" onclick="toggleHierarchyBox('addBox');" title="Add a hierarchy, folder or item." /><br />
      -->
      <% If (Logix.UserRoles.DeleteFromHierarchy) Then%>
      <input type="button" class="regular" id="btnDelete" name="btnDelete" value="<% Sendb(Copient.PhraseLib.Lookup("hierarchy.delete", LanguageID)) %>" onclick="deleteFromHierarchy();" title="<%Sendb(Copient.PhraseLib.Lookup("phierarchy.DeleteMessage", LanguageID)) %>" /><br />
      <%End If
      End If%>
    </div>
    <% End If%>
  </form>
</div>

<script runat="server">
  Public MyCommon As New Copient.CommonInc
  
  Sub GenerateNodeDiv(ByVal nodeId As Integer, ByVal level As Integer, ByVal searchPathIDs As String())
    Dim dt As DataTable
    Dim row As DataRow
    Dim newLevel As Integer = level + 1
    Dim newLeft As Integer
    Dim newNodeId As Integer
    Dim hierId As Integer
    Dim x As Integer
    Dim name As String = ""
    Dim sQuery As String = ""
    Dim IdType As String = ""
    
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
      MyCommon.Open_LogixRT()
      
      sQuery = "select NodeID, HierarchyID, NodeName = "
      IdType = MyCommon.Fetch_SystemOption(63)
      Select Case IdType
        Case "0" ' None
          sQuery += "NodeName "
        Case "1" ' ExternalID
          sQuery += "   case  " & _
                    "       when ExternalID is NULL then NodeName " & _
                    "       when ExternalID = '' then NodeName " & _
                    "       when ExternalID not like '%' + NodeName + '%' then ExternalID + '-' + NodeName " & _
                    "       else ExternalID " & _
                    "   end "
        Case "2" ' DisplayID
          sQuery += "   case  " & _
                    "       when DisplayID is NULL or DisplayID='' then NodeName " & _
                    "       else DisplayID + '-' + NodeName " & _
                    "   end "
        Case Else ' default to none
          sQuery += "NodeName "
      End Select
      sQuery += " from LHNodes with (nolock) "
      
      If (level = 1) Then
        sQuery += "where HierarchyID = " & nodeId & " and ParentId = 0 order by NodeName;"
      Else
        sQuery += "where parentId = " & nodeId & " order by NodeName;"
      End If
      MyCommon.QueryStr = sQuery
      
      dt = MyCommon.LRT_Select()
      For Each row In dt.Rows
        newNodeId = MyCommon.NZ(row.Item("NodeID"), 0)
        hierId = MyCommon.NZ(row.Item("HierarchyID"), 0)
        newLeft = level * 17
        name = MyCommon.NZ(row.Item("NodeName"), "[No Name]")
        name = name.Replace("&", "&amp;")
        name = name.Replace("<", "&lt;")
        name = name.Replace(">", "&gt;")
        
        Send("<span id=""node" & newNodeId & """>")
        Send(" <span id=""indent" & newNodeId & """ style=""line-height:18px;"">")
        Send("  <img src=""../images/clear.png"" style=""height:1px;width:" & newLeft & "px;"" />")
        Send("  <a href=""javascript:toggleNode(" & newNodeId & ", " & newLevel & ");""><img id=""img" & newNodeId & """ src=""../images/plus.png"" alt="""" border=""0"" alt="""" /></a>")
        Send("  <a href=""#"" ondblclick=""toggleNode(" & newNodeId & ", " & newLevel & ");""><img src=""../images/folder.png"" alt="""" border=""0"" alt="""" /></a>")
        Send("  <span id=""name" & newNodeId & """ style=""left:5px;"" onclick=""highlightNode(" & newNodeId & ", " & newLevel & ")"" ondblclick=""toggleNode(" & newNodeId & ", " & newLevel & ");"">" & name & "</span>")
        Send("  <input type=""hidden"" id=""hierarchy" & newNodeId & """ name=""hierarchy" & newNodeId & """ value=""" & hierId & """ />")
        Send("  <br  class=""zero"" />")
        Send(" </span>")
        
        If (Not searchPathIDs Is Nothing AndAlso searchPathIDs.GetUpperBound(0) >= level AndAlso newNodeId = searchPathIDs(level)) Then
          Send(" <span id=""hId" & newNodeId & """ style=""display:inline;"">")
          GenerateNodeDiv(newNodeId, level + 1, searchPathIDs)
        Else
          Send(" <span id=""hId" & newNodeId & """ style=""display:none;"">")
        End If
        
        Send(" </span>")
        Send("</span>")
      Next
    Catch ex As Exception
      Send(Copient.PhraseLib.Lookup("hierarchy.unabletoloadnodes", LanguageID))
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  Function GetParentNodeList(ByVal NodeId As Integer) As String
    Dim ParentNodeIdList As String = ""
    Dim ParentID As Integer
    
    ParentID = GetParentNode(NodeId)
    While (ParentID > 0)
      ParentNodeIdList = ParentID & "," & ParentNodeIdList
      ParentID = GetParentNode(ParentID)
    End While
    ParentNodeIdList = GetHierarchyID(NodeId) & "," & ParentNodeIdList
    
    If (ParentNodeIdList <> "") Then
      ParentNodeIdList = ParentNodeIdList.Substring(0, ParentNodeIdList.LastIndexOf(","))
    End If
    
    Return ParentNodeIdList
  End Function
  
  Function GetHierarchyID(ByVal NodeId As Integer) As Integer
    Dim HierarchyID As Integer = 0
    Dim MyCommon As New Copient.CommonInc
    Dim dt As DataTable
    
    Try
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select HierarchyID from LHNodes with (NoLock) where NodeId =" & NodeId
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
  
  Function GetParentNode(ByVal NodeId As Integer) As Integer
    Dim ParentID As Integer = 0
    Dim MyCommon As New Copient.CommonInc
    Dim dt As DataTable
    
    Try
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select ParentID from LHNodes with (NoLock) where NodeId =" & NodeId
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0)
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
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select NodeID from LHContainer with (NoLock) where PKID =" & PKID
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

<%If (searchSelNodeID <> "") Then%>

<script type="text/javascript" language="javascript">
    highlightNode(<% Sendb(searchSelNodeID) %>, <%Sendb(iif(searchPathIDs.length=1, 2, searchPathIDs.length))%>);
    scrollToNode(<% Sendb(searchSelNodeID) %>);
</script>

<%End If%>

<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
