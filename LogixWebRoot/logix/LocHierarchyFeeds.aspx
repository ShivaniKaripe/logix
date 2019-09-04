<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim AdminUserID As Long
  Dim action As String = ""
  Dim nodeId As Integer
  Dim hierId As Integer
  Dim level As Integer
  Dim nodeName As String = ""
  Dim items As String = ""
  Dim ParentID As Integer
  Dim IsTopLevel As Boolean = False
  Dim PreCount As Integer = 0
  Dim PostCount As Integer = 0
  Dim foundItems As Boolean = False
  Dim LogMsg As String = ""
  
  MyCommon.AppName = "LocHierarchyFeeds.aspx"  
  
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  action = Request.QueryString("action")
  nodeId = MyCommon.Extract_Val(Request.QueryString("node"))
  level = MyCommon.Extract_Val(Request.QueryString("level"))
  nodeName = Request.QueryString("nodeName")
  hierId = MyCommon.Extract_Val(Request.QueryString("hierId"))
  items = Request.QueryString("items")
  
  If (nodeId > 0 AndAlso action <> "") Then
    Response.Expires = 0
    Response.Clear()
    Response.ContentType = "text/html"
    Select Case action
      Case "expand"
        GenerateNodeDiv(nodeId, level)
      Case "openFolder"
        GenerateNodeDiv(nodeId, level)
      Case "showItems"
        GenerateItemsDiv(nodeId, level, nodeName, Request.QueryString("lg"))
      Case "removeAll"
        RemoveAllFromLocationGroups(Request.QueryString("lg"), AdminUserID)
      Case Else
        Send(Request.RawUrl)
    End Select
  ElseIf (action = "assignNodes") Then
    PreCount = GetProductGroupItemCount(Request.Form("lg"))
    TransmitGroups(Request.Form("ids"), Request.Form("lg"), foundItems, True)
    If (Request.Form("sel") <> "") Then
      Integer.TryParse(Request.Form("sel"), ParentID)
    End If
    If (ParentID <= 0) Then ParentID = FindParentNodeID(Request.Form("ids"), IsTopLevel)
    GenerateItemsDiv(ParentID, IIf(IsTopLevel, 1, 0), "", Request.Form("lg"))
    PostCount = GetProductGroupItemCount(Request.Form("lg"))
    Send(GetAssignedTrailer(PreCount, PostCount, foundItems))
  ElseIf (action = "unassignNodes") Then
    PreCount = GetProductGroupItemCount(Request.Form("lg"))
    TransmitGroups(Request.Form("ids"), Request.Form("lg"), foundItems, False)
    If (Request.Form("sel") <> "") Then
      Integer.TryParse(Request.Form("sel"), ParentID)
    End If
    If (ParentID <= 0) Then ParentID = FindParentNodeID(Request.Form("ids"), IsTopLevel)
    GenerateItemsDiv(ParentID, IIf(IsTopLevel, 1, 0), "", Request.Form("lg"))
    PostCount = GetProductGroupItemCount(Request.Form("lg"))
    Send(GetAssignedTrailer(PreCount, PostCount, foundItems))
  ElseIf (action = "removeAll") Then
    RemoveAllFromLocationGroups(Request.QueryString("lg"), AdminUserID)
  ElseIf (action = "findMatches") Then
    FindSearchItems(Request.QueryString("search"), MyCommon.Extract_Val(Request.QueryString("stype")), _
                    Request.QueryString("lg"), MyCommon.Extract_Val(Request.QueryString("nodeid")), _
                    MyCommon.Extract_Val(Request.QueryString("hierid")))
  ElseIf (action = "handleSearchItemAdjust") Then
    handleSearchItemAdjust(Request.QueryString("type"), Request.QueryString("hID"), Request.QueryString("lg"), Request.QueryString("store"))
  ElseIf (action = "handleSearchNodeAdjust") Then
    handleSearchNodeAdjust(Request.QueryString("type"), Request.QueryString("hID"), Request.QueryString("lg"), Request.QueryString("nodeID"))
  ElseIf (action = "addToHierarchy") Then
    'handleAddToHierarchy(Request.QueryString("level"), Request.QueryString("prntID"), Request.QueryString("extID"), Request.QueryString("desc"))
  ElseIf (action = "delFromHierarchy") Then
    handleDelFromHierarchy(Request.QueryString("nodeID"), Request.QueryString("itemID"))
  Else
    Send("action: " & action)
    Send("node: " & nodeId)
    Send("level: " & level)
    Send("nodeName: " & nodeName)
    Send(Request.RawUrl)
    Send("<b>" & Copient.PhraseLib.Lookup("term.noarguments", LanguageID) & "</b>")
  End If
  Response.Flush()
  Response.End()
%>

<script runat="server">
  Public DefaultLanguageID
  Public MyCommon As New Copient.CommonInc
  Public callCount As Integer = 0
  Private Enum SORT_CODES
    SORT_ID_ASC = 1
    SORT_ID_DESC = 2
    DESCRIPTION_ASC = 3
    DESCRIPTION_DESC = 4
  End Enum
  
  Sub GenerateNodeDiv(ByVal nodeId As Integer, ByVal level As Integer)
    Dim dt As DataTable
    Dim row As DataRow
    Dim newLevel As Integer = level + 1
    Dim newLeft As Integer
    Dim newNodeId As Integer
    Dim hierId As Integer
    Dim name As String = ""
    Dim sQuery As String = ""
    Dim IdType As String = ""
    
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
        MyCommon.Open_LogixRT()
      End If
      
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
        sQuery += "where HierarchyID = " & nodeId & " and ParentId = 0 order by ExternalID;"
      Else
        sQuery += "where parentId = " & nodeId & " order by ExternalID;"
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
                      
        Send("<span id=""node" & newNodeId & """ name=""node" & newNodeId & """>")
        Send(" <span id=""indent" & newNodeId & """ style=""line-height:18px;"">")
        Send("  <img src=""../images/clear.png"" style=""height:1px;width:" & newLeft & "px;"" />")
        Send("  <a href=""javascript:toggleNode(" & newNodeId & ", " & newLevel & ");""><img id=""img" & newNodeId & """ src=""../images/plus.png"" alt="""" border=""0"" /></a>")
        Send("  <a href=""#"" ondblclick=""toggleNode(" & newNodeId & ", " & newLevel & ");""><span onclick=""javascript:highlightNode(" & newNodeId & ", " & newLevel & ")""""><img src=""../images/folder.png"" alt="""" border=""0"" /></span></a>")
        Send("  <span id=""name" & newNodeId & """ style=""left:5px;"" onclick=""highlightNode(" & newNodeId & ", " & newLevel & ")"" ondblclick=""toggleNode(" & newNodeId & ", " & newLevel & ");"">" & name & "</span>")
        Send("  <input type=""hidden"" id=""hierarchy" & newNodeId & """ name=""hierarchy" & newNodeId & """ value=""" & hierId & """ />")
        Send("  <br class=""zero"" />")
        Send(" </span>")
        Send(" <span id=""hId" & newNodeId & """ style=""display:none;""></span>")
        Send("</span>")
      Next
    Catch ex As Exception
      Send(Copient.PhraseLib.Lookup("hierarchy.unabletoloadnodes", LanguageID))
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  Sub GenerateItemsDiv(ByVal nodeId As Integer, ByVal level As Integer, ByVal nodeName As String, ByVal lgID As String)
    Dim dt As DataTable
    Dim row As DataRow
    Dim ExtLocCode As String = ""
    Dim LocationID As Integer
    Dim LocName As String = ""
    Dim iconFileName As String = ""
    Dim LevelType As Integer = -1
    Dim PKID As Integer
    Dim i As Integer = 0
    Dim SelectedItem As Boolean = False
    Dim sQuery As String = ""
    Dim highlightedRow As Integer = -1
    Dim totalItemCt As Integer = 0
    Dim totalSelCt As Integer = 0
    Dim EngineID As Integer = 2 ' default to CPE
    Dim DisplayID As String = ""
    Dim IdType As String = ""
    Dim CurrentSortCode As Integer = SORT_CODES.SORT_ID_ASC
    Dim SortCol As String = "ExternalID"
    Dim SortDir As String = "ASC"
    Dim Col1Icon As String = "&nbsp;"
    Dim Col2Icon As String = "&nbsp;"
    Dim IdColName As String = "ExternalID"
    Dim BannersEnabled As Boolean = False
    Dim AdminUserID As Long
    Dim Logix As New Copient.LogixInc
    
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
      MyCommon.Open_LogixRT()
      
      AdminUserID = Verify_AdminUser(MyCommon, Logix)
      BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
        
      ' find the EngineID for this location group
      MyCommon.QueryStr = "select EngineId from LocationGroups with (NoLock) where LocationGroupID=" & lgID
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineId"), 2)
      Else
        ' find the default engine and use this as the engineID
        MyCommon.QueryStr = "select EngineId from PromoEngines with (NoLock) where DefaultEngine=1 and Installed=1;"
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
          EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineId"), 2)
        End If
      End If
      
      ' figure out the sorted column and its order
      IdType = MyCommon.Fetch_SystemOption(62)
      IdColName = IIf(IdType = "2", "DisplayID", "ExternalID")
      If Not Integer.TryParse(Request.QueryString("sort"), CurrentSortCode) Then CurrentSortCode = SORT_CODES.SORT_ID_ASC
      SortCol = IIf(CurrentSortCode = SORT_CODES.SORT_ID_ASC OrElse CurrentSortCode = SORT_CODES.SORT_ID_DESC, IdColName, "LocationName")
      SortDir = IIf(CurrentSortCode = SORT_CODES.SORT_ID_ASC OrElse CurrentSortCode = SORT_CODES.DESCRIPTION_ASC, "ASC", "DESC")
      If (SortCol = IdColName AndAlso SortDir = "ASC") Then Col1Icon = "&#9650;"
      If (SortCol = IdColName AndAlso SortDir = "DESC") Then Col1Icon = "&#9660;"
      If (SortCol = "LocationName" AndAlso SortDir = "ASC") Then Col2Icon = "&#9650;"
      If (SortCol = "LocationName" AndAlso SortDir = "DESC") Then Col2Icon = "&#9660;"
      Col1Icon = "<span style=""color:#808080;padding-left:5px;width:20px;"">" & Col1Icon & "</span>"
      Col2Icon = "<span style=""color:#808080;padding-left:10px;width:20px;"">" & Col2Icon & "</span>"
      
      If (level = 1) Then
        sQuery = "select 1 as LevelType, NodeID as PKID, -1 as LocationID, ExternalID as ExtLocationCode, DisplayID, NodeName as LocationName, 0 as Selected " & _
                 "from lhnodes with (NoLock) where parentid=0 and hierarchyid=" & nodeId & " " & _
                 "order by " & SortCol & " " & SortDir & ";"
      Else
        sQuery = "select 1 as LevelType, NodeID as PKID, -1 as LocationID, ExternalID as ExtLocationCode, DisplayID, NodeName as LocationName, 0 as Selected " & _
                  "from lhnodes with (NoLock) where parentid =" & nodeId & " " & _
                  "union " & _
                  "select 2 as LevelType, con.pkid, loc.LocationID, loc.ExtLocationCode, '' as DisplayID, loc.LocationName,  " & _
                  "  (select count(*) from LocGroupItems with (NoLock) where LocationGroupID=" & lgID & " and LocationID = loc.LocationID and Deleted=0) as Selected  " & _
                  "from LHContainer con with (nolock) left join Locations loc with (nolock) on con.LocationID = loc.LocationID  " & _
                  "where con.nodeID =" & nodeId & " and loc.EngineID=" & EngineID & " "
        
        If (BannersEnabled) Then
          sQuery &= "and (BannerID =0 or BannerID Is Null or BannerID in ( " & _
                    "				select AUB.BannerID from Banners BAN with (NoLock) " & _
                    "				inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                    "				where AUB.AdminUserID = " & AdminUserID & ") ) "
        End If
        sQuery &= "order by " & SortCol & " " & SortDir & ";"
      End If
      MyCommon.QueryStr = sQuery
      dt = MyCommon.LRT_Select()
      
      'Send(MyCommon.QueryStr)
      Send("<table style=""width:100%;font-family:arial;font-size:11px;"" summary=""" & Copient.PhraseLib.Lookup("term.items", LanguageID) & """>")
      Send("<thead>")
      Send("<tr>")
      If (lgID <> "" AndAlso CInt(lgID) > 0) Then Send("<th scope=""col"" style=""background-color:#e0e0e0;width:30px;""><input type=""checkbox"" id=""chkAll"" name=""chkAll"" title=""" & Copient.PhraseLib.Lookup("hierarchy.SelectAllItems", LanguageID) & """ onclick=""handleAllItems();"" /></th>")
      Send("<th scope=""col"" style=""background-color:#e0e0e0;width:30px;"">&nbsp;</th>")
      Send("<th scope=""col"" style=""background-color:#e0e0e0;"" onclick=""sortByColumn(1," & CurrentSortCode & ");"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & Col1Icon & "</th>")
      Send("<th scope=""col"" style=""background-color:#e0e0e0;"" onclick=""sortByColumn(2," & CurrentSortCode & ");"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & Col2Icon & "</th>")
      Send("</tr>")
      Send("</thead>")
      Send("<tbody>")
      
      If (dt.Rows.Count = 0) Then
        If (lgID <> "" AndAlso CInt(lgID) > 0) Then
          Send("<tr><td style=""width:30px;""></td><td style=""width:30px;""></td><td style=""width:100px;""></td><td style=""width:200px;""></td></tr>")
          Send("<tr><td colspan=""4"">" & Copient.PhraseLib.Detokenize("hierarchy.NoItemsInNode", LanguageID, nodeName) & "</td></tr>")
        Else
          Send("<tr><td style=""width:30px;""></td><td style=""width:100px;""></td><td style=""width:200px;""></td></tr>")
          Send("<tr><td colspan=""3"">" & Copient.PhraseLib.Detokenize("hierarchy.NoItemsInNode", LanguageID, nodeName) & "</td></tr>")
        End If
      Else
        For Each row In dt.Rows
          PKID = MyCommon.NZ(row.Item("pkid"), 0)
          LocationID = MyCommon.NZ(row.Item("LocationID"), -1)
          ExtLocCode = MyCommon.NZ(row.Item("ExtLocationCode"), "")
          DisplayID = MyCommon.NZ(row.Item("DisplayID"), "")
          LocName = MyCommon.NZ(row.Item("LocationName"), "[No Name]")
          SelectedItem = IIf(MyCommon.NZ(row.Item("Selected"), 0) > 0, True, False)
          If (SelectedItem) Then
            totalSelCt += 1
          End If
          
          LevelType = MyCommon.NZ(row.Item("LevelType"), -1)
          Select Case LevelType
            Case 1
              iconFileName = "folder.png"
            Case Else
              iconFileName = "store" & IIf(SelectedItem, "-on", "") & ".png"
              totalItemCt += 1
          End Select
          
          If (LevelType = 1) Then
            Send("<tr id=""itemRow" & i & """ onclick=""highlightItem(" & i & ");"" ondblclick=""handleItemDblClick(" & PKID & ", " & (level + 1) & ");"">")
          Else
            Send("<tr id=""itemRow" & i & """ onclick=""highlightItem(" & i & ");"">")
          End If
                          
          If (lgID <> "" AndAlso CInt(lgID) > 0) Then Send("<td><input type=""checkbox"" id=""chk" & i & """ name=""chk" & i & """ value=""" & IIf(LevelType = 1, "N" & PKID, "I" & LocationID) & """  /></td>")
          Sendb("<td>")
          
          If (lgID = "" OrElse CInt(lgID <= 0)) Then
            Send("<input type=""hidden"" id=""PKID" & i & """ name=""PKID" & i & """ value=""" & IIf(LevelType = 1, "N" & PKID, "I" & LocationID) & """ /> ")
          End If
          
          Send("<img src=""../images/" & iconFileName & """ />" & "</td>")
          
          If (MyCommon.NZ(row.Item("LevelType"), -1) = 1) Then
            ' work-around in case the name is found in the external id then remove the name
            If (ExtLocCode.IndexOf(LocName) > -1) Then
              ExtLocCode = ExtLocCode.Replace("-" & LocName, "")
            End If
          End If
          
          Send("<td>" & GetIdText(ExtLocCode, DisplayID, LevelType) & "</td>")
          Send("<td nowrap>" & LocName & "</td>")
          Send("</tr>")
          
          If (Request.QueryString("itemPK") <> "" AndAlso PKID = Integer.Parse(Request.QueryString("itemPK"))) Then
            highlightedRow = i
          End If
          
          i += 1
        Next
      End If
      
      Send("</tbody>")
      Send("</table>")
      
      'If (highlightedRow > -1) Then
      Send("<trailer>" & highlightedRow & "," & totalItemCt & "," & totalSelCt & "</trailer>")
      'End If
      
    Catch ex As Exception
      Send(Copient.PhraseLib.Lookup("hierarchy.unabletoloaditems", LanguageID))
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  Private Function GetIdText(ByVal ExternalID As String, ByVal DisplayID As String, ByVal LevelType As String)
    Dim IdText As String = ""
    Dim IdType As String = ""
    
    IdType = MyCommon.Fetch_SystemOption(63)
    'Is this a folder? (i.e. 1) - if so determine which ID to display
    If (LevelType = "1") Then
      Select Case IdType
        Case "0" ' none
          IdText = "&nbsp;"
        Case "1" ' external Id
          IdText = IIf(ExternalID = "-1", "&nbsp;", ExternalID)
        Case "2"
          IdText = DisplayID
        Case Else
          IdText = "&nbsp;"
      End Select
    Else
      ' otherwise, show the store external id
      IdText = ExternalID
    End If
    
    Return IdText
  End Function
  
  Sub TransmitGroups(ByVal groupIDs As String, ByVal lg As String, ByRef foundItems As Boolean, Optional ByVal add As Boolean = True)
    Dim sQuery As String = ""
    Dim dt As DataTable
    Dim dtStores As DataTable
    Dim dtName As DataTable
    Dim row As DataRow
    Dim insertRow As DataRow
    Dim groupArray As String() = Nothing
    Dim itemList As New ArrayList(100)
    Dim childList As ArrayList
    Dim temp As String
    Dim i As Integer
    Dim nodeList As String()
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    Dim nodeName As String = ""
    Dim hierarchyName As String = ""
    Dim LogMsg As String = ""
    Dim SelTreeNode As Integer = -1
    Dim itemString As String
    Dim recordsAffected As Integer = -1
    Dim EngineID As Integer = 2 ' Default to CPE
    Dim m_OfferService As IOffer
    
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
      MyCommon.Open_LogixRT()
      AdminUserID = Verify_AdminUser(MyCommon, Logix)
      CurrentRequest.Resolver.AppName = MyCommon.AppName
      m_OfferService = CurrentRequest.Resolver.Resolve(Of IOffer)()
      ' find the EngineID for this location group
      MyCommon.QueryStr = "select EngineId from LocationGroups with (NoLock) where LocationGroupID=" & lg
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineId"), 2)
      End If

      groupArray = groupIDs.Split(",")
      
      dtStores = New DataTable()
      dtStores.Columns.Add("LocationGroupID", System.Type.GetType("System.Int64"))
      dtStores.Columns.Add("LocationID", System.Type.GetType("System.Int64"))
      dtStores.Columns.Add("LastUpdate", System.Type.GetType("System.DateTime"))
      dtStores.Columns.Add("Deleted", System.Type.GetType("System.Boolean"))
      dtStores.Columns.Add("StatusFlag", System.Type.GetType("System.Int32"))
      dtStores.Columns.Add("TCRMAStatusFlag", System.Type.GetType("System.Int32"))
      
      For i = 0 To groupArray.GetUpperBound(0)
        temp = groupArray(i)
        ' if it's a store simply save it to add later; otherwise, find all the stores in the tree
        ' and add them now.
        If (temp <> "" AndAlso temp.Substring(0, 1) = "I") Then
          itemList.Add(temp.Substring(1))
          foundItems = True
        ElseIf (temp <> "") Then
          childList = New ArrayList(100)
          childList = GetChildNodes(temp.Substring(1))
          childList.Add(temp.Substring(1))
          nodeList = Array.ConvertAll(Of Object, String)(childList.ToArray, New Converter(Of Object, String)(AddressOf cString))
          If (nodeList.Length > 0) Then
            ' find stores at the node level
            MyCommon.QueryStr = "select LHC.LocationID from LHContainer LHC with (NoLock) " & _
                                "inner join Locations LOC on LOC.LocationID = LHC.LocationID " & _
                                "where NodeID in (" & String.Join(",", nodeList) & ") and EngineID=" & EngineID
            dt = MyCommon.LRT_Select
            'Send(MyCommon.QueryStr)
            If (dt.Rows.Count > 0) Then
              dtStores.Clear()
              ' add the store to the database for this product group
              For Each row In dt.Rows
                insertRow = dtStores.NewRow()
                insertRow.ItemArray = New Object() {Long.Parse(lg), row.Item("LocationID"), Date.Now, 0, 2, 2}
                dtStores.Rows.Add(insertRow)
              Next
              If (dtStores.Rows.Count > 0) Then
                foundItems = True
                ' get the selected node name for logging purposes
                MyCommon.QueryStr = "select NodeName = case when LHN.ExternalID is NULL then LHN.NodeName when LHN.ExternalID = '' then LHN.NodeName  " & _
                                    "when LHN.ExternalID not like '%' + LHN.NodeName + '%' then LHN.ExternalID + '-' + LHN.NodeName else LHN.ExternalID end  " & _
                                    ", LH.Name as HierarchyName from LHNodes LHN with (nolock) inner join LocationHierarchies LH with (NoLock) on LH.HierarchyID=LHN.HierarchyID " & _
                                    "where NodeID = " & temp.Substring(1) & "; "
                dtName = MyCommon.LRT_Select
                If (dtName.Rows.Count > 0) Then
                  nodeName = dtName.Rows(0).Item("NodeName")
                  hierarchyName = dtName.Rows(0).Item("HierarchyName")
                End If
                LogMsg = IIf(add, "Added ", "Removed ") & nodeName & " ( " & dtStores.Rows.Count & " locations) from hierarchy " & hierarchyName
                If (add) Then
                  recordsAffected = BatchInsert(dtStores, dtStores.Rows.Count)
                Else
                  recordsAffected = BatchDelete(dtStores, lg)
                End If
                If (recordsAffected > 0) Then
                  If (MyCommon.Fetch_UE_SystemOption(191) = "1") Then 
                      m_OfferService.ProcessOfferCollisionDetectionStoreGroupChanges(lg) 
                  End If 
                  RecordSelectedNode(lg, temp.Substring(1), add)
                  MyCommon.Activity_Log(11, CLng(lg), AdminUserID, LogMsg)
                End If
              End If
            End If
          End If
        End If
      Next
      
      ' add all the individual stores that are at the selected node level
      If (itemList.Count > 0) Then
        dtStores.Clear()
        For i = 0 To itemList.Count - 1
          insertRow = dtStores.NewRow()
          insertRow.ItemArray = New Object() {Long.Parse(lg), Long.Parse(itemList.Item(i)), Date.Now, 0, 2, 2}
          dtStores.Rows.Add(insertRow)
        Next
        If (dtStores.Rows.Count > 0) Then
          foundItems = True
          If (add) Then
            recordsAffected = BatchInsert(dtStores, dtStores.Rows.Count)
          Else
            recordsAffected = BatchDelete(dtStores, lg)
          End If
          
           If (recordsAffected > 0 And MyCommon.Fetch_UE_SystemOption(191) = "1") Then                   
                m_OfferService.ProcessOfferCollisionDetectionStoreGroupChanges(lg) 
           End If
          
          ' get the selected node name for logging purposes
          If (Request.Form("sel") <> "") Then
            Integer.TryParse(Request.Form("sel"), SelTreeNode)
          End If
          MyCommon.QueryStr = "select NodeName = case when LHN.ExternalID is NULL then LHN.NodeName when LHN.ExternalID = '' then LHN.NodeName  " & _
                              "when LHN.ExternalID not like '%' + LHN.NodeName + '%' then LHN.ExternalID + '-' + LHN.NodeName else LHN.ExternalID end  " & _
                              ", LH.Name as HierarchyName from LHNodes LHN with (nolock) inner join LocationHierarchies LH on LH.HierarchyID=LHN.HierarchyID " & _
                              "where NodeID = " & SelTreeNode & "; "
          dtName = MyCommon.LRT_Select
          If (dtName.Rows.Count > 0) Then
            nodeName = dtName.Rows(0).Item("NodeName")
            hierarchyName = dtName.Rows(0).Item("HierarchyName")
          End If
          ' log the items that were added or removed
          itemString = IIf(dtStores.Rows.Count = 1, " location ", " locations ")
          LogMsg = IIf(add, "Added ", "Removed ") & dtStores.Rows.Count & itemString & " from " & nodeName & " within hierarchy " & hierarchyName
          MyCommon.Activity_Log(11, CLng(lg), AdminUserID, LogMsg)
        End If
      End If
      'If (Not foundItems) Then
      '    Send("No Items were found at or below this level in the hierarchy.")
      'End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  Function cString(ByVal obj As Object) As String
    Return obj.ToString
  End Function
  
  Function GetChildNodes(ByVal NodeId As String) As ArrayList
    Dim childList As New ArrayList(100)
    Dim dt As DataTable
    Dim row As DataRow
    
    ' bail out if this is a runaway recursive procedure
    callCount += 1
    If (callCount > 100000) Then
      Return childList
      Send(Copient.PhraseLib.Lookup("hierarchy.TooManyNodeSelected", LanguageID))
    End If
    
    MyCommon.QueryStr = "select NodeID from LHNodes with (NoLock) where ParentID=" & NodeId
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      For Each row In dt.Rows
        childList.Add(row.Item("NodeID"))
        childList.AddRange(GetChildNodes(row.Item("NodeID")))
      Next
    End If
    
    Return childList
  End Function
  
  Public Function BatchInsert(ByVal dataTable As DataTable, ByVal batchSize As Int32) As Integer
    Dim adapter As New SqlDataAdapter()
    Dim recordsAffected As Integer
    
    'Set the INSERT command and parameters.
    adapter.InsertCommand = New SqlCommand( _
      "Insert into LocGroupItems with (RowLock) (LocationGroupID, LocationID, LastUpdate, Deleted, StatusFlag, TCRMAStatusFlag) " & _
      "values (@LocationGroupID, @LocationID, getdate(), 0, 2, 2) ", _
      MyCommon.LRTadoConn)
    adapter.InsertCommand.Parameters.Add("@LocationGroupID", _
      SqlDbType.BigInt, 8, "LocationGroupID")
    adapter.InsertCommand.Parameters.Add("@LocationID", _
      SqlDbType.BigInt, 8, "LocationID")
    adapter.InsertCommand.UpdatedRowSource = UpdateRowSource.None
    
    ' Set the batch size.
    batchSize = IIf(batchSize > 1000, batchSize = 1000, batchSize)
    adapter.UpdateBatchSize = batchSize
    
    recordsAffected = adapter.Update(dataTable)
    
    Return recordsAffected
  End Function
  
  Public Function BatchDelete(ByVal dataTable As DataTable, ByVal LocationGroupID As String) As Integer
    Dim LocIdList As String() = Nothing
    Dim IdClause As String = ""
    Dim row As DataRow
    Dim i As Integer = 0
    Dim recordsAffected As Integer = -1
    
    If (Not dataTable Is Nothing AndAlso dataTable.Rows.Count > 0) Then
      ReDim LocIdList(dataTable.Rows.Count - 1)
      
      For Each row In dataTable.Rows
        LocIdList(i) = MyCommon.NZ(row.Item("LocationID"), "-1")
        i += 1
      Next
      
      If (LocIdList.Length > 0) Then
        IdClause = String.Join(",", LocIdList)
        MyCommon.QueryStr = "Delete from LocGroupItems with (RowLock) where LocationGroupID=" & LocationGroupID & " " & _
                            "and Deleted=1 and LocationID in (" & IdClause & ");"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "Update LocGroupItems with (RowLock) set Deleted=1, LastUpdate=getdate(), StatusFlag=2, TCRMAStatusFlag=2 where Deleted=0 and LocationGroupID=" & LocationGroupID & " " & _
                            "and LocationID in (" & IdClause & ");"
        'Send(MyCommon.QueryStr)
        MyCommon.LRT_Execute()
        recordsAffected = MyCommon.RowsAffected
      End If
    End If
    Return recordsAffected
  End Function
  
  Public Function FindParentNodeID(ByVal groupIDs As String, ByRef IsTopLevel As Boolean) As Long
    Dim ParentID As Long = -1
    Dim temp As String
    Dim dt As DataTable
    Dim idList As String()
    
    If (groupIDs.Length > 0) Then
      idList = groupIDs.Split(",")
      temp = idList(0)
      If (temp.Length > 0) Then
        If (temp.Substring(0, 1) = "I") Then
          MyCommon.QueryStr = "Select NodeID from LHContainer with (NoLock) where LocationID = " & temp.Substring(1)
          'Send(MyCommon.QueryStr)
          dt = MyCommon.LRT_Select
          If (dt.Rows.Count > 0) Then
            ParentID = MyCommon.NZ(dt.Rows(0).Item("NodeID"), -1)
          End If
        Else
          MyCommon.QueryStr = "Select ParentID, HierarchyID from LHNodes with (NoLock) where NodeID = " & temp.Substring(1)
          'Send(MyCommon.QueryStr)
          dt = MyCommon.LRT_Select
          If (dt.Rows.Count > 0) Then
            ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), -1)
            If (ParentID = 0) Then
              ParentID = MyCommon.NZ(dt.Rows(0).Item("HierarchyID"), -1)
              IsTopLevel = True
            End If
          End If
        End If
      End If
    End If
    
    Return ParentID
  End Function
  
  Function GetProductGroupItemCount(ByVal lgID As String) As Integer
    Dim AssignedCount As Integer = 0
    Dim dt As DataTable
    Try
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "Select count(*) as AssignedCount from LocGroupItems with (NoLock) where LocationGroupID=" & lgID & " and Deleted=0;"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        AssignedCount = MyCommon.NZ(dt.Rows(0).Item("AssignedCount"), 0)
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
    End Try
    Return AssignedCount
  End Function
  
  Function GetAssignedTrailer(ByVal PreCount As Integer, ByVal PostCount As Integer, ByRef foundItems As Boolean) As String
    Dim msgBuf As New StringBuilder()
    Dim assignedChange As Integer
    Dim displayValue As Integer
    
    assignedChange = PostCount - PreCount
    displayValue = Math.Abs(assignedChange)
    
    msgBuf.Append("<trailer>")
    msgBuf.Append("<message>")
    
    If (foundItems) Then
      If (assignedChange < 0) Then
        msgBuf.Append(Copient.PhraseLib.Detokenize("hierarchy.ItemsRemovedFromLG", LanguageID, displayValue))
      ElseIf (assignedChange = 0) Then
        If (Request.QueryString("action") = "unassignNodes" OrElse Request.QueryString("type") = "REMOVE") Then
          msgBuf.Append(Copient.PhraseLib.Lookup("hierarchy.NoItemsToRemove", LanguageID))
        Else
          msgBuf.Append(Copient.PhraseLib.Lookup("hierarchy.NoItemsToAdd", LanguageID))
        End If
      Else
        msgBuf.Append(Copient.PhraseLib.Detokenize("hierarchy.ItemsAddedToLG", LanguageID, displayValue))
      End If
    Else
      msgBuf.Append(Copient.PhraseLib.Lookup("hierarchy.NoItemsBelow", LanguageID))
    End If
    msgBuf.Append("</message>")
    msgBuf.Append("<count>")
    msgBuf.Append(PostCount)
    msgBuf.Append("</count>")
    msgBuf.Append("</trailer>")
    Return msgBuf.ToString
  End Function
  
  Sub RemoveAllFromLocationGroups(ByVal LocationGroupID As String, ByVal AdminUserID As Long)
    Dim rowCt As Integer = 0
    Dim dt As DataTable
    Dim preDeleteCt As Integer = 0
    Dim postDeleteCt As Integer = 0
    Dim LogMsg As String = ""
    Dim EngineID As Integer = 2 ' Default to CPE
    
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
      MyCommon.Open_LogixRT()
      
      MyCommon.QueryStr = "Select count(*) as LocGroupCt from LocGroupItems with (NoLock) where Deleted=0 and LocationGroupID=" & LocationGroupID
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        preDeleteCt = MyCommon.NZ(dt.Rows(0).Item("LocGroupCt"), 0)
      End If
      
      MyCommon.QueryStr = "Delete from LocGroupItems with (RowLock) where LocationGroupID=" & LocationGroupID & " and Deleted=1;"
      MyCommon.LRT_Execute()
      
      MyCommon.QueryStr = "Update LocGroupItems with (RowLock) set Deleted=1, LastUpdate=getdate(), StatusFlag=2, TCRMAStatusFlag=3 where LocationGroupID=" & LocationGroupID & " and Deleted=0;"
      MyCommon.LRT_Execute()
      
      MyCommon.QueryStr = "Select count(*) as LocGroupCt from LocGroupItems with (NoLock) where Deleted=0 and LocationGroupID=" & LocationGroupID
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        postDeleteCt = MyCommon.NZ(dt.Rows(0).Item("LocGroupCt"), 0)
      End If
      
      LogMsg = Copient.PhraseLib.Lookup("lhierarchy.remove-all", LanguageID) & " (" & Math.Abs(postDeleteCt - preDeleteCt) & _
         " " & Copient.PhraseLib.Lookup("term.items", LanguageID) & ")"
      MyCommon.Activity_Log(11, CLng(LocationGroupID), AdminUserID, LogMsg)
      
      If (postDeleteCt = 0) Then
        MyCommon.QueryStr = "delete from LocationGroupNodes with (RowLock) where LocationGroupID = " & LocationGroupID
        MyCommon.LRT_Execute()
      End If
      
      Send("|" & preDeleteCt & "," & postDeleteCt)
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  Sub FindSearchItems(ByVal SearchString As String, ByVal SearchType As Integer, ByVal LocationGroupID As String, _
                      ByVal SelectedNodeID As Integer, ByVal SelectedHierID As Integer)
    Dim dt As DataTable
    Dim row As DataRow
    Dim Level As Integer
    Dim LevelDisplay As String = ""
    Dim NodeID As Integer
    Dim Name As String = ""
    Dim ExternalID As String = ""
    Dim LocationID As String = ""
    Dim ShowNoItemMsg As Boolean = True
    Dim Shaded As String = "shaded"
    Dim iconFileName As String = ""
    Dim SelectedItem As Boolean = False
    Dim HierarchyID As Integer = -1
    Dim AnchorID As String = ""
    Dim ImgID As String = ""
    Dim LinkTitle As String = ""
    Dim rowCt As Integer = 0
    Dim EngineID As Integer = 2 ' Default to CPE
        Dim OrigSearchString = SearchString
        
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
      If (SearchString <> "") Then
        Send("<table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """ style=""width:100%;"">")
        Send("<thead>")
        Send("  <tr>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.select", LanguageID) & "</th>")
        If (LocationGroupID <> "" And CInt(LocationGroupID) > 0) Then Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.action", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.level", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.hierarchy", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "</th>")
        Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
        Send("  </tr>")
        Send("</thead>")
        Send("<tbody>")
        MyCommon.Open_LogixRT()
        
        ' find the EngineID for this location group
        MyCommon.QueryStr = "select EngineId from LocationGroups with (NoLock) where LocationGroupID=" & LocationGroupID
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
          EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineId"), 2)
        Else
          If MyCommon.IsEngineInstalled(0) Then
            EngineID = 0
          ElseIf MyCommon.IsIntegrationInstalled(2) Then
            EngineID = 2
          ElseIf MyCommon.IsIntegrationInstalled(9) Then
            EngineID = 9
          End If
        End If
        Select Case SearchType
          Case 0 ' contains
            SearchString = "%" & MyCommon.Parse_Quotes(SearchString) & "%"
          Case 1 ' starts with
            SearchString = MyCommon.Parse_Quotes(SearchString) & "%"
          Case 2 ' ends with
            SearchString = "%" & MyCommon.Parse_Quotes(SearchString)
          Case Else ' use contains as default
            SearchString = "%" & MyCommon.Parse_Quotes(SearchString) & "%"
        End Select
        
        If SelectedNodeID > 0 Then
          ' search for all the matches in and under the selected node.
          MyCommon.QueryStr = "dbo.pa_LHA_SearchFromNode"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@NodeID", SqlDbType.BigInt).Value = SelectedNodeID
          MyCommon.LRTsp.Parameters.Add("@SearchString", SqlDbType.NVarChar, 50).Value = SearchString
          MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
          MyCommon.LRTsp.Parameters.Add("@LGID", SqlDbType.BigInt).Value = CLng(LocationGroupID)
          dt = MyCommon.LRTsp_select
          MyCommon.Close_LRTsp()
        ElseIf SelectedHierID > 0 Then
          ' search the selected hierarchy for matches
          MyCommon.QueryStr = "select top 500 * from ( " & _
                              "select 1 as LevelType, NodeID as ID, NodeName as LocationName, lhn.ExternalID as ExternalID, -1 as LocationID,lh.Name as HierarchyName, lh.HierarchyID, 0 as Selected " & _
                              "from LHNodes lhn with (NoLock) " & _
                              "inner join LocationHierarchies lh with (NoLock) on lh.HierarchyID = lhn.HierarchyID and lh.HierarchyID=" & SelectedHierID & " " & _
                              "Where lhn.ExternalID like '" & SearchString & "' or lhn.NodeName like '" & SearchString & "' " & _
                              "union " & _
                              "select 2 as LevelType, PKID as ID, loc.LocationName, loc.ExtLocationCode as ExternalID, loc.LocationID,lh.Name as HierarchyName, lh.HierarchyID,  " & _
                              "	(select count(*) from LocGroupItems with (NoLock) where LocationGroupID=" & LocationGroupID & " and LocationID = loc.LocationID and Deleted=0) as Selected  " & _
                              "from Locations loc with (NoLock) inner join LHContainer con on loc.LocationID= con.LocationID " & _
                              "inner join LHNodes lhn with (NoLock) on con.NodeID = lhn.NodeID " & _
                              "inner join LocationHierarchies lh with (NoLock) on lh.HierarchyId = lhn.HierarchyID and lh.HierarchyID=" & SelectedHierID & " " & _
                              "Where loc.Deleted=0 and loc.EngineID=" & EngineID & " and loc.ExtLocationCode like '" & SearchString & "' or loc.LocationName like '" & SearchString & "' " & _
                              ") ResultsTable " & _
                              "order by HierarchyID, LevelType, ExternalID, LocationName; "
          dt = MyCommon.LRT_Select
          
        Else
          ' search all the hierarchies for a match
          MyCommon.QueryStr = "select top 500 * from ( " & _
                              "select 1 as LevelType, NodeID as ID, NodeName as LocationName, lhn.ExternalID as ExternalID, -1 as LocationID,lh.Name as HierarchyName, lh.HierarchyID, 0 as Selected " & _
                              "from LHNodes lhn with (NoLock) " & _
                              "inner join LocationHierarchies lh with (NoLock) on lh.HierarchyID = lhn.HierarchyID  " & _
                              "Where lhn.ExternalID like '" & SearchString & "' or lhn.NodeName like '" & SearchString & "' " & _
                              "union " & _
                              "select 2 as LevelType, PKID as ID, loc.LocationName, loc.ExtLocationCode as ExternalID, loc.LocationID,lh.Name as HierarchyName, lh.HierarchyID,  " & _
                              "	(select count(*) from LocGroupItems with (NoLock) where LocationGroupID=" & LocationGroupID & " and LocationID = loc.LocationID and Deleted=0) as Selected  " & _
                              "from Locations loc with (NoLock) inner join LHContainer con on loc.LocationID= con.LocationID " & _
                              "inner join LHNodes lhn with (NoLock) on con.NodeID = lhn.NodeID " & _
                              "inner join LocationHierarchies lh with (NoLock) on lh.HierarchyId = lhn.HierarchyID " & _
                              "Where loc.Deleted=0 and loc.EngineID=" & EngineID & " and loc.ExtLocationCode like '" & SearchString & "' or loc.LocationName like '" & SearchString & "' " & _
                              ") ResultsTable " & _
                              "order by HierarchyID, LevelType, ExternalID, LocationName; "
          dt = MyCommon.LRT_Select
        End If
        
        rowCt = dt.Rows.Count
        If (dt.Rows.Count > 0) Then
          For Each row In dt.Rows
            NodeID = MyCommon.NZ(row.Item("ID"), 0)
            Name = MyCommon.NZ(row.Item("LocationName"), "")
            ExternalID = MyCommon.NZ(row.Item("ExternalID"), "&nbsp;")
            LocationID = MyCommon.NZ(row.Item("LocationID"), "&nbsp;")
            HierarchyID = MyCommon.NZ(row.Item("HierarchyID"), -1)
            Level = MyCommon.NZ(row.Item("LevelType"), 2)
            SelectedItem = IIf(MyCommon.NZ(row.Item("Selected"), 0) > 0, True, False)
            If (Level = 1) Then
              iconFileName = "folder.png"
            Else
              iconFileName = "store" & IIf(SelectedItem, "-on", "") & ".png"
            End If
            ShowNoItemMsg = False
            Send("<tr class=""" & Shaded & """>")
            Send("  <td style=""width:50px;""><a href=""javascript:locateItem('L" & Level & "ID" & NodeID & "');"" alt=""Locate this store within its hierarchy"" title=""Locate this store within its hierarchy"">Locate</a></td>")
            AnchorID = "link" & IIf(LocationID > -1, LocationID, NodeID) & "H" & HierarchyID
            ImgID = "img" & IIf(LocationID > -1, LocationID, NodeID) & "H" & HierarchyID
            If SelectedItem Then
              LinkTitle = Copient.PhraseLib.Detokenize("hierarchy.RemoveFromLocationGroup", LanguageID, LocationGroupID)
            Else
              LinkTitle = Copient.PhraseLib.Detokenize("hierarchy.AddToLocationGroup", LanguageID, LocationGroupID)
            End If
            
            If (LocationGroupID <> "" And CInt(LocationGroupID) > 0) Then
              If (LocationID = -1) Then
                Send("  <td style=""width:55px;""><a href=""javascript:handleSearchNodeAdjust('" & NodeID & "', '" & AnchorID & "'," & HierarchyID & ", 'ADD');"" id=""" & AnchorID & """ title=""" & Copient.PhraseLib.Detokenize("hierarchy.AddToLocationGroup", LanguageID, LocationGroupID) & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a><br />")
                Send("  <a href=""javascript:handleSearchNodeAdjust('" & NodeID & "', '" & AnchorID & "'," & HierarchyID & ", 'REMOVE');"" id=""" & AnchorID & """ title=""" & Copient.PhraseLib.Detokenize("hierarchy.RemoveFromLocationGroup", LanguageID, LocationGroupID) & """>" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & "</a></td>")
              Else
                If (SelectedItem) Then
                  Send("  <td style=""width:55px;""><a href=""javascript:handleSearchItemAdjust('" & LocationID & "', '" & AnchorID & "'," & HierarchyID & ",'" & MyCommon.NZ(row.Item("ExternalID"), "0") & "');"" id=""" & AnchorID & """ alt=""" & LinkTitle & """ title=""" & LinkTitle & """>" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & "</a></td>")
                Else
                  Send("  <td style=""width:55px;""><a href=""javascript:handleSearchItemAdjust('" & LocationID & "', '" & AnchorID & "'," & HierarchyID & ",'" & MyCommon.NZ(row.Item("ExternalID"), "0") & "');"" id=""" & AnchorID & """ alt=""" & LinkTitle & """ title=""" & LinkTitle & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a></td>")
                End If
              End If

            End If
            Send("  <td style=""width:35px;text-align:center;""><img src=""../images/" & iconFileName & """  id=""" & ImgID & """ />" & "</td>")
            Send("  <td>" & MyCommon.NZ(row.Item("HierarchyName"), "&nbsp;") & "</td>")
            Send("  <td>" & HighlightMatches(ExternalID, SearchString) & "</td>")
            Send("  <td>" & HighlightMatches(Name, SearchString) & "</td>")
            Send("</tr>")
            Shaded = IIf(Shaded = "shaded", "", "shaded")
          Next
        Else
          SearchString = SearchString.Replace("''", "&#39;")
                    Send("<tr><td colspan=""4""><center><i>" & Copient.PhraseLib.Lookup("phierarchy.noresultsfound", LanguageID) & " '" & OrigSearchString & "'</i></center></td></tr>")
          'Send("<tr><td colspan=""4"">" & MyCommon.QueryStr & "</td></tr>")
          ShowNoItemMsg = False
        End If
        If (ShowNoItemMsg) Then
          SearchString = SearchString.Replace("''" , "&#39;")
                    Send("<tr><td colspan=""4""><center><i>" & Copient.PhraseLib.Lookup("phierarchy.noresultsfound", LanguageID) & " '" & OrigSearchString & "'</i></center></td></tr>")
          'Send("<tr><td colspan=""4"">" & MyCommon.QueryStr & "</td></tr>")
        End If
      End If
      Send("</tbody>")
      Send("</table>")
      Send("<trailer>" & rowCt & "</trailer>")
    Catch ex As Exception
      Send(ex.ToString())
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  Function HighlightMatches(ByVal ColValue As String, ByVal Search As String) As String
    Dim FormattedCol As String = ColValue
    If (Not ColValue Is Nothing) Then
      FormattedCol = ColValue.Replace(Search, "<span class=""red"">" & Search & "</span>")
    End If
    Return FormattedCol
  End Function
  
  Public Sub handleSearchItemAdjust(ByVal ActionType As String, ByVal HierarchyID As Integer, ByVal LocationGroupID As String, ByVal LocationID As String)
    Dim dt As DataTable
    Dim retVal As String = ""
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
      MyCommon.Open_LogixRT()
      AdminUserID = Verify_AdminUser(MyCommon, Logix)
      If (ActionType = "add") Then
        MyCommon.QueryStr = "Select LocationID from LocGroupItems with (NoLock) where LocationGroupID=" & LocationGroupID & _
                            " and LocationID=" & LocationID & " and Deleted=0;"
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count = 0) Then
          MyCommon.QueryStr = "Insert into LocGroupItems with (RowLock) (LocationGroupID, LocationID, LastUpdate, Deleted, StatusFlag, TCRMAStatusFlag) " & _
                              "values (" & LocationGroupID & ", " & LocationID & ", getdate(), 0, 2, 2);"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(11, CLng(LocationGroupID), AdminUserID, "Added location " & Request.QueryString("extID") & " through hierarchy search")
        End If
        retVal = "ADDED"
      ElseIf (ActionType = "remove") Then
        MyCommon.QueryStr = "Delete from LocGroupItems with (RowLock) where LocationGroupID=" & LocationGroupID & " and LocationID=" & LocationID & " and Deleted=1;"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "Update LocGroupItems with (RowLock) set Deleted=1, StatusFlag=2, TCRMAStatusFlag=2 where Deleted=0 and LocationGroupID=" & LocationGroupID & " " & _
                            "and LocationID = " & LocationID & ";"
        MyCommon.LRT_Execute()

        If (MyCommon.RowsAffected > 0) Then
          MyCommon.Activity_Log(11, CLng(LocationGroupID), AdminUserID, "Removed location " & Request.QueryString("extID") & " through hierarchy search")
        End If
        retVal = "REMOVED"
      End If
      Sendb(retVal)
    Catch ex As Exception
      Send(ex.ToString())
    Finally
      MyCommon.Close_LogixRT()
    End Try
  End Sub
  
  Public Sub handleSearchNodeAdjust(ByVal ActionType As String, ByVal HierarchyID As Integer, ByVal LocationGroupID As String, ByVal NodeID As String)
    Dim PreCount, PostCount As Integer
    Dim foundItems As Boolean = False
    Try
      Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
      PreCount = GetProductGroupItemCount(LocationGroupID)
      TransmitGroups("N" & NodeID, LocationGroupID, foundItems, IIf(ActionType = "ADD", True, False))
      PostCount = GetProductGroupItemCount(LocationGroupID)
      Send(GetAssignedTrailer(PreCount, PostCount, foundItems))
    Catch ex As Exception
      Send(ex.ToString)
    Finally
    End Try
  End Sub
  
  Public Sub RecordSelectedNode(ByVal LgID As String, ByVal NodeID As String, ByVal AddToLG As Boolean)
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    Dim ChildList As ArrayList
    Dim NodeList As String() = Nothing
    
    Try
      If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
      AdminUserID = Verify_AdminUser(MyCommon, Logix)
      
      MyCommon.QueryStr = "dbo.pa_LocationGroupNodes_Update"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@LgId", SqlDbType.BigInt).Value = CLng(LgID)
      MyCommon.LRTsp.Parameters.Add("@NodeID", SqlDbType.BigInt).Value = CLng(NodeID)
      MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.BigInt).Value = AdminUserID
      MyCommon.LRTsp.Parameters.Add("@LinkDate", SqlDbType.DateTime).Value = Date.Now
      MyCommon.LRTsp.Parameters.Add("@AddToLG", SqlDbType.Bit).Value = IIf(AddToLG, 1, 0)
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      
      ' remove all child nodes of this node
      ChildList = GetChildNodes(NodeID)
      If (ChildList.Count > 0) Then
        NodeList = Array.ConvertAll(Of Object, String)(ChildList.ToArray, New Converter(Of Object, String)(AddressOf cString))
        MyCommon.QueryStr = "delete from LocationGroupNodes with (RowLock) where NodeID in (" & String.Join(",", NodeList) & ")"
        MyCommon.LRT_Execute()
      End If
    Catch ex As Exception
      Send(ex.ToString)
    End Try
  End Sub
  
  Sub Remove_Node(ByVal ExtNodeID As String, ByVal HierarchyID As Integer)

    Dim NodeID As Long
    Dim dst As DataTable
    Dim OrphanRow As DataRow


    'fetch the ID of the node to be removed
    MyCommon.QueryStr = "dbo.pa_LHA_GetNodeIDFE"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@HierarchyID", SqlDbType.Int).Value = HierarchyID
    MyCommon.LRTsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 100).Value = ExtNodeID
    dst = MyCommon.LRTsp_select
    MyCommon.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      NodeID = dst.Rows(0).Item("NodeID")
    End If

    'remove the node in question
    MyCommon.QueryStr = "dbo.pa_LHA_RemoveNode"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@NodeID", SqlDbType.BigInt).Value = NodeID
    MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()

    'now look for any oprhaned nodes
    MyCommon.QueryStr = "dbo.pa_LHA_OrphanedNodes"
    MyCommon.Open_LRTsp()
    dst = MyCommon.LRTsp_select()
    MyCommon.Close_LRTsp()

    'as long as there are any orphaned nodes, just keep deleting them until they are all gone
    While dst.Rows.Count > 0
      For Each OrphanRow In dst.Rows
        NodeID = OrphanRow.Item("NodeID")
        'delete the orphaned node
        MyCommon.QueryStr = "dbo.pa_LHA_RemoveNode"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@NodeID", SqlDbType.BigInt).Value = NodeID
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      Next
      dst = Nothing
      'look for any additional nodes
      MyCommon.QueryStr = "dbo.pa_LHA_OrphanedNodes"
      MyCommon.Open_LRTsp()
      dst = MyCommon.LRTsp_select()
      MyCommon.Close_LRTsp()
    End While

  End Sub

  Sub Remove_Hierarchy(ByVal ExtHierarchyID As String)

    ' clear the hierarchy (leaves, nodes, and top-level root node) with the external id of ExtHierarchyID
    MyCommon.QueryStr = "dbo.pa_LHA_RemoveHierarchy"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@ExtHierarchyID", SqlDbType.NVarChar, 100).Value = ExtHierarchyID
    MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()

  End Sub
  
  Sub HandleDelFromHierarchy(ByVal NodeID As String, ByVal ItemID As String)
    Dim isLocation As Boolean
    Dim dt As DataTable
    Dim HierID As Integer
    Dim ParentID As Integer
    Dim ExtID As String
    Dim DelNodeID As String
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    Dim refreshType As Integer = 0
    
    Try
      If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
      AdminUserID = Verify_AdminUser(MyCommon, Logix)

      If (Request.QueryString("isHier") = "1") Then
        MyCommon.QueryStr = "select ExternalId from LocationHierarchies with (NoLock) where HierarchyID =" & NodeID & ";"
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
          ExtID = MyCommon.NZ(dt.Rows(0).Item("ExternalId"), "")
          Remove_Hierarchy(ExtID)
          Send("<refresh>2</refresh>")
          Send("<node>" & NodeID & "</node>")
        Else
          Send(Copient.PhraseLib.Lookup("hierarchy.UnableToDelete", LanguageID))
        End If
      Else
        isLocation = (ItemID <> "" AndAlso ItemID.Length > 0 AndAlso ItemID.Substring(0, 1) = "I")
    
        If (isLocation) Then
          MyCommon.QueryStr = "delete from LHContainer with (RowLock) where nodeid = " & NodeID & " and locationid = " & ItemID.Substring(1)
          MyCommon.LRT_Execute()
          Send("<refresh>1</refresh>")
          Send("<node>" & NodeID & "</node>")
        Else
          If (ItemID = "") Then
            DelNodeID = NodeID
            refreshType = 3
          ElseIf (ItemID.Length > 0 AndAlso ItemID.Substring(0, 1) = "N") Then
            DelNodeID = ItemID.Substring(1)
            refreshType = 0
          Else
            DelNodeID = -1
            refreshType = 0
          End If
      
          MyCommon.QueryStr = "select HierarchyID, ExternalID, ParentID from LHNodes where nodeid = " & DelNodeID
          dt = MyCommon.LRT_Select
          If (dt.Rows.Count > 0) Then
            HierID = MyCommon.NZ(dt.Rows(0).Item("HierarchyID"), -1)
            ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), -1)
            ExtID = MyCommon.NZ(dt.Rows(0).Item("ExternalID"), "")
        
            Remove_Node(ExtID, HierID)
            Send("<refresh>" & refreshType & "</refresh>")
            Send("<node>" & DelNodeID & "</node>")
            Send("<parent>" & ParentID & "</parent>")
          Else
            Send(Copient.PhraseLib.Lookup("hierarchy.UnableToDelete", LanguageID))
          End If
      
        End If
      End If
    Catch ex As Exception
      Send(ex.ToString)
    End Try

  End Sub
  
</script>
