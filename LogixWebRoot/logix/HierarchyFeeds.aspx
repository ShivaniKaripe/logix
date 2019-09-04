 <%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    Dim action As String = ""
    Dim nodeId As Long
    Dim hierId As Integer
    Dim level As Integer
    Dim nodeName As String = ""
    Dim items As String = "" 
    Dim ParentID As Long
    Dim IsTopLevel As Boolean = False
    Dim PreCount As Integer = 0
    Dim PostCount As Integer = 0
    Dim foundItems As Boolean = False
    Dim LogMsg As String = ""
    Dim PromoID As Long
    Dim PromoSetID As Long
    Dim StartIndex As Integer = 0
    Dim itemSelectedPK As Integer = 0
    Dim locate As Boolean = False
    Dim PABFlag As String = ""
    Dim idList As String = ""
    Dim NodeIdList As String = ""
    Dim IncentiveProdGroupID As Integer = 0
    Dim SelectedNodeIDs As String = ""
    
    MyCommon.AppName = "HierarchyFeeds.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    action = Request.QueryString("action")
    NodeIdList = Request.Form("nodeidl")
    nodeId = MyCommon.Extract_Val(GetCgiValue("node"))
    level = MyCommon.Extract_Val(Request.QueryString("level"))
    nodeName = Request.QueryString("nodeName")
    hierId = MyCommon.Extract_Val(Request.QueryString("hierId"))
    items = Request.QueryString("items")
    PromoID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    PromoSetID = MyCommon.Extract_Val(Request.QueryString("PromoSetID"))
    StartIndex = MyCommon.Extract_Val(Request.QueryString("StartIndex"))
    itemSelectedPK = MyCommon.Extract_Val(Request.QueryString("itemSelectedPK"))
    locate = IIf(Request.QueryString("locate") <> "", True, False)
    PABFlag = Request.QueryString("PAB")
    SelectedNodeIDs = Request.QueryString("SelectedNodeIDs")
    
    If (nodeId > 0 AndAlso action <> "") Then
        Response.Expires = 0
        Response.Clear()
        Response.ContentType = "text/html"
        Select Case action
            Case "expand"
                GenerateDrilledNodeDiv(nodeId, level, PromoID, PromoSetID)
            Case "drill"
                GenerateDrilledNodeDiv(nodeId, level, PromoID, PromoSetID)
            Case "openFolder"
                GenerateDrilledNodeDiv(nodeId, level, PromoID, PromoSetID)
            Case "showItems"
                GenerateItemsDiv(nodeId, level, nodeName, Request.QueryString("pg"), Request.QueryString("buyerid"), StartIndex, itemSelectedPK, PABFlag, SelectedNodeIDs)
            Case "addNode"
                AddNode(level = 0, hierId, nodeId, nodeName)
            Case "deleteNode"
                DeleteNode(level = 0, hierId, nodeId)
            Case "deleteItems"
                DeleteItems(nodeId, items)
            Case "showAvailItems"
                GenerateAvailableItems(nodeId, nodeName)
            Case "removeAll"
                RemoveAllFromProductGroups(Request.QueryString("pg"), AdminUserID)
                UpdateProductGroup(Request.Form("pg"), True)
            Case "getCount"
                UpdateProductCount(NodeIdList)
            Case Else
                Send(Request.RawUrl)
        End Select
    ElseIf (action = "getCount") Then
        UpdateProductCount(NodeIdList)
    ElseIf (nodeId = 0 AndAlso action = "drill") Then
        GenerateDrilledNodeDiv(nodeId, level, PromoID, PromoSetID)
    ElseIf (nodeId <= 0 AndAlso action = "showItems") Then
        GenerateItemsDiv(nodeId, level, nodeName, Request.QueryString("pg"), Request.QueryString("buyerid"), StartIndex, itemSelectedPK, PABFlag, SelectedNodeIDs)
    ElseIf (nodeId = 0 AndAlso action = "addNode") Then
        AddNode(True, 0, 0, nodeName)
    ElseIf (action = "assignNodes") Then
        PreCount = GetProductGroupItemCount(Request.Form("pg"))
        TransmitGroups(Request.Form("ids"), Request.Form("pg"), foundItems, True)
        If (Request.Form("sel") <> "") Then
            Long.TryParse(Request.Form("sel"), ParentID)
        End If
        If (ParentID <= 0) Then ParentID = FindParentNodeID(Request.Form("ids"), IsTopLevel)

        GenerateItemsDiv(ParentID, IIf(IsTopLevel, 1, 0), "", Request.Form("pg"), "-1", PABFlag, SelectedNodeIDs)
        PostCount = GetProductGroupItemCount(Request.Form("pg"))
        UpdateProductGroup(Request.Form("pg"), False)
        Send(GetAssignedTrailer(PreCount, PostCount, foundItems))
    ElseIf (action = "unassignNodes") Then
        PreCount = GetProductGroupItemCount(Request.Form("pg"))
        TransmitGroups(Request.Form("ids"), Request.Form("pg"), foundItems, False)
        If (Request.Form("sel") <> "") Then
            Long.TryParse(Request.Form("sel"), ParentID)
        End If
        If (ParentID <= 0) Then ParentID = FindParentNodeID(Request.Form("ids"), IsTopLevel)

        GenerateItemsDiv(ParentID, IIf(IsTopLevel, 1, 0), "", Request.Form("pg"), "-1", PABFlag, SelectedNodeIDs)
        PostCount = GetProductGroupItemCount(Request.Form("pg"))
        UpdateProductGroup(Request.Form("pg"), False)
        Send(GetAssignedTrailer(PreCount, PostCount, foundItems))
    ElseIf (action = "removeAll") Then
        RemoveAllFromProductGroups(Request.QueryString("pg"), AdminUserID)
        UpdateProductGroup(Request.Form("pg"), True)
    ElseIf (action = "findMatches") Then
        FindSearchItems(Request.QueryString("search"), MyCommon.Extract_Val(Request.QueryString("stype")), _
                        Request.QueryString("pg"), MyCommon.Extract_Val(Request.QueryString("nodeid")), _
                        MyCommon.Extract_Val(Request.QueryString("hierid")), PABFlag)
    ElseIf (action = "findMatchesItemAttrb") Then
        FindSearchItemAttributes(Request.QueryString("ItemAttribVal1"), Request.QueryString("Attr1Selection"), Request.QueryString("low1Value"), Request.QueryString("high1Value"), _
                        Request.QueryString("ItemAttribVal2"), Request.QueryString("Attr2Selection"), Request.QueryString("low2Value"), Request.QueryString("high2Value"), _
                        Request.QueryString("ItemAttribVal3"), Request.QueryString("Attr3Selection"), Request.QueryString("low3Value"), Request.QueryString("high3Value"), _
                        Request.QueryString("pg"), MyCommon.Extract_Val(Request.QueryString("nodeid")), _
                        MyCommon.Extract_Val(Request.QueryString("hierid")))
    ElseIf (action = "CheckItemAttrb") Then
        BindDstostartvalues(Request.QueryString("ItemAttrbValue"), Request.QueryString("pg"), MyCommon.Extract_Val(Request.QueryString("nodeid")), _
                            MyCommon.Extract_Val(Request.QueryString("hierid")), Request.QueryString("AttributeId"))
    ElseIf (action = "LastItemAttrbSearchCriteria") Then
        AssignPrevValues(Request.QueryString("pg"), MyCommon.Extract_Val(Request.QueryString("nodeid")), _
                       MyCommon.Extract_Val(Request.QueryString("hierid")), Request.QueryString("AttributeId"))
    ElseIf (action = "BindlastItemAttributevalues") Then
        SendLastItemAttrbValues()
    ElseIf (action = "LinkMatchestoHierarchy") Or (action = "RemoveMatchesFromHierarchy") Then
        UpdateItemAttributeProductsToHierarchy(action, AdminUserID)
    ElseIf (action = "handleSearchItemAdjust") Then
        handleSearchItemAdjust(Request.QueryString("type"), Request.QueryString("hID"), Request.QueryString("pg"), Request.QueryString("productID"))
    ElseIf (action = "handleSearchNodeAdjust") Then
		If (Not String.IsNullOrEmpty(Request.QueryString("Linking"))) Then
			IsLinkedPage = IIf(Request.QueryString("Linking") = "1", True, False)
		End If
        handleSearchNodeAdjust(Request.QueryString("type"), Request.QueryString("hID"), Request.QueryString("pg"), Request.QueryString("nodeID"))
    ElseIf (action = "delFromHierarchy") Then
        HandleDelFromHierarchy(Request.QueryString("nodeID"), Request.QueryString("itemID"))
    ElseIf (action = "linkToGroup") Then
        HandleLinkToGroup(MyCommon.Extract_Val(Request.Form("pg")), Request.Form("ids"), AdminUserID)
    ElseIf (action = "removeLinkToGroup") Then
        HandleRemoveLinkToGroup(MyCommon.Extract_Val(Request.Form("pg")), Request.Form("ids"), AdminUserID)
    ElseIf (action = "excludeFromGroup") Then
        If (Request.Form("sel") <> "") Then
            Long.TryParse(Request.Form("sel"), nodeId)
        End If
        HandleExcludeFromGroup(MyCommon.Extract_Val(Request.Form("pg")), nodeId, Request.Form("ids"), AdminUserID)
    ElseIf (action = "removeExclusion") Then
        HandleRemoveExclusion(MyCommon.Extract_Val(Request.Form("pg")), Request.Form("ids"), AdminUserID)
    ElseIf (action = "AssignSigns") Then
        If (Request.Form("sel") <> "") Then
            Long.TryParse(Request.Form("sel"), nodeId)
        End If
        HandleAssignSigns(Request.Form("hiesign"), nodeId, Request.QueryString("hID"), AdminUserID)
    ElseIf (action = "GenerateDivForSigns") Then
        If (Request.Form("sel") <> "") Then
            Long.TryParse(Request.Form("sel"), nodeId)
        End If
        HandleGenerateDivForSigns(nodeId, Request.QueryString("hID"))
    Else
        Send("action: " & action)
        Send("node: " & nodeId)
        Send("level: " & level)
        Send("nodeName: " & nodeName)
        Send(Request.RawUrl)
        Send("<b>" & Copient.PhraseLib.Lookup("term.noarguments", LanguageID) & "</b>")
    End If
    Try
    Response.Flush()
    Response.End()
    Catch ex As Exception        
        Select Case action
            Case "getCount"
                'We can safely ignore this case, as writing a Response is not required when the request is aborted 
        End Select
    End Try
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
	Dim IsLinkedPage As Boolean = False

    Sub GenerateDrilledNodeDiv(ByVal nodeId As Long, ByVal level As Integer, ByVal PromoID As Long, ByVal PromoSetID As Long, Optional ByVal searchPathIDs As String() = Nothing, Optional ByVal ProductGroupID As Long = 0)
        Dim dt As DataTable
        Dim row As DataRow
        Dim newLevel As Integer = level
        Dim newLeft As Integer
        Dim newNodeId As Long
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
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            Dim Logix As New Copient.LogixInc
            Dim AdminUserID = Verify_AdminUser(MyCommon, Logix)
            Dim ViewRestricted As Boolean = Logix.UserRoles.ViewRestrictedHierarchyNodes

            ProductGroupID = MyCommon.Extract_Val(Request.Form("pg"))
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
                sQuery &= IIf(ViewRestricted, " where HierarchyID=" & nodeId & ";", " where HierarchyID=" & nodeId & " AND (Restricted < 1 or Restricted IS NULL);")
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
                Send(" <span id=""indent" & IIf(level <= 1, "H", "") & newNodeId & """ style=""line-height:18px;"" onclick=""highlightNode(" & newNodeId & ", " & newLevel & ")"">")
                Send("  <img src=""/images/clear.png"" style=""height:1px;width:" & newLeft & "px;"" />")
                Send("  <a href=""#""><img id=""imgfldr" & IIf(level <= 1, "H", "") & newNodeId & """ src=""/images/" & IIf(level <= 1, "hierarchy", "folder" & iconType(excludeIndex)) & ".png"" alt=""" & FolderAltText & """ title=""" & FolderAltText & """ border=""0"" /></a>")
                Send("  <span id=""name" & IIf(level <= 1, "H", "") & newNodeId & """ style=""left:5px;"">" & name & "</span>")
                Send("  <input type=""hidden"" id=""hierarchy" & IIf(level <= 1, "H", "") & newNodeId & """ name=""hierarchy" & IIf(level <= 1, "H", "") & newNodeId & """ value=""" & hierId & """ />")
                Send("  <br class=""zero"" />")
                Send(" </span>")

                If (Not searchPathIDs Is Nothing AndAlso searchPathIDs.GetUpperBound(0) >= level AndAlso newNodeId = searchPathIDs(level)) Then
                    Send(" <span id=""hId" & IIf(level <= 1, "H", "") & newNodeId & """ style=""display:inline;"">")
                    GenerateDrilledNodeDiv(newNodeId, level + 1, PromoID, PromoSetID, searchPathIDs, ProductGroupID)
                    Send(" </span>")
                Else
                    Send(" <span id=""hId" & IIf(level <= 1, "H", "") & newNodeId & """ style=""display:none;"">")
                    Send(" </span>")
                End If

                Send("</span>")
            Next
            'Send("<label style=""display:none"">productcount:" & GetProductCount(nodeId.ToString()) & "</label>")

        Catch ex As Exception
            Send(Copient.PhraseLib.Lookup("hierarchy.unabletoloadnodes", LanguageID))
            MyCommon.Error_Processor(, ex.ToString(), "HierarchyFeeds.aspx", , )
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub GenerateNodeDiv(ByVal nodeId As Long, ByVal level As Integer, ByVal PromoID As Long, ByVal PromoSetID As Long)
        Dim dt As DataTable
        Dim row As DataRow
        Dim newLevel As Integer = level + 1
        Dim newLeft As Integer
        Dim newNodeId As Long
        Dim hierId As Integer
        Dim name As String = ""
        Dim sQuery As String = ""
        Dim IdType As String = ""
        Dim iconType() As String = {"", "-green", "-red", "-down", "-purple"}
        Dim excludeIndex As Integer = 0
        Dim Resyncer As New Copient.HierarchyResync(MyCommon, "HierarchyFeeds", "Hierarchy.txt")
        Dim ExtHierarchyID As String = ""
        Dim ExtNodeID As String = ""
        Dim ProductGroupID As Long = 0
        Dim FolderAltText As String = ""

        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            Dim Logix As New Copient.LogixInc
            Dim AdminUserID = Verify_AdminUser(MyCommon, Logix)
            Dim ViewRestricted As Boolean = Logix.UserRoles.ViewRestrictedHierarchyNodes
            ProductGroupID = MyCommon.Extract_Val(Request.Form("pg"))
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
            If (level = 0) Then
                sQuery &= " from ProdHierarchies with (NoLock)" & IIf(ViewRestricted, "", " where (Restricted < 1 or Restricted IS NULL)")
            ElseIf (level = 1) Then
                sQuery &= ", NodeID "
                sQuery &= " from PHNodes with (NoLock) "
                sQuery &= "where HierarchyID=" & nodeId & " and ParentID=0 "
            Else
                sQuery &= ", NodeID "
                sQuery &= " from PHNodes with (NoLock) "
                sQuery &= "where ParentID=" & nodeId & " "
            End If
            sQuery &= "order by ExternalID;"

            MyCommon.QueryStr = sQuery
            dt = MyCommon.LRT_Select()
            For Each row In dt.Rows
                newNodeId = row.Item("NodeID")
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

                Send("<span class=""hrow"" id=""node" & newNodeId & """ name=""node" & newNodeId & """>")
                Send(" <span id=""indent" & newNodeId & """ style=""line-height:18px;"" onclick=""highlightNode(" & newNodeId & ", " & newLevel & ")"">")
                Send("  <img src=""/images/clear.png"" style=""height:1px;width:" & newLeft & "px;"" />")
                Send("  <a href=""#""><img id=""imgfldr" & newNodeId & """ src=""/images/" & IIf(level <= 1, "hierarchy", "folder" & iconType(excludeIndex)) & ".png"" alt=""" & FolderAltText & """ title=""" & FolderAltText & """ border=""0"" /></a>")
                Send("  <span id=""name" & newNodeId & """ style=""left:5px;"" onclick=""highlightNode(" & newNodeId & ", " & newLevel & ")"">" & name & "</span>")
                Send("  <input type=""hidden"" id=""hierarchy" & newNodeId & """ name=""hierarchy" & newNodeId & """ value=""" & hierId & """ />")
                Send("  <br class=""zero"" />")
                Send(" </span>")
                Send(" <span id=""hId" & newNodeId & """ style=""display:none;""></span>")
                Send("</span>")
            Next
        Catch ex As Exception
            Send(Copient.PhraseLib.Lookup("hierarchy.unabletoloadnodes", LanguageID))
            MyCommon.Error_Processor(, ex.ToString(), "HierarchyFeeds.aspx", , )
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub


    Sub GenerateItemsDiv(ByVal nodeId As Long, ByVal level As Integer, ByVal nodeName As String, ByVal pgID As String, ByVal BuyerID As String, Optional ByVal StartIndex As Integer = 1, Optional ByVal itemSelectedPK As Integer = 0, Optional ByVal PABFlag As String = "", Optional ByVal SelectedNodeIDs As String = "")
        Dim dt As DataTable
        'Dim dt1 As DataTable
        Dim row As DataRow
        Dim ExtProductID As String = ""
        Dim ProductID As Integer
        Dim ProductDesc As String = ""
        Dim iconFileName As String = ""
        Dim ProductTypeID As Integer = -1
        Dim PKID As Long
        Dim i As Integer = 0
        Dim SelectedItem As Boolean = False
        Dim sQuery As String = ""
        Dim highlightedRow As Integer = -1
        Dim totalItemCt As Integer = 0
        Dim totalSelCt As Integer = 0
        Dim DisplayID As String = ""
        Dim IdType As String = ""
        Dim CurrentSortCode As Integer = SORT_CODES.SORT_ID_ASC
        Dim SortCol As String = "ExtProductID"
        Dim SortDir As String = "ASC"
        Dim SortDirReverse As String = "DESC"
        Dim Col1Icon As String = "&nbsp;"
        Dim Col2Icon As String = "&nbsp;"
        Dim IdColName As String = "ExtProductID"
        Dim Resyncer As New Copient.HierarchyResync(MyCommon, "HierarchyFeeds", "Hierarchy.txt")
        Dim ExtHierarchyID As String = ""
        Dim ExtNodeID As String = ""
        Dim FolderAltText As String = ""
        Dim ResultsCount As Integer = 0
        Dim ResultsPerPage As Integer = 100


        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            Dim Logix As New Copient.LogixInc
            Dim AdminUserID = Verify_AdminUser(MyCommon, Logix)
            Dim ViewRestricted As Boolean = Logix.UserRoles.ViewRestrictedHierarchyNodes
            Dim BID As Integer = Convert.ToInt32(BuyerID)
            BID = IIf(BID < 0, 0, BID)
            ''get node id's associated with this buyer
            'MyCommon.QueryStr = " SELECT DepartmentID From BuyersDepts Where BuyerID=" & BuyerID
            'dt1 = MyCommon.LRT_Select()
            ' figure out the sorted column and its order
            IdType = MyCommon.Fetch_SystemOption(62)
            IdColName = IIf(IdType = "2", "DisplayID", "ExtProductID")
            If Not Integer.TryParse(Request.QueryString("sort"), CurrentSortCode) Then CurrentSortCode = SORT_CODES.SORT_ID_ASC
            SortCol = IIf(CurrentSortCode = SORT_CODES.SORT_ID_ASC OrElse CurrentSortCode = SORT_CODES.SORT_ID_DESC, IdColName, "Description")
            SortDir = IIf(CurrentSortCode = SORT_CODES.SORT_ID_ASC OrElse CurrentSortCode = SORT_CODES.DESCRIPTION_ASC, "ASC", "DESC")
            SortDirReverse = IIf(SortDir = "ASC", "DESC", "ASC")
            If (SortCol = IdColName AndAlso SortDir = "ASC") Then Col1Icon = "&#9650;"
            If (SortCol = IdColName AndAlso SortDir = "DESC") Then Col1Icon = "&#9660;"
            If (SortCol = "Description" AndAlso SortDir = "ASC") Then Col2Icon = "&#9650;"
            If (SortCol = "Description" AndAlso SortDir = "DESC") Then Col2Icon = "&#9660;"
            Col1Icon = "<span style=""color:#808080;padding-left:5px;width:20px;"">" & Col1Icon & "</span>"
            Col2Icon = "<span style=""color:#808080;padding-left:10px;width:20px;"">" & Col2Icon & "</span>"
            Dim IsBuyerAssociated As Boolean

            If (level = 0) Then
                MyCommon.QueryStr = "SELECT COUNT(*) AS ResultsCount FROM ProdHierarchies WITH (NoLock);"
                dt = MyCommon.LRT_Select()
                ResultsCount = dt.Rows(0).Item("ResultsCount")
                sQuery = "SELECT * FROM (" & _
                          "	SELECT ROW_NUMBER() OVER (ORDER BY ExternalID) AS RowNumber, 1 AS ItemType, HierarchyID AS PKID, -1 AS ProductID, ExternalID AS ExtProductID, DisplayID, -2 AS ProductTypeID, Name AS Description, 0 AS Selected" & _
                          IIf(ViewRestricted, "	FROM ProdHierarchies WITH (NoLock)", "	FROM ProdHierarchies WITH (NoLock) where (Restricted < 1 or Restricted IS NULL)") & _
                          "	) AS P1 " & _
                          "WHERE P1.RowNumber BETWEEN " & StartIndex & " AND " & (StartIndex + (ResultsPerPage - 1)) & " " & _
                              "ORDER BY " & SortCol & " " & SortDir & ";"
                MyCommon.QueryStr = sQuery
                dt = MyCommon.LRT_Select()
            ElseIf (level = 1) Then
                MyCommon.QueryStr = "SELECT COUNT(*) AS ResultsCount FROM PHNodes WITH (NoLock) WHERE ParentID=0 AND HierarchyID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, "") & ";"
                dt = MyCommon.LRT_Select()
                ResultsCount = dt.Rows(0).Item("ResultsCount")
                sQuery = "SELECT * FROM (" & _
                          "	SELECT ROW_NUMBER() OVER (ORDER BY ExternalID) AS RowNumber, 1 AS ItemType, NodeID AS PKID, -1 AS ProductID, ExternalID AS ExtProductID, DisplayID, -1 AS ProductTypeID, Name AS Description, 0 AS Selected" & _
                          "	FROM PHNodes WITH (NoLock)" & _
                          IIf(ViewRestricted, "	WHERE ParentID=0 AND HierarchyID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, ""), "	WHERE ParentID=0 AND (Restricted < 1 or Restricted IS NULL) AND HierarchyID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, "")) & _
                          "	) AS P1 " & _
                          "WHERE P1.RowNumber BETWEEN " & StartIndex & " AND " & (StartIndex + (ResultsPerPage - 1)) & " " & _
                          "ORDER BY ProductTypeID, " & SortCol & " " & SortDir & ";"
                MyCommon.QueryStr = sQuery
                dt = MyCommon.LRT_Select()
            Else

                If PABFlag = "1" Then
                    MyCommon.QueryStr = "SELECT COUNT(*) As ResultsCount FROM PHNodes WITH (NoLock) WHERE ParentID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, "")
                    dt = MyCommon.LRT_Select()
                    ResultsCount = dt.Rows(0).Item("ResultsCount")
                    sQuery = "SELECT * FROM ( " & _
                              "	SELECT ROW_NUMBER() OVER (ORDER BY " & IIf(SortCol = "Description", "Name", "ExternalID") & " " & SortDir & ") AS RowNumber, 1 AS ItemType, NodeID AS PKID, -1 AS ProductID, ExternalID AS ExtProductID, DisplayID, -1 AS ProductTypeID, Name AS Description, 0 AS Selected " & _
                              "	FROM PHNodes WITH (NoLock) " & _
                              IIf(ViewRestricted, "	WHERE ParentID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, ""), "	WHERE (Restricted < 1 or Restricted IS NULL) AND ParentID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, "")) & _
                              IIf(itemSelectedPK <> 0, " AND NodeID=" & itemSelectedPK, "") & _
                     "	) AS P1 " & _
                     "WHERE P1.RowNumber BETWEEN " & StartIndex & " AND " & (StartIndex + (ResultsPerPage - 1)) & " " & _
                     "ORDER BY ProductTypeID, " & SortCol & " " & SortDir & ";"
                Else
                    MyCommon.QueryStr = "SELECT " & _
                                "(SELECT COUNT(*) NodeID FROM PHNodes WITH (NoLock) WHERE ParentID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, "") & ") " & _
                                " + " & _
                                "(SELECT COUNT(*) PKID FROM PHContainer WITH (NoLock) WHERE NodeID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, "") & ") AS ResultsCount;"
                    dt = MyCommon.LRT_Select()
                    ResultsCount = dt.Rows(0).Item("ResultsCount")

                    sQuery = "SELECT * FROM ( " & _
                        "	SELECT ROW_NUMBER() OVER (ORDER BY " & IIf(SortCol = "Description", "Name", "ExternalID") & " " & SortDir & ") AS RowNumber, 1 AS ItemType, NodeID AS PKID, -1 AS ProductID, ExternalID AS ExtProductID, DisplayID, -1 AS ProductTypeID, Name AS Description, 0 AS Selected " & _
                        "	FROM PHNodes WITH (NoLock) " & _
                        IIf(ViewRestricted, "	WHERE ParentID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, ""), "	WHERE (Restricted < 1 or Restricted IS NULL) AND ParentID=" & nodeId & IIf(itemSelectedPK > 0, " AND NodeID=" & itemSelectedPK, "")) & _
                        IIf(itemSelectedPK <> 0, " AND NodeID=" & itemSelectedPK, "") & _
                  "	UNION " & _
                  "	SELECT ROW_NUMBER() OVER (ORDER BY " & IIf(SortCol = "Description", "Description", "ExtProductID") & " " & SortDir & ") AS RowNumber, 2 AS ItemType, CON.PKID, PROD.ProductID, PROD.ExtProductID, '' AS DisplayID, PROD.ProductTypeID, PROD.Description, (SELECT COUNT(*) FROM ProdGroupItems WITH (NoLock) WHERE ProductGroupID=" & pgID & " AND ProductID=PROD.ProductID AND Deleted=0) AS Selected " & _
                  "	FROM PHContainer CON WITH (NoLock) " & _
                  "	LEFT JOIN Products PROD WITH (NoLock) ON CON.ProductID=PROD.ProductID " & _
                  IIf(ViewRestricted, "	WHERE CON.NodeID=" & nodeId, "	WHERE (CON.Restricted < 1 or CON.Restricted IS NULL) AND CON.NodeID=" & nodeId) & _
                  IIf(itemSelectedPK <> 0, " AND CON.PKID=" & itemSelectedPK, "") & _
                          "	) AS P1 " & _
                          "WHERE P1.RowNumber BETWEEN " & StartIndex & " AND " & (StartIndex + (ResultsPerPage - 1)) & " " & _
                          "ORDER BY ProductTypeID, " & SortCol & " " & SortDir & ";"
                End If
                MyCommon.QueryStr = sQuery
                dt = MyCommon.LRT_Select()
            End If


            'Draw the pagination box
            Dim CurrentPage As Integer = Math.Ceiling(StartIndex / ResultsPerPage)
            Dim TotalPages As Integer = Math.Ceiling(ResultsCount / ResultsPerPage)
            Dim LastWholePage As Integer = Math.Floor(ResultsCount / ResultsPerPage)
            Dim FinalIndex As Integer = 0
            'Dim IsBuyerAssociated As Boolean = True
            'If nodeId <> 0 AndAlso BID > 0 Then
            '    IsBuyerAssociated = MyCommon.IsParentNodeAssociatedWithBuyer(nodeId, BID)
            'End If
            FinalIndex = IIf(TotalPages > LastWholePage, (LastWholePage * ResultsPerPage) + 1, (LastWholePage * ResultsPerPage) - (ResultsPerPage - 1))
            FinalIndex = (FinalIndex - StartIndex)


            Send("<div style=""background-color:#cccccc;border-bottom:1px solid white;text-align:center;position:absolute;padding:1px 0;width:100%;"">")
            If itemSelectedPK > 0 And ResultsCount <= 1 Then
                'Single searched item, so replace unnecessary pagination controls with message
                Send("  Only the located item is shown below.  Click a folder above to see its full contents.")
            Else
                'Normal pagination
                If (StartIndex <= 1) Then
                    Send("  <span class=""grey"">|◄ " & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</span>")
                    Send("  <span class=""grey"">◄ " & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</span>")
                Else
                    Send("  <a href=""javascript:pageResults(" & nodeId & "," & level & ",-9999999);"">|◄ " & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a>")
                    Send("  <a href=""javascript:pageResults(" & nodeId & "," & level & ",-" & ResultsPerPage & ");"">◄ " & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a>")
                End If
                Send("  &nbsp;[" & IIf(ResultsCount > 0, StartIndex, "0") & " - " & IIf(CurrentPage >= TotalPages, ResultsCount, (StartIndex + (ResultsPerPage - 1))) & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & ResultsCount & "]&nbsp; ")
                If (CurrentPage >= TotalPages) Then
                    Send("  <span class=""grey"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & " ►</span>")
                    Send("  <span class=""grey"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & " ►|</span>")
                Else
                    Send("  <a href=""javascript:pageResults(" & nodeId & "," & level & "," & ResultsPerPage & ");"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & " ►</a>")
                    Send("  <a href=""javascript:pageResults(" & nodeId & "," & level & "," & FinalIndex & ");"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & " ►|</a>")
                End If
            End If
            Send("</div>")

            'Draw the results
            Send("<table style=""width:100%;font-family:arial;font-size:11px;padding-top:16px;"" summary=""" & Copient.PhraseLib.Lookup("term.items", LanguageID) & """>")
            Send("<thead>")
            Send("<tr>")
            If (pgID <> "" AndAlso CInt(pgID) > 0 Or PABFlag = "1") Then
                Send("<th scope=""col"" style=""background-color:#e0e0e0;width:30px;""><input type=""checkbox"" id=""chkAll"" name=""chkAll"" title=""" & Copient.PhraseLib.Lookup("hierarchy.SelectAllItems", LanguageID) & """ onclick=""handleAllItems(" & level & ");""" & IIf(level = 0, " disabled=""disabled""", "") & " /></th>")
            End If
            Send("<th scope=""col"" style=""background-color:#e0e0e0;width:30px;"">&nbsp;</th>")
            Send("<th scope=""col"" style=""background-color:#e0e0e0;cursor:pointer;""" & " onclick=""sortByColumn(1," & CurrentSortCode & ");""" & ">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & Col1Icon & "</th>")
            Send("<th scope=""col"" style=""background-color:#e0e0e0;cursor:pointer;""" & " onclick=""sortByColumn(2," & CurrentSortCode & ");""" & ">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & Col2Icon & "</th>")
            Send("</tr>")
            Send("</thead>")
            Send("<tbody>")

            If (dt.Rows.Count = 0) Then
                If (pgID <> "" AndAlso CInt(pgID) > 0 Or PABFlag = "1") Then
                    Send("<tr><td style=""width:30px;""></td><td style=""width:30px;""></td><td style=""width:100px;""></td><td style=""width:200px;""></td></tr>")
                    Send("<tr><td colspan=""4"">" & Copient.PhraseLib.Detokenize("hierarchy.NoChildNodes", LanguageID, nodeName) & "</td></tr>") 'No items assigned for node <br />{0}
                Else
                    Send("<tr></td><td style=""width:30px;""></td><td style=""width:100px;""></td><td style=""width:200px;""></td></tr>")
                    Send("<tr><td colspan=""3"">" & Copient.PhraseLib.Detokenize("hierarchy.NoChildNodes", LanguageID, nodeName) & "</td></tr>") 'No items assigned for node <br />{0}
                End If
            Else
                Dim SelectedNodeIDsArray As String()
                If (Not SelectedNodeIDs Is Nothing AndAlso SelectedNodeIDs <> "") Then
                    SelectedNodeIDsArray = SelectedNodeIDs.Split(",")

                End If

                Dim dtBuyers As New DataTable()
                If BID > 0 Then
                    'Find all the Nodes associated with the Buyer
                    If IsNothing(Session("AllBuyerNodes" + BuyerID.ToString())) Then
                        dtBuyers.Columns.Add("NodeID")
                        MyCommon.AppName = "phierarchytree.aspx"
                        MyCommon.QueryStr = "dbo.pa_GetNodesAssociatedWithBuyer"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@BuyerId", SqlDbType.Int).Value = BID
                        dtBuyers = MyCommon.LRTsp_select
                        If dtBuyers.Rows.Count > 0 Then
                            Session("AllBuyerNodes" + BuyerID.ToString()) = dtBuyers
                        End If
                    Else
                        dtBuyers = Session("AllBuyerNodes" + BuyerID.ToString())
                    End If
                End If

                For Each row In dt.Rows
                    PKID = MyCommon.NZ(row.Item("pkid"), 0)

                    If BID > 0 Then
                        If dtBuyers.Rows.Count > 0 AndAlso dtBuyers.Select("NodeID = " & PKID).Length > 0 Then
                            IsBuyerAssociated = True
                        Else
                            IsBuyerAssociated = False
                        End If
                    Else
                        IsBuyerAssociated = True
                    End If

                    ProductID = MyCommon.NZ(row.Item("ProductID"), -1)
                    ExtProductID = MyCommon.NZ(row.Item("ExtProductID"), "")
                    DisplayID = MyCommon.NZ(row.Item("DisplayID"), "")
                    ProductDesc = MyCommon.NZ(row.Item("Description"), "[No Description]")
                    ProductTypeID = MyCommon.NZ(row.Item("ProductTypeID"), -1)
                    SelectedItem = IIf(MyCommon.NZ(row.Item("Selected"), 0) > 0, True, False)
                    FolderAltText = ""

                    If (SelectedItem) Then
                        totalSelCt += 1
                    End If

                    Select Case MyCommon.NZ(row.Item("ProductTypeID"), 0)
                        Case -2
                            iconFileName = "hierarchy.png"
                        Case -1
                            iconFileName = "folder.png"
                            ' determine if this is linked or exclude folder.
                            If pgID > 0 Or PABFlag = "1" Then
                                Resyncer.FindHierarchyExtIDs(PKID, ExtHierarchyID, ExtNodeID)
                                If Resyncer.IsNodeLinked(pgID, ExtHierarchyID, ExtNodeID, True) Then
                                    iconFileName = "folder-green.png"
                                    FolderAltText = Copient.PhraseLib.Detokenize("hierarchy.LinkedToProductGroup", LanguageID, pgID)
                                ElseIf Resyncer.IsNodeLinkedWithAttribute(pgID, ExtHierarchyID, ExtNodeID) Then
                                    iconFileName = "folder-purple.png"
                                    FolderAltText = Copient.PhraseLib.Detokenize("hierarchy.LinkedToProductGroup", LanguageID, pgID)
                                ElseIf Resyncer.IsNodeExcluded(pgID, ExtHierarchyID, ExtNodeID) Then
                                    iconFileName = "folder-red.png"
                                    FolderAltText = Copient.PhraseLib.Detokenize("hierarchy.ExcludedFromProductGroup", LanguageID, pgID)
                                ElseIf Resyncer.IsChildNodeLinked(pgID, ExtHierarchyID, ExtNodeID) Then
                                    iconFileName = "folder-down.png"
                                    FolderAltText = Copient.PhraseLib.Lookup("hierarchy.ChildLinkedToProductGroup", LanguageID)
                                End If
                            End If
                        Case 2
                            If Resyncer.IsItemExcluded(pgID, nodeId, ProductID) Then
                                iconFileName = "dept-red.png"
                            Else
                                iconFileName = "dept" & IIf(SelectedItem, "-on", "") & ".png"
                            End If
                        Case Else
                            If Resyncer.IsItemExcluded(pgID, nodeId, ProductID) Then
                                iconFileName = "upc-red.png"
                            Else
                                iconFileName = "upc" & IIf(SelectedItem, "-on", "") & ".png"
                            End If
                            totalItemCt += 1
                    End Select

                    If (ProductTypeID = -1 Or ProductTypeID = -2) Then
                        Send("<tr class=""hrow""" & IIf((level = 0 OrElse Not IsBuyerAssociated), " style=""color:Gray;font-style:italic;""", "") & "id=""itemRow" & i & """ onclick=""highlightItem(" & i & ");"" ondblclick=""handleItemDblClick(" & PKID & ", " & (level + 1) & ", " & GetParentNode(PKID, IIf(ProductTypeID = -2, True, False)) & ");"">")
                    Else
                        Send("<tr class=""hrow""" & IIf((level = 0 OrElse Not IsBuyerAssociated), " style=""color:Gray;font-style:italic;display:" & IIf(PABFlag = "1", "none", "block") & """", "") & " id=""itemRow" & i & """ onclick=""highlightItem(" & i & ");"">")
                    End If


                    If (MyCommon.NZ(row.Item("ItemType"), -1) = 1) Then
                        ' work-around in case the name is found in the external id then remove the name
                        If (ExtProductID.IndexOf(ProductDesc) > -1) Then
                            ExtProductID = ExtProductID.Replace("-" & ProductDesc, "")
                        End If
                    End If
                    Dim ExternalIDText As String = Server.HtmlEncode(GetIdText(ExtProductID, DisplayID, ProductTypeID, IdType))
                    If (pgID <> "" AndAlso CInt(pgID) > 0 Or PABFlag = "1") Then

                        Send("  <td><input type=""checkbox"" " & IIf(Not SelectedNodeIDsArray Is Nothing AndAlso SelectedNodeIDsArray.Contains(PKID.ToString()), " checked=""checked"" ", "") & " id=""chk" & i & """ name=""chk" & ExternalIDText & """ value=""" & IIf(ProductTypeID = -1, "N" & PKID, "I" & ProductID) & """  onclick=""updateIdList(this," & PKID & ",'" & ExternalIDText & "');""" & IIf(level = 0 OrElse Not IsBuyerAssociated, " disabled=""disabled""", "") & " style=""display:" & IIf(PABFlag = "1" AndAlso ProductTypeID > 0, "none", "block") & """ /></td>")
                    End If


                    Sendb("  <td style=""display:" & IIf(PABFlag = "1" AndAlso ProductTypeID > 0, "none", "block") & """>")
                    If (pgID = "" OrElse CInt(pgID <= 0)) Then
                        Send("<input type=""hidden"" id=""PKID" & i & """ name=""PKID" & i & """ value=""" & IIf(ProductTypeID = -1, "N" & PKID, "I" & ProductID) & """ /> ")
                    End If
                    Send("    <img id=""imgitem" & IIf(ProductTypeID = -1, "N", "I") & PKID & """ src=""/images/" & iconFileName & """ alt=""" & FolderAltText & """ title=""" & FolderAltText & """  />" & "</td>")


                    Send("  <td id=""ExternalID"">" & ExternalIDText & "</td>")
                    Send("  <td nowrap>" & HttpUtility.HtmlEncode(ProductDesc) & "</td>")
                    Send("</tr>")
                    If (Request.QueryString("itemPK") <> "" AndAlso PKID = Long.Parse(Request.QueryString("itemPK"))) Then
                        highlightedRow = i
                    End If
                    i += 1
                Next
            End If
            Send("</tbody>")
            Send("</table>")
            'If (highlightedRow > -1) Then
            Send("<trailer>" & highlightedRow & "," & totalItemCt & "," & totalSelCt & "</trailer>")
            'Send("<label style=""display:none"">productcount:" & GetProductCount(nodeId.ToString()) & "</label>")
            'End If
        Catch ex As Exception
            Send(Copient.PhraseLib.Lookup("hierarchy.unabletoloaditems", LanguageID))
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub
    Public Sub UpdateProductCount(ByVal NodeId As String)
        Dim PCount As Integer = GetProductCount(NodeId)
        Send(PCount)
    End Sub
    Function GetProductCount(ByVal NodeID As String) As Integer
        Dim MyCommon As New Copient.CommonInc
        Dim dt As DataTable
        Dim IsClosedRT As Boolean = False
        Try
            If (String.IsNullOrEmpty(NodeID)) Then
                Return 0
            Else

                ' ensure everything we need is opened
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    IsClosedRT = True
                    MyCommon.Open_LogixRT()
                End If

                Dim dtNodeIDs As DataTable = New DataTable()
                dtNodeIDs.Columns.Add("NodeID")

                For Each IDs In NodeID.Trim(",").Split(",")
                    dtNodeIDs.Rows.Add(Convert.ToInt64(IDs))
                Next

                MyCommon.AppName = "phierarchytree.aspx"
                MyCommon.QueryStr = "dbo.pt_Product_In_NodeHierarchy"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@NodeIDList", SqlDbType.Structured).Value = dtNodeIDs
                dt = MyCommon.LRTsp_select
                If (dt.Rows.Count > 0) Then
                    Return Convert.ToInt32(dt.Rows(0)(0))
                Else
                    Return 0

                End If
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            If IsClosedRT Then
                If Not (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Close_LogixRT()
            End If
            MyCommon = Nothing
        End Try
        Return 0
    End Function
    Private Function GetIdText(ByVal ExternalID As String, ByVal DisplayID As String, ByVal LevelType As String, ByVal IDType As String) As String
        Dim IdText As String = ""

        'Is this a folder? (i.e. 1) - if so determine which ID to display
        If (LevelType = "-1") Then
            Select Case IDType
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

    Sub AddNode(ByVal TopLevel As Boolean, ByVal HierarchyId As Long, ByVal ParentNodeId As Long, ByVal NodeName As String)
        Dim NodeId As Long
        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            If (TopLevel) Then
                MyCommon.QueryStr = "dbo.pt_ProdHierarchies_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ExternalId", SqlDbType.NVarChar, 20).Value = ""
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = NodeName
                MyCommon.LRTsp.Parameters.Add("@DisplayID", SqlDbType.NVarChar, 100).Value = ""
                MyCommon.LRTsp.Parameters.Add("@Restricted", SqlDbType.Bit).Value = False
                MyCommon.LRTsp.Parameters.Add("@HierarchyId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                NodeId = MyCommon.LRTsp.Parameters("@HierarchyId").Value
                MyCommon.Close_LRTsp()
                Send(vbTab & "<span id=""node" & NodeId & """ name=""node" & NodeId & """><span id=""indent" & NodeId & """ style=""line-height:18px;padding-left:0px;"">")
                Send(vbTab & "<a href=""#""><img src=""/images/folder.png"" alt="""" /></a>")
                Send(vbTab & "<span id=""name" & NodeId & """ style=""padding-left:5px;"" onclick=""highlightNode(" & NodeId & ",1)"">" & NodeName & "</span>")
                Send(vbTab & "<input type=""hidden"" id=""hierarchy" & NodeId & """ name=""hierarchy" & NodeId & """ value=""" & NodeId & """ />")
                Send(vbTab & "<br /></span>")
                Send(vbTab & "<span id=""hId" & NodeId & """ style=""display:none;""></span></span>")
            Else
                If (ParentNodeId = HierarchyId) Then
                    ParentNodeId = 0
                End If
                MyCommon.QueryStr = "dbo.pt_PHNodes_InsertNode"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@HierarchyId", SqlDbType.BigInt, 8).Value = HierarchyId
                MyCommon.LRTsp.Parameters.Add("@ExternalId", SqlDbType.NVarChar, 120).Value = ""
                MyCommon.LRTsp.Parameters.Add("@ParentId", SqlDbType.BigInt, 8).Value = ParentNodeId
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = NodeName
                MyCommon.LRTsp.Parameters.Add("@DisplayID", SqlDbType.NVarChar, 100).Value = ""
                MyCommon.LRTsp.Parameters.Add("@Restricted", SqlDbType.Bit).Value = False
                MyCommon.LRTsp.Parameters.Add("@NodeId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                NodeId = MyCommon.LRTsp.Parameters("@NodeId").Value
                MyCommon.Close_LRTsp()
            End If
            MyCommon.Close_LogixRT()
        Catch ex As Exception
            Send(Copient.PhraseLib.Lookup("hierarchy.unabletoaddnode", LanguageID))
        End Try
    End Sub

    Sub DeleteNode(ByVal TopLevel As Boolean, ByVal HierarchyId As Long, ByVal NodeId As Long)
        Dim ParentNode As Integer = 0
        Dim dt As DataTable
        Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")

        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        MyCommon.AppName = "HierarchyFeeds.aspx"
        MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select ParentId, HierarchyID from PHNodes with (NoLock) where NodeId=" & NodeId
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
            ParentNode = MyCommon.NZ(dt.Rows(0).Item("ParentId"), 0)
            If (ParentNode = 0) Then
                ParentNode = MyCommon.NZ(dt.Rows(0).Item("HierarchyID"), 0)
            End If
        Else
            ParentNode = -1
        End If
        Dim infoMessage As String = ""
        infoMessage = MyResyncer.RemoveNodeAndAllChildren(TopLevel, NodeId)
        If (infoMessage = "") Then
            Send("OK")
            Send("ParentID=" & ParentNode)
        Else
            Send(Copient.PhraseLib.Lookup(infoMessage, LanguageID))
        End If
        MyCommon.Close_LogixRT()
    End Sub

    Sub DeleteItems(ByVal nodeId As Long, ByVal Items As String)
        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            If (Items <> "") Then
                If (Items.Substring(Items.Length - 1, 1) = ",") Then
                    Items = Items.Substring(0, Items.Length - 1)
                End If
                MyCommon.QueryStr = "delete from PHContainer with (RowLock) where pkid in (" & Items & ");"
                MyCommon.LRT_Execute()
            End If
            GenerateItemsDiv(nodeId, 1, "", "-1", "-1")
        Catch ex As Exception
            Send(Copient.PhraseLib.Lookup("hierarchy.unabletodeleteitems", LanguageID))
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub HandleGenerateDivForSigns(ByVal NodeId As Long, ByVal HierarchyID As Long)
        Dim dt, dtAvailableSigns, dtSelectedSign As DataTable
        Dim index As Integer
        Dim row As DataRow
        Dim ExternalId As String
        Dim ParentID As Long
        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            Send("<div class=""foldertitlebar"">")
            Send(" <span class=""dialogtitle"">Select/Assign Signs</span>")
            Send(" <span class=""dialogclose"" onclick=""toggleDialog('divAssignSigns', false);"">X</span>")
            Send("</div>")
            Send("<div class=""dialogcontents"">")
            Send(" <br class=""half""/>")
            If HierarchyID = -1 Then
                'MyCommon.QueryStr = "select signnumber,DefaultSign from HierarchySigns HS with (nolock) left outer join phnodes PN  with (nolock) on HS.hierarchyid=PN.hierarchyid where pn.nodeid=" & NodeId & " or HS.hierarchyid=0;"
                MyCommon.QueryStr = "Select ParentID,ExternalID From PHNodes with (nolock) where NodeID=" & NodeId
                dt = MyCommon.LRT_Select
                ExternalId = MyCommon.NZ(dt.Rows(0).Item("ExternalID"), "")
                ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0)

                If ParentID = 0 Then
                    MyCommon.QueryStr = "select signnumber, DefaultSign, ExternalID from HierarchySigns with (nolock) where ExternalID='" & ExternalId & "' or ExternalId='0';"
                    dtAvailableSigns = MyCommon.LRT_Select
                Else
                    Do While (ParentID <> 0)
                        MyCommon.QueryStr = "Select ParentID,ExternalID From PHNodes with (nolock) where NodeID=" & ParentID
                        dt = MyCommon.LRT_Select
                        ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0)
                        ExternalId = MyCommon.NZ(dt.Rows(0).Item("ExternalID"), "")
                    Loop
                    MyCommon.QueryStr = "select signnumber, DefaultSign, ExternalID from HierarchySigns with (nolock) where ExternalID='" & ExternalId & "' or ExternalId='0';"
                    dtAvailableSigns = MyCommon.LRT_Select
                End If
            Else
                MyCommon.QueryStr = "select signnumber, DefaultSign, ExternalID from HierarchySigns with (nolock) where hierarchyid = '" & HierarchyID & "' or ExternalId='0';"
                dtAvailableSigns = MyCommon.LRT_Select
            End If
            MyCommon.QueryStr = "Select SelectedSign from HierarchySelectedSigns Where ExternalID='" & ExternalId & "'"
            dtSelectedSign = MyCommon.LRT_Select
            If dtSelectedSign.Rows.Count > 0 Then
                Send("<table>")
                Send("<tr>")
                Send("<td valign=""top"">")
                Send("<label for=""lblsignselected"">Selected Sign:</label><br />")
                Send("<select name=""SignSelected"" id=""SignSelected"" disabled=""disabled"" >")
                Send("<option value=""0"">" & dtSelectedSign.Rows(0).Item("SelectedSign") & "</option>")
                Send("</select>")
                Send("  </td>")
                Send("</tr>")
                Send(" </table>")
            End If
            Send("<div id=""AvailableSigns"" >")
            Send("<label for=""lblsigns"">Please select to change Sign:</label><br />")
            Send("<table>")
            Send("<tr>")
            Send("<td valign=""top"">")
            Send("<select name=""Signs"" id=""Signs"" >")
            Send("<option value=""-1"">[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]</option>")
            index = 0
            For Each row In dtAvailableSigns.Rows
                Send(" <option value=""" & index & """" & IIf(MyCommon.NZ(row.Item("DefaultSign"), 0) = 1, " selected=""selected""", "") & " >" & MyCommon.NZ(row.Item("signnumber"), "") & "</option>")
                index = index + 1
            Next
            Send("</select>")
            Send("  </td>")
            Send("</tr>")
            Send(" </table>")
            Send("<br /> ")
            Send("<input type=""button"" name=""btnSignstoHierarchy"" id=""btnSignstoHierarchy"" value=""Save"" onclick=""javascript:AssignSigns();"" />")
            Send(" </div>")
            Send(" </div>")
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub


    Sub GenerateAvailableItems(ByVal nodeId As Long, ByVal nodeName As String)
        Dim sQuery As String = ""
        Dim dtAvailable As DataTable
        Dim dr As DataRow
        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            Send("<center>")
            Send("<h2>Select items to add to " & nodeName & " node</h2>")
            Send("<select id=""itemstoadd"" name=""itemstoadd"" multiple=""multiple"" size=""15"" style=""width:220px;"">")
            sQuery = "select ProductId as Id,ExtProductId as Code,Description as Name from Products"
            sQuery += " with (NoLock) where ProductId not in ("
            sQuery += " select a.ProductId from Products a with (NoLock), PHContainer b"
            sQuery += " with (NoLock) where a.ProductId = b.ProductId and b.NodeId = " & nodeId & ")"
            MyCommon.QueryStr = sQuery
            dtAvailable = MyCommon.LRT_Select()
            If (dtAvailable.Rows.Count > 0) Then
                For Each dr In dtAvailable.Rows
                    If dr.Item("Name") = "" Then
                        Send("<option value=""" & dr.Item("Id") & """>" & dr.Item("Code") & "</option>")
                    Else
                        Send("<option value=""" & dr.Item("Id") & """>" & dr.Item("Code") & " - " & dr.Item("Name") & "</option>")
                    End If
                Next
            Else
                Send("<option value=""-1"">" & Copient.PhraseLib.Lookup("hierarchy.noitems", LanguageID) & "</option>")
            End If
            Send("</select><br /><br />")
            Send("<input type=""button"" id=""btnAddItems"" name=""btnAddItems"" value=""" & Copient.PhraseLib.Lookup("term.additems", LanguageID) & """ onclick=""AddItemsToNode();"" />")
            Send("<input type=""button"" id=""btnAddItemCancel"" name=""btnAddItemCancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""showAddItem(false);"" />")
            Send("</center>")
        Catch ex As Exception
            Send(Copient.PhraseLib.Lookup("hierarchy.unabletoadditems", LanguageID))
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub TransmitGroups(ByVal groupIDs As String, ByVal pg As String, ByRef foundItems As Boolean, Optional ByVal add As Boolean = True)
        Dim sQuery As String = ""
        Dim dt As DataTable
        Dim dtProducts As DataTable
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

        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            AdminUserID = Verify_AdminUser(MyCommon, Logix)


            groupArray = groupIDs.Split(",")
            dtProducts = New DataTable()
            dtProducts.Columns.Add("ProductGroupID", System.Type.GetType("System.Int64"))
            dtProducts.Columns.Add("ProductID", System.Type.GetType("System.Int64"))
            dtProducts.Columns.Add("Manual", System.Type.GetType("System.Int32"))
            dtProducts.Columns.Add("Deleted", System.Type.GetType("System.Boolean"))
            dtProducts.Columns.Add("CMOAStatusFlag", System.Type.GetType("System.Int32"))
            dtProducts.Columns.Add("TCRMAStatusFlag", System.Type.GetType("System.Int32"))

            For i = 0 To groupArray.GetUpperBound(0)
                temp = groupArray(i)
                ' if it's a upc or department simply save it to add later; otherwise, find all the upcs in the tree
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
                        ' find UPCs and department at the node level
                        MyCommon.QueryStr = "select P.ProductID from PHContainer PHC with (NoLock) inner join Products P with (NoLock) on P.ProductID = PHC.ProductID " & _
                               "where NodeID in (" & String.Join(",", nodeList) & ")"
                        dt = MyCommon.LRT_Select
                        'Send(MyCommon.QueryStr)
                        If (dt.Rows.Count > 0) Then
                            dtProducts.Clear()
                            ' add the UPCs to the database for this product group
                            For Each row In dt.Rows
                                insertRow = dtProducts.NewRow()
                                insertRow.ItemArray = New Object() {Long.Parse(pg), row.Item("ProductID"), 1, 0, 2, 2}
                                dtProducts.Rows.Add(insertRow)
                            Next
                            If (dtProducts.Rows.Count > 0) Then
                                foundItems = True
                                ' get the selected node name for logging purposes
                                MyCommon.QueryStr = "select NodeName = case when PHN.ExternalID is NULL then PHN.Name when PHN.ExternalID = '' then PHN.Name  " & _
                                                    "when PHN.ExternalID not like '%' + PHN.Name + '%' then PHN.ExternalID + '-' + PHN.Name else PHN.ExternalID end  " & _
                                                    ", PH.Name as HierarchyName from PHNodes PHN with (nolock) inner join ProdHierarchies PH with (NoLock) on PH.HierarchyID=PHN.HierarchyID " & _
                                                    "where NodeID = " & temp.Substring(1) & "; "
                                dtName = MyCommon.LRT_Select
                                If (dtName.Rows.Count > 0) Then
                                    nodeName = dtName.Rows(0).Item("NodeName")
                                    hierarchyName = dtName.Rows(0).Item("HierarchyName")
                                End If
                                LogMsg = IIf(add, "Added ", "Removed ") & nodeName & " ( " & dtProducts.Rows.Count & " items) from hierarchy " & hierarchyName
								Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
								Dim RetMsg As String = ""
                                If (add) Then
                                    If(IsLinkedPage) Then
										RetMsg = MyResyncer.LinkNodesToGroup(Convert.ToInt64(pg), groupIDs)
									else
										recordsAffected = BatchInsert(dtProducts, dtProducts.Rows.Count)
									End If
                                Else
                                    If(IsLinkedPage) Then
										RetMsg = MyResyncer.UnlinkNodesFromGroup(Convert.ToInt64(pg), groupIDs)
									else
										recordsAffected = BatchDelete(dtProducts, pg)	
									End If
                                End If
                                If (recordsAffected > 0) Then
                                    RecordSelectedNode(pg, temp.Substring(1), add)
                                    MyCommon.Activity_Log(5, CLng(pg), AdminUserID, LogMsg)
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            ' add all the individual upc that are at the selected node level
            If (itemList.Count > 0) Then
                dtProducts.Clear()
                For i = 0 To itemList.Count - 1
                    insertRow = dtProducts.NewRow()
                    insertRow.ItemArray = New Object() {Long.Parse(pg), Long.Parse(itemList.Item(i)), 1, 0, 2, 2}
                    dtProducts.Rows.Add(insertRow)
                Next
                If (dtProducts.Rows.Count > 0) Then
                    foundItems = True
                    If (add) Then
                        BatchInsert(dtProducts, dtProducts.Rows.Count)
                    Else
                        BatchDelete(dtProducts, pg)
                    End If
                    ' get the selected node name for logging purposes
                    If (Request.Form("sel") <> "") Then
                        Integer.TryParse(Request.Form("sel"), SelTreeNode)
                    End If
                    MyCommon.QueryStr = "select NodeName = case when PHN.ExternalID is NULL then PHN.Name when PHN.ExternalID = '' then PHN.Name  " & _
                                        "when PHN.ExternalID not like '%' + PHN.Name + '%' then PHN.ExternalID + '-' + PHN.Name else PHN.ExternalID end  " & _
                                        ", PH.Name as HierarchyName from PHNodes PHN with (nolock) inner join ProdHierarchies PH with (NoLock) on PH.HierarchyID=PHN.HierarchyID " & _
                                        "where NodeID = " & SelTreeNode & "; "
                    dtName = MyCommon.LRT_Select
                    If (dtName.Rows.Count > 0) Then
                        nodeName = dtName.Rows(0).Item("NodeName")
                        hierarchyName = dtName.Rows(0).Item("HierarchyName")
                    End If
                    ' log the items that were added or removed
                    itemString = IIf(dtProducts.Rows.Count = 1, " item ", " items ")
                    LogMsg = IIf(add, "Added ", "Removed ") & dtProducts.Rows.Count & itemString & " from " & nodeName & " within hierarchy " & hierarchyName
                    MyCommon.Activity_Log(5, CLng(pg), AdminUserID, LogMsg)
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

        MyCommon.QueryStr = "select NodeID from PHNodes with (NoLock) where ParentID=" & NodeId
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
          "Insert into ProdGroupItems with (RowLock) (ProductGroupID, ProductID, Manual, Deleted, CMOAStatusFlag, CPEStatusFlag, UEStatusFlag, TCRMAStatusFlag, LastUpdate) " & _
          "values (@ProductGroupID, @ProductID, 1, 0, 2, 2, 2, 2, getdate()) ", _
          MyCommon.LRTadoConn)
        adapter.InsertCommand.Parameters.Add("@ProductGroupID", _
          SqlDbType.BigInt, 8, "ProductGroupID")
        adapter.InsertCommand.Parameters.Add("@ProductID", _
          SqlDbType.BigInt, 8, "ProductID")
        adapter.InsertCommand.UpdatedRowSource = UpdateRowSource.None

        ' Set the batch size.
        batchSize = IIf(batchSize > 1000, 1000, batchSize)
        adapter.UpdateBatchSize = batchSize
        recordsAffected = adapter.Update(dataTable)

        Return recordsAffected
    End Function

    Public Function BatchDelete(ByVal dataTable As DataTable, ByVal ProductGroupID As String) As Integer
        Dim ProductIdList As String() = Nothing
        Dim IdClause As String = ""
        Dim row As DataRow
        Dim i As Integer = 0
        Dim recordsAffected As Integer = -1

        If (Not dataTable Is Nothing AndAlso dataTable.Rows.Count > 0) Then
            ReDim ProductIdList(dataTable.Rows.Count - 1)
            For Each row In dataTable.Rows
                ProductIdList(i) = MyCommon.NZ(row.Item("ProductID"), "-1")
                i += 1
            Next
            If (ProductIdList.Length > 0) Then
                Dim plist As New List(Of String)(ProductIdList)
                Do While plist.Count > 0
                    IdClause = String.Join(",", plist.Take(1000))
                    MyCommon.QueryStr = "Delete from ProdGroupItems with (RowLock) where ProductGroupID=" & ProductGroupID & " " & _
                                        "and Deleted=1 and ProductID in (" & IdClause & ");"
                    MyCommon.LRT_Execute()

                    MyCommon.QueryStr = "Update ProdGroupItems with (RowLock) set Deleted=1, Manual=1, CMOAStatusFlag=2, TCRMAStatusFlag=2, CPEStatusFlag=2, UEStatusFlag=2, " & _
                                        "  LastUpdate=getdate() where Deleted=0 and ProductGroupID=" & ProductGroupID & " " & _
                                        "and ProductID in (" & IdClause & ");"
                    'Send(MyCommon.QueryStr)
                    MyCommon.LRT_Execute()
                    recordsAffected = recordsAffected + MyCommon.RowsAffected
                    plist.RemoveRange(0, IdClause.Split(",").Length)
                Loop
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
                    MyCommon.QueryStr = "Select NodeID from PHContainer with (NoLock) where ProductID = " & temp.Substring(1)
                    'Send(MyCommon.QueryStr)
                    dt = MyCommon.LRT_Select
                    If (dt.Rows.Count > 0) Then
                        ParentID = MyCommon.NZ(dt.Rows(0).Item("NodeID"), -1)
                    End If
                Else
                    MyCommon.QueryStr = "Select ParentID, HierarchyID from PHNodes with (NoLock) where NodeID = " & temp.Substring(1)
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

    Public Function GetParentNode(ByVal NodeID As Long, Optional ByVal IsHierarchy As Boolean = False) As Long
        Dim ParentID As Long = 0
        Dim dt As DataTable

        Try
            If NodeID <= 0 OrElse IsHierarchy = True Then
                ParentID = 0
            Else
                MyCommon.QueryStr = "SELECT ParentID, HierarchyID FROM PHNodes WITH (NoLock) WHERE NodeID=" & NodeID & ";"
                dt = MyCommon.LRT_Select
                If (dt.Rows.Count > 0) Then
                    If (MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0) = 0) Then
                        ParentID = MyCommon.NZ(dt.Rows(0).Item("HierarchyID"), -1)
                    Else
                        ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0)
                    End If
                End If
            End If
        Catch ex As Exception
            Send(ex.ToString)
        Finally
        End Try

        Return ParentID
    End Function

    Function GetProductGroupItemCount(ByVal pgID As String) As Integer
        Dim AssignedCount As Integer = 0
        Dim dt As DataTable
        Try
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "Select count(*) as AssignedCount from ProdGroupItems with (NoLock) where ProductGroupID=" & pgID & " and Deleted=0;"
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
                msgBuf.Append(Copient.PhraseLib.Detokenize("hierarchy.ItemsRemovedFromPG", LanguageID, displayValue))
            ElseIf (assignedChange = 0) Then
                If (Request.QueryString("action") = "unassignNodes" OrElse Request.QueryString("type") = "REMOVE") Then
                    msgBuf.Append(Copient.PhraseLib.Lookup("hierarchy.NoItemsToRemove", LanguageID))
                    'Else
                    '    msgBuf.Append(Copient.PhraseLib.Lookup("hierarchy.NoItemsToAdd", LanguageID))
                End If
            Else
                msgBuf.Append(Copient.PhraseLib.Detokenize("hierarchy.ItemsAddedToPG", LanguageID, displayValue))
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

    Sub RemoveAllFromProductGroups(ByVal ProductGroupID As String, ByVal AdminUserID As Long)
        Dim rowCt As Integer = 0
        Dim dt As DataTable
        Dim preDeleteCt As Integer = 0
        Dim postDeleteCt As Integer = 0
        Dim LogMsg As String = ""

        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "Select count(*) as ProdGroupCt from ProdGroupItems with (NoLock) where Deleted=0 and ProductGroupID=" & ProductGroupID
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                preDeleteCt = MyCommon.NZ(dt.Rows(0).Item("ProdGroupCt"), 0)
            End If

            MyCommon.QueryStr = "dbo.pt_ProdGroupitems_Delete"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = Long.Parse(ProductGroupID)
            MyCommon.LRTsp.ExecuteNonQuery()

            MyCommon.QueryStr = "Select count(*) as ProdGroupCt from ProdGroupItems with (NoLock) where Deleted=0 and ProductGroupID=" & ProductGroupID
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                postDeleteCt = MyCommon.NZ(dt.Rows(0).Item("ProdGroupCt"), 0)
            End If

            LogMsg = Copient.PhraseLib.Lookup("phierarchy.removedall", LanguageID) & " (" & Math.Abs(postDeleteCt - preDeleteCt) & _
                     " " & Copient.PhraseLib.Lookup("term.items", LanguageID) & ")"
            MyCommon.Activity_Log(5, CLng(ProductGroupID), AdminUserID, LogMsg)

            If (postDeleteCt = 0) Then
                MyCommon.QueryStr = "delete from ProductGroupNodes with (RowLock) where ProductGroupID = " & ProductGroupID
                MyCommon.LRT_Execute()
            End If

            Send("|" & preDeleteCt & "," & postDeleteCt)
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub UpdateProductGroup(ByVal ProductGroupID As String, ByVal SetStatusFlag As Boolean)
        Dim sQuery As String = ""
        Try
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            sQuery = "update productgroups with (RowLock) set updatelevel=updatelevel+1"
            If (SetStatusFlag) Then
                sQuery += ", CPEStatusFlag=2, UEStatusFlag=2 "
            End If
            sQuery += " where ProductGroupID = " & ProductGroupID
            MyCommon.QueryStr = sQuery
            MyCommon.LRT_Execute()
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub FindSearchItems(ByVal SearchString As String, ByVal SearchType As Integer, ByVal ProductGroupID As String, _
                          ByVal SelectedNodeID As Long, ByVal SelectedHierID As Integer, ByVal PAB As String)
        Dim dt As DataTable
        Dim row As DataRow
        Dim Level As Integer
        Dim LevelDisplay As String = ""
        Dim NodeID As Long
        Dim Name As String = ""
        Dim ExternalID As String = ""
        Dim ProductID As String = ""
        Dim ShowNoItemMsg As Boolean = True
        Dim Shaded As String = "shaded"
        Dim iconFileName As String = ""
        Dim SelectedItem As Boolean = False
        Dim HierarchyID As Integer = -1
        Dim AnchorID As String = ""
        Dim ImgID As String = ""
        Dim LinkTitle As String = ""
        Dim rowCt As Integer = 0
        Dim ProductName As String = ""
        Dim PhnodesName As String = ""
        Dim PhnodeExternalId As String = ""
        Dim OrigSearchString As String = SearchString
        Dim TOPSearchLimit As Integer = 1000
        Dim ProductExternalId As String  =""
        Try
            If (SearchString <> "") Then
                Send("<table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """ style=""width:100%;word-break:break-all; word-wrap:break-word;margin-top:44px;"">")
                Send("<thead>")
                Send("  <tr>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.select", LanguageID) & "</th>")
                If (ProductGroupID <> "" And CInt(ProductGroupID) > 0) Then If (PAB <> "1") Then Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.action", LanguageID) & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.level", LanguageID) & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.hierarchy", LanguageID) & "</th>")
                ''Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.Product", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "&nbsp;&nbsp;" & Copient.PhraseLib.Lookup("term.Name", LanguageID) & ")" & "</th>")
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term.Node", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "&nbsp;&nbsp;" & Copient.PhraseLib.Lookup("term.Name", LanguageID) & ")" & "</th>")
                Send("  </tr>")
                Send("</thead>")
                Send("<tbody>")
                SearchString = Server.HtmlDecode(SearchString)
                MyCommon.AppName = "HierarchyFeeds.aspx"
                MyCommon.Open_LogixRT()
                Dim tempSearchString = MyCommon.Parse_Quotes(SearchString)
                SearchString = IIf(SearchString.IndexOf("%") = -1, MyCommon.Parse_Quotes(SearchString), MyCommon.Parse_Quotes(SearchString.Replace("%", "[%]")))
                Select Case SearchType
                    Case 0 ' contains
                        SearchString = "%" & SearchString & "%"
                    Case 1 ' starts with
                        SearchString = SearchString & "%"
                    Case 2 ' ends with
                        SearchString = "%" & SearchString
                    Case Else ' use contains as default
                        SearchString = "%" & SearchString & "%"
                End Select


                If SelectedNodeID > 0 Then
                    ' search for all the matches in and under the selected node.
                    MyCommon.QueryStr = "dbo.pa_PHA_SearchFromNode"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@NodeID", SqlDbType.BigInt).Value = SelectedNodeID
                    MyCommon.LRTsp.Parameters.Add("@SearchString", SqlDbType.NVarChar, 120).Value = SearchString
                    MyCommon.LRTsp.Parameters.Add("@PGID", SqlDbType.BigInt).Value = CLng(ProductGroupID)
                    MyCommon.LRTsp.Parameters.Add("@PABFlag", SqlDbType.Bit).Value = IIf(PAB = "1", True, False)
                    MyCommon.LRTsp.Parameters.Add("@TOP", SqlDbType.SmallInt).Value = TOPSearchLimit

                    dt = MyCommon.LRTsp_select
                    MyCommon.Close_LRTsp()
                ElseIf SelectedHierID > 0 Then
                    ' search the selected hierarchy for matches
                    MyCommon.QueryStr = "select top " & cString(TOPSearchLimit) & " * from ( " & _
                                          "select 1 as LevelType, NodeID as ID,''AS ProductNAME,''AS ProductExternalID,   -1 as ProductID,-1 as ProductTypeID,ph.Name as HierarchyName, ph.HierarchyID,phn.Name as PhnodesName,phn.ExternalID as PhnodesExternalId, 0 as Selected " & _
                                          "from PHNodes phn with (NoLock) " & _
                                          "inner join ProdHierarchies ph with (NoLock) on ph.HierarchyID = phn.HierarchyID  and ph.HierarchyID=" & SelectedHierID & " " & _
                      "Where phn.ExternalID like '" & SearchString & "' or phn.name like '" & SearchString & "' "
                    If (PAB <> "1") Then
                        MyCommon.QueryStr &=
                                  "union " & _
                                  "select 2 as LevelType, PKID as ID, p.Description as ProductNAME, p.extProductID as ProductExternalID, p.ProductID as ProductID,p.ProductTypeID,ph.Name as HierarchyName, ph.HierarchyID,phn.Name as PhnodesName ,phn.ExternalID as PhnodesExternalId,   " & _
                                  "	(select count(*) from ProdGroupItems with (NoLock) where ProductGroupID=" & ProductGroupID & " and ProductID = p.ProductID and Deleted=0) as Selected  " & _
                                  "from Products p with (NoLock) inner join PHContainer c with (NoLock) on p.productid= c.productid  " & _
                                  "inner join PHNodes phn with (NoLock) on phn.NodeID = c.NodeID  " & _
                                  "inner join ProdHierarchies ph with (NoLock) on ph.HierarchyID = phn.HierarchyID and ph.HierarchyID=" & SelectedHierID & " " & _
                                  "Where p.extProductID like '" & SearchString & "' "
                        If (SearchType = 1) Then
                            MyCommon.QueryStr &= " or Substring(p.extProductID, PATINDEX('%[^0]%', p.extProductID), Len(p.extProductID) - PATINDEX('%[^0]%', p.extProductID)) like '" & SearchString & "' "
                        End If
                        MyCommon.QueryStr &= " or p.Description like '" & SearchString & "' "

                    End If
                    MyCommon.QueryStr &= ") ResultsTable " & _
                                                   "order by  HierarchyID,LevelType, ProductTypeID, ProductExternalID; "

                    dt = MyCommon.LRT_Select
                Else
                    ' search all the hierarchies for a match
                    MyCommon.QueryStr = "select top 500 * from ( " & _
                                "select 1 as LevelType, NodeID as ID, ''AS ProductNAME,''AS ProductExternalID, -1 as ProductID,-1 as ProductTypeID,ph.Name as HierarchyName, ph.HierarchyID,phn.Name as PhnodesName,phn.ExternalID as PhnodesExternalId , 0 as Selected " & _
                                "from PHNodes phn with (NoLock) " & _
                                "inner join ProdHierarchies ph with (NoLock) on ph.HierarchyID = phn.HierarchyID  " & _
            "Where phn.ExternalID like '" & SearchString & "' or phn.name like '" & SearchString & "' "

                    If (PAB <> "1") Then
                        MyCommon.QueryStr &=
                                  "union " & _
                                  "select 2 as LevelType, PKID as ID, p.Description as ProductNAME, p.extProductID as ProductExternalID, p.ProductID as ProductID,p.ProductTypeID,ph.Name as HierarchyName, ph.HierarchyID, phn.Name as PhnodesName ,phn.ExternalID as PhnodesExternalId, " & _
                                  "	(select count(*) from ProdGroupItems with (NoLock) where ProductGroupID=" & ProductGroupID & " and ProductID = p.ProductID and Deleted=0) as Selected  " & _
                                  "from Products p with (NoLock) inner join PHContainer c with (NoLock) on p.productid= c.productid  " & _
                                  "inner join PHNodes phn with (NoLock) on phn.NodeID = c.NodeID  " & _
                                  "inner join ProdHierarchies ph with (NoLock) on ph.HierarchyID = phn.HierarchyID   " & _
                                  "Where p.extProductID like '" & SearchString & "' "

                        If (SearchType = 1) Then
                            MyCommon.QueryStr &= " or Substring(p.extProductID, PATINDEX('%[^0]%', p.extProductID), Len(p.extProductID) - PATINDEX('%[^0]%', p.extProductID)) like '" & SearchString & "' "
                        End If
                        MyCommon.QueryStr &= " or p.Description like '" & SearchString & "' "

                    End If
                    MyCommon.QueryStr &= ") ResultsTable " & _
                                                   "order by  HierarchyID,LevelType, ProductTypeID, ProductExternalID; "

                    dt = MyCommon.LRT_Select
                End If


                rowCt = dt.Rows.Count
                If (dt.Rows.Count > 0) Then
                    If (dt.Rows.Count >= TOPSearchLimit) Then
                        Send("<div id=""infobar"" class=""green-background"" style=""width:100%"">" & String.Format(Copient.PhraseLib.Lookup("phierarchy.MaxResults", LanguageID), TOPSearchLimit) & "</div>")
                    End If
                    'Send("" & String.Format(Copient.PhraseLib.Lookup("phierarchy.MaxResults", LanguageID), TOPSearchLimit) & "</b>")
                    For Each row In dt.Rows
                        SelectedItem = IIf(MyCommon.NZ(row.Item("Selected"), 0) > 0, True, False)
                        Select Case MyCommon.NZ(row.Item("ProductTypeID"), "")
                            Case -1
                                iconFileName = "folder.png"
                            Case 2
                                iconFileName = "dept" & IIf(SelectedItem, "-on", "") & ".png"
                            Case Else
                                iconFileName = "upc" & IIf(SelectedItem, "-on", "") & ".png"
                        End Select

                        NodeID = MyCommon.NZ(row.Item("ID"), 0)
                        ProductExternalId = MyCommon.NZ(row.Item("ProductExternalID"), "&nbsp;")
                        ProductName = ProductExternalId & "&nbsp;&nbsp;&nbsp;" & MyCommon.NZ(row.Item("ProductNAME"), "")
                        ProductID = MyCommon.NZ(row.Item("ProductID"), "&nbsp;")
                        HierarchyID = MyCommon.NZ(row.Item("HierarchyID"), -1)
                        Level = MyCommon.NZ(row.Item("LevelType"), 2)
                        ' PhnodesName  is  Node Name   and Node External ID  
                        PhnodeExternalId = MyCommon.NZ(row.Item("PhnodesExternalId"), "&nbsp;")
                        PhnodesName = PhnodeExternalId & "&nbsp;&nbsp;&nbsp;" & MyCommon.NZ(row.Item("PhnodesName"), "")
                        ShowNoItemMsg = False
                        Send("<tr class=""" & Shaded & """>")
                        Send("  <td style=""width:50px;""><a href=""javascript:locateItem('L" & Level & "ID" & NodeID & "');"" alt=""" & Copient.PhraseLib.Lookup("hierarchy.LocateItemInHierarchy", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("hierarchy.LocateItemInHierarchy", LanguageID) & """ >" & Copient.PhraseLib.Lookup("term.locate", LanguageID) & "</a></td>")
                        AnchorID = "link" & IIf(ProductID > -1, ProductID, NodeID) & "H" & HierarchyID
                        ImgID = "img" & IIf(ProductID > -1, ProductID, NodeID) & "H" & HierarchyID
                        If SelectedItem Then
                            LinkTitle = Copient.PhraseLib.Detokenize("hierarchy.RemoveFromProductGroup", LanguageID, ProductGroupID)  'Remove from product group {0}
                        Else
                            LinkTitle = Copient.PhraseLib.Detokenize("hierarchy.AddToProductGroup", LanguageID, ProductGroupID)  'Add to product group {0}
                        End If

                        If (ProductGroupID <> "" And CInt(ProductGroupID) > 0 And PAB <> "1") Then
                            If (ProductID = -1) Then
                                Send("  <td style=""width:55px;""><a href=""javascript:handleSearchNodeAdjust('" & NodeID & "', '" & AnchorID & "'," & HierarchyID & ", 'ADD');"" id=""" & AnchorID & """ title=""" & Copient.PhraseLib.Detokenize("hierarchy.AddToProductGroup", LanguageID, ProductGroupID) & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a><br />")
                                Send("  <a href=""javascript:handleSearchNodeAdjust('" & NodeID & "', '" & AnchorID & "'," & HierarchyID & ", 'REMOVE');"" id=""" & AnchorID & """ title=""" & Copient.PhraseLib.Detokenize("hierarchy.RemoveFromProductGroup", LanguageID, ProductGroupID) & """>" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & "</a></td>")
                            Else
                                If (SelectedItem) Then
                                    Send("  <td style=""width:55px;""><a href=""javascript:handleSearchItemAdjust('" & ProductID & "', '" & AnchorID & "'," & HierarchyID & ",'" & MyCommon.NZ(row.Item("ProductExternalID"), "0") & "');"" id=""" & AnchorID & """ alt=""" & LinkTitle & """ title=""" & LinkTitle & """>" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & "</a></td>")
                                Else
                                    Send("  <td style=""width:55px;""><a href=""javascript:handleSearchItemAdjust('" & ProductID & "', '" & AnchorID & "'," & HierarchyID & ",'" & MyCommon.NZ(row.Item("ProductExternalID"), "0") & "');"" id=""" & AnchorID & """ alt=""" & LinkTitle & """ title=""" & LinkTitle & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a></td>")
                                End If
                            End If
                        End If

                        Send("  <td style=""width:35px;text-align:center;""><img src=""/images/" & iconFileName & """  id=""" & ImgID & """ />" & "</td>")
                        Send("  <td>" & MyCommon.NZ(row.Item("HierarchyName"), "&nbsp;") & "</td>")
                        'Send("  <td>" & HighlightMatches(ExternalID, SearchString) & "</td>")
                        Send("  <td>" & HighlightMatches(ProductName, SearchString) & "</td>")
                        Send("  <td>" & HighlightMatches(PhnodesName, SearchString) & "</td>")
                        Send("</tr>")
                        Shaded = IIf(Shaded = "shaded", "", "shaded")
                    Next
                Else
                    'SearchString = HttpUtility.HtmlEncode(SearchString.Replace("''", "&#39;"))
                    Send("<tr><td id=""noresultmsg"" colspan=""4""><center><i>" & Copient.PhraseLib.Lookup("phierarchy.noresultsfound", LanguageID) & " '" & OrigSearchString & "'</i></center></td></tr>")
                    'Send("<tr><td colspan=""4"">" & MyCommon.QueryStr & "</td></tr>")
                    ShowNoItemMsg = False
                End If
                If (ShowNoItemMsg) Then
                    'SearchString = HttpUtility.HtmlEncode(SearchString.Replace("''", "&#39;"))
                    Send("<tr><td id=""noresultmsg"" colspan=""4""><center><i>" & Copient.PhraseLib.Lookup("phierarchy.noresultsfound", LanguageID) & " '" & OrigSearchString & "'</i></center></td></tr>")
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

    Sub FindSearchItemAttributes(ByVal ItemAttribVal1 As String, ByVal Attr1Selection As Integer, ByVal low1Value As String, ByVal high1Value As String, _
                               ByVal ItemAttribVal2 As String, ByVal Attr2Selection As Integer, ByVal low2Value As String, ByVal high2Value As String, _
                               ByVal ItemAttribVal3 As String, ByVal Attr3Selection As Integer, ByVal low3Value As String, ByVal high3Value As String, _
                               ByVal ProductGroupID As String, ByVal SelectedNodeID As Long, ByVal SelectedHierID As Integer)
        Dim dt As DataTable
        Dim dtNodes As DataTable
        Dim rowNodes As DataRow
        Dim strNodes As String = ""
        Dim dtExtHierDetails As DataTable
        Dim SelectedNode As Integer = -1
        Dim rowCt As Integer = 0
        Dim row As DataRow
        Dim SearchCriteria1 As String = ""
        Dim SearchCriteria2 As String = ""
        Dim SearchCriteria3 As String = ""
        'Dim MainSearchCriteria as String= ""
        Dim ExternalHierarchy As String = "00"
        Dim HierarchyName As String = ""
        Dim AdminUserID As Long

        Dim Logix As New Copient.LogixInc
        If (ItemAttribVal1 <> "-1" And ItemAttribVal1 <> "") Then
            If low1Value = "" And high1Value = "" Then
                SearchCriteria1 = ""
            Else
                If Attr1Selection = 1 Then
                    SearchCriteria1 = "(HierAttribID = " + ItemAttribVal1 + " and (HierAttribValue = '" + low1Value + "' or HierAttribValue = '" + high1Value + "'))"
                Else
                    If (low1Value <= high1Value) Then
                        SearchCriteria1 = "(HierAttribID = " + ItemAttribVal1 + " and HierAttribValue >= " + low1Value + " and HierAttribValue <= " + high1Value + ")"
                    Else
                        SearchCriteria1 = ""
                    End If
                End If
            End If
        End If


        If (ItemAttribVal2 <> "-1" And ItemAttribVal2 <> "") Then
            If low2Value = "" And high2Value = "" Then
                SearchCriteria2 = ""
            Else
                If Attr2Selection = 1 Then
                    SearchCriteria2 = "(HierAttribID = " + ItemAttribVal2 + " and (HierAttribValue = '" + low2Value + "' or HierAttribValue = '" + high2Value + "'))"
                Else
                    If (low2Value <= high2Value) Then
                        SearchCriteria2 = "(HierAttribID = " + ItemAttribVal2 + " and HierAttribValue >= " + low2Value + " and HierAttribValue <= " + high2Value + ")"
                    Else
                        SearchCriteria2 = ""
                    End If
                End If
            End If
        End If

        If (ItemAttribVal3 <> "-1" And ItemAttribVal3 <> "") Then
            If low3Value = "" And high3Value = "" Then
                SearchCriteria3 = ""
            Else
                If Attr3Selection = 1 Then
                    SearchCriteria3 = "(HierAttribID = " + ItemAttribVal3 + " and (HierAttribValue = '" + low3Value + "' or HierAttribValue = '" + high3Value + "'))"
                Else
                    If (low3Value <= high3Value) Then
                        SearchCriteria3 = "(HierAttribID = " + ItemAttribVal3 + " and HierAttribValue >= " + low3Value + " and HierAttribValue <= " + high3Value + ")"
                    Else
                        SearchCriteria3 = ""
                    End If
                End If
            End If
        End If

        Try
            MyCommon.AppName = "HierarchyFeeds.aspx"
            If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
            AdminUserID = Verify_AdminUser(MyCommon, Logix)
            If SelectedHierID > 0 Then
                MyCommon.QueryStr = "Select ExternalID,Name from ProdHierarchies with (NoLock) where HierarchyID = " & SelectedHierID
            ElseIf SelectedNodeID > 0 Then
                MyCommon.QueryStr = "Select ExternalID,Name from PHnodes with (NoLock) where NodeID = " & SelectedNodeID
            End If
            dtExtHierDetails = MyCommon.LRT_Select
            If dtExtHierDetails.Rows.Count > 0 Then
                ExternalHierarchy = dtExtHierDetails.Rows(0)(0)
                HierarchyName = dtExtHierDetails.Rows(0)(1)
            End If

            MyCommon.QueryStr = "delete from ItemAttribSearchHistory"
            MyCommon.LRT_Execute()

            MyCommon.QueryStr = "insert into ItemAttribSearchHistory with (RowLock) " & _
                                "(AttribNo,ProdHierAttibuteID,Selection,LowValue,HighValue,SearchedDate,UserID,ExternalID,ExtHierName) values" & _
                                "(1," & ItemAttribVal1 & "," & Attr1Selection & ",'" & low1Value & "','" & high1Value & "'," & _
                                "'" & Now.ToString() & "','" & AdminUserID & "','" & ExternalHierarchy & "','" & HierarchyName & "')"
            MyCommon.LRT_Execute()

            MyCommon.QueryStr = "insert into ItemAttribSearchHistory with (RowLock) " & _
                                "(AttribNo,ProdHierAttibuteID,Selection,LowValue,HighValue,SearchedDate,UserID,ExternalID,ExtHierName) values" & _
                                "(2," & ItemAttribVal2 & "," & Attr2Selection & ",'" & low2Value & "','" & high2Value & "'," & _
                                "'" & Now.ToString() & "','" & AdminUserID & "','" & ExternalHierarchy & "','" & HierarchyName & "')"
            MyCommon.LRT_Execute()

            MyCommon.QueryStr = "insert into ItemAttribSearchHistory with (RowLock) " & _
                                "(AttribNo,ProdHierAttibuteID,Selection,LowValue,HighValue,SearchedDate,UserID,ExternalID,ExtHierName) values" & _
                                "(3," & ItemAttribVal3 & "," & Attr3Selection & ",'" & low3Value & "','" & high3Value & "'," & _
                                "'" & Now.ToString() & "','" & AdminUserID & "','" & ExternalHierarchy & "','" & HierarchyName & "')"
            MyCommon.LRT_Execute()


            If (SearchCriteria1 <> "" Or SearchCriteria2 <> "" Or SearchCriteria3 <> "") Then
                Send("<table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """ style=""width:100%;"">")
                Send("<tbody>")

                If (SearchCriteria1 <> "" And SearchCriteria2 <> "" And SearchCriteria3 <> "") Then
                    MyCommon.QueryStr = "Select Distinct AttrbSearch1.HierID from (Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria1 & ")AttrbSearch1 Inner Join " & _
                                  "((Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria2 & ")AttrbSearch2 Inner Join " & _
                      "(Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria3 & ")AttrbSearch3 on AttrbSearch2.HierID=AttrbSearch3.HierID) " & _
                      " On AttrbSearch1.HierID=AttrbSearch2.HierID"
                ElseIf (SearchCriteria1 <> "" And SearchCriteria2 = "" And SearchCriteria3 = "") Then
                    MyCommon.QueryStr = "Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria1
                ElseIf (SearchCriteria1 = "" And SearchCriteria2 <> "" And SearchCriteria3 = "") Then
                    MyCommon.QueryStr = "Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria2
                ElseIf (SearchCriteria1 = "" And SearchCriteria2 = "" And SearchCriteria3 <> "") Then
                    MyCommon.QueryStr = "Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria3
                ElseIf (SearchCriteria1 <> "" And SearchCriteria2 <> "" And SearchCriteria3 = "") Then
                    MyCommon.QueryStr = "Select Distinct AttrbSearch1.HierID from (Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria1 & ")AttrbSearch1 Inner Join " & _
                                  " (Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria2 & ")AttrbSearch2 On AttrbSearch1.HierID=AttrbSearch2.HierID"
                ElseIf (SearchCriteria1 <> "" And SearchCriteria2 = "" And SearchCriteria3 <> "") Then
                    MyCommon.QueryStr = "Select Distinct AttrbSearch1.HierID from (Select distinct HierID from BannerHierAttribValues with (NoLock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria1 & ")AttrbSearch1 Inner Join " & _
                                 "(Select distinct HierID from BannerHierAttribValues where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria3 & ")AttrbSearch3 On AttrbSearch1.HierID=AttrbSearch3.HierID"
                ElseIf (SearchCriteria1 = "" And SearchCriteria2 <> "" And SearchCriteria3 <> "") Then
                    MyCommon.QueryStr = "Select Distinct AttrbSearch2.HierID from (Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria2 & ")AttrbSearch2 Inner Join " & _
                                  "(Select distinct HierID from BannerHierAttribValues with (nolock) where HierID like '" & ExternalHierarchy & "%'" & _
                                           " and " & SearchCriteria3 & ")AttrbSearch3 On AttrbSearch2.HierID=AttrbSearch3.HierID"
                End If
                dt = MyCommon.LRT_Select
                rowCt = dt.Rows.Count
                MyCommon.QueryStr = "delete from ItemAttribSearchResult"
                MyCommon.LRT_Execute()
                If rowCt > 0 Then
                    strNodes = ""
                    For Each row In dt.Rows
                        MyCommon.QueryStr = "select NodeID from phnodes with (nolock) where ExternalID = '" & row.Item("HierID") & "'"
                        dtNodes = MyCommon.LRT_Select
                        If dtNodes.Rows.Count > 0 Then
                            For Each rowNodes In dtNodes.Rows
                                If strNodes = "" Then
                                    strNodes = "N" & rowNodes.Item("NodeID")
                                Else
                                    strNodes = strNodes & ",N" & rowNodes.Item("NodeID")
                                End If
                            Next
                        End If
                    Next
                    MyCommon.QueryStr = "Insert into ItemAttribSearchResult with (RowLock) (ProdGroupID,NodeIDs) values" & _
                                              "(" & ProductGroupID & ",'" & strNodes & "')"
                    MyCommon.LRT_Execute()
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

    Sub BindDstostartvalues(ByVal ItemAttrbValue As String, ByVal ProductGroupID As String, _
                            ByVal SelectedNodeID As Long, ByVal SelectedHierID As Integer, ByVal AttrbId As String)

        Dim dt As DataTable
        Dim row As DataRow
        Dim dtExtHierDetails As DataTable
        Dim ExternalHierarchy As String = "00"
        Dim StartRangeId, EndRangeId, RbName, RbRangeName As String
        Dim optValMaxCharLen As Integer = 0
        Dim OrgCharLen As Integer = 120
        If AttrbId = "ItemAttrib1" Then
            StartRangeId = "StartRange1"
            EndRangeId = "EndRange1"
            RbName = "Rb1"
            RbRangeName = "RbRange1"
        ElseIf AttrbId = "ItemAttrib2" Then
            StartRangeId = "StartRange2"
            EndRangeId = "EndRange2"
            RbName = "Rb2"
            RbRangeName = "RbRange2"
        ElseIf AttrbId = "ItemAttrib3" Then
            StartRangeId = "StartRange3"
            EndRangeId = "EndRange3"
            RbName = "Rb3"
            RbRangeName = "RbRange3"
        End If
        Try
            If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
            Send("<table class=""list""><thead></thead><tbody><tr>")
            If ItemAttrbValue <= 0 Then
                Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """/></td>")
                Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
            Else
                MyCommon.QueryStr = "select SupportRange from HierAttribDefinition with (NoLock) where HierAttribID = " & ItemAttrbValue
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    If dt.Rows(0)(0) = True Then
                        Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """/></td>")
                        Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
                    Else
                        Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
                        Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """  disabled=""disabled""/></td>")
                    End If
                Else
                    Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
                    Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """ disabled=""disabled""/></td>")
                End If
            End If
            'response.write("select SupportRange from HierAttribDefinition with (NoLock) where HierAttribID = " & ItemAttrbValue )
            Send("<td style=""width: 120px;"">")
            If SelectedNodeID <= 0 And SelectedHierID <= 0 Then
                Send("<select name=""" + StartRangeId + """ id=""" + StartRangeId + """ style=""width: 120px;"">")
                Send("<option value=""""></option>")
                Send("</select>")
                Send("</td>")
            Else
                If SelectedHierID > 0 Then
                    MyCommon.QueryStr = "Select ExternalID from ProdHierarchies with (NoLock) where HierarchyID = " & SelectedHierID
                ElseIf SelectedNodeID > 0 Then
                    MyCommon.QueryStr = "Select ExternalID from PHnodes with (NoLock) where NodeID = " & SelectedNodeID
                End If
                dtExtHierDetails = MyCommon.LRT_Select
                If dtExtHierDetails.Rows.Count > 0 Then
                    ExternalHierarchy = dtExtHierDetails.Rows(0)(0)
                End If
                MyCommon.QueryStr = "Select distinct HierAttribValue from BannerHierAttribValues with (NoLock) where " & _
                                    "HierAttribID = " & ItemAttrbValue & " and HierID like '" & ExternalHierarchy & "%' and HierAttribValue<>'' order by HierAttribValue"
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    OrgCharLen = 120
                    For Each row In dt.Rows
                        optValMaxCharLen = Len(MyCommon.NZ(row.Item("HierAttribValue"), "")) * 8 + 8
                        If optValMaxCharLen > OrgCharLen Then
                            OrgCharLen = optValMaxCharLen
                        End If
                    Next
                    If Request.Browser.Type.Contains("IE") = True And Request.Browser.Version < 9 Then
                        Send("<select name=""" + StartRangeId + """ id=""" + StartRangeId + """ onmouseover=""javascript:dynamicWidth(" + StartRangeId + ",true," & OrgCharLen & ");""  onload=""javascript:dynamicWidth(" + StartRangeId + ",false,120);"" onblur=""javascript:dynamicWidth(" + StartRangeId + ",false,120);"" onchange=""javascript:dynamicWidth(" + StartRangeId + ",false,120);"" style=""width: 120px;"">")
                    Else
                        Send("<select name=""" + StartRangeId + """ id=""" + StartRangeId + """ style=""width: 120px;"">")
                    End If
                    Send("<option value=""""></option>")
                    For Each row In dt.Rows
                        Send("<option value=""" & MyCommon.NZ(row.Item("HierAttribValue"), "") & """>" & MyCommon.NZ(row.Item("HierAttribValue"), "") & "</option>")
                    Next
                Else
                    Send("<select name=""" + StartRangeId + """ id=""" + StartRangeId + """ style=""width: 120px;"">")
                    Send("<option value=""""></option>")
                End If
                Send("</select>")
                Send("</td>")
            End If
            Send("<td style=""width: 120px;"">")
            If SelectedNodeID <= 0 And SelectedHierID <= 0 Then
                Send("<select name=name=""" + EndRangeId + """ id=""" + EndRangeId + """ style=""width: 120px;"">")
                Send("<option value=""""></option>")
                Send("</select>")
                Send("</td>")
            Else
                If dt.Rows.Count > 0 Then
                    If Request.Browser.Type.Contains("IE") = True And Request.Browser.Version < 9 Then
                        Send("<select name=name=""" + EndRangeId + """ id=""" + EndRangeId + """ onmouseover=""javascript:dynamicWidth(" + EndRangeId + ",true," & OrgCharLen & ");"" onload=""javascript:dynamicWidth(" + EndRangeId + ",false,120);"" onblur=""javascript:dynamicWidth(" + EndRangeId + ",false,120);"" onchange=""javascript:dynamicWidth(" + EndRangeId + ",false,120);"" style=""width: 120px;"">")
                    Else
                        Send("<select name=name=""" + EndRangeId + """ id=""" + EndRangeId + """ style=""width: 120px;"">")
                    End If
                    Send("<option value=""""></option>")
                    For Each row In dt.Rows
                        Send(" <option value=""" & MyCommon.NZ(row.Item("HierAttribValue"), "") & """>" & MyCommon.NZ(row.Item("HierAttribValue"), "") & "</option>")
                    Next
                Else
                    Send("<select name=name=""" + EndRangeId + """ id=""" + EndRangeId + """ style=""width: 120px;"">")
                    Send("<option value=""""></option>")
                End If
                Send("</select>")
                Send("</td>")
            End If
            Send("</tr></tbody></table>")
        Catch ex As Exception
            Send(ex.ToString())
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub AssignPrevValues(ByVal ProductGroupID As String, ByVal SelectedNodeID As Long, _
                         ByVal SelectedHierID As Integer, ByVal AttrbId As String)
        Try
            Dim ItemAttrb As Integer = 0
            Dim selectedtype As Boolean = False
            Dim lowvalue As String = ""
            Dim highvalue As String = ""
            Dim attribserchid As Integer
            Dim dt As DataTable
            Dim row As DataRow
            Dim SupportRange As Boolean = False
            Dim StartRangeId, EndRangeId, RbName, RbRangeName As String
            Dim dtExtHierDetails As DataTable
            Dim ExternalHierarchy As String = "00"
            Dim optValMaxCharLen As Integer = 0
            Dim OrgCharLen As Integer = 120
            If AttrbId = "ItemAttrib1" Then
                StartRangeId = "StartRange1"
                EndRangeId = "EndRange1"
                RbName = "Rb1"
                RbRangeName = "RbRange1"
                attribserchid = 1
            ElseIf AttrbId = "ItemAttrib2" Then
                StartRangeId = "StartRange2"
                EndRangeId = "EndRange2"
                RbName = "Rb2"
                RbRangeName = "RbRange2"
                attribserchid = 2
            ElseIf AttrbId = "ItemAttrib3" Then
                StartRangeId = "StartRange3"
                EndRangeId = "EndRange3"
                RbName = "Rb3"
                RbRangeName = "RbRange3"
                attribserchid = 3
            End If

            If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
            Dim dtvalues As DataTable
            MyCommon.QueryStr = "Select top 1 AttribNo,ProdHierAttibuteID,Selection,isnull(LowValue,'')LowValue,isnull(HighValue,'')HighValue," & _
                                " ExternalID,ExtHierName from ItemAttribSearchHistory with (NoLock) where AttribNo=" & attribserchid.ToString()
            dtvalues = MyCommon.LRT_Select
            If dtvalues.Rows.Count > 0 Then
                ItemAttrb = dtvalues.Rows(0)("ProdHierAttibuteID")
                selectedtype = dtvalues.Rows(0)("Selection")
                lowvalue = dtvalues.Rows(0)("LowValue")
                highvalue = dtvalues.Rows(0)("HighValue")
            End If
            Send("<table class=""list""><thead></thead><tbody><tr>")
            If ItemAttrb <= 0 Then
                SupportRange = True
            Else
                MyCommon.QueryStr = "select SupportRange from HierAttribDefinition with (NoLock) where HierAttribID = " & ItemAttrb
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    If dt.Rows(0)(0) = True Then
                        SupportRange = True
                    End If
                End If
            End If

            If selectedtype = False Then
                If SupportRange = False Then
                    Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
                    Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """  disabled=""disabled""/></td>")
                Else
                    Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """/></td>")
                    Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
                End If
            Else
                If SupportRange = False Then
                    Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
                    Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """  disabled=""disabled""/></td>")
                Else
                    Send("<td style=""width: 30px;""><input id=""" + RbName + """ type=""radio"" name=""" + RbName + """ checked=""CHECKED""/></td>")
                    Send("<td style=""width: 40px;""><input id=""" + RbRangeName + """ type=""radio"" name=""" + RbName + """/></td>")
                End If
            End If

            Send("<td style=""width: 120px;"">")
            If SelectedHierID > 0 Then
                MyCommon.QueryStr = "Select ExternalID from ProdHierarchies with (NoLock) where HierarchyID = " & SelectedHierID
            ElseIf SelectedNodeID > 0 Then
                MyCommon.QueryStr = "Select ExternalID from PHnodes with (NoLock) where NodeID = " & SelectedNodeID
            End If
            dtExtHierDetails = MyCommon.LRT_Select
            If dtExtHierDetails.Rows.Count > 0 Then
                ExternalHierarchy = dtExtHierDetails.Rows(0)(0)
            End If

            MyCommon.QueryStr = "Select HierAttribValue from(select distinct HierAttribValue from BannerHierAttribValues with (NoLock) where " & _
                             "HierAttribID = " & ItemAttrb.ToString & " and HierID like '" & ExternalHierarchy & "%' Union " & _
                             "select '" & lowvalue & "' 'HierAttribValue' Union " & _
                                "select '" & highvalue & "' 'HierAttribValue' )HierTable where HierAttribValue<>'' order by HierAttribValue"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                OrgCharLen = 120
                For Each row In dt.Rows
                    optValMaxCharLen = Len(MyCommon.NZ(row.Item("HierAttribValue"), "")) * 8 + 8
                    If optValMaxCharLen > OrgCharLen Then
                        OrgCharLen = optValMaxCharLen
                    End If
                Next
                If Request.Browser.Type.Contains("IE") = True And Request.Browser.Version < 9 Then
                    Send("<select name=""" + StartRangeId + """ id=""" + StartRangeId + """ onmouseover=""javascript:dynamicWidth(" + StartRangeId + ",true," & OrgCharLen & ");""  onload=""javascript:dynamicWidth(" + StartRangeId + ",false,120);"" onblur=""javascript:dynamicWidth(" + StartRangeId + ",false,120);"" onchange=""javascript:dynamicWidth(" + StartRangeId + ",false,120);"" style=""width: 120px;"">")
                Else
                    Send("<select name=""" + StartRangeId + """ id=""" + StartRangeId + """ style=""width: 120px;"">")
                End If
                Send("<option value=""""></option>")
                For Each row In dt.Rows
                    If lowvalue = row.Item("HierAttribValue") Then
                        Send(" <option value=""" & MyCommon.NZ(row.Item("HierAttribValue"), "") & """ selected=""selected"">" & MyCommon.NZ(row.Item("HierAttribValue"), "") & "</option>")
                    Else
                        Send(" <option value=""" & MyCommon.NZ(row.Item("HierAttribValue"), "") & """>" & MyCommon.NZ(row.Item("HierAttribValue"), "") & "</option>")
                    End If
                Next
            Else
                Send("<select name=""" + StartRangeId + """ id=""" + StartRangeId + """ style=""width: 120px;"">")
                Send("<option value=""""></option>")
            End If
            Send("</select>")
            Send("</td>")
            Send("<td style=""width: 120px;"">")
            If dt.Rows.Count > 0 Then
                If Request.Browser.Type.Contains("IE") = True And Request.Browser.Version < 9 Then
                    Send("<select name=name=""" + EndRangeId + """ id=""" + EndRangeId + """ onmouseover=""javascript:dynamicWidth(" + EndRangeId + ",true," & OrgCharLen & ");"" onload=""javascript:dynamicWidth(" + EndRangeId + ",false,120);"" onblur=""javascript:dynamicWidth(" + EndRangeId + ",false,120);"" onchange=""javascript:dynamicWidth(" + EndRangeId + ",false,120);"" style=""width: 120px;"">")
                Else
                    Send("<select name=name=""" + EndRangeId + """ id=""" + EndRangeId + """ style=""width: 120px;"">")
                End If
                Send("<option value=""""></option>")
                For Each row In dt.Rows
                    If highvalue = row.Item("HierAttribValue") Then
                        Send(" <option value=""" & MyCommon.NZ(row.Item("HierAttribValue"), "") & """ selected=""selected"">" & MyCommon.NZ(row.Item("HierAttribValue"), "") & "</option>")
                    Else
                        Send(" <option value=""" & MyCommon.NZ(row.Item("HierAttribValue"), "") & """>" & MyCommon.NZ(row.Item("HierAttribValue"), "") & "</option>")
                    End If
                Next
            Else
                Send("<select name=name=""" + EndRangeId + """ id=""" + EndRangeId + """ style=""width: 120px;"">")
                Send("<option value=""""></option>")
            End If
            Send("</select>")
            Send("</td>")
            Send("</tr></tbody></table>")
        Catch ex As Exception
            Send("<errorstr>" & ex.ToString() & "</errorstr>")
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub SendLastItemAttrbValues()
        Try
            Dim ItemAttrb1 As String = "-1"
            Dim ItemAttrb2 As String = "-1"
            Dim ItemAttrb3 As String = "-1"

            If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
            Dim dt As DataTable
            Dim row As DataRow
            MyCommon.QueryStr = "Select AttribNo,ProdHierAttibuteID from ItemAttribSearchHistory with (NoLock)"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                For Each row In dt.Rows
                    Select Case row.Item("AttribNo")
                        Case 1
                            ItemAttrb1 = row.Item("ProdHierAttibuteID")
                        Case 2
                            ItemAttrb2 = row.Item("ProdHierAttibuteID")
                        Case 3
                            ItemAttrb3 = row.Item("ProdHierAttibuteID")
                    End Select
                Next
            End If
            Send("<ItemAttrb1>" & ItemAttrb1 & "</ItemAttrb1>")
            Send("<ItemAttrb2>" & ItemAttrb2 & "</ItemAttrb2>")
            Send("<ItemAttrb3>" & ItemAttrb3 & "</ItemAttrb3>")
            Send("<errorstr></errorstr>")
        Catch ex As Exception
            Send("<errorstr>" & ex.ToString() & "</errorstr>")
        Finally
            MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub UpdateItemAttributeProductsToHierarchy(ByVal ActionType As String, ByVal AdminUserID As Long)
        Try
            Dim ProductGroupID As String = Nothing
            Dim NodeIDs As String = Nothing

            If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
            Dim dtvalues As DataTable
            MyCommon.QueryStr = "SELECT ProdGroupID, NodeIDs FROM ItemAttribSearchResult with (NoLock)"
            dtvalues = MyCommon.LRT_Select
            If dtvalues.Rows.Count > 0 Then
                ProductGroupID = dtvalues.Rows(0)("ProdGroupID")
                NodeIDs = dtvalues.Rows(0)("NodeIDs")
            Else
                'No matches were found	 
                Send("<table class=""list""><thead></thead><tbody><tr>")
            End If

            If ActionType = "LinkMatchestoHierarchy" Then
                HandleLinkToGroup(CLng(ProductGroupID), NodeIDs, AdminUserID, True)
                Send("<Status>Success</Status>")
            ElseIf ActionType = "RemoveMatchesFromHierarchy" Then
                HandleRemoveLinkToGroup(CLng(ProductGroupID), NodeIDs, AdminUserID)
                Send("<Status>Success</Status>")
            End If
        Catch ex As Exception
            Send("<errorstr>" & ex.ToString() & "</errorstr>")
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

    Public Sub handleSearchItemAdjust(ByVal ActionType As String, ByVal HierarchyID As Integer, ByVal ProductGroupID As String, ByVal ProductID As String)
        Dim dt As DataTable
        Dim retVal As String = ""
        Dim Logix As New Copient.LogixInc
        Dim AdminUserID As Long

        Try
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            AdminUserID = Verify_AdminUser(MyCommon, Logix)
            If (ActionType = "add") Then
                MyCommon.QueryStr = "Select ProductID from ProdGroupItems with (NoLock) where ProductGroupID=" & ProductGroupID & _
                                    " and ProductID=" & ProductID & " and Deleted=0;"
                dt = MyCommon.LRT_Select
                If (dt.Rows.Count = 0) Then
                    MyCommon.QueryStr = "Insert into ProdGroupItems with (RowLock) (ProductGroupID, ProductID, Manual, Deleted, CMOAStatusFlag, TCRMAStatusFlag, CPEStatusFlag, UEStatusFlag) " & _
                                        "values (" & ProductGroupID & ", " & ProductID & ", 1, 0, 2, 2, 2, 2);"
                    MyCommon.LRT_Execute()
                    MyCommon.Activity_Log(5, CLng(ProductGroupID), AdminUserID, Copient.PhraseLib.Detokenize("history.hierarchy-AddedItem", LanguageID, Request.QueryString("extID")))
                End If
                retVal = "ADDED"
            ElseIf (ActionType = "remove") Then
                MyCommon.QueryStr = "Delete from ProdGroupItems with (RowLock) where ProductGroupID=" & ProductGroupID & " and ProductID=" & ProductID & " and Deleted=1;"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "Update ProdGroupItems with (RowLock) set Deleted=1, Manual=1, CMOAStatusFlag=2, TCRMAStatusFlag=2, CPEStatusFlag=2, UEStatusFlag=2 where Deleted=0 and ProductGroupID=" & ProductGroupID & " " & _
                                    "and ProductID = " & ProductID & ";"
                MyCommon.LRT_Execute()
                If (MyCommon.RowsAffected > 0) Then
                    MyCommon.Activity_Log(5, CLng(ProductGroupID), AdminUserID, Copient.PhraseLib.Detokenize("history.hierarchy-RemovedItem", LanguageID, Request.QueryString("extID")))
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

    Public Sub handleSearchNodeAdjust(ByVal ActionType As String, ByVal HierarchyID As Integer, ByVal ProductGroupID As String, ByVal NodeID As String)
        Dim PreCount, PostCount As Integer
        Dim foundItems As Boolean = False
        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            PreCount = GetProductGroupItemCount(ProductGroupID)
            TransmitGroups("N" & NodeID, ProductGroupID, foundItems, IIf(ActionType = "ADD", True, False))
            PostCount = GetProductGroupItemCount(ProductGroupID)
            Send(GetAssignedTrailer(PreCount, PostCount, foundItems))
        Catch ex As Exception
            Send(ex.ToString)
        Finally
        End Try
    End Sub

    Public Sub RecordSelectedNode(ByVal PgID As String, ByVal NodeID As String, ByVal AddToPG As Boolean)
        Dim Logix As New Copient.LogixInc
        Dim AdminUserID As Long
        Dim ChildList As ArrayList
        Dim NodeList As String() = Nothing

        Try
            MyCommon.AppName = "HierarchyFeeds.aspx"
            If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
            AdminUserID = Verify_AdminUser(MyCommon, Logix)

            MyCommon.QueryStr = "dbo.pa_ProductGroupNodes_Update"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@PgId", SqlDbType.BigInt).Value = CLng(PgID)
            MyCommon.LRTsp.Parameters.Add("@NodeID", SqlDbType.BigInt).Value = CLng(NodeID)
            MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.BigInt).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@LinkDate", SqlDbType.DateTime).Value = Date.Now
            MyCommon.LRTsp.Parameters.Add("@AddToPG", SqlDbType.Bit).Value = IIf(AddToPG, 1, 0)
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()

            ' remove all child nodes of this node
            ChildList = GetChildNodes(NodeID)
            If (ChildList.Count > 0) Then
                NodeList = Array.ConvertAll(Of Object, String)(ChildList.ToArray, New Converter(Of Object, String)(AddressOf cString))
                MyCommon.QueryStr = "delete from ProductGroupNodes with (RowLock) where NodeID in (" & String.Join(",", NodeList) & ")"
                MyCommon.LRT_Execute()
            End If

        Catch ex As Exception
            Send(ex.ToString)
        End Try
    End Sub

    Sub HandleDelFromHierarchy(ByVal NodeID As String, ByVal ItemID As String)
        Dim isLocation As Boolean
        Dim dt As DataTable
        Dim HierID As Integer
        Dim ParentID As Long
        Dim ExtID As String
        Dim DelNodeID As String
        Dim Logix As New Copient.LogixInc
        Dim AdminUserID As Long
        Dim refreshType As Integer = 0
        Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
        Dim isPABEnabled As Boolean = (MyCommon.Fetch_UE_SystemOption(157) = "1")
        Dim AllowProductNodeDeletion As Boolean = (MyCommon.Fetch_UE_SystemOption(179) = "1")
        Dim HierarchyID As Integer
        Try
            MyCommon.AppName = "HierarchyFeeds.aspx"
            If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
            AdminUserID = Verify_AdminUser(MyCommon, Logix)

            If (Request.QueryString("isHier") = "1") Then
                MyCommon.QueryStr = "select ExternalId from ProdHierarchies with (NoLock) where HierarchyID =" & NodeID & ";"
                dt = MyCommon.LRT_Select
                If (dt.Rows.Count > 0) Then
                    ExtID = MyCommon.NZ(dt.Rows(0).Item("ExternalId"), "")
                    If isPABEnabled AndAlso Not AllowProductNodeDeletion AndAlso MyResyncer.IsPHierarchyHasAnyProductGroups(ExtID) Then
                        Send(Copient.PhraseLib.Lookup("term.hierarchyinuse", LanguageID))
                    Else
                        MyResyncer.Remove_Hierarchy(ExtID)
                        Send("<refresh>2</refresh>")
                        Send("<node>" & NodeID & "</node>")
                    End If
                Else
                    Send(Copient.PhraseLib.Lookup("hierarchy.UnableToDelete", LanguageID))
                End If
            Else
                isLocation = (ItemID <> "" AndAlso ItemID.Length > 0 AndAlso ItemID.Substring(0, 1) = "I")

                If (isLocation) Then
                    MyCommon.QueryStr = "delete from PHContainer with (RowLock) where nodeid = " & NodeID & " and productid = " & ItemID.Substring(1)
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

                    MyCommon.QueryStr = "select HierarchyID, ExternalID, ParentID from PHNodes where nodeid = " & DelNodeID
                    dt = MyCommon.LRT_Select
                    If (dt.Rows.Count > 0) Then
                        HierarchyID = dt.Rows(0).Item("HierarchyID")
                        If isPABEnabled AndAlso Not AllowProductNodeDeletion AndAlso MyResyncer.IsNodeAssociatedwithAnyAttributeBasedProductGroups(MyCommon.NZ(dt.Rows(0).Item("ExternalID"), ""), HierarchyID) Then
                            Send(Copient.PhraseLib.Lookup("term.nodeinuse", LanguageID))
                        Else
                            HierID = MyCommon.NZ(dt.Rows(0).Item("HierarchyID"), -1)
                            ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), -1)
                            ExtID = MyCommon.NZ(dt.Rows(0).Item("ExternalID"), "")

                            MyResyncer.Remove_Node(ExtID, HierID)
                            Send("<refresh>" & refreshType & "</refresh>")
                            Send("<node>" & DelNodeID & "</node>")
                            Send("<parent>" & ParentID & "</parent>")
                        End If

                    End If
                End If
            End If
        Catch ex As Exception
            Send(ex.ToString)
        End Try

    End Sub



    Sub HandleLinkToGroup(ByVal ProductGroupID As Long, ByVal IDList As String, ByVal AdminUserID As Long)
        Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
        Dim RetMsg As String = ""
        Dim count As Integer = 0

        RetMsg = MyResyncer.LinkNodesToGroup(ProductGroupID, IDList)

        If RetMsg = "" Then
            MyCommon.Activity_Log2(5, 2, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.hierarchy-Linked", LanguageID))
        End If

        SendLinkingXML(ProductGroupID, RetMsg, False)
    End Sub

    Sub HandleLinkToGroup(ByVal ProductGroupID As Long, ByVal IDList As String, ByVal AdminUserID As Long, ByVal ItemAttributeSearch As Boolean)
        'used for Product Item Attributes search
        Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
        Dim RetMsg As String = ""
        Dim count As Integer = 0

        MyResyncer.SetItemAttributeSearchState(ItemAttributeSearch)
        RetMsg = MyResyncer.LinkNodesToGroup(ProductGroupID, IDList)

        If RetMsg = "" Then
            MyCommon.Activity_Log2(5, 2, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.hierarchy-Linked", LanguageID))
        End If

        SendLinkingXML(ProductGroupID, RetMsg, False)
    End Sub

    Sub HandleRemoveLinkToGroup(ByVal ProductGroupID As Long, ByVal IDList As String, ByVal AdminUserID As Long)
        Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
        Dim RetMsg As String = ""

        RetMsg = MyResyncer.UnlinkNodesFromGroup(ProductGroupID, IDList)

        If RetMsg = "" Then
            MyCommon.Activity_Log2(5, 2, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.hierarchy-Delinked", LanguageID))
        End If

        SendLinkingXML(ProductGroupID, RetMsg, True)
    End Sub

    Sub HandleExcludeFromGroup(ByVal ProductGroupID As Long, ByVal SelectedNode As Long, ByVal IDList As String, ByVal AdminUserID As Long)
        Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
        Dim RetMsg As String = ""

        RetMsg = MyResyncer.ExcludeFromGroup(ProductGroupID, SelectedNode, IDList)
        If RetMsg = "" Then
            MyCommon.Activity_Log2(5, 2, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.hierarchy-LinkedExcluded", LanguageID))
        End If

        SendLinkingXML(ProductGroupID, RetMsg, False)
    End Sub

    Sub HandleRemoveExclusion(ByVal ProductGroupID As Long, ByVal IDList As String, ByVal AdminUserID As Long)
        Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
        Dim RetMsg As String = ""

        RetMsg = MyResyncer.RemoveExclusionFromGroup(ProductGroupID, IDList)

        If RetMsg = "" Then
            MyCommon.Activity_Log2(5, 2, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.hierarchy-DelinkedExcluded", LanguageID))
        End If

        SendLinkingXML(ProductGroupID, RetMsg, False)
    End Sub

    Sub HandleAssignSigns(ByVal HierarchySign As String, ByVal NodeId As Integer, ByVal HierarchyID As Long, ByVal AdminUserID As Long)
        'Dim SignDescription As String
        Dim ExternalID As String
        Dim ParentId As Long
        Dim dt As DataTable
        'Dim RetMsg As String = ""
        Try
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            MyCommon.AppName = "HierarchyFeeds.aspx"
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "Select ParentID,ExternalId From PHNodes with (nolock) where NodeID=" & NodeId
            dt = MyCommon.LRT_Select
            ParentId = MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0)
            If ParentId = 0 Then
                ExternalID = MyCommon.NZ(dt.Rows(0).Item("ExternalId"), 0)
            Else
                Do While (ParentId <> 0)
                    MyCommon.QueryStr = "Select ParentID,ExternalID From PHNodes with (nolock) where NodeID=" & ParentId
                    dt = MyCommon.LRT_Select
                    ParentId = MyCommon.NZ(dt.Rows(0).Item("ParentID"), 0)
                    ExternalID = MyCommon.NZ(dt.Rows(0).Item("ExternalID"), "")
                Loop
            End If
            MyCommon.QueryStr = "Select ExternalId From HierarchySelectedSigns with (nolock) where ExternalID='" & ExternalID & "'"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                MyCommon.QueryStr = "Update HierarchySelectedSigns with (rowlock) set SelectedSign='" & HierarchySign & "',LastUpdated=getdate() where ExternalID='" & ExternalID & "'"
                MyCommon.LRT_Execute()
            Else
                MyCommon.QueryStr = "Insert into HierarchySelectedSigns with (rowlock) (ExternalID,SelectedSign,LastUpdated) values ('" & ExternalID & "','" & HierarchySign & "',getdate());"
                MyCommon.LRT_Execute()
            End If
            Sendb("OK")
        Catch ex As Exception
            Sendb("NO")
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
    End Sub

    Sub SendLinkingXML(ByVal ProductGroupID As Long, ByVal Message As String, ByVal WriteLinkCount As Boolean)
        Dim SelNode As Integer
        Dim dt As DataTable
        Dim LinkedChildCount As Integer

        Sendb("<trailer>")

        If Message IsNot Nothing AndAlso Message <> "" Then
            Sendb("<message>" & Message & "</message>")
        End If

        Sendb("<count>")
        Sendb(GetProductGroupItemCount(ProductGroupID))
        Sendb("</count>")

        If WriteLinkCount AndAlso Request.Form("sel") <> "" Then
            SelNode = MyCommon.Extract_Val(Request.Form("sel"))
            If SelNode > 0 Then
                MyCommon.QueryStr = "select COUNT(*) as LinkedChildCount from ProdGroupHierarchies as PGH with (NoLock) " & _
                                    "inner join (select PH.ExternalID as ExtHierarchyID, PHN.ExternalID as ExtNodeID " & _
                                    "from PHNodes as PHN with (NoLock) " & _
                                    "inner join ProdHierarchies as PH with (NoLock) on PH.HierarchyID = PHN.HierarchyID " & _
                                    "where PHN.ParentID = " & SelNode & ") as t1 on t1.ExtHierarchyID = PGH.ExtHierarchyID and t1.ExtNodeID = PGH.ExtNodeID "
                dt = MyCommon.LRT_Select
                LinkedChildCount = dt.Rows(0).Item("LinkedChildCount")
                Sendb("<linked>" & LinkedChildCount & "</linked>")
            End If
        End If

        Sendb("</trailer>")
    End Sub

</script>
