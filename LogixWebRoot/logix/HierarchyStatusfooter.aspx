 <%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%  ' version:6.0.1.91435.Official Build (SUSDAY10083) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    Dim dt As DataTable = Nothing
    Dim SelectedNodeIDs As String = ""

    Dim AssignedCount As Integer = 0
    Dim SelectedCount As Integer = 0
    Dim pgID As Integer = 0 
    Dim SelectedNotAssignedCount As Integer = 0
    Dim LastSelectedCount = 0
    Dim ChkboxChecked As Integer = 0
    Dim SelectedItemsAreAssigned As String = " -- Items are already be assigned"
    Dim NotProdGroup As String = ""
    Dim Level As Integer = 0

    MyCommon.AppName = "HierarchyStatusfooter.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix) 
    SelectedNodeIDs = Request.QueryString("nodelist") 
    pgID = Request.QueryString("prodgroupid")

    AssignedCount = Request.QueryString("assignedCt")
    LastSelectedCount = Request.QueryString("lastSelectedCt")
    NotProdGroup = Request.QueryString("notProdGroup")
    ChkboxChecked = Request.QueryString("chkboxChecked")
    Level = Request.QueryString("level")

    If (pgID > 0) Then
    ' Get assigned count
      MyCommon.QueryStr = "Select count(*) as AssignedCount from ProdGroupItems with (NoLock) where ProductGroupID=" & pgID & " and Deleted=0;"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        AssignedCount = MyCommon.NZ(dt.Rows(0).Item("AssignedCount"), 0)
      End If

    ' Get selected count
      SelectedCount = GetProductCount(SelectedNodeIDs)

    ' Get selected count minus assigned count
    If NotProdGroup = "1" Then
       If ChkboxChecked = 1 Then
         SelectedNotAssignedCount =  LastSelectedCount + 1
       Else
         SelectedNotAssignedCount =  LastSelectedCount - 1
       End If
       If SelectedNotAssignedCount < 0 Then
          SelectedNotAssignedCount = 0
       End If
    Else
       If Level = 7 Then 'Selected All Items from prpoduct ID list 
         If ChkboxChecked = 1 Then
           SelectedNotAssignedCount = GetListCount(SelectedNodeIDs)
         Else
           SelectedNotAssignedCount = 0
         End If
       Else
         SelectedNotAssignedCount = GetSelectedNotAssignedCount(SelectedNodeIDs, pgID)
       End If
    End If

    ' below line display only selected but not assigned
    '  If (SelectedNotAssignedCount > LastSelectedCount And ChkboxChecked = 1) OR ChkboxChecked = 0 Then 
        Send("<span>" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & ": <span id=""prodgroupID"">" & pgID & "</span></span>&nbsp;|&nbsp;<span>" & Copient.PhraseLib.Lookup("term.Assigned", LanguageID) & ": <span id=""assignedCt"">" & AssignedCount & "</span></span>&nbsp;|&nbsp;<span>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ": <span id=""snaCt"">" & SelectedNotAssignedCount & "</span></span>")   
    '  Else
    '    Send("<span>" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & ": <span id=""prodgroupID"">" & pgID & "</span></span>&nbsp;|&nbsp;<span>" & Copient.PhraseLib.Lookup("term.Assigned", LanguageID) & ": <span id=""assignedCt"">" & AssignedCount & "</span></span>&nbsp;|&nbsp;<span>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ": <span id=""snaCt"">" & SelectedNotAssignedCount & SelectedItemsAreAssigned & "</span></span>")
    '  End If 
    End If
    Send(" <input type=""hidden"" id=""selectedNotAssignedCt"" name=""selectedNotAssignedCt"" value=""" & SelectedNotAssignedCount & """ />")

    MyCommon.Close_LogixRT()
    Logix = Nothing
    MyCommon = Nothing
%>

<script runat="server">

    Function GetProductCount(ByVal NodeID As String) As Integer
        Dim MyCommon As New Copient.CommonInc
        Dim dt As DataTable
        Try
            If (String.IsNullOrEmpty(NodeID)) Then
                Return 0
            Else
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
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
        Return 0
    End Function

    Function GetSelectedNotAssignedCount(ByVal NodeID As String, ProductGroupID As Integer) As Integer
        Dim MyCommon As New Copient.CommonInc
        Dim dt As DataTable
        Try
            If (String.IsNullOrEmpty(NodeID)) Then
                Return 0
            Else
                Dim dtNodeIDs As DataTable = New DataTable()
                dtNodeIDs.Columns.Add("NodeID")
                
                For Each IDs In NodeID.Trim(",").Split(",")
                    dtNodeIDs.Rows.Add(Convert.ToInt64(IDs))
                Next
                
                MyCommon.AppName = "phierarchytree.aspx"
                MyCommon.QueryStr = "dbo.pt_Product_In_NodeHierarchy_Not_Assigned"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@NodeIDList", SqlDbType.Structured).Value = dtNodeIDs
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
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
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
        End Try
        Return 0
    End Function

    Function GetListCount(Byval ProdIDList) As Integer
       Dim ProdIds As String() = Split(ProdIDList, ",", -1, CompareMethod.Text)
       Return uBound(ProdIds)
    End Function

</script>