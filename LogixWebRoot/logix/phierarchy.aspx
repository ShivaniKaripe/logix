<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%
    ' *****************************************************************************
    ' * FILENAME: phierarchy.aspx 
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
    Dim NodeId As Long
    Dim NodeIds() As String
    Dim ParentId As String
    Dim dr As DataRow
    Dim dtParents As DataTable = Nothing
    Dim dtParents1 As DataTable = Nothing
    Dim dtChildren As DataTable = Nothing
    Dim dtSelected As DataTable = Nothing
    Dim dtAvailable As DataTable = Nothing
    Dim sQuery As String
    Dim SelectedCount As Integer
    Dim AvailableCount As Integer
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim ParentNodeIdList As String
    Dim SelectedNodeId As String
    Dim NodeName As String
    Dim ProductList As String
    Dim Products() As String
    Dim bSave As Boolean
    Dim bDelete As Boolean
    Dim bCreate As Boolean
    Dim bUp As Boolean
    Dim bDown As Boolean
    Dim bRemove As Boolean
    Dim bAdd As Boolean
    Dim i As Integer
    Dim ItemPKID As Integer = -1
    Dim SelectedOption As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim MyResyncer As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
    Dim isPABEnabled As Boolean = (MyCommon.Fetch_UE_SystemOption(157) = "1")
    Dim AllowProductNodeDeletion As Boolean = (MyCommon.Fetch_UE_SystemOption(179) = "1")

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "phierarchy.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    Try
        If Request.RequestType = "GET" Then
            ParentNodeIdList = Request.QueryString("ParentNodeIdList")
            SelectedNodeId = Request.QueryString("SelectedNodeId")
            NodeName = Request.QueryString("NodeName")
            If Request.QueryString("NewNode") = "" Then
                bSave = False
            Else
                bSave = True
            End If
            If Request.QueryString("DeleteNode") = "" Then
                bDelete = False
            Else
                bDelete = True
            End If
            If Request.QueryString("mode") = "" Then
                bCreate = False
            Else
                bCreate = True
            End If
            If Request.QueryString("up") = "" Then
                bUp = False
            Else
                bUp = True
            End If
            If Request.QueryString("down") = "" Then
                bDown = False
            Else
                bDown = True
            End If
            If Request.QueryString("removestore") = "" Then
                bRemove = False
            Else
                bRemove = True
            End If
            If Request.QueryString("selectstore") = "" Then
                bAdd = False
            Else
                bAdd = True
            End If
        Else
            ParentNodeIdList = Request.Form("ParentNodeIdList")
            SelectedNodeId = Request.Form("SelectedNodeId")
            NodeName = Request.Form("NodeName")
            If Request.Form("NewNode") = "" Then
                bSave = False
            Else
                bSave = True
            End If
            If Request.Form("DeleteNode") = "" Then
                bDelete = False
            Else
                bDelete = True
            End If
            If Request.Form("mode") = "" Then
                bCreate = False
            Else
                bCreate = True
            End If
            If Request.Form("up") = "" Then
                bUp = False
            Else
                bUp = True
            End If
            If Request.Form("down") = "" Then
                bDown = False
            Else
                bDown = True
            End If
            If Request.Form("removestore") = "" Then
                bRemove = False
            Else
                bRemove = True
            End If
            If Request.Form("selectstore") = "" Then
                bAdd = False
            Else
                bAdd = True
            End If
        End If


        Send_HeadBegin("term.producthierarchy")
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 8)
        Send_Subtabs(Logix, 8, 4)

        If (Logix.UserRoles.EditSystemConfiguration = False) Then
            Send_Denied(1, "perm.admin-configuration")
            GoTo done
        End If

        If bDown Then
            If ParentNodeIdList = "" Then
                ParentNodeIdList = SelectedNodeId
            Else
                ParentNodeIdList += "," & SelectedNodeId
            End If
            SelectedNodeId = ""
            NodeName = ""
        End If

        If bUp Then
            If ParentNodeIdList <> "" Then
                Dim n As Integer
                n = ParentNodeIdList.LastIndexOf(",")
                If n > 0 Then
                    SelectedNodeId = ParentNodeIdList.Substring(n + 1)
                    ParentNodeIdList = ParentNodeIdList.Substring(0, n)
                Else
                    SelectedNodeId = ParentNodeIdList
                    ParentNodeIdList = ""
                End If
            End If
            NodeName = ""
        End If

        If bDelete AndAlso SelectedNodeId <> "" Then
            ' delete current node and all children
            infoMessage = MyResyncer.RemoveNodeAndAllChildren(ParentNodeIdList = "", SelectedNodeId)
            If infoMessage <> "" Then
                infoMessage = Copient.PhraseLib.Lookup(infoMessage, LanguageID)
            End If
        End If
        If ParentNodeIdList = "" Then
            ' No hierarchy is selected, so no parents
            ' existing Hierarchies makeup children
            ParentId = ""
            'sQuery = "select HierarchyId as Id, Name from ProdHierarchies with (NoLock)"
            sQuery = "select HierarchyId as Id,Name = " & _
                    "   case  " & _
                    "       when ExternalID is NULL then Name " & _
                    "       when ExternalID = '' then Name " & _
                    "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                    "       else ExternalID " & _
                    "   end " & _
                    "from ProdHierarchies with (NoLock);"
        Else
            NodeIds = ParentNodeIdList.Split(",")
            'MyCommon.QueryStr = "select HierarchyId as Id, Name from ProdHierarchies with (NoLock) where HierarchyId = " & NodeIds(0)
            MyCommon.QueryStr = "select HierarchyId as Id,Name = " & _
                                "   case  " & _
                                "       when ExternalID is NULL then Name " & _
                                "       when ExternalID = '' then Name " & _
                                "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                                "       else ExternalID " & _
                                "   end " & _
                                "from ProdHierarchies with (NoLock) where HierarchyID=" & NodeIds(0)
            dtParents = MyCommon.LRT_Select()
            If NodeIds.Length = 1 Then
                ' No node is selected, selected hierarchy is parent
                ' existing root nodes for this hierarchy makeup children
                ParentId = "0"
                'sQuery = "select NodeId as Id, Name from PHNodes with (NoLock) where ParentId = 0 and HierarchyId = " & NodeIds(0)
                sQuery = "select NodeId as Id, Name = " & _
                        "   case  " & _
                        "       when ExternalID is NULL then Name " & _
                        "       when ExternalID = '' then Name " & _
                        "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                        "       else ExternalID " & _
                        "   end " & _
                        "from PHNodes with (NoLock) where ParentId = 0 and HierarchyId = " & NodeIds(0)
            Else
                For i = 1 To NodeIds.Length - 1
                    'MyCommon.QueryStr = "select NodeId as Id, Name from PHNodes with (NoLock) where NodeId = " & NodeIds(i)
                    MyCommon.QueryStr = "select NodeId as Id, Name= " & _
                                        "   case  " & _
                                        "       when ExternalID is NULL then Name " & _
                                        "       when ExternalID = '' then Name " & _
                                        "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                                        "       else ExternalID " & _
                                        "   end " & _
                                        "from PHNodes with (NoLock) where NodeId = " & NodeIds(i)
                    dtParents1 = MyCommon.LRT_Select()
                    dtParents.Merge(dtParents1)
                Next
                ' parents consist of hierarchy and listed nodes
                ' chidren madeup of nodes with parentId = last parent
                ParentId = dtParents.Rows(dtParents.Rows.Count - 1).Item("id").ToString
                'sQuery = "select NodeId as Id, Name from PHNodes with (NoLock) where ParentId = " & ParentId
                sQuery = "select NodeId as Id,Name= " & _
                        "   case  " & _
                        "       when ExternalID is NULL then Name " & _
                        "       when ExternalID = '' then Name " & _
                        "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                        "       else ExternalID " & _
                        "   end " & _
                        "from PHNodes with (NoLock) where ParentId = " & ParentId
            End If
        End If

        MyCommon.QueryStr = sQuery
        dtChildren = MyCommon.LRT_Select()

        If bSave Then
            If dtChildren.Rows.Count > 0 Then
                For Each dr In dtChildren.Rows
                    If NodeName.ToUpper = dr.Item("Name").ToString.ToUpper Then
                        bSave = False
                        Exit For
                    End If
                Next
            End If
            If bSave Then
                If ParentId = "" Then
                    ' Hierachy
                    MyCommon.QueryStr = "dbo.pt_ProdHierarchies_Insert"
                    MyCommon.Open_LRTsp()
                    NodeName = Logix.TrimAll(NodeName)
                    MyCommon.LRTsp.Parameters.Add("@ExternalId", SqlDbType.NVarChar, 20).Value = ""
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = NodeName
                    MyCommon.LRTsp.Parameters.Add("@DisplayID", SqlDbType.NVarChar, 100).Value = ""
                    MyCommon.LRTsp.Parameters.Add("@Restricted", SqlDbType.Bit).Value = False
                    MyCommon.LRTsp.Parameters.Add("@HierarchyId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    If (NodeName = "") Then
                        infoMessage = Copient.PhraseLib.Lookup("hierarchy.noname", LanguageID)
                    Else
                        MyCommon.LRTsp.ExecuteNonQuery()
                    End If
                    NodeId = MyCommon.LRTsp.Parameters("@HierarchyId").Value
                    MyCommon.Close_LRTsp()
                Else
                    ' Nodes
                    MyCommon.QueryStr = "dbo.pt_PHNodes_InsertNode"
                    MyCommon.Open_LRTsp()
                    NodeName = Logix.TrimAll(NodeName)
                    MyCommon.LRTsp.Parameters.Add("@HierarchyId", SqlDbType.BigInt, 8).Value = NodeIds(0)
                    MyCommon.LRTsp.Parameters.Add("@ExternalId", SqlDbType.NVarChar, 120).Value = ""
                    MyCommon.LRTsp.Parameters.Add("@ParentId", SqlDbType.BigInt, 8).Value = ParentId
                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = NodeName
                    MyCommon.LRTsp.Parameters.Add("@DisplayID", SqlDbType.NVarChar, 100).Value = ""
                    MyCommon.LRTsp.Parameters.Add("@Restricted", SqlDbType.Bit).Value = False
                    MyCommon.LRTsp.Parameters.Add("@NodeId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    If (NodeName = "") Then
                        infoMessage = Copient.PhraseLib.Lookup("hierarchy.noname", LanguageID)
                    Else
                        MyCommon.LRTsp.ExecuteNonQuery()
                    End If
                    NodeId = MyCommon.LRTsp.Parameters("@NodeId").Value
                    MyCommon.Close_LRTsp()
                End If
                NodeName = ""
                SelectedNodeId = NodeId.ToString
                MyCommon.QueryStr = sQuery
                dtChildren = MyCommon.LRT_Select()
            End If
        End If

        If SelectedNodeId = "" AndAlso dtChildren.Rows.Count > 0 Then
            SelectedNodeId = dtChildren.Rows(0).Item("id")
        End If

        If bAdd Then
            ProductList = Request.QueryString("level-avail")
            If ProductList <> "" Then
                Products = ProductList.Split(",")
                For i = 0 To Products.Length - 1
                    MyCommon.QueryStr = "dbo.pt_PHContainer_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@NodeId", SqlDbType.BigInt, 8).Value = SelectedNodeId
                    MyCommon.LRTsp.Parameters.Add("@ProductId", SqlDbType.BigInt, 8).Value = Products(i)
                    MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                Next
            End If
        End If

        If bRemove Then
            ProductList = Request.QueryString("level-select")
            If ProductList <> "" Then
                sQuery = "delete from PHContainer with (RowLock) where NodeId = " & SelectedNodeId
                sQuery += " and ProductId in (" & ProductList & ")"
                MyCommon.QueryStr = sQuery
                MyCommon.LRT_Execute()
            End If
        End If

        If Not dtParents Is Nothing AndAlso dtParents.Rows.Count > 0 AndAlso dtChildren.Rows.Count > 0 Then
            sQuery = "select a.ProductId as Id,a.ExtProductId as Code,a.Description as Name, b.PKID from Products a with (NoLock), PHContainer b"
            sQuery += " with (NoLock) where a.ProductId = b.ProductId and b.NodeId = " & SelectedNodeId
            MyCommon.QueryStr = sQuery
            dtSelected = MyCommon.LRT_Select()
            SelectedCount = dtSelected.Rows.Count
        Else
            SelectedCount = 0
        End If

        If Not dtParents Is Nothing AndAlso dtParents.Rows.Count > 0 AndAlso dtChildren.Rows.Count > 0 Then
            sQuery = "select ProductId as Id,ExtProductId as Code,Description as Name from Products"
            sQuery += " with (NoLock) where ProductId not in ("
            sQuery += " select a.ProductId from Products a with (NoLock), PHContainer b"
            sQuery += " with (NoLock) where a.ProductId = b.ProductId and b.NodeId = " & SelectedNodeId & ")"
            MyCommon.QueryStr = sQuery
            dtAvailable = MyCommon.LRT_Select()
            AvailableCount = dtAvailable.Rows.Count
        Else
            AvailableCount = 0
        End If

        If (Request.QueryString("ItemPK") <> "") Then
            Integer.TryParse(Request.QueryString("ItemPK"), ItemPKID)
        End If
%>

<script language="JavaScript" type="text/javascript">
  function SubmitForm() {
    document.mainform.submit();
  }
  
  function VerifyNodeName() {
    if(document.mainform.nodename.value == "")
    {
	    alert('<% Sendb(Copient.PhraseLib.Lookup("hierarchy.mustname", LanguageID)) %>');
	    return false;
	  }
	  else
	  {
      //get a reference to the SelectNodeId object
      var select = document.mainform.SelectedNodeId;
      for (var i = 0; i < select.options.length; i ++)
      {
        if (select.options[i].text == document.mainform.nodename.value)
        { 
	        alert('<% Sendb(Copient.PhraseLib.Lookup("hierarchy.nameused", LanguageID)) %>');
          return false;
        }
      }
	    return true;
	  }
	}
	
	function launchSearch() {
	    openPopup('phierarchy-search.aspx');
	}
</script>

<form id="mainform" name="mainform" action="#">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.producthierarchies", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        Send("<input type=""hidden"" id=""ParentNodeIdList"" name=""ParentNodeIdList"" value=""" & ParentNodeIdList & """ />")
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(23, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <% Sendb(Copient.PhraseLib.Lookup("phierarchy.main", LanguageID))%>
      <br />
      <br />
      <%
        i = 0
        If Not dtParents Is Nothing Then
          For Each dr In dtParents.Rows
            If i > 0 Then
              Send("<br /><br class=""half"" />")
              If i > 1 Then
                Dim j As Integer = (i - 1) * 20
                Send("<img src=""../images/l.png"" alt="""" style=""margin-left: " & j.ToString & "px;"" />")
              Else
                Send("<img src=""../images/l.png"" alt="""" />")
              End If
            End If
            Sendb("<select id=""level" & i.ToString & """ name=""level" & i.ToString & """ disabled=""disabled"">")
            Send("    <option value=""1"" selected=""selected"">" & dr.Item("Name") & "</option>")
            Send("</select>")
            i += 1
          Next
        End If
        If i > 0 Then
          Send("<br /><br class=""half"" />")
          If i > 1 Then
            Dim j As Integer = (i - 1) * 20
            Send("<img src=""../images/l.png"" alt="""" style=""margin-left: " & j.ToString & "px;"" />")
          Else
            Send("<img src=""../images/l.png"" alt="""" />")
          End If
        End If
        If dtChildren.Rows.Count > 0 Then
          Send("<select id=""SelectedNodeId"" name=""SelectedNodeId"" onchange=""SubmitForm();"">")
          For Each dr In dtChildren.Rows
            If dr.Item("Id") = SelectedNodeId Then
              Send("    <option value=""" & dr.Item("Id") & """ selected=""selected"">" & dr.Item("Name") & "</option>")
            Else
              Send("    <option value=""" & dr.Item("Id") & """>" & dr.Item("Name") & "</option>")
            End If
          Next
          Send("</select>")
        Else
          If dtParents Is Nothing Then
            Send(Copient.PhraseLib.Lookup("hierarchy.none", LanguageID))
          Else
            Send(Copient.PhraseLib.Lookup("hierarchy.nonodes", LanguageID) & " " & dtParents.Rows(dtParents.Rows.Count - 1).Item("Name"))
          End If
        End If
        
        Sendb("<input type=""submit"" class=""up"" id=""up"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.up", LanguageID) & """ ")
        If dtParents Is Nothing Then
          Send(" value=""&#9650;"" disabled=""disabled"" />")
        Else
          Send(" value=""&#9650;"" />")
        End If
        
        Sendb("<input type=""submit"" class=""down"" id=""down"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.down", LanguageID) & """ ")
        If dtChildren.Rows.Count > 0 Then
          Send(" value=""&#9660;"" />")
        Else
          Send(" value=""&#9660;"" disabled=""disabled"" />")
        End If
        
        If (Logix.UserRoles.EditProductGroups = True) Then
          Send("<br /><br class=""half"" />")
          Sendb("<input type=""submit"" class=""regular"" id=""newnode"" name=""newnode"" style=""margin-left:60px;"" maxlength=""100"" value=""" & Copient.PhraseLib.Lookup("hierarchy.createnode", LanguageID) & """ onclick=""return VerifyNodeName();"" />")
          Sendb("&nbsp;<label for=""nodename"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label>")
          Sendb("<input type=""text"" id=""nodename"" name=""nodename"" class=""long"" maxlength=""100"" value=""" & NodeName & """ />")
          Send("<br /><br class=""half"" />")
          Sendb("<input type=""submit"" class=""regular"" id=""deletenode"" name=""deletenode"" style=""margin-left: 60px;"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("hierarchy.confirmdelete", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("hierarchy.deletenode", LanguageID) & """ ")
          If dtChildren.Rows.Count > 0 Then
            Sendb(" /><br />")
          Else
            Sendb(" disabled=""disabled"" /><br />")
          End If
        End If
        Send("<br class=""half"" />")
        Send("<input type=""button"" class=""regular"" id=""search"" name=""search"" style=""margin-left: 60px;"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""launchSearch();"" />")
      %>
      <hr class="hidden" />
      <br />
      <div class="box" id="products">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.products", LanguageID))%>
          </span>
        </h2>
        <div style="float: left;">
          <% Sendb(Copient.PhraseLib.Lookup("phierarchy.products", LanguageID))%>
          <br />
        </div>
        <br clear="all" />
        <div style="float: left;">
          <label for="level-select"><b><% Sendb(Copient.PhraseLib.Lookup("phierarchy.selected", LanguageID))%></b></label>
          <br />
          <select class="wideselector" multiple="multiple" id="level-select" name="level-select">
            <%
              If SelectedCount > 0 Then
                i = 0
                For Each dr In dtSelected.Rows
                  SelectedOption = IIf(MyCommon.NZ(dr.Item("PKID"), 0) = ItemPKID, "selected=""selected""", "")
                  If dr.Item("Name") = "" Then
                    Send("<option value=""" & dr.Item("Id") & """ " & SelectedOption & ">" & dr.Item("Code") & "</option>")
                  Else
                    Send("<option value=""" & dr.Item("Id") & """ " & SelectedOption & ">" & dr.Item("Code") & " - " & dr.Item("Name") & "</option>")
                  End If
                Next
              End If
            %>
          </select>
        </div>
        <div style="float: left; padding: 90px 2px 1px 2px;">
          <%
            Sendb("<input type=""submit"" class=""arrowadd"" id=""selectstore"" name=""selectstore"" value=""&#171;"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
            If (Logix.UserRoles.EditProductGroups = True) And (AvailableCount > 0) Then
              Sendb(" />")
            Else
              Sendb(" disabled=""disabled"" />")
            End If
          %>
          <br />
          <br class="half" />
          <%
            Sendb("<input type=""submit"" class=""arrowrem"" id=""removestore"" name=""removestore"" value=""&#187;"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """")
            If (Logix.UserRoles.EditProductGroups = True) And (SelectedCount > 0) Then
              Sendb(" />")
            Else
              Sendb(" disabled=""disabled"" />")
            End If
          %>
        </div>
        <div style="float: left;">
          <label for="level-avail"><b><% Sendb(Copient.PhraseLib.Lookup("phierarchy.available", LanguageID))%></b></label>
          <br />
          <select class="wideselector" multiple="multiple" id="level-avail" name="level-avail">
            <%
              If AvailableCount > 0 Then
                i = 0
                For Each dr In dtAvailable.Rows
                  If MyCommon.NZ(dr.Item("Name"), "") = "" Then
                    Send("<option value=""" & dr.Item("Id") & """>" & dr.Item("Code") & "</option>")
                  Else
                    Send("<option value=""" & dr.Item("Id") & """>" & dr.Item("Code") & " - " & dr.Item("Name") & "</option>")
                  End If
                Next
              End If
            %>
          </select>
        </div>
        <br clear="left" />
        <br class="zero" />
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(23, 0, AdminUserID)
    End If
  End If
done:
  ' Catch ex As Exception
  ' MyCommon.Error_Processor("Catch", ex.Message, "phierarchy.aspx", "Products")
  ' Throw ex
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "NodeName")
Logix = Nothing
MyCommon = Nothing
%>
