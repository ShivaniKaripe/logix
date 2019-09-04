<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: lhierarchy-search.aspx 
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
  Dim SearchString As String = ""
  Dim LevelDisplay As String = ""
  Dim Name As String = ""
  Dim ExternalID As String = ""
  Dim LocationID As String = ""
  Dim ID As Integer
  Dim NodeID As Integer
  Dim Level As Integer
  Dim i As Integer = 0
  Dim SelectedOption As String = ""
  Dim levelPos, idPos As Integer
  Dim ParentNodeIdList As String = ""
  Dim SelectedNodeId As String = ""
  Dim qryStr As String = ""
  Dim ParentID As Integer
  Dim Shaded As String = "shaded"
  Dim LocationGroupID As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "lhierarchy-search.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  SearchString = Request.QueryString("searchString")
  SelectedOption = Request.QueryString("selected")
  LocationGroupID = Request.QueryString("LocationGroupID")
  
  If (SelectedOption <> "" AndAlso Request.QueryString("Find") = "") Then
    levelPos = SelectedOption.IndexOf("L")
    idPos = SelectedOption.IndexOf("ID")
    If (levelPos > -1 AndAlso idPos > -1) Then
      Integer.TryParse(SelectedOption.Substring(levelPos + 1, idPos - 1), Level)
      Integer.TryParse(SelectedOption.Substring(idPos + 2), ID)
      qryStr = "LocationGroupID=" & LocationGroupID
      ' Find parents based on the level
      Select Case Level
        Case 0 ' Root Level
          NodeID = ID
          ParentNodeIdList = ""
          SelectedNodeId = NodeID
          qryStr += "&ParentNodeIdList=&SelectedNodeId=" & NodeID & "&NodeName="
        Case 1  ' Node Level
          NodeID = ID
          ParentNodeIdList = GetParentNodeList(NodeID)
          SelectedNodeId = NodeID
          qryStr += "&ParentNodeIdList=" & ParentNodeIdList & "&SelectedNodeId=" & NodeID & "&NodeName="
        Case 2 ' Location Level
          NodeID = GetLocationNodeID(ID)
          ParentNodeIdList = GetParentNodeList(NodeID)
          SelectedNodeId = NodeID
          qryStr += "&ParentNodeIdList=" & ParentNodeIdList & "&SelectedNodeId=" & NodeID & "&NodeName="
          qryStr += "&ItemPK=" & ID
      End Select
    End If
  End If
  
  Send_HeadBegin("term.offer", "term.locationhierarchy")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>

<script type="text/javascript" language="javascript">
<% If (qryStr <> "") Then %>
  if (opener != null && opener.document != null) {
      var pageName = getOpenerPage(opener.document.location.href);
      if (pageName.indexOf("lhierarchy.aspx") > -1) {
          pageName = 'lhierarchy.aspx?<% Sendb(qryStr) %>'
      } else if (pageName.indexOf("lgroup-edit.aspx") > -1) {
            pageName = 'lgroup-edit.aspx?<% Sendb(qryStr) %>'
      } else {
          pageName = 'lhierarchy.aspx?<% Sendb(qryStr) %>'
      }
      opener.document.location = pageName;
      window.close();
  }        
<%  End If %>

function select_Click() {
    if (validateEntry()) {
            document.mainform.submit();
    }
}

function validateEntry() {
    var elemOpt = document.mainform.selected;
    var selID = "";
    var optSelected = false;
    
    if (elemOpt != null) {
        if (isArray(elemOpt)) {
            for (var i=0; i < elemOpt.length; i++) {
                if (elemOpt[i].checked) {
                    optSelected =true;
                    break;
                }
            }
        } else {
            optSelected = elemOpt.checked;
        }
       if (!optSelected) {
        alert('<% Sendb(Copient.PhraseLib.Lookup("lhierarchy.selectfromlist", LanguageID))%>');
       } 
    }
    return optSelected;            
}

function close_Click() {
    window.close();
}

function optSelected_Click(opt) {
}

function isArray(obj) {
    return(typeof(obj.length)=="undefined") ? false:true;
}

function getOpenerPage(strHref) {
    var namer = "";
    if (strHref.lastIndexOf('/') !=-1) {
        var firstpos=strHref.lastIndexOf('/')+1;
        var lastpos=strHref.length;
        namer=strHref.substring(firstpos,lastpos);
    }
    return namer;
}
</script>

<%
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(2, "perm.admin-configuration")
    GoTo done
  End If
%>
<form action="lhierarchy-search.aspx" id="mainform" name="mainform">
  <input type="hidden" id="LocationGroupID" name="LocationGroupID" value="<% Sendb(LocationGroupID)%>" />
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.locationhierarchies", LanguageID))%>
    </h1>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>
          </span>
        </h2>
        <div style="float: left;">
          <input type="text" class="medium" id="searchString" name="searchString" maxlength="100" value="<% Sendb(SearchString) %>" />
          <input type="submit" class="regular" id="find" name="find" value="<% Sendb(Copient.PhraseLib.Lookup("term.find", LanguageID))%>" />
        </div>
        <div style="float: right; text-align: right;">
          <input type="button" id="select" name="select" value="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" onclick="select_Click();" />
          <input type="button" id="close" name="close" value="<% Sendb(Copient.PhraseLib.Lookup("term.close", LanguageID)) %>" onclick="close_Click();" />
        </div>
        <br clear="all" />
        <br />
        <div style="height: 420px; overflow: auto; border: solid 1px #d0d0d0;">
          <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.results", LanguageID))%>" style="width: 100%;">
            <thead>
              <tr>
                <th class="th-select" scope="col">
                  <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>
                </th>
                <th class="th-level" scope="col">
                  <% Sendb(Copient.PhraseLib.Lookup("term.level", LanguageID))%>
                </th>
                <th class="th-name" scope="col">
                  <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
                </th>
                <th class="th-longid" scope="col">
                  <% Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID))%>
                </th>
                <th class="th-longid" scope="col">
                  <% Sendb(Copient.PhraseLib.Lookup("term.storeid", LanguageID))%>
                </th>
              </tr>
            </thead>
            <tbody>
              <%
                If (SearchString <> "") Then
                  MyCommon.QueryStr = "dbo.pa_Location_Hierarchy_Search"
                  MyCommon.Open_LRTsp()
                  MyCommon.LRTsp.Parameters.Add("@SearchString", SqlDbType.NVarChar, 100).Value = SearchString
                  rst = MyCommon.LRTsp_select
                  If (rst.Rows.Count > 0) Then
                    For Each row In rst.Rows
                      Level = MyCommon.NZ(row.Item("Level"), "")
                      Select Case Level
                        Case 0
                          LevelDisplay = Copient.PhraseLib.Lookup("term.root", LanguageID)
                        Case 1
                          LevelDisplay = Copient.PhraseLib.Lookup("term.node", LanguageID)
                        Case 2
                          LevelDisplay = Copient.PhraseLib.Lookup("term.item", LanguageID)
                        Case Else
                          LevelDisplay = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
                      End Select
                      NodeID = MyCommon.NZ(row.Item("ID"), 0)
                      Name = MyCommon.NZ(row.Item("Name"), "")
                      ExternalID = MyCommon.NZ(row.Item("ExternalID"), "&nbsp;")
                      LocationID = MyCommon.NZ(row.Item("LocationID"), "&nbsp;")
                      Send("<tr class=""" & Shaded & """>")
                      Send("  <td><input type=""radio"" id=""selectedL" & Level & "ID" & NodeID & """ name=""selected"" value=""L" & Level & "ID" & NodeID & """ onclick=""optSelected_Click(this);"" /></td>")
                      Send("  <td>" & LevelDisplay & "</td>")
                      Send("  <td>" & HighlightMatches(Name, SearchString) & "</td>")
                      Send("  <td>" & HighlightMatches(ExternalID, SearchString) & "</td>")
                      Send("  <td>" & HighlightMatches(LocationID, SearchString) & "</td>")
                      Send("</tr>")
                      Shaded = IIf(Shaded = "shaded", "", "shaded")
                    Next
                  Else
                    Send("<tr>")
                    Send("  <td colspan=""5""><center><i>" & Copient.PhraseLib.Lookup("phierarchy.noresultsfound", LanguageID) & " '" & SearchString & "'</i></center></td>")
                    Send("</tr>")
                  End If
                End If
              %>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
</form>

<script type="text/javascript" language="javascript">
    var elem = document.getElementById("searchString");
    if (elem != null) {
        elem.focus();
        elem.select();
    }
</script>

<script runat="server">
  Function HighlightMatches(ByVal ColValue As String, ByVal Search As String) As String
    Dim FormattedCol As String = ColValue
    
    If (Not ColValue Is Nothing) Then
      FormattedCol = ColValue.Replace(Search, "<span class=""red"">" & Search & "</span>")
    End If
    
    Return FormattedCol
  End Function
    
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
  
  Function GetLocationNodeID(ByVal PKID As Integer) As Integer
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

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
