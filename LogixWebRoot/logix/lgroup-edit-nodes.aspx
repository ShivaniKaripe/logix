<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: lgroup-edit-nodes.aspx 
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
  Dim LocationGroupID As Long
  Dim Name As String = ""
  Dim row As DataRow
  Dim rstOrder As DataTable
  Dim rstNew As DataTable
  Dim historyString As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim ExcludingStr As String = ""
  Dim NodeList As New SortedList(50)
  Dim ParentID As Integer
  Dim aRows As New ArrayList(10)
  Dim MainList As New ArrayList(10)
  Dim Shaded As String = ""
  Dim td1Style As String = ""
  Dim i As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "lgroup-edit-nodes.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If Not Int64.TryParse(Request.QueryString("LocationGroupID"), LocationGroupID) Then
    LocationGroupID = 0
  End If
  Name = HttpUtility.HtmlEncode(Request.QueryString("Name"))
  
  MyCommon.QueryStr = "dbo.pa_LocationGroupNodes_Select"
  MyCommon.Open_LRTsp()
  MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = LocationGroupID
  rstOrder = MyCommon.LRTsp_select()
  
  For Each row In rstOrder.Rows
    If (MyCommon.NZ(row.Item("Excluded"), False)) Then
      ' find the match for excluded with the first parent in the tree
      ParentID = GetParentNodeID(MyCommon, MyCommon.NZ(row.Item("NodeID"), -1))
      While ParentID > 0
        If (NodeList.Contains(ParentID.ToString)) Then
          ' associate this excluded row to its parent row
          aRows = MainList.Item(NodeList.Item(ParentID.ToString))
          aRows.Add(row)
          MainList.Item(NodeList.Item(ParentID.ToString)) = aRows
          ParentID = 0
        Else
          ' keep looking for the parent
          ParentID = GetParentNodeID(MyCommon, ParentID)
        End If
      End While
    Else
      ' store the nodeid
      If Not (NodeList.Contains(MyCommon.NZ(row.Item("NodeID"), "-1").ToString)) Then
        aRows = New ArrayList(10)
        aRows.Add(row)
        NodeList.Add(MyCommon.NZ(row.Item("NodeID"), "-1").ToString, i)
        MainList.Insert(i, aRows)
        i += 1
      End If
    End If
  Next
  
  rstNew = rstOrder.Clone
  For i = 0 To MainList.Count - 1
    aRows = MainList.Item(i)
    For Each row In aRows
      rstNew.ImportRow(row)
    Next
  Next
  
  Send_HeadBegin("term.storegroup", "term.nodes", LocationGroupID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.EditProductGroups = False) Then
    Send_Denied(2, "perm.lgroup-edit")
    GoTo done
  End If
  
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  Send("    opener.location = 'lgroup-edit.aspx?LocationGroupID=" & LocationGroupID & "'; ")
  Send("    } ")
  Send("</script>")
%>
<form action="#" id="mainform" name="mainform">
  <input type="hidden" id="LocationGroupID" name="LocationGroupID" value="<% sendb(LocationGroupID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
  <div id="intro">
    <h1 id="title">
      <%
        If LocationGroupID <> 0 Then
          Sendb(Copient.PhraseLib.Lookup("term.storegroup", LanguageID) & " #" & LocationGroupID & ": " & MyCommon.TruncateString(Name, 35))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.newstoregroup", LanguageID))
        End If
      %>
    </h1>
    <div id="controls">
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <div class="box" id="nodes">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.nodes", LanguageID))%>
          </span>
        </h2>
        <br class="half" />
        <%  
          If (rstNew.Rows.Count > 0) Then
        %>
        <table summary="">
          <thead>
            <tr>
              <th><% Sendb(Copient.PhraseLib.Lookup("term.foldername", LanguageID))%></th>
              <th><% Sendb(Copient.PhraseLib.Lookup("term.HierarchyName", LanguageID))%></th>
              <th><% Sendb(Copient.PhraseLib.Lookup("term.SelectionDate", LanguageID))%></th>
              <th><% Sendb(Copient.PhraseLib.Lookup("term.user", LanguageID))%></th>
            </tr>
          </thead>
          <%                    
            For Each row In rstNew.Rows
              If (MyCommon.NZ(row.Item("Excluded"), False) = True) Then
                ExcludingStr = "<i>" & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & ": </i>"
                td1Style = "padding-left:10px;"
              Else
                ExcludingStr = ""
                td1Style = ""
                Shaded = IIf(Shaded = "shaded", "", "shaded")
              End If
              Send("<tr class=""" & Shaded & """>")
              Send("  <td style=""" & td1Style & """>" & ExcludingStr & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("NodeName"), "&nbsp;"), 30) & "</td>")
              Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("HierarchyName"), "&nbsp;"), 30) & "</td>")
              If IsDBNull(row.Item("LinkDate")) Then
                Send("  <td>&nbsp;</td>")
              Else
                Send("  <td>" & Logix.ToShortDateTimeString(row.Item("LinkDate"), MyCommon) & "</td>")
              End If
              Send("  <td>" & MyCommon.NZ(row.Item("UserName"), "&nbsp;") & "</td>")
              Send("</tr>")
            Next
          %>
        </table>
        <% 
        End If
        MyCommon.Close_LRTsp()
        %>
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>

<script runat="server">
  Function GetParentNodeID(ByRef MyCommon As Copient.CommonInc, ByVal NodeID As Integer) As Integer
    Dim dt As DataTable
    Dim ParentID As Integer = -1
    MyCommon.QueryStr = "Select ParentID from LHNodes with (NoLock) where NodeID = @NodeID"
    MyCommon.DBParameters.Add("NodeID", SqlDbType.BigInt).Value = NodeID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If (dt.Rows.Count > 0) Then
      ParentID = MyCommon.NZ(dt.Rows(0).Item("ParentID"), -1)
    End If
    Return ParentID
  End Function
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  MyCommon = Nothing
  Logix = Nothing
%>
