<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: layout-cell.aspx 
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
  Dim dt As System.Data.DataTable = Nothing
  Dim row As System.Data.DataRow = Nothing
  Dim CellID As Integer
  Dim LayoutID As Integer
  Dim CellName As String = ""
  Dim ContentsID As Integer
  Dim X As Integer
  Dim Y As Integer
  Dim Width As Integer = 1
  Dim Height As Integer = 1
  Dim BackgroundImg As Integer
  Dim Title As String = ""
  Dim NumRecs As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "layout-cell.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  LayoutID = MyCommon.Extract_Decimal(Request.QueryString("LayoutID"), MyCommon.GetAdminUser.Culture)
  CellID = MyCommon.Extract_Decimal(Request.QueryString("CellID"), MyCommon.GetAdminUser.Culture)
  
  Send_HeadBegin("term.layouts")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 6)
  Send_Subtabs(Logix, 62, 3, , LayoutID)
  
  If (Logix.UserRoles.AccessLayouts = False) Then
    Send_Denied(1, "perm.layouts-access")
    GoTo done
  End If
  
  
  If Not (LayoutID <> 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "layout-edit.aspx")
  End If
  
  If (Request.QueryString("save") <> "") Then
    CellName = Left(Logix.TrimAll(MyCommon.Strip_Quotes(Request.QueryString("cellName"))), 255)
    X = MyCommon.Extract_Decimal(Request.QueryString("xpos"), MyCommon.GetAdminUser.Culture)
    Y = MyCommon.Extract_Decimal(Request.QueryString("ypos"), MyCommon.GetAdminUser.Culture)
    Width = MyCommon.Extract_Decimal(Request.QueryString("width"), MyCommon.GetAdminUser.Culture)
    Height = MyCommon.Extract_Decimal(Request.QueryString("height"), MyCommon.GetAdminUser.Culture)
    ContentsID = MyCommon.Extract_Decimal(Request.QueryString("contentsid"), MyCommon.GetAdminUser.Culture)
    BackgroundImg = MyCommon.Extract_Decimal(Request.QueryString("backgroundImg"), MyCommon.GetAdminUser.Culture)
    
    If LayoutID = 0 Then infoMessage = Copient.PhraseLib.Lookup("layout.invalid", LanguageID)
    If infoMessage = "" And X < 0 Then infoMessage = Copient.PhraseLib.Lookup("layout.invalid-xpos", LanguageID)
    If infoMessage = "" And Y < 0 Then infoMessage = Copient.PhraseLib.Lookup("layout.invalid-ypos", LanguageID)
    If infoMessage = "" And Width < 1 Then infoMessage = Copient.PhraseLib.Lookup("layout.width-minimum", LanguageID)
    If infoMessage = "" And Height < 1 Then infoMessage = Copient.PhraseLib.Lookup("layout.height-minimum", LanguageID)
    If infoMessage = "" And CellName = "" Then infoMessage = Copient.PhraseLib.Lookup("layout.nocellname", LanguageID)
    If infoMessage = "" Then
      NumRecs = 0
      'make sure the cell name isn't already in use in this layout
      MyCommon.QueryStr = "select count(*) as NumRecs from ScreenCells with (NoLock) where Name='" & MyCommon.Parse_Quotes(CellName) & "' and LayoutID=" & LayoutID & " and CellID<>" & CellID & " and deleted=0;"
      dt = MyCommon.LRT_Select()
      If (dt.Rows.Count > 0) Then
        NumRecs = MyCommon.NZ(dt.Rows(0).Item("NumRecs"), 0)
      End If
      If NumRecs > 0 Then infoMessage = Copient.PhraseLib.Lookup("layout.duplicatecell", LanguageID)
    End If
    If infoMessage = "" Then
      If CellID = 0 Then 'this is a new cell so we must do an insert
        MyCommon.Activity_Log(13, LayoutID, AdminUserID, Copient.PhraseLib.Lookup("history.layout-cell-create", LanguageID))
        MyCommon.QueryStr = "insert into ScreenCells with (RowLock) (LayoutID, ContentsID, Name, X, Y, Width, Height, BackgroundImg, Deleted, LastUpdate) values " & _
                            "(" & LayoutID & ", " & ContentsID & ", N'" & MyCommon.Parse_Quotes(CellName) & "', " & X & ", " & Y & ", " & Width & ", " & Height & ", 0, 0, getdate());"
      Else  'this is an existing cell so do an update
        MyCommon.Activity_Log(13, LayoutID, AdminUserID, Copient.PhraseLib.Lookup("history.layout-cell-edit", LanguageID))
        MyCommon.QueryStr = "update ScreenCells with (RowLock) set ContentsID=" & ContentsID & ", Name=N'" & MyCommon.Parse_Quotes(CellName) & "', X=" & X & ", Y=" & Y & ", Width=" & Width & ", Height=" & Height & ", BackgroundImg=" & BackgroundImg & ", LastUpdate=getdate() where CellID=" & CellID & " and Deleted=0;"
      End If
      MyCommon.LRT_Execute()
      
      ' send the changed graphic back down to the stores.
      MyCommon.QueryStr = "update OnScreenAds with (RowLock) set CPEStatusFlag=2, UEStatusFlag=2, LastUpdate=getdate() where OnScreenAdID=" & BackgroundImg
      MyCommon.LRT_Execute()
      
      ' return user back to the main layout edit screen to view change or addition
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "layout-edit.aspx?LayoutID=" & LayoutID)
      GoTo done
    End If
  ElseIf (Request.QueryString("close") = "Close") Then
    ' return user back to the main layout edit screen to view change or addition
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "layout-edit.aspx?LayoutID=" & LayoutID)
    GoTo done
  End If
  
  If Not (CellID = 0) And infoMessage = "" Then
    MyCommon.QueryStr = "select Name, ContentsID, X, Y, Width, Height, BackgroundImg from ScreenCells with (NoLock) where CellID=" & CellID & " and LayoutID=" & LayoutID & " and Deleted=0;"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      CellName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
      ContentsID = MyCommon.NZ(dt.Rows(0).Item("ContentsID"), 0)
      X = MyCommon.NZ(dt.Rows(0).Item("X"), 0)
      Y = MyCommon.NZ(dt.Rows(0).Item("Y"), 0)
      Width = MyCommon.NZ(dt.Rows(0).Item("Width"), 1)
      Height = MyCommon.NZ(dt.Rows(0).Item("Height"), 1)
      BackgroundImg = MyCommon.NZ(dt.Rows(0).Item("BackgroundImg"), 0)
    End If
  End If
  
  Title = IIf((CellID = 0), Copient.PhraseLib.Lookup("term.layout", LanguageID) & " #" & LayoutID & ": " & Copient.PhraseLib.Lookup("term.newcell", LanguageID), Copient.PhraseLib.Lookup("term.layout", LanguageID) & " #" & LayoutID & " " & StrConv(Copient.PhraseLib.Lookup("term.cell", LanguageID), VbStrConv.Lowercase) & ": " & MyCommon.TruncateString(CellName, 30))
%>
<form action="#" method="get" id="mainform" name="mainform">
  <input type="hidden" id="LayoutID" name="LayoutID" value="<% Sendb(LayoutID) %>" />
  <input type="hidden" id="CellID" name="CellID" value="<% Sendb(CellID) %>" />
  <div id="intro">
    <h1 id="title">
      <% Sendb(Title)%>
    </h1>
    <div id="controls">
      <%
        Send_Close()
        If (Logix.UserRoles.EditLayouts) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (LayoutID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(12, LayoutID, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="editCellBox">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.celledit", LanguageID))%>
          </span>
        </h2>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.cell", LanguageID))%>">
          <tr>
            <td align="right">
              <label for="cellName"><% Sendb(Copient.PhraseLib.Lookup("term.cellname", LanguageID))%>:</label>
            </td>
            <td>
              <% If (CellName Is Nothing) Then CellName = ""%>
              <input type="text" class="medium" id="cellName" name="cellName" maxlength="100" value="<% Sendb(CellName.Replace("""", "&quot;")) %>" />
            </td>
          </tr>
          <tr>
            <td align="right">
              <label for="contentsid"><% Sendb(Copient.PhraseLib.Lookup("term.contents", LanguageID))%>:</label>
            </td>
            <td>
              <select name="contentsid" id="contentsid">
                <% 
                  MyCommon.QueryStr = "select ContentsID, Name, DefaultValue, PhraseID from ScreenCellContents with (NoLock) order by ContentsID;"
                  dt = MyCommon.LRT_Select
                  If (dt.Rows.Count > 0) Then
                    For Each row In dt.Rows
                      Sendb("<option value=""" & MyCommon.NZ(row.Item("ContentsID"), 0) & """")
                      If ContentsID = 0 And MyCommon.NZ(row.Item("DefaultValue"), False) = True Then Sendb(" selected=""selected""")
                      If ContentsID = MyCommon.NZ(row.Item("ContentsID"), 999) Then Sendb(" selected=""selected""")
                      Sendb(">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("Name"), "")))
                    Next
                  End If
                %>
              </select>
            </td>
          </tr>
          <tr>
            <td align="right">
              <label for="xpos"><% Sendb(Copient.PhraseLib.Lookup("term.xpos", LanguageID))%>:</label>
            </td>
            <td>
              <input type="text" class="short" id="xpos" name="xpos" maxlength="5" value="<% Sendb(X) %>" />
            </td>
          </tr>
          <tr>
            <td align="right">
              <label for="ypos"><% Sendb(Copient.PhraseLib.Lookup("term.ypos", LanguageID))%>:</label>
            </td>
            <td>
              <input type="text" class="short" id="ypos" name="ypos" maxlength="5" value="<% Sendb(Y) %>" />
            </td>
          </tr>
          <tr>
            <td align="right">
              <label for="width"><% Sendb(Copient.PhraseLib.Lookup("term.width", LanguageID))%>:</label>
            </td>
            <td>
              <input type="text" class="short" id="width" name="width" maxlength="5" value="<% Sendb(Width) %>" />
            </td>
          </tr>
          <tr>
            <td align="right">
              <label for="height"><% Sendb(Copient.PhraseLib.Lookup("term.height", LanguageID))%>:</label>
            </td>
            <td>
              <input type="text" class="short" id="height" name="height" maxlength="5" value="<% Sendb(Height) %>" />
            </td>
          </tr>
          <tr>
            <td align="right">
              <label for="backgroundImg"><% Sendb(Copient.PhraseLib.Lookup("term.background", LanguageID))%>:</label>
            </td>
            <td>
              <select class="long" name="backgroundImg" id="backgroundImg">
                <% 
                  MyCommon.QueryStr = "select Name, OnScreenAdID from OnScreenAds with (NoLock) where Width=" & Width & " and height=" & Height & " and deleted=0 order by Name;"
                  dt = MyCommon.LRT_Select
                  If (dt.Rows.Count > 0) Then
                    Sendb("<option value=""0""")
                    If BackgroundImg = 0 Then Send(" selected=""selected""")
                    Send(">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
                    For Each row In dt.Rows
                      Sendb("<option value=""" & row.Item("OnScreenAdID") & """")
                      If row.Item("OnScreenAdID") = BackgroundImg Then Sendb(" selected=""selected""")
                      Send(">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                    Next
                  Else
                    Send("<option value=""0"">" & Copient.PhraseLib.Lookup("term.noneavailable", LanguageID) & "</option>")
                  End If
                %>
              </select>
            </td>
          </tr>
        </table>
      </div>
    </div>
  </div>
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (LayoutID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(12, LayoutID, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd("mainform", "cellName")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
