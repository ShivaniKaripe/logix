<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-cashier-inquiry-options.aspx 
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
  ' * Version : 5.10b1.0 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim rst2 As System.Data.DataTable
  Dim row2 As System.Data.DataRow
  Dim FieldID As Integer
  Dim FieldName As String
  Dim tempstr As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim OptionObj As Copient.SystemOption = Nothing
  Dim HistoryStr As String = ""
  Dim OpenTagEscape As String = "<>"
  Dim sCheckBox1 As String = ""
  Dim sCheckBox2 As String = ""
  Dim iOldValue As Integer
  Dim iNewValue As Integer
  Const sDots As String = "................"
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CM-cashier-inquiry-options.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Request.QueryString("save") <> "") Then
    ' someone clicked save lets get to it
    MyCommon.QueryStr = "select FieldID, FieldName, Display, AllowEdit from CM_Cashier_Inquiry_Options with (NoLock) order by FieldID;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        FieldID = MyCommon.NZ(row.Item("FieldID"), 0)
        tempstr = Request.QueryString("fid1-" & FieldID)
        tempstr = Logix.TrimAll(tempstr)
        
        ' Replace tag escape put into place due to request failure caused by scripting validation exception
        ' encountered with tags being sent over the querystring.
        If (Not tempstr Is Nothing AndAlso tempstr.IndexOf(OpenTagEscape) > -1) Then
          tempstr = tempstr.Replace(OpenTagEscape, "<")
        End If
        
        If tempstr = "on" Then
          iNewValue = -1
        Else
          iNewValue = 0
        End If
        iOldValue = MyCommon.NZ(row.Item("Display"), 0)

        If iNewValue <> iOldValue Then
          MyCommon.QueryStr = "Update CM_Cashier_Inquiry_Options with (RowLock) set Display=" & iNewValue & ", LastUpdate=getdate() where FieldID=" & FieldID
          MyCommon.LRT_Execute()
            
          'If (MyCommon.RowsAffected > 0) Then
          '  HistoryStr = Copient.PhraseLib.Lookup("history.edit-cmsetting", LanguageID) & " '" & MyCommon.NZ(row.Item("OptionName"), "") & "'" & _
          '               " from: " & OptionObj.GetOldValue() & " to: " & OptionObj.GetNewValue()
          '  MyCommon.Activity_Log(24, 0, AdminUserID, HistoryStr)
          'End If
        End If

        tempstr = Request.QueryString("fid2-" & FieldID)
        tempstr = Logix.TrimAll(tempstr)
        
        ' Replace tag escape put into place due to request failure caused by scripting validation exception
        ' encountered with tags being sent over the querystring.
        If (Not tempstr Is Nothing AndAlso tempstr.IndexOf(OpenTagEscape) > -1) Then
          tempstr = tempstr.Replace(OpenTagEscape, "<")
        End If

        If tempstr = "on" Then
          iNewValue = -1
        Else
          iNewValue = 0
        End If
        iOldValue = MyCommon.NZ(row.Item("AllowEdit"), 0)

        If iNewValue <> iOldValue Then
          MyCommon.QueryStr = "Update CM_Cashier_Inquiry_Options with (RowLock) set AllowEdit=" & iNewValue & ", LastUpdate=getdate() where FieldID=" & FieldID
          MyCommon.LRT_Execute()
            
          'If (MyCommon.RowsAffected > 0) Then
          '  HistoryStr = Copient.PhraseLib.Lookup("history.edit-cmsetting", LanguageID) & " '" & MyCommon.NZ(row.Item("OptionName"), "") & "'" & _
          '               " from: " & OptionObj.GetOldValue() & " to: " & OptionObj.GetNewValue()
          '  MyCommon.Activity_Log(24, 0, AdminUserID, HistoryStr)
          'End If
        End If

      Next
    End If
  End If
  
  Send_HeadBegin("term.cmsettings")
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
  
  If (Logix.UserRoles.AccessSystemSettings = False) Then
    Send_Denied(1, "perm.admin-settings")
    GoTo done
  End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.cashiersettings", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemConfiguration = True) Then
          Send_Save()
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <table width="350 px" style="margin-top:5px" summary="<% Sendb(Copient.PhraseLib.Lookup("term.cashiersettings", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" style="width:54%" scope="col">
          <a>
            <% Sendb(Copient.PhraseLib.Lookup("term.fieldname", LanguageID))%>
          </a>
        </th>
        <th align="left" style="width:23%" scope="col">
          <a>
            <% Sendb(Copient.PhraseLib.Lookup("term.displayfield", LanguageID))%>
          </a>
        </th>
        <th align="left" style="width:23%" scope="col">
          <a>
            <% Sendb(Copient.PhraseLib.Lookup("term.allowedit", LanguageID))%>
          </a>
        </th>
      </tr>
    </thead>
      <%
        '
        MyCommon.QueryStr = "select CIO.FieldID, CIO.FieldName, CIO.Display, CIO.AllowEdit, CIO.PhraseID, PT.Phrase from CM_Cashier_Inquiry_Options as CIO with (NoLock) " & _
                            "inner join PhraseText as PT with (NoLock) on PT.PhraseID=CIO.PhraseID where PT.LanguageID=" & LanguageID & " " & _
                            "order by FieldID;"
        'MyCommon.QueryStr = "select CIO.FieldID, CIO.FieldName, CIO.Display, CIO.AllowEdit, CIO.PhraseID from CM_Cashier_Inquiry_Options as CIO with (NoLock) order by FieldID;"
        rst = MyCommon.LRT_Select
        Dim Counter As Integer = 0
        If (rst.Rows.Count > 0) Then
          For Each row In rst.Rows
            FieldID = MyCommon.NZ(row.Item("FieldID"), 0)
            If (row.Item("Display") <> 0) Then sCheckBox1 = " checked=""checked""" Else sCheckBox1 = ""
            If (row.Item("AllowEdit") <> 0) Then sCheckBox2 = " checked=""checked""" Else sCheckBox2 = ""
            FieldName = MyCommon.NZ(row.Item("Phrase"), "Unknown")
            'FieldName = MyCommon.NZ(row.Item("FieldName"), "Unknown")
            Send("")
            Send("<tr>")
            Send("  <td>")
            Send("    <label for=""fid" & FieldID & """ >" & FieldName & "</label>")
            Send("  </td>")
            Send("  <td>")
            Send("    <input class=""checkbox"" id=""fid1-" & FieldID & """ name=""fid1-" & FieldID & """ type=""checkbox""" & sCheckBox1 & " />")
            Send("  </td>")
            Send("  <td>")
            Send("    <input class=""checkbox"" id=""fid2-" & FieldID & """ name=""fid2-" & FieldID & """ type=""checkbox""" & sCheckBox2 & " />")
            Send("  </td>")
            Send("</tr>")
            sCheckBox1 = ""
            sCheckBox2 = ""
          Next
        End If
      %>
    </table>
  </div>
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(29, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
