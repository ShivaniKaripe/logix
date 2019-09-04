<%-- version:7.3.1.138972 --%>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: WEBsettings.aspx 
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
  Dim rst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim rst2 As System.Data.DataTable
  Dim row2 As System.Data.DataRow
  Dim OptionID As Integer
  Dim tempstr As String
  Dim infoMessage As String = ""
  Dim OpenTagEscape As String = "<>"
  Dim OptionValue As String = ""
  Dim Handheld As Boolean = False
  Dim opt7, opt6 As Boolean
  Dim DisableEdit As String = " disabled=""disabled"""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "WEBsettings.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  If Request.QueryString("oid7") <> "" Then
    If IsNumeric(Request.QueryString("oid7")) Then
      opt7 = True
    Else
      opt7 = False
    End If
  End If
  
  If Request.QueryString("oid6") <> "" Then
    If IsNumeric(Request.QueryString("oid6")) Then
      opt6 = True
    Else
      opt6 = False
    End If
  End If
  
  If (Request.QueryString("save") <> "") Then
    If opt7 AndAlso opt6 Then
      ' someone clicked save lets get to it
      MyCommon.QueryStr = "select OptionID from Web_Options with (NoLock) where Visible=1  order by OptionID;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          tempstr = MyCommon.Parse_Quotes(Request.QueryString("oid" & row.Item("OptionID")))
          tempstr = Logix.TrimAll(tempstr)
          ' Replace tag escape put into place due to request failure caused by scripting validation exception
          ' encountered with tags being sent over the querystring.
          If (Not tempstr Is Nothing AndAlso tempstr.IndexOf("<>") > -1) Then
            tempstr = tempstr.Replace(OpenTagEscape, "<")
          End If
          MyCommon.QueryStr = "Update Web_Options with (RowLock) set OptionValue=@NewValue, LastUpdate=getdate() where OptionID=@OptionID;"
          MyCommon.DBParameters.Add("@NewValue", SqlDbType.NVarChar, 255).Value = tempstr
          MyCommon.DBParameters.Add("@OptionID", SqlDbType.Int).Value = row.Item("OptionID")
          MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        Next
      End If
      
      ' Refresh cache with new system options
      CMS.AMS.CurrentRequest.Resolver.AppName = MyCommon.AppName
      Dim cacheData As CMS.AMS.Contract.ICacheData = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.ICacheData)()
      cacheData.ClearAllSystemOptionsCache()
      Copient.SystemOptionsCache.RemoveCache(System.Web.HttpContext.Current.Request.Url.Host)
      MyCommon.Activity_Log(33, 0, AdminUserID, Copient.PhraseLib.Lookup("history.settings", LanguageID))
    Else
      infoMessage = Copient.PhraseLib.Lookup("CPEoffer-con-tender.positivevalue", LanguageID)
    End If
  End If
  
  Send_HeadBegin("term.websettings")
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
      <% Sendb(Copient.PhraseLib.Lookup("term.websettings", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemOptions = True OrElse Logix.UserRoles.AccessSystemSettings = True) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(31, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID))%>">
      <%
        If Logix.UserRoles.ViewHiddenOptions = True Then
          MyCommon.QueryStr = "select SO.OptionName, SO.OptionID, SO.OptionValue, SO.PhraseID, SO.Visible, SO.OptionType, SO.OptionTypePhraseID, IsNull(PT.Phrase, SO.OptionName) as Phrase, PT.LanguageID from Web_Options as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID order by OptionName;"
        Else
          MyCommon.QueryStr = "select SO.OptionName, SO.OptionID, SO.OptionValue, SO.PhraseID, SO.Visible, SO.OptionType, SO.OptionTypePhraseID, IsNull(PT.Phrase, SO.OptionName) as Phrase, PT.LanguageID from Web_Options as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID where visible=1 order by OptionName;"
        End If
        MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        Dim Counter As Integer = 0
        If (rst.Rows.Count > 0) Then
          For Each row In rst.Rows
            OptionID = MyCommon.NZ(row.Item("OptionID"), 0)
            If MyCommon.NZ(row.Item("OptionType"), 0) > Counter Then
              Send("<tr>")
              Send("  <td><h2>" & Copient.PhraseLib.Lookup(row.Item("OptionTypePhraseID"), LanguageID) & "</h2></td>")
              Send("</tr>")
            End If
            Send("")
            Send("<tr>")
            Send("  <td" & IIf(row.Item("Visible"), "", " style=""color:red;""") & "><label for=""oid" & OptionID & """>" & MyCommon.NZ(row.Item("Phrase"), "") & ":</label>")
            MyCommon.QueryStr = "select SOV.OptionValue, SOV.Description, SOV.PhraseID, IsNull(PT.Phrase, SOV.Description) as Phrase, PT.LanguageID from Web_OptionsValues as SOV with (NoLock) left join PhraseText as PT with (NoLock) on SOV.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID where OptionID=@OptionID order by OptionValue;"
            MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
            MyCommon.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
            rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (rst2.Rows.Count > 0) Then
              Send("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("Visible"), "", DisableEdit) & ">")
              For Each row2 In rst2.Rows
                OptionValue = MyCommon.NZ(row2.Item("OptionValue"), "")
                OptionValue = OptionValue.Replace("<", OpenTagEscape)
                Sendb("      <option value=""" & OptionValue & """")
                If MyCommon.NZ(row2.Item("OptionValue"), "") = MyCommon.NZ(row.Item("OptionValue"), "") Then Sendb(" selected=""selected""")
                Send(">" & MyCommon.NZ(row2.Item("Phrase"), "") & "</option>")
              Next
              Send("    </select>")
            Else
              Send("    <input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """" & IIf(row.Item("Visible"), "", DisableEdit) & """ />")
            End If
            Send("  </td>")
            Send("</tr>")
            Counter = MyCommon.NZ(row.Item("OptionType"), 0)
          Next
        End If
      %>
    </table>
  </div>
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(31, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
