<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-settings.aspx 
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
  Dim MyCryptLib As New Copient.CryptLib
  Dim rst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim rst2 As System.Data.DataTable
  Dim row2 As System.Data.DataRow
  Dim OptionID As Integer
  Dim tempstr As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim OptionObj As Copient.SystemOption = Nothing
  Dim HistoryStr As String = ""
  Dim OpenTagEscape As String = "<>"
  Dim DisableEdit As String = " disabled=""disabled"""
  Dim user As String
  Dim euser As String
  Dim pwd As String
  Dim epwd As String
  Dim upos As Integer
  Dim ppos As Integer
  Dim pend As Integer
  Dim sDB2 As String 
  Dim tmp As String
  Dim buser As Boolean
  Dim bpwd As Boolean

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CM-settings.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Request.QueryString("save") <> "") Then
    ' someone clicked save lets get to it
    MyCommon.QueryStr = "select OptionID, OptionName, OptionValue from CM_SystemOptions with (NoLock) where Visible=1 and OptionID <> 5 order by OptionID;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        tempstr = Request.QueryString("oid" & MyCommon.NZ(row.Item("OptionID"), 0))
        tempstr = Logix.TrimAll(tempstr)
        
        ' Replace tag escape put into place due to request failure caused by scripting validation exception
        ' encountered with tags being sent over the querystring.
        If (Not tempstr Is Nothing AndAlso tempstr.IndexOf(OpenTagEscape) > -1) Then
          tempstr = tempstr.Replace(OpenTagEscape, "<")
        End If

        OptionObj = New Copient.SystemOption(MyCommon.NZ(row.Item("OptionID"), 0), MyCommon.NZ(row.Item("OptionValue"), ""))
        OptionObj.SetNewValue(tempstr)

        If OptionObj.IsModified Then
          MyCommon.QueryStr = "Update CM_SystemOptions with (RowLock) set OptionValue=@NewValue, LastUpdate=getdate() where OptionID=@OptionID;"
          MyCommon.DBParameters.Add("@NewValue", SqlDbType.NVarChar, 255).Value = OptionObj.GetNewValue()
          MyCommon.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID()
          MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            
          If (MyCommon.RowsAffected > 0) Then
            HistoryStr = Copient.PhraseLib.Lookup("history.edit-cmsetting", LanguageID) & " '" & MyCommon.NZ(row.Item("OptionName"), "") & "'" & _
                         " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & ": " & OptionObj.GetOldValue() & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & ": " & OptionObj.GetNewValue()
            MyCommon.Activity_Log(31, 0, AdminUserID, HistoryStr)
          End If
        End If
        
      Next
      
      If (Logix.UserRoles.EditEncryptDB2 = True) Then
        tempstr = Request.QueryString("DB2")
        tempstr = Logix.TrimAll(tempstr)
        sDB2 = MyCommon.Fetch_CM_SystemOption(5)
        buser = tempstr.Contains("UID=****;")
        bpwd  = tempstr.Contains("PWD=****;")
        
        Try
          If (tempstr = "") Then
            OptionObj = New Copient.SystemOption(5, sDB2)
            OptionObj.SetNewValue(tempstr)
          Else If ((tempstr.Contains("UID=")) AndAlso (tempstr.Contains(";PWD=")) AndAlso (tempstr.Contains(";host"))) Then
            
              If(buser AndAlso bpwd) Then
                tmp = sDB2
              Else 
               tmp = tempstr
                If (buser) Then
                  upos = InStr(sDB2, "UID=")
                  ppos = InStr(sDB2, ";PWD=")
                  euser = sDB2.Substring(upos+3, ppos-upos-4)
                Else If (bpwd) Then
                  ppos = InStr(sDB2, ";PWD=")
                  pend = InStr(sDB2, ";host")
                  epwd = sDB2.Substring(ppos+4, pend-ppos-5)
                End If
              End If
              
              upos = InStr(tmp, "UID=")
              ppos = InStr(tmp, ";PWD=")
              pend = InStr(tmp, ";host")
              
              If ((ppos-upos-4) > 0) Then 
                user = tmp.Substring(upos+3, ppos-upos-4)
              Else
                Throw New ApplicationException(Copient.PhraseLib.Lookup("term.db2-nouser", LanguageID))
              End If
              If ((pend-ppos-5) > 0) Then
                pwd = tmp.Substring(ppos+4, pend-ppos-5)
              Else
                Throw New ApplicationException(Copient.PhraseLib.Lookup("term.db2-nopassword", LanguageID))
              End If         
              
              If(buser) Then
                euser = IIF(bpwd, user, euser)
                tempstr = tempstr.Replace("UID=****", "UID=" & euser)
              Else
                euser = MyCryptLib.SQL_StringEncrypt(user)
                tempstr = tempstr.Replace("UID=" & user, "UID=" & euser)
              End If
              If(bpwd) Then
                epwd = IIF(buser, pwd, epwd)
                tempstr = tempstr.Replace("PWD=****", "PWD=" & epwd)
              Else
                epwd = MyCryptLib.SQL_StringEncrypt(pwd)
                tempstr = tempstr.Replace("PWD=" & pwd, "PWD=" & epwd)
              End If
              
            
            OptionObj = New Copient.SystemOption(5, sDB2)
            OptionObj.SetNewValue(tempstr)
          
          Else
            tempstr = Request.QueryString("DB2")
            infomessage = Copient.PhraseLib.Lookup("term.db2-invalidformat", LanguageID)
          End If
          

        If OptionObj.IsModified Then
          MyCommon.QueryStr = "Update CM_SystemOptions with (RowLock) set OptionValue=@NewValue, LastUpdate=getdate() where OptionID=@OptionID;"
          MyCommon.DBParameters.Add("@NewValue", SqlDbType.NVarChar, 255).Value = OptionObj.GetNewValue()
          MyCommon.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID()
          MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            
          If (MyCommon.RowsAffected > 0) Then
            HistoryStr = Copient.PhraseLib.Lookup("history.edit-cmsetting", LanguageID) & " '" & Copient.PhraseLib.Lookup("perm.admin-editDB2", LanguageID) & "'" & _
                         " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & ": " & OptionObj.GetOldValue() & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & ": " & OptionObj.GetNewValue()
            MyCommon.Activity_Log(31, 0, AdminUserID, HistoryStr)
          End If
        End If
          
        Catch exApp As ApplicationException
          infomessage = exApp.Message
        Finally
          If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
            MyCommon.Close_LogixRT()
          End If
        End Try
      End If
      
      ' Refresh cache with new system options
      CMS.AMS.CurrentRequest.Resolver.AppName = MyCommon.AppName
      Dim cacheData As CMS.AMS.Contract.ICacheData = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.ICacheData)()
            cacheData.ClearAllSystemOptionsCache()
            Copient.SystemOptionsCache.RemoveCache(System.Web.HttpContext.Current.Request.Url.Host)
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
      <% Sendb(Copient.PhraseLib.Lookup("term.cmsettings", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemOptions = True or Logix.UserRoles.AccessSystemSettings) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(29, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>">
      <%
        If Logix.UserRoles.ViewHiddenOptions = True Then
          MyCommon.QueryStr = "select SO.OptionID, SO.OptionName, SO.OptionValue, SO.Visible, SO.PhraseID, SO.OptionType, SO.OptionTypePhraseID, PT.Phrase from CM_SystemOptions as SO with (NoLock) inner join PhraseText as PT with (NoLock) on PT.PhraseID=SO.PhraseID where PT.LanguageID=@LanguageID and OptionID <> 5  order by OptionName;"
        Else
          MyCommon.QueryStr = "select SO.OptionID, SO.OptionName, SO.OptionValue, SO.Visible, SO.PhraseID, SO.OptionType, SO.OptionTypePhraseID, PT.Phrase from CM_SystemOptions as SO with (NoLock) inner join PhraseText as PT with (NoLock) on PT.PhraseID=SO.PhraseID where Visible=1 and PT.LanguageID=@LanguageID and OptionID <> 5 order by OptionName;"
        End If
        MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        Dim Counter As Integer = 0
        If (rst.Rows.Count > 0) Then
          For Each row In rst.Rows
            OptionID = MyCommon.NZ(row.Item("OptionID"), 0)
            If MyCommon.NZ(row.Item("OptionType"), 0) > Counter Then
              Send("      <tr>")
              Send("        <td>")
              Send("          <h2>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("OptionTypePhraseID"), 0), LanguageID) & "</h2>")
              Send("        </td>")
              Send("      </tr>")
            End If
            Send("")
            Send("<tr>")
            Send("  <td " & IIf(row.Item("Visible"), "", " style=""color:red;""") & ">")
            Send("<label for=""oid" & OptionID & """>" & MyCommon.NZ(row.Item("Phrase"), "") & ":</label>")
            MyCommon.QueryStr = "select SOV.OptionValue, SOV.Description, SOV.PhraseID, PT.Phrase from CM_SystemOptionValues as SOV with (NoLock) inner join PhraseText as PT with (NoLock) on PT.PhraseID=SOV.PhraseID where PT.LanguageID=@LanguageID and OptionID=@OptionID order by OptionValue;"
            MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
            MyCommon.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
            rst2 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (rst2.Rows.Count > 0) Then
              Send("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("Visible"), "", DisableEdit) & ">")
              For Each row2 In rst2.Rows
                Sendb("      <option value=""" & MyCommon.NZ(row2.Item("OptionValue"), "") & """")
                If MyCommon.NZ(row2.Item("OptionValue"), "") = MyCommon.NZ(row.Item("OptionValue"), "") Then Sendb(" selected=""selected""")
                Send(">" & MyCommon.NZ(row2.Item("Phrase"), "") & "</option>")
              Next
              Send("    </select>")
            Else
              Send("    <input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """" & IIf(row.Item("Visible"), "", DisableEdit) & " />")
            End If
            Send("  </td>")
            Send("</tr>")
            Counter = MyCommon.NZ(row.Item("OptionType"), 0)
          Next
          If (Logix.UserRoles.EditEncryptDB2 = True) Then
            Send("")
            Send("<tr>")
            Send("  <td  style=""color:red;"" >")
            Send("<label for=""DB2"">" & Copient.PhraseLib.Lookup("perm.admin-editDB2", LanguageID) & ":</label>")
            tempstr = MyCommon.Fetch_CM_SystemOption(5)
            If((tempstr.Contains("UID=")) AndAlso (tempstr.Contains(";PWD=")) AndAlso (tempstr.Contains(";host"))) Then 
              upos = InStr(tempstr, "UID=")
              ppos = InStr(tempstr, ";PWD=")
              pend = InStr(tempstr, ";host")
              user = tempstr.Substring(upos+3, ppos-upos-4)
              pwd = tempstr.Substring(ppos+4, pend-ppos-5)
              tempstr = tempstr.Replace("UID=" & user, "UID=****")
              tempstr = tempstr.Replace("PWD=" & pwd, "PWD=****")
            Else
              tempstr = Copient.PhraseLib.Lookup("term.db2-dberror", LanguageID)
            End If
            Send("  <input type=""text"" class=""longest"" id=""DB2"" name=""DB2""  size=""100"" value=""" & tempStr & """ ""  /></td>")
            Send("</tr>")
          End If
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
