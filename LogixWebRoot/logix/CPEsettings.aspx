<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEsettings.aspx 
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
  Dim rst3 As System.Data.DataTable
    Dim row3 As System.Data.DataRow
    Dim rst4 As System.Data.DataTable
    Dim row4 As System.Data.DataRow
  Dim OptionID As Integer
  Dim tempstr As String
  Dim infoMessage As String = ""
  Dim OpenTagEscape As String = "<>"
  Dim Handheld As Boolean = False
  Dim OptionObj As Copient.SystemOption = Nothing
  Dim HistoryStr As String = ""
  Dim DisableEdit As String = " disabled=""disabled"""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.AppName = "CPEsettings.aspx"
  MyCommon.Open_LogixRT()
  
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Response.Expires = 0
  
  If (Request.QueryString("save") <> "") Then
    ' someone clicked save lets get to it
    MyCommon.QueryStr = "select OptionID, OptionName, OptionValue from CPE_SystemOptions with (NoLock) where visible=1 order by OptionID;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        tempstr = MyCommon.Parse_Quotes(Request.QueryString("oid" & MyCommon.NZ(row.Item("OptionID"), 0)))
        tempstr = Logix.TrimAll(tempstr)
        ' Replace tag escape put into place due to request failure caused by scripting validation exception
        ' encountered with tags being sent over the querystring.
        If (Not tempstr Is Nothing AndAlso tempstr.IndexOf(OpenTagEscape) > -1) Then
          tempstr = tempstr.Replace(OpenTagEscape, "<")
        End If

        OptionObj = New Copient.SystemOption(MyCommon.NZ(row.Item("OptionID"), 0), MyCommon.NZ(row.Item("OptionValue"), ""))
        OptionObj.SetNewValue(tempstr)

        If OptionObj.IsModified Then
          If (IsValidEntry(OptionObj.GetOptionID, MyCommon.NZ(row.Item("OptionName"), ""), OptionObj.GetNewValue, MyCommon, infoMessage)) Then
            MyCommon.QueryStr = "Update CPE_SystemOptions with (RowLock) set OptionValue=@NewValue, LastUpdate=getdate() where OptionID=@OptionID;"
            MyCommon.DBParameters.Add("@NewValue", SqlDbType.NVarChar, 255).Value = OptionObj.GetNewValue()
            MyCommon.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID()
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            
            If (MyCommon.RowsAffected > 0) Then
              HistoryStr = Copient.PhraseLib.Lookup("history.edit-cpesetting", LanguageID) & " '" & MyCommon.NZ(row.Item("OptionName"), "") & "'" & _
                           " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & ": " & OptionObj.GetOldValue() & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & ": " & OptionObj.GetNewValue()
              MyCommon.Activity_Log(32, 0, AdminUserID, HistoryStr)
            End If
          End If
        End If
      Next
      
      ' Refresh cache with new system options
      CMS.AMS.CurrentRequest.Resolver.AppName = MyCommon.AppName
      Dim cacheData As CMS.AMS.Contract.ICacheData = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.ICacheData)()
            cacheData.ClearAllSystemOptionsCache()
            Copient.SystemOptionsCache.RemoveCache(System.Web.HttpContext.Current.Request.Url.Host)
    End If
  End If
  
  Send_HeadBegin("term.cpesettings")
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
      <% Sendb(Copient.PhraseLib.Lookup("term.cpesettings", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemOptions = True or Logix.UserRoles.AccessSystemSettings) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(30, 0, AdminUserID)
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
          MyCommon.QueryStr = "select SO.OptionName, SO.OptionID, SO.OptionValue, SO.PhraseID, SO.Visible, SO.OptionType, SO.OptionTypePhraseID, IsNull(PT.Phrase, SO.OptionName) as Phrase, PT.LanguageID from CPE_SystemOptions as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID order by OptionType, OptionName;"
        Else
          MyCommon.QueryStr = "select SO.OptionName, SO.OptionID, SO.OptionValue, SO.PhraseID, SO.Visible, SO.OptionType, SO.OptionTypePhraseID, IsNull(PT.Phrase, SO.OptionName) as Phrase, PT.LanguageID from CPE_SystemOptions as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID where visible=1 order by OptionType, OptionName;"
        End If
        MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        Dim Counter As Integer = 0
        If (rst.Rows.Count > 0) Then
          For Each row In rst.Rows
            OptionID = MyCommon.NZ(row.Item("OptionID"), 0)
            If MyCommon.NZ(row.Item("OptionType"), 0) > Counter Then
              Send("<tr>")
              Send("  <td>")
              Send("    <h2>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("OptionTypePhraseID"), 0), LanguageID) & "</h2>")
              Send("  </td>")
              Send("</tr>")
            End If
                  
            '----------------------------------------------------------'
            Select Case OptionID
              Case 116, 117, 118 ' default chargeback departments
                'only display the selection for default chargeback departments if banners are not enabled
                'if banners are enabled, then each banner could have different chargeback departments associated with it this messes up the concept of system wide defaults
                If MyCommon.Fetch_SystemOption(66) = "0" Then  'Banners enabled?
                  Send("")
                  Send("<tr>")
                  Send("  <td " & IIf(row.Item("Visible"), "", " style=""color:red;""") & ">")
                  Send("    <label for=""oid" & OptionID & """>" & MyCommon.NZ(row.Item("Phrase"), "") & ":</label>")
                  
                  MyCommon.QueryStr = "select distinct convert(nvarchar,ChargebackDeptID) as ChargebackDeptID,  " & _
                                      "  case when (ISNULL(PT.PhraseID,0) > 0 and ISNULL(ExternalID, '') <> '') then  ExternalID + ' - ' + convert(nvarchar(100),PT.Phrase) " & _
                                      "       when (ISNULL(PT.PhraseID,0) > 0 and ISNULL(ExternalID, '') =  '') then  convert(nvarchar(100),PT.Phrase) " & _
                                      "       when (ISNULL(PT.PhraseID,0) = 0 and ISNULL(ExternalID, '') <> '') then  ExternalID + ' - ' + CBD.Name " & _
                                      "       else CBD.Name " & _
                                      "  end as OptionText " & _
                                      "from ChargeBackDepts as CBD with (NoLock) " & _
                                      "left join UIPhrases as UIP with (NoLock) on UIP.PhraseID = CBD.PhraseID  " & _
                                      "left join PhraseText as PT with (NoLock) on PT.PhraseID = UIP.PhraseID and LanguageID = @LanguageID " & _
                                      "where Deleted=0 and ISNULL(BannerID,0)=0"
                  If OptionID = 118 Then ' basket level
                    MyCommon.QueryStr &= " and ChargeBackDeptID<>0 "
                  ElseIf OptionID = 117 Then ' dept level
                    MyCommon.QueryStr &= " and ChargeBackDeptID<>0 and ChargeBackDeptID<>14 "
                  ElseIf OptionID = 116 Then 'item level
                    MyCommon.QueryStr &= " and ChargeBackDeptID<>10 "
                  End If
                  MyCommon.QueryStr &= " order by OptionText;"

                  MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
                  rst3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                  If rst3.Rows.Count = 0 Then
                    Send("<input type=""text"" id=""option" & OptionID & """ name=""option" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """" & IIf(row.Item("Visible"), "", DisableEdit) & " />")
                  Else
                    Send("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("Visible"), "", DisableEdit) & ">")
                    For Each row3 In rst3.Rows
                      If row3.Item("ChargebackDeptID") = MyCommon.NZ(row.Item("OptionValue"), "") Then
                        Send("      <option value=""" & MyCommon.NZ(row3.Item("ChargebackDeptID"), "") & """ selected=""selected"">" & MyCommon.NZ(row3.Item("OptionText"), "") & "</option>")
                      Else
                        Send("      <option value=""" & MyCommon.NZ(row3.Item("ChargebackDeptID"), "") & """>" & MyCommon.NZ(row3.Item("OptionText"), "") & "</option>")
                      End If
                    Next
                    Send("    </select>")
                  End If
                  Send("  </td>")
                  Send("</tr>")
                Else
                  'banners are enabled, so we need to send hidden form values to prevent these from being set to blank in the CPE_SystemOptions table
                  Send("<input type=""hidden"" id=""oid" & OptionID & """ & name=""oid" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
                End If  'banners not enabled
                         
              Case 128  'Enable Preference data distribution for CPE
                'only display the option to enable/disable preference data distribution if the EPM integration is enabled
                If Not (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER)) Then
                  Send("<input type=""hidden"" id=""oid" & OptionID & """ & name=""oid" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
                Else
                  Send_Option(MyCommon, MyCommon.NZ(row.Item("Phrase"), ""), OptionID, MyCommon.NZ(row.Item("OptionValue"), ""), OpenTagEscape, row.Item("Visible"))
                End If

              Case 188 'RLM Name of Scorecard Preference
               If  (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER)) Then
               MyCommon.Close_LogixRT()
               MyCommon.Open_PrefManRT()
               'only display the selection for Scorecard Preference if RLM integration enabled
               'if RLM integration enabled, then each Scorecards get a preference associated with it 
               If MyCommon.Fetch_CPE_SystemOption(163) = "1" Then  'RLM Integration
                 Send("")
                 Send("<tr>")
                 Send("  <td " & IIf(row.Item("Visible"), "", " style=""color:red;""") & ">")
                 Send("    <label for=""oid" & OptionID & """>" & MyCommon.NZ(row.Item("Phrase"), "") & ":</label>")
                 
                 MyCommon.QueryStr = "select distinct Name from Preferences where DataTypeID=1 and Deleted=0"
             
               ' MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
                 rst3 = MyCommon.ExecuteQuery(Copient.DataBases.PrefManRT)
                 If rst3.Rows.Count = 0 Then
                   Send("<input type=""text"" id=""option" & OptionID & """ name=""option" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """" & IIf(row.Item("Visible"), "", DisableEdit) & " />")
                 Else
                   Send("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("Visible"), "", DisableEdit) & ">")
                   Send("      <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
                   For Each row3 In rst3.Rows
                     If row3.Item("Name") = MyCommon.NZ(row.Item("OptionValue"), "") Then
                       Send("      <option value=""" & MyCommon.NZ(row3.Item("Name"), "") & """ selected=""selected"">" & MyCommon.NZ(row3.Item("Name"), "") & "</option>")
                     Else
                       Send("      <option value=""" & MyCommon.NZ(row3.Item("Name"), "") & """>" & MyCommon.NZ(row3.Item("Name"), "") & "</option>")
                     End If
                   Next
                   Send("    </select>")
                 End If
                 Send("  </td>")
                 Send("</tr>")
               Else
                 'banners are enabled, so we need to send hidden form values to prevent these from being set to blank in the CPE_SystemOptions table
                 Send("<input type=""hidden"" id=""oid" & OptionID & """ & name=""oid" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
               End If  'RLM Integration not enabled
               
               MyCommon.Close_PrefManRT()
               MyCommon.Open_LogixRT()
                          End If
                      Case 181 'External Source to require AltID with PIN
                          Send("")
                          Send("<tr>")
                          Send("  <td " & IIf(row.Item("Visible"), "", " style=""color:red;""") & ">")
                          Send("    <label for=""oid" & OptionID & """>" & MyCommon.NZ(row.Item("Phrase"), "") & ":</label>")
                  
                          MyCommon.QueryStr = "select ExtInterfaceID, Name from ExtCRMInterfaces where Active=1     and Deleted=0;"
                          MyCommon.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
                          rst4 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
             
                          If rst4.Rows.Count = 0 Then
                              Send("<input type=""text"" id=""option" & OptionID & """ name=""option" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """" & IIf(row.Item("Visible"), "", DisableEdit) & " />")
                          Else
                              Send("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("Visible"), "", DisableEdit) & ">")
                              For Each row4 In rst4.Rows
                                  If row4.Item("ExtInterfaceID") = MyCommon.NZ(row.Item("OptionValue"), "") Then
                                      Send("      <option value=""" & MyCommon.NZ(row4.Item("ExtInterfaceID"), "") & """ selected=""selected"">" & MyCommon.NZ(row4.Item("ExtInterfaceID"), "") & " - " & MyCommon.NZ(row4.Item("Name"), "") & "</option>")
                                  Else
                                      Send("      <option value=""" & MyCommon.NZ(row4.Item("ExtInterfaceID"), "") & """>" & MyCommon.NZ(row4.Item("ExtInterfaceID"), "") & " - " & MyCommon.NZ(row4.Item("Name"), "") & "</option>")
                                  End If
                              Next
                              Send("    </select>")
                          End If
                          Send("  </td>")
                          Send("</tr>")
			   
              Case Else
                Send_Option(MyCommon, MyCommon.NZ(row.Item("Phrase"), ""), OptionID, MyCommon.NZ(row.Item("OptionValue"), ""), OpenTagEscape, row.Item("Visible"))
                
            End Select

            Counter = MyCommon.NZ(row.Item("OptionType"), 0)
          Next
        End If
      %>
    </table>
  </div>
</form>

<script runat="server">
  Private Function IsValidEntry(ByVal OptionID As Integer, ByVal OptionName As String, ByVal OptionValue As String, _
                                ByRef MyCommon As Copient.CommonInc, ByRef infoMessage As String) As Boolean
    Dim ValidEntry As Boolean = True
    
    'If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    
    ' Add OptionIDs validation routines to the case statement as validation is needed
    Select Case OptionID
      Case Else
        ValidEntry = True
    End Select

    Return ValidEntry
  End Function
  
  '------------------------------------------------------------------------------------------------------------------
  
  Private Sub Send_Option(ByRef Common As Copient.CommonInc, ByVal OptionPhrase As String, ByVal OptionID As Integer, ByVal OptionValue As String, ByVal OpenTagEscape As String, ByVal Visible As Boolean)
    Send("")
    Send("<tr>")
    Send("  <td " & IIf(Visible, "", " style=""color:red;""") & ">")
    Send("    <label for=""oid" & OptionID & """>" & OptionPhrase & ":</label>")
    Send_Option_Value(Common, OptionID, OptionValue, OpenTagEscape, Visible)
    Send("  </td>")
    Send("</tr>")

  End Sub
  
  '------------------------------------------------------------------------------------------------------------------
  
  Private Sub Send_Option_Value(ByRef Common As Copient.CommonInc, ByVal OptionID As Integer, ByVal OptionValue As String, ByVal OpenTagEscape As String, ByVal Visible As Boolean)
    
    Dim dst As DataTable
    Dim row As DataRow
    Dim EscapedValue As String
    Dim DisableEdit As String = " disabled=""disabled"""
    
    Common.QueryStr = "select SOV.OptionValue, SOV.Description, SOV.PhraseID, IsNull(PT.Phrase, SOV.Description) as Phrase, PT.LanguageID from CPE_SystemOptionValues as SOV with (NoLock) left join PhraseText as PT with (NoLock) on SOV.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID where OptionID=@OptionID order by OptionValue;"
    Common.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
    Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
    dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                  
    If (dst.Rows.Count > 0) Then
      Send("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(Visible, "", DisableEdit) & ">")
      For Each row In dst.Rows
        EscapedValue = Common.NZ(row.Item("OptionValue"), "")
        EscapedValue = EscapedValue.Replace("<", OpenTagEscape)
        Sendb("      <option value=""" & EscapedValue & """")
        If Common.NZ(row.Item("OptionValue"), "") = OptionValue Then Sendb(" selected=""selected""")
        Send(">" & Common.NZ(row.Item("Phrase"), "") & "</option>")
      Next
      Send("    </select>")
    Else
      Send("    <input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & OptionValue & """" & IIf(Visible, "", DisableEdit) & " />")
    End If

  End Sub
  
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(30, 0, AdminUserID)
    End If
  End If
  
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
