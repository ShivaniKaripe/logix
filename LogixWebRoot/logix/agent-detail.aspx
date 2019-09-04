<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%
  ' *****************************************************************************
  ' * FILENAME: agent-detail.aspx 
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
%>

<script runat="server">
  Dim infoMessage As String = ""
  Dim AdminUserID As Long
  Dim programID As Integer = 0
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  
  '---------------------------------------------------------------------
  
  Sub Save_Settings()
    Dim AppID As Long
    Dim SendAlert As Integer
    Dim RunFreq As Integer
    
    AppID = MyCommon.Extract_Val(Request.QueryString("appid"))
    RunFreq = MyCommon.Extract_Decimal(GetCgiValue("runfreq"), MyCommon.GetAdminUser.Culture)
    SendAlert = MyCommon.Extract_Val(Request.QueryString("sendalert"))
    If RunFreq < 1 Then RunFreq = 1
        If (AppID = 110) Then
            MyCommon.QueryStr = "Update LastSync with (RowLock) set  SendAlert=" & SendAlert & " where AppID=" & AppID & ";"
        Else
            MyCommon.QueryStr = "Update LastSync with (RowLock) set RunFreq=" & RunFreq & ", SendAlert=" & SendAlert & " where AppID=" & AppID & ";"
        End If
        
    MyCommon.LRT_Execute()

    Save_Options(AppID)
    End Sub
  
  '---------------------------------------------------------------------
  
  Sub Save_Options(ByVal AppID As Long)
    Dim dt As DataTable
    Dim row As DataRow
    Dim TempStr As String
    Dim HistoryStr As String
    Dim OptionObj As Copient.SystemOption
        Dim oldPromoVarValue As Integer = 0
        Dim oldDiscountRateValue As Single = 0.0F
        Dim newPromoVarValue As Integer = 0
        Dim newDiscountRateValue As Single = 0.0F
        
        Dim LogFile As String = String.Format("ErrorLog.{0}.txt", MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2))
        
        If (MyCommon.Fetch_CM_SystemOption(143) = "1" AndAlso AppID = 110) Then
            
            Dim option145Name As String = ""
            Dim option146Name As String = ""
            MyCommon.QueryStr = "select OptionID, OptionName, OptionValue, PhraseID " & _
                        "from CM_Systemoptions with (NoLock) where OptionID In(145,146) order by OptionID;"
            dt = MyCommon.LRT_Select()
    
            If dt.Rows.Count > 0 Then
                For Each row In dt.Rows
                    If (MyCommon.NZ(row.Item("OptionID"), 0) = 145) Then
                        oldPromoVarValue = MyCommon.NZ(row.Item("OptionValue"), 0)
                        option145Name = MyCommon.NZ(row.Item("OptionName"), "")
                    End If
                    If (MyCommon.NZ(row.Item("OptionID"), 0) = 146) Then
                        oldDiscountRateValue = MyCommon.NZ(row.Item("OptionValue"), 0)
                        option146Name = MyCommon.NZ(row.Item("OptionName"), "")
                    End If
                Next
            End If
            
            Dim TempStrPromoVar As String
            Dim TempStrDiscountRate As String

            TempStrPromoVar = MyCommon.Parse_Quotes(Request.QueryString("option145"))
            TempStrPromoVar = Logix.TrimAll(TempStrPromoVar)
                
            TempStrDiscountRate = MyCommon.Parse_Quotes(Request.QueryString("option146"))
            TempStrDiscountRate = Logix.TrimAll(TempStrDiscountRate)

            'MyCommon.QueryStr = "Select LinkID from PromoVariables With (NoLock) where Deleted=0 and PromoVarID=" & TempStrPromoVar
            'dt = MyCommon.LXS_Select()
            'If (dt.Rows.Count > 0) Then
            '    programID = MyCommon.NZ(dt.Rows(0)("LinkID"), 0)
            'End If
            
            If (Not String.IsNullOrEmpty(TempStrPromoVar)) Then
                newPromoVarValue = Convert.ToInt64(TempStrPromoVar)
            End If
                
                
            If (Not String.IsNullOrEmpty(TempStrDiscountRate)) Then
                newDiscountRateValue = Convert.ToSingle(TempStrDiscountRate)
            End If
                
                        
            If (programID = 0) Then
                newPromoVarValue = oldPromoVarValue
                MyCommon.Write_Log(LogFile, "Cannot update CM System option #145 (Employee Promotion Variable ID ) as no valid ProgramID exists for promotion variable" & newPromoVarValue)
            Else
                MyCommon.QueryStr = "Update CM_SystemOptions with (RowLock) set OptionValue=N'" & newPromoVarValue & "', LastUpdate=getdate() where OptionID=145"
                MyCommon.LRT_Execute()
            
                If (MyCommon.RowsAffected > 0) Then
                    HistoryStr = Copient.PhraseLib.Lookup("history.edit-agentsetting", LanguageID) & " '" & option145Name & "'" & _
                                 " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & ": " & oldPromoVarValue & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & ": " & newPromoVarValue
                    MyCommon.Activity_Log(24, 0, AdminUserID, HistoryStr)
                End If
            End If
            
            MyCommon.QueryStr = "Update CM_SystemOptions with (RowLock) set OptionValue=N'" & newDiscountRateValue & "', LastUpdate=getdate() where OptionID=146"
            MyCommon.LRT_Execute()
    
            If (MyCommon.RowsAffected > 0) Then
                HistoryStr = Copient.PhraseLib.Lookup("history.edit-agentsetting", LanguageID) & " '" & option146Name & "'" & _
                             " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & ": " & oldDiscountRateValue & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & ": " & newDiscountRateValue
                MyCommon.Activity_Log(24, 0, AdminUserID, HistoryStr)
            End If
            'Update in Points table if any of the values are updated
            If (oldPromoVarValue <> newPromoVarValue) AndAlso (oldDiscountRateValue <> newDiscountRateValue) Then
                If (programID > 0) Then
                    MyCommon.QueryStr = "Update Points  with (RowLock) set PromoVarID=" & newPromoVarValue & ", ProgramID=" & programID & _
                        "  , Amount=" & newDiscountRateValue & " Where PromoVarID = " & oldPromoVarValue
                    MyCommon.LXS_Execute()
                End If
            ElseIf (oldPromoVarValue <> newPromoVarValue) Then
                MyCommon.QueryStr = "Update Points  with (RowLock) set PromoVarID=" & newPromoVarValue & ", ProgramID=" & programID & _
                        " Where PromoVarID = " & oldPromoVarValue
                MyCommon.LXS_Execute()
            ElseIf (oldDiscountRateValue <> newDiscountRateValue) Then
                MyCommon.QueryStr = "Update Points  with (RowLock) set  Amount=" & newDiscountRateValue & " Where PromoVarID = " & oldPromoVarValue
                MyCommon.LXS_Execute()
            End If
        Else
    MyCommon.QueryStr = "select OptionID, OptionName, OptionValue, PhraseID " & _
                        "from InterfaceOptions with (NoLock) " & _
                        "where AppID = " & AppID & " and Visible=1 " & _
                        "order by OptionName;"
    dt = MyCommon.LRT_Select
    
    If dt.Rows.Count > 0 Then
      For Each row In dt.Rows
        TempStr = MyCommon.Parse_Quotes(Request.QueryString("option" & MyCommon.NZ(row.Item("OptionID"), 0)))
        TempStr = Logix.TrimAll(TempStr)

        OptionObj = New Copient.SystemOption(MyCommon.NZ(row.Item("OptionID"), 0), MyCommon.NZ(row.Item("OptionValue"), ""))
        OptionObj.SetNewValue(TempStr)

        If OptionObj.IsModified Then
          MyCommon.QueryStr = "Update InterfaceOptions with (RowLock) set OptionValue=N'" & OptionObj.GetNewValue() & "', LastUpdate=getdate() where OptionID=" & OptionObj.GetOptionID()
          MyCommon.LRT_Execute()
            
          If (MyCommon.RowsAffected > 0) Then
            HistoryStr = Copient.PhraseLib.Lookup("history.edit-agentsetting", LanguageID) & " '" & MyCommon.NZ(row.Item("OptionName"), "") & "'" & _
                         " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & ": " & OptionObj.GetOldValue() & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & ": " & OptionObj.GetNewValue()
            MyCommon.Activity_Log(24, 0, AdminUserID, HistoryStr)
          End If
        End If
      Next
    End If
        End If
    Send_Detail_Page()
  End Sub
  
  '---------------------------------------------------------------------

  Sub Send_Detail_Page()
    Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
    
    Dim dst, dst2 As DataTable
    Dim AppID As Long
    Dim LogFileID As integer
    Dim Handheld As Boolean = False
    Dim shaded As Boolean = True
    
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
      Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
    
    'AdminUserID = Verify_AdminUser(MyCommon, Logix)
    AppID = MyCommon.Extract_Val(Request.QueryString("appid"))
    If AppID = 0 Then AppID = MyCommon.Extract_Val(Request.QueryString("appid"))
    
    Send_HeadBegin("term.agent", , AppID)
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
    Send_Subtabs(Logix, 8, 1)
    
    If (Logix.UserRoles.AccessSystemHealth = False) Then
      Send_Denied(1, "perm.admin-health")
      GoTo done
    End If
    
    Send("<form method=""get"" action=""agent-detail.aspx"" id=""mainform"" name=""mainform"">")
    Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""save"" />")
    Send("<input type=""hidden"" id=""appid"" name=""appid"" value=""" & AppID & """ />")
    Send("")
    
    MyCommon.QueryStr = "select AppName, LastLaunchTime, LastStartTime, LastEndTime, LastTouchTime, SendAlert, RunFreq, DescriptionPhraseID, getdate() as CurrentTime from LastSync with (NoLock) where AppID=" & AppID & ";"
    dst = MyCommon.LRT_Select
    
    Send("<div id=""intro"">")
    Sendb(" <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.agent", LanguageID))
    If dst.Rows.Count > 0 Then
      Sendb(" #" & AppID & ": " & MyCommon.NZ(dst.Rows(0).Item("AppName"), ""))
    End If
    Sendb("</h1>")
    If (Logix.UserRoles.EditSystemConfiguration = True) Then
      Send(" <div id=""controls"">")
      Send_Save()
      If MyCommon.Fetch_SystemOption(75) Then
        If (AppID > 0 And Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(16, AppID, AdminUserID)
        End If
      End If
      Send(" </div>")
    End If
    Send("</div>")
    Send("")
    Send("<div id=""main"">")
    If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    Send("<div id=""column1"">")
    
    If (dst.Rows.Count > 0) Then
      Send(" <div class=""box"">")
      Send(" <h2><span>" & Copient.PhraseLib.Lookup("term.application", LanguageID) & "</span></h2>")
      Send(" <table border=""0"" cellpadding=""0"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.health", LanguageID) & """>")
      Send("  <tr>")
      Send("   <td align=""right"">" & Copient.PhraseLib.Lookup("term.application", LanguageID) & ":</td>")
      Send("   <td><b>" & MyCommon.NZ(dst.Rows(0).Item("AppName"), "") & "</b></td>")
      Send("  </tr>")
      
      Send("  <tr>")
      Send("   <td align=""right"">" & Copient.PhraseLib.Lookup("term.lastlaunch", LanguageID) & ":</td>")
      If (Not IsDBNull(dst.Rows(0).Item("LastLaunchTime"))) Then
        Send("   <td><b>" & Logix.ToLongDateTimeString(dst.Rows(0).Item("LastLaunchTime"), MyCommon) & "</b></td>")
      Else
        Send("   <td><b>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</b></td>")
      End If
      Send("  </tr>")
      Send("  <tr>")
      Send("   <td align=""right"">" & Copient.PhraseLib.Lookup("term.laststarttime", LanguageID) & ":</td>")
      If (Not IsDBNull(dst.Rows(0).Item("LastStartTime"))) Then
        Send("   <td><b>" & Logix.ToLongDateTimeString(dst.Rows(0).Item("LastStartTime"), MyCommon) & "</b></td>")
      Else
        Send("   <td><b>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</b></td>")
      End If
      Send("  </tr>")
      Send("  <tr>")
      Send("   <td align=""right"">" & Copient.PhraseLib.Lookup("term.lasttouchtime", LanguageID) & ":</td>")
      If (Not IsDBNull(dst.Rows(0).Item("LastTouchTime"))) Then
        Send("   <td><b>" & Logix.ToLongDateTimeString(dst.Rows(0).Item("LastTouchTime"), MyCommon) & "</b></td>")
      Else
        Send("   <td><b>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</b></td>")
      End If
      Send("  </tr>")
      Send("  <tr>")
      Send("   <td align=""right"">" & Copient.PhraseLib.Lookup("term.lastendtime", LanguageID) & ":</td>")
      If (Not IsDBNull(dst.Rows(0).Item("LastEndTime"))) Then
        Send("   <td><b>" & Logix.ToLongDateTimeString(dst.Rows(0).Item("LastEndTime"), MyCommon) & "</b></td>")
      Else
        Send("   <td><b>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</b></td>")
      End If
      Send("  </tr>")
      Send("  <tr>")
      Send("   <td align=""right"">" & Copient.PhraseLib.Lookup("term.currenttime", LanguageID) & ":</td>")
      If (Not IsDBNull(dst.Rows(0).Item("CurrentTime"))) Then
        Send("   <td><b>" & Logix.ToLongDateTimeString(dst.Rows(0).Item("CurrentTime"), MyCommon) & "</b></td>")
      Else
        Send("   <td><b>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</b></td>")
      End If
      Send("  </tr>")
      Send("  <tr>")
      Send("   <td align=""right""><label for=""runfreq"">" & Copient.PhraseLib.Lookup("term.runfrequency", LanguageID) & ":</label></td>")
      If (MyCommon.NZ(dst.Rows(0).Item("RunFreq"), "").ToString <> "") Then
        Send("   <td><input type=""text"" class=""shorter"" id=""runfreq"" maxlength=""9"" name=""runfreq"" value=""" & MyCommon.NZ(dst.Rows(0).Item("RunFreq"), "") & """ /> " & StrConv(Copient.PhraseLib.Lookup("term.seconds", LanguageID), VbStrConv.Lowercase) & " </td>")
      Else
        Send("<td>" & Copient.PhraseLib.Lookup("term.na", LanguageID) & "</td>")
      End If
      Send("  </tr>")
      Send("  <tr>")
      Send("   <td align=""right""><label for=""sendalert"">" & Copient.PhraseLib.Lookup("health.enablealerts", LanguageID) & ":</label></td>")
      Sendb("   <td><input type=""checkbox"" id=""sendalert"" name=""sendalert"" value=""1""")
      If MyCommon.NZ(dst.Rows(0).Item("SendAlert"), False) = True Then Sendb(" checked=""checked""")
      Sendb(" /></td>")
      Send("  </tr>")
      If (Logix.UserRoles.AccessLogs = True) Then
        If Not (AppID = 15 AndAlso MyCommon.Fetch_InterfaceOption(26) = 0) Then
          MyCommon.QueryStr = "Select LogFileID from ApplicationLogFiles with (NoLock) where AppID = " & AppID
          dst2 = MyCommon.LRT_Select
          If dst2.Rows.Count > 0 Then
            LogFileID = dst2.Rows(0).Item(0)
            Send("  <tr>")
            Send("   <td></td>")
            Send("   <td><a href=""log-view.aspx?filetype=" & LogFileID & "&amp;fileyear=" & Year(Today) & "&amp;fileday=" & Day(Today) & "&amp;filemonth=" & Month(Today) & """ target=""_blank"">" & Copient.PhraseLib.Lookup("health.viewlog", LanguageID) & "</a></td>")
            Send("  </tr>")
          End If
        End If
      End If
      Send(" </table>")
      If MyCommon.NZ(dst.Rows(0).Item("DescriptionPhraseID"), 0) <> 0 Then
        Send(" <br class=""half"" />")
        Send(" <div style=""margin:0 10px 0 10px;"">")
        Send("  <hr />")
        Send("  <b>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</b><br />")
        Send("  <p>" & Copient.PhraseLib.Lookup(dst.Rows(0).Item("DescriptionPhraseID"), LanguageID) & "</p>")
        Send(" </div>")
      End If
      Send(" </div>")
    End If
    Send("</div>")
    
    Send("<div id=""gutter"">")
    Send("</div>")
    
        If (MyCommon.Fetch_CM_SystemOption(143) = "1" AndAlso AppID = 110) Then
            
            MyCommon.QueryStr = "select OptionID, OptionName, OptionValue, PhraseID " & _
                               "from CM_SystemOptions with (NoLock) " & _
                               "where OptionID in(145,146) order by OptionID; " 
            dst = MyCommon.LRT_Select
            Send("<div id=""column2"">")
            Send("  <div class=""box"">")
            Send("   <h2><span>" & Copient.PhraseLib.Lookup("term.options", LanguageID) & "</span></h2>")
            If (dst.Rows.Count > 0) Then
                For Each row As DataRow In dst.Rows
                    Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("OptionName"), "")) & ":")
                    Send("  <br />")
                    Send("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
                    Send("<input type=""text"" id=""option" & MyCommon.NZ(row.Item("OptionID"), 0) & """ name=""option" & MyCommon.NZ(row.Item("OptionID"), 0) & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
                    Send("  <hr />")
                Next
                Send("<input type=""hidden"" id=""option145ProgramID"" name=""option145ProgramID""  />")
            End If
          
            Send("</div>")
            Send("</div>")
        Else
    ' load up all the options for this connector; if none, then don't show the option box
    MyCommon.QueryStr = "select OptionID, OptionName, OptionValue, PhraseID " & _
                        "from InterfaceOptions with (NoLock) " & _
                        "where AppID = " & AppID & " and Visible=1 " & _
                        "order by OptionName;"
    dst = MyCommon.LRT_Select

    If dst.Rows.Count > 0 Then
      Send("<div id=""column2"">")
      Send("  <div class=""box"">")
      Send("   <h2><span>" & Copient.PhraseLib.Lookup("term.options", LanguageID) & "</span></h2>")
      For Each row As DataRow In dst.Rows
        Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("OptionName"), "")) & ":")
        Send("  <br />")

        MyCommon.QueryStr = "select OptionValue, Description, PhraseID from InterfaceOptionValues with (NoLock) " & _
                            "where OptionID=" & MyCommon.NZ(row.Item("OptionID"), 0)
        dst2 = MyCommon.LRT_Select
        Send("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
        If dst2.Rows.Count = 0 Then
          Send("<input type=""text"" id=""option" & MyCommon.NZ(row.Item("OptionID"), 0) & """ name=""option" & MyCommon.NZ(row.Item("OptionID"), 0) & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
        Else
          Send("<select id=""option" & MyCommon.NZ(row.Item("OptionID"), 0) & """ name=""option" & MyCommon.NZ(row.Item("OptionID"), 0) & """>")
          For Each row2 As DataRow In dst2.Rows
            If MyCommon.NZ(row2.Item("OptionValue"), "") = MyCommon.NZ(row.Item("OptionValue"), "") Then
              Send("  <option value=""" & MyCommon.NZ(row2.Item("OptionValue"), "") & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Description"), "")) & "</option>")
            Else
              Send("  <option value=""" & MyCommon.NZ(row2.Item("OptionValue"), "") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Description"), "")) & "</option>")
            End If
          Next
          Send("</select>")
        End If
        Send("  <hr />")
      Next
      Send("</div>")
      Send("</div>")
    End If
        End If


    Send("<br clear=""all"" />")

    ' load up all the currently running versions of this agent
    MyCommon.QueryStr = "SELECT LastTouchTime,LastStartTime,LastEndTime,ProcessingServer FROM LastSyncDetails WHERE APPID=" & AppID
    dst = MyCommon.LRT_Select

    If dst.Rows.Count > 0 Then


      Send("<div id=""column"">")
      Send("   <div class=""box"" id=""conditions"">")
      Send("     <h2>")
      Send("       <span>")
      Send("         " & Copient.PhraseLib.Lookup("term.agents", LanguageID))
      Send("       </span>")
      Send("     </h2>")
      Send("     <table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & """>")
      Send("       <thead>")
      Send("         <tr>")
      Send("           <th align=""left"" scope=""col"" >")
      Sendb(Copient.PhraseLib.Lookup("term.server", LanguageID))
      Send("           </th>")
      Send("           <th align=""left"" scope=""col"" class=""th-datetime"">")
      Sendb(Copient.PhraseLib.Lookup("term.laststarted", LanguageID))
      Send("           </th>")
      Send("           <th align=""left"" scope=""col"" class=""th-datetime"">")
      Sendb(Copient.PhraseLib.Lookup("term.lasttouch", LanguageID))
      Send("           </th>")
      Send("           <th align=""left"" scope=""col"" class=""th-datetime"">")
      Sendb(Copient.PhraseLib.Lookup("term.lastfinished", LanguageID))
      Send("           </th>")
      Send("         </tr>")
      Send("       </thead>")
      Send("       <tbody>")



      For Each row As DataRow In dst.Rows

        If (shaded) Then
          Send("<tr class=""shaded"">")
          shaded = False
        Else
          Send("<tr>")
          shaded = True
        End If
        
        Send("<TD>" & MyCommon.NZ(row.Item("ProcessingServer"), "") & "</TD>")
      
        If (Not IsDBNull(row.Item("LastStartTime"))) Then
          Send("  <td>" & Logix.ToLongDateTimeString(MyCommon.NZ(row.Item("LastStartTime"), New Date(1900, 1, 1)), MyCommon) & "</td>")
        Else
          Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
        End If
        If (Not IsDBNull(row.Item("LastTouchTime"))) Then
          Send("  <td>" & Logix.ToLongDateTimeString(MyCommon.NZ(row.Item("LastTouchTime"), New Date(1900, 1, 1)), MyCommon) & "</td>")
        Else
          Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
        End If
        If (Not IsDBNull(row.Item("LastEndTime"))) Then
          Send("  <td>" & Logix.ToLongDateTimeString(MyCommon.NZ(row.Item("LastEndTime"), New Date(1900, 1, 1)), MyCommon) & "</td>")
        Else
          Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
        End If
        Send("</TR>")
      Next
      
      
         
      Send("       </tbody>")
      Send("     </table>")
      Send("   </div>")
      
      Send("   </div>")
      
    End If





    Send("   </div>")

    Send("")
    Send("</form>")

    dst = Nothing
    If MyCommon.Fetch_SystemOption(75) Then
      If (AppID > 0 And Logix.UserRoles.AccessNotes) Then
        Send_Notes(16, AppID, AdminUserID)
      End If
    End If
done:
    Send_BodyEnd("mainform", "runfreq")
  End Sub
  
  '---------------------------------------------------------------------
</script>

<%
  Dim Mode As String
  Dim AppID As Long
    
  Response.Expires = 0
  MyCommon.AppName = "agent-detail.aspx"
  On Error GoTo ErrorTrap
  
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  AppID = MyCommon.Extract_Val(Request.QueryString("appid"))
  Mode = ""
  Mode = Request.QueryString("mode")
  If Mode = "" Then Mode = Request.QueryString("mode")
  
  Select Case UCase(Mode)
    Case "SAVE"
            Dim TempStrPromoVar As String = "0"
            If (MyCommon.Fetch_CM_SystemOption(143) = "1" AndAlso AppID = 110) Then
                If (Not String.IsNullOrEmpty(Request.QueryString("option145"))) Then
                    TempStrPromoVar = MyCommon.Parse_Quotes(Request.QueryString("option145"))
                    TempStrPromoVar = Logix.TrimAll(TempStrPromoVar)
                    MyCommon.QueryStr = "Select LinkID from PromoVariables With (NoLock) where Deleted=0 and PromoVarID=" & TempStrPromoVar
                    Dim dt As DataTable = MyCommon.LXS_Select()
                    If (dt.Rows.Count > 0) Then
                        programID = MyCommon.NZ(dt.Rows(0)("LinkID"), 0)
                    End If
                End If
                
                If (programID > 0) Then
                    Save_Settings()
                    MyCommon.Activity_Log(24, AppID, AdminUserID, Copient.PhraseLib.Lookup("history.settings", LanguageID))
                Else
                    infoMessage = "Cannot update 'Employee Promotion Variable ID' as no valid ProgramID exists for promotion variable :" & TempStrPromoVar
                    Send_Detail_Page()
                End If
            Else
                Save_Settings()
                MyCommon.Activity_Log(24, AppID, AdminUserID, Copient.PhraseLib.Lookup("history.settings", LanguageID))
            End If
    Case Else
      Send_Detail_Page()
  End Select

  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
  Response.End()
  
ErrorTrap:
  MyCommon.Error_Processor()
%>
