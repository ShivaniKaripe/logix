<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<script runat="server">
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim LogFile As String = "ExternalRewardsFeed." & Format(Now(), "yyyyMMdd") & ".txt"
</script>
<%
  
    ' *****************************************************************************
    ' * FILENAME: ExternalRewardsFeeds.aspx 
    ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' * Copyright © 2002 - 2013.  All rights reserved by:
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
    ' * Version : 6.0.1.92814 
    ' *
    ' *****************************************************************************
  
    Dim CopientFileName As String = "ExternalRewardsFeeds.aspx"
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
  
    Dim AdminUserID As Long
    Dim rst As DataTable
    Dim row As DataRow
    Dim Category As String = "0"
    Dim Description As String = ""
    Dim CategoryDesc As String = ""
    Dim OfferName As String = ""
    Dim MonthAbbrs(-1) As String
  
    Dim ProductGroupID As String = ""
    Dim Products As String = ""
    Dim OpertaionType As Integer = -1
    Dim ProductType As Integer = -1
  
    
    MyCommon.AppName = "ExternalRewardsFeeds.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    If (LanguageID = 0) Then
        LanguageID = MyCommon.Extract_Val(Request.QueryString("LanguageID"))
    End If
  
 
    If (Request.Form("GenerateExternalRewardsReport") <> "") Then
        Dim partner As String = MyCommon.NZ(Request.Form("partner"), "")
        Dim operation As String = MyCommon.NZ(Request.Form("operationtype"), "")
        Dim startDate As String = MyCommon.NZ(Request.Form("startdate"), "")
        Dim endDate As String = MyCommon.NZ(Request.Form("enddate"), "")
        Dim frequency As Integer = MyCommon.NZ(Request.Form("frequency"), -1)
        Dim day As String = MyCommon.NZ(Request.Form("day"), "")
        If (String.IsNullOrEmpty(partner) OrElse String.IsNullOrEmpty(operation) OrElse String.IsNullOrEmpty(startDate) OrElse String.IsNullOrEmpty(endDate) OrElse (frequency = 3 AndAlso String.IsNullOrEmpty(day))) Then
            Dim errMsg As String = ""
            If (String.IsNullOrEmpty(partner)) Then
                errMsg &= Copient.PhraseLib.Lookup("term.partner", LanguageID) & ","
            End If
            If (String.IsNullOrEmpty(operation)) Then
                errMsg &= Copient.PhraseLib.Lookup("term.operation-type", LanguageID) & ","
            End If
            If (String.IsNullOrEmpty(startDate) OrElse String.IsNullOrEmpty(endDate)) Then
                errMsg &= Copient.PhraseLib.Lookup("term.daterange", LanguageID) & ","
            End If
            If (frequency = 3 AndAlso String.IsNullOrEmpty(day)) Then
                errMsg &= Copient.PhraseLib.Lookup("externalrewards.error-invaliddayofweek", LanguageID) & ","
            End If
            If (Not String.IsNullOrEmpty(errMsg)) Then
                errMsg = ": " & Copient.PhraseLib.Lookup("term.select", LanguageID) & errMsg
            End If
            Send("<b>" & Copient.PhraseLib.Lookup("term.invalid", LanguageID) & " " & Copient.PhraseLib.Lookup("term.criteria", LanguageID) & errMsg & "</b>")
        Else
            GenerateExternalRewardsReport(partner, operation, startDate, endDate, frequency, day)
        End If
    Else
        Send("<b>" & Copient.PhraseLib.Lookup("term.noarguments", LanguageID) & "</b>")
    End If
        
    MyCommon.Close_LogixRT()
    Logix = Nothing
    MyCommon = Nothing
    Response.Flush()
    Response.End()
%>
<script runat="server">
  
    Sub GenerateExternalRewardsReport(ByVal partner As String, ByVal operation As String, ByVal startdate As String, ByVal enddate As String, ByVal frequency As Integer, ByVal weekdays As String)
        Dim dt As DataTable
        Dim row As DataRow
        Dim ouputSB As New StringBuilder()
        Dim ShowLimitNote As Boolean = False
        Dim RECORD_LIMIT As Integer = 0
        Dim RecordLimitValue As String = ""
        
        Try
            RecordLimitValue = MyCommon.Fetch_SystemOption(176)
            If RecordLimitValue <> "" Then
                If IsNumeric(RecordLimitValue) Then
                    If CInt(RecordLimitValue) > 0 Then
                        RECORD_LIMIT = RecordLimitValue
                    End If
                End If
            End If
            
            MyCommon.Open_Logix3P()
      
            Dim _unnamed = Copient.PhraseLib.Lookup("term.unnamed", LanguageID)
            
            If Session("RptDataTable") IsNot Nothing Then
                Session.Remove("RptDataTable")
            End If
            
            Dim query As String = " SELECT " & IIf(RECORD_LIMIT > 0, " TOP(" & RECORD_LIMIT & ")", "") & " ip.Name,ot.OperationName,os.OperationDate,os.TotalCount,os.FailCount  "
            Dim joinQry = " FROM ExternalRewards_OperationSummary as os with (NoLock) INNER JOIN ExternalRewards_InternalPartner ip with (NoLock) ON ip.InternalPartnerId=os.InternalPartnerId INNER JOIN ExternalRewards_OperationType ot with (NoLock) ON os.OperationTypeChar=ot.OperationTypeChar "
            Dim whereQry As String = " WHERE os.OperationDate BETWEEN '" & startdate & "' AND '" & enddate & "'" & _
                "AND ip.InternalPartnerId IN (" & partner & ") AND os.OperationTypeChar IN(" & operation & ")"
            Dim gpQry As String = " GROUP BY os.OperationDate,ip.Name,ot.OperationName,os.TotalCount,os.FailCount ORDER BY os.OperationDate desc,ip.Name,ot.OperationName "
           
            
            ouputSB.Append("<table id=""list"" style=""width: 98%;"" summary=" & Copient.PhraseLib.Lookup("term.reports", LanguageID) & ">")
            ouputSB.Append("<thead>")
            ouputSB.Append("<tr>")
            ouputSB.Append("    <th id=""th1"" scope=""col"">" & Copient.PhraseLib.Lookup("term.partner", LanguageID) & "</th>")
            ouputSB.Append("    <th id=""th2"" scope=""col"">" & Copient.PhraseLib.Lookup("term.operation-type", LanguageID) & "</th>")
            If (frequency <> 1 AndAlso frequency <> 2) Then
                ouputSB.Append("    <th id=""th3""  scope=""col"" >" & Copient.PhraseLib.Lookup("term.operation", LanguageID) & " " & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
            End If
            ouputSB.Append("    <th id=""th4"" scope=""col"" >" & Copient.PhraseLib.Lookup("term.total", LanguageID) & " " & Copient.PhraseLib.Lookup("term.count", LanguageID) & "</th>")
            ouputSB.Append("    <th id=""th5"" scope=""col"" >" & Copient.PhraseLib.Lookup("term.fail", LanguageID) & " " & Copient.PhraseLib.Lookup("term.count", LanguageID) & "</th>")
            ouputSB.Append("  </tr>")
            ouputSB.Append("</thead>")
            Select Case frequency
                Case 0 'Days
                    MyCommon.QueryStr = query & joinQry & whereQry & gpQry
                    dt = MyCommon.L3P_Select
                    
                    Dim strBuf As New StringBuilder
                    ouputSB.Append("<tbody>")
                    If dt.Rows.Count > 0 Then
                        Session("RptDataTable") = dt
                        
                        ShowLimitNote = (dt.Rows.Count >= RECORD_LIMIT)
                        If (ShowLimitNote) Then
                            ouputSB.Insert(0, "<center><i>* " & Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & " </i></center>")
                        End If
                        For Each row In dt.Rows
                            strBuf.Append("  <tr>" & _
                          "    <td>" & MyCommon.NZ(row.Item("Name"), " ") & "</td>" & _
                          "    <td>" & MyCommon.NZ(row.Item("OperationName"), " ") & "</td>" & _
                           "   <td>" & MyCommon.NZ(row.Item("OperationDate"), " ") & "</td>" & _
                          "    <td>" & MyCommon.NZ(row.Item("TotalCount"), 0) & "</td>" & _
                          "    <td>" & MyCommon.NZ(row.Item("FailCount"), 0) & "</td>" & _
                          "  </tr>")
                        Next
                    Else
                        strBuf.Append("  <tr >" & _
                       "    <td colspan=""5"" align=""center""> " & Copient.PhraseLib.Lookup("reports.nodata", LanguageID) & "</td>" & _
                       "  </tr>")
                    End If
                    
                    ouputSB.Append(strBuf.ToString())
                    ouputSB.Append("</tbody>")
                    ouputSB.Append("</table>")
                    Sendb(ouputSB.ToString())
                  
                Case 1 'Week
                  
                    Dim weekStr As String = String.Empty
                    dt = RollupReportWeek(partner, operation, startdate, enddate, MyCommon, RECORD_LIMIT)
                    Dim strBuf As New StringBuilder
                    
                    Dim distinctWeeks As List(Of String) = dt.AsEnumerable() _
                                               .Select(Function(r) r.Field(Of String)("WeekOf")) _
                                               .Distinct() _
                                               .ToList()
                    ouputSB.Append("<tbody>")
                    If (dt.Rows.Count > 0 AndAlso distinctWeeks.Count > 0) Then
                        Session("RptDataTable") = dt
                        ShowLimitNote = (dt.Rows.Count >= RECORD_LIMIT)
                        If (ShowLimitNote) Then
                            ouputSB.Insert(0, "<center><i>* " & Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & " </i></center>")
                        End If
                        For Each weekOf As String In distinctWeeks
                            Dim result() As DataRow
                            result = dt.Select("WeekOf='" & weekOf & "'")
                            If (result.Count > 0) Then
                                strBuf.Append("  <tr class=""shaded""> <td colspan=""4"" align=""center"" style=""font-size:14px""> <b>" & Copient.PhraseLib.Lookup("term.week", LanguageID) & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & weekOf & "</b></td></tr>")
                                For Each row In result
                                    strBuf.Append("  <tr>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("Name"), " ") & "</td>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("OperationName"), " ") & "</td>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("TotalCount"), 0) & "</td>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("FailCount"), 0) & "</td>" & _
                                  "  </tr>")
                                Next
                            End If
                        Next
                    Else
                        strBuf.Append("  <tr >" & _
                       "    <td colspan=""4"" align=""center""> " & Copient.PhraseLib.Lookup("reports.nodata", LanguageID) & "</td>" & _
                       "  </tr>")
                    End If
                    ouputSB.Append(strBuf.ToString())
                    ouputSB.Append("</tbody>")
                    ouputSB.Append("</table>")
                    Sendb(ouputSB.ToString())
                    
                Case 2 'Month
                    query = "SELECT " & IIf(RECORD_LIMIT > 0, " TOP(" & RECORD_LIMIT & ")", "") & " ip.Name,ot.OperationName,SUM(os.TotalCount) As TotalCount ,SUM(os.FailCount) as FailCount,REPLACE(RIGHT(CONVERT(VARCHAR(11), os.OperationDate, 106), 8), ' ', '-') as MonthYear  "
                    gpQry = " GROUP BY DATEPART(Year, OperationDate), DATEPART(Month, OperationDate),REPLACE(RIGHT(CONVERT(VARCHAR(11), os.OperationDate, 106), 8), ' ', '-') ,ip.Name,ot.OperationName  " & _
                            " ORDER BY DATEPART(yyyy, OperationDate) desc, DATEPART(mm, OperationDate) desc,ip.Name,ot.OperationName "
                    MyCommon.QueryStr = query & " " & joinQry & whereQry & gpQry
                    dt = MyCommon.L3P_Select
                    
                    Dim strBuf As New StringBuilder
                    
                    Dim distinctMonths As List(Of String) = dt.AsEnumerable() _
                                               .Select(Function(r) r.Field(Of String)("MonthYear")) _
                                               .Distinct() _
                                               .ToList()
                    ouputSB.Append("<tbody>")
                    If (dt.Rows.Count > 0 AndAlso distinctMonths.Count > 0) Then
                        Session("RptDataTable") = dt
                        ShowLimitNote = (dt.Rows.Count >= RECORD_LIMIT)
                        If (ShowLimitNote) Then
                            ouputSB.Insert(0, "<center><i>* " & Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & " </i></center>")
                        End If
                        
                        For Each monthOf As String In distinctMonths
                            Dim result() As DataRow
                            result = dt.Select("MonthYear='" & monthOf & "'")
                            If (result.Count > 0) Then
                                strBuf.Append("  <tr class=""shaded""> <td colspan=""4"" align=""center"" style=""font-size:14px""> <b>" & Copient.PhraseLib.Lookup("term.month", LanguageID) & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & monthOf & "</b></td></tr>")
                                For Each row In result
                                    strBuf.Append("  <tr>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("Name"), " ") & "</td>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("OperationName"), " ") & "</td>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("TotalCount"), 0) & "</td>" & _
                                  "    <td>" & MyCommon.NZ(row.Item("FailCount"), 0) & "</td>" & _
                                  "  </tr>")
                                Next
                            End If
                        Next
                    Else
                        strBuf.Append("  <tr >" & _
                       "    <td colspan=""4"" align=""center""> " & Copient.PhraseLib.Lookup("reports.nodata", LanguageID) & "</td>" & _
                       "  </tr>")
                    End If
                    ouputSB.Append(strBuf.ToString())
                    ouputSB.Append("</tbody>")
                    ouputSB.Append("</table>")
                    Sendb(ouputSB.ToString())
                    
                Case 3 'Days of week
                    query &= " ,DATENAME(weekday, os.OperationDate) As DayOfWeek"
                    whereQry &= " AND DATENAME(weekday, os.OperationDate)  IN( " & weekdays & ")"
                    gpQry &= " ,DATENAME(weekday, os.OperationDate)"
                    MyCommon.QueryStr = query & joinQry & whereQry & gpQry
                    dt = MyCommon.L3P_Select
                    Dim strBuf As New StringBuilder
                    ouputSB.Append("<tbody>")
                    If dt.Rows.Count > 0 Then
                        Session("RptDataTable") = dt
                        
                        ShowLimitNote = (dt.Rows.Count >= RECORD_LIMIT)
                        If (ShowLimitNote) Then
                            ouputSB.Insert(0, "<center><i>* " & Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & " </i></center>")
                        End If
                        
                        For Each row In dt.Rows
                            strBuf.Append("  <tr>" & _
                          "    <td>" & MyCommon.NZ(row.Item("Name"), " ") & "</td>" & _
                          "    <td>" & MyCommon.NZ(row.Item("OperationName"), " ") & "</td>" & _
                          "    <td>" & MyCommon.NZ(row.Item("OperationDate"), " ") & " (" & Convert.ToDateTime(row.Item("OperationDate")).ToString("ddd") & ")" & "</td>" & _
                          "    <td>" & MyCommon.NZ(row.Item("TotalCount"), 0) & "</td>" & _
                          "    <td>" & MyCommon.NZ(row.Item("FailCount"), 0) & "</td>" & _
                          "  </tr>")
                        Next
                    Else
                        strBuf.Append("  <tr >" & _
                       "    <td colspan=""5"" align=""center""> " & Copient.PhraseLib.Lookup("reports.nodata", LanguageID) & "</td>" & _
                       "  </tr>")
                    End If
                        
                    ouputSB.Append(strBuf.ToString())
                    ouputSB.Append("</tbody>")
                    ouputSB.Append("</table>")
                    Sendb(ouputSB.ToString())
                Case Else
                    'default case
            End Select
            
          
        Catch ex As Exception
            Send(ex.ToString)
        Finally
            MyCommon.Close_Logix3P()
        End Try
       
    End Sub
    
    Function RollupReportWeek(ByVal partner As String, ByVal operations As String, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date, ByRef MyCommon As Copient.CommonInc, ByVal record_limit As Integer) As DataTable
        Dim dstWeek As DataTable = Nothing
        Dim dt As DataTable
        Dim CurrentStart As Date
        Dim CurrentEnd As Date
        Dim ReportWeeks As Integer
        Dim row As DataRow
      
        CurrentStart = ReportEndDate.AddDays(-6)
        CurrentEnd = ReportEndDate
        ReportWeeks = DateDiff(DateInterval.Day, ReportStartDate, ReportEndDate) / 7
      
        For i = 0 To ReportWeeks
            If (DateTime.Compare(ReportEndDate, CurrentStart) >= 0) Then
                Dim query As String = " SELECT  ip.Name,ot.OperationName, '" & CurrentStart.ToShortDateString() & "' as WeekOf,SUM(os.TotalCount) As TotalCount ,SUM(os.FailCount) as FailCount FROM ExternalRewards_OperationSummary as os " & _
                                    " with (NoLock) INNER JOIN ExternalRewards_InternalPartner ip with (NoLock) ON  ip.InternalPartnerId=os.InternalPartnerId INNER JOIN ExternalRewards_OperationType ot " & _
                                    " with (NoLock) ON os.OperationTypeChar=ot.OperationTypeChar   WHERE os.OperationDate BETWEEN '" & CurrentStart.ToShortDateString() & "' AND '" & CurrentEnd.ToShortDateString() & "' " & _
                                     "AND ip.InternalPartnerId IN (" & partner & ") AND os.OperationTypeChar IN(" & operations & ")" & _
                                    " GROUP BY ip.Name,ot.OperationName order by ip.Name,ot.OperationName "
                MyCommon.QueryStr = query
                dt = MyCommon.L3P_Select
                If (i = 0) Then
                    dstWeek=dt.Clone()
                End If
                If (dt.Rows.Count > 0) Then
                    For Each row In dt.Rows
                        dstWeek.ImportRow(row)
                    Next
                End If
            End If
            
            CurrentEnd = CurrentStart.AddDays(-1)
            CurrentStart = CurrentEnd.AddDays(-6)
        Next
        If (dstWeek.Rows.Count >= record_limit) Then
            dstWeek = dstWeek.AsEnumerable().Take(record_limit).CopyToDataTable()
        End If
        Return dstWeek
    End Function
</script>
