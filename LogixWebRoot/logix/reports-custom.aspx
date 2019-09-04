<%@ Page ValidateRequest="false" Language="vb" Debug="true" CodeFile="LogixCB.vb"
  Inherits="LogixCB" %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: status.aspx 
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
  ' * MODULE  : Preference Manager
  ' *
  ' * PURPOSE : 
  ' *
  ' * NOTES   : 
  ' *
  ' * Version : 1.0b1.0 
  ' *
  ' *****************************************************************************
%>
<script runat="server">
    Dim CopientFileVersion As String = "6.0.1.86599"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""

    Dim dtResult As New DataTable
    Dim Common As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
    Dim PageNum As Integer = 0
    Dim MorePages As Boolean
    Dim linesPerPage As Integer = 20
    Dim sizeOfData As Integer
    Dim i As Integer = 0
    Dim SortText As String = "OfferID"
    Dim SortDirection As String = "ASC"
    Dim NextSortDirection As String = "DESC"
    Dim ShowExpired As String = "FALSE"
    Dim ShowReportList As Boolean = False
    Dim Shaded As String = "shaded"
    Dim OCDate As New DateTime
    Dim OfferIds As String = ""
    Dim OfferIdOptions As String = ""

    Dim RptTypeFromFolder As String = ""
    Dim RptStartDateFromFolder As String = ""
    Dim RptEndDateFromFolder As String = ""
    Dim Impressions As String = ""
    Dim ImpressionType As String = ""
    Dim Redemptions As String = ""
    Dim RedemptionType As String = ""
    Dim Transactions As String = ""
    Dim TransactionType As String = ""
    Dim MarkDowns As String = ""
    Dim MarkDownType As String
    Dim ReportingStartDate As String = ""
    Dim ReportingType As String = ""
    Dim ReportingEndDate As String = ""
    Dim DisplayRptEnd As String = "display:none;"
    Dim StyleDownloadBtn As String = "visibility:hidden;"
    Dim ShowLimitNote As Boolean = False
    Dim RECORD_LIMIT As Integer = 0
    Dim RecordLimitValue As String = ""
    Dim DefaultToEnhancedCustomReport As Integer = 0
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False

    Dim ImpressionChecked As Boolean = False
    Dim DisplayHint As String = ""
    Dim ColspanHint As String = ""
    Dim LogFile As String = "CustomReport." & Format(Now(), "yyyyMMdd") & ".txt"

    Const MaxBigIntLength As Integer = 19 'OfferID for OfferReporting table is big int, max=9223372036854775808 or 19 digits. SQL will allow up to 38 digits
    '-------------------------------------------------------------------------------------------------------------  

    Sub GenerateReportBox()

        Send("<div id=""loading"" style=""display:none;"">")
        Send("  <div id=""loadingwrap"" style=""width:420px;"">")
        Send("    <div class=""box"" id=""loadingbox"" style=""height:380px;"">")
        Send("      <h2><span>Custom Reports</span></h2>")
        Send("      <div id=""collisionsContent"" style=""height:325px;overflow-y:auto;width:100%;"">")
        Send("        <p>Report is being generated<p>")
        Send("        <p style=""text-align:center;padding-top:80px;""><img src=""../images/loadingAnimation.gif"" alt=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ /></p>")
        Send("      </div>")
        Send("    </div>")
        Send("  </div>")
        Send("</div>")

    End Sub

    '------------------------------------------------------------------------------------------------------------- 
    ' From reports-detail below
    Public DefaultLanguageID
    ' Generate enhanced report for downloading
    Sub GenerateReport(ByVal ReportID As String, OptList As String())
        Dim builder As StringBuilder = New StringBuilder()
        Dim bParsed As Boolean = False
        Dim LanguageID As Integer = 1
        Dim Frequency As String = 1

        If (Request.Form("lang") <> "") Then
            bParsed = Integer.TryParse(Request.Form("lang"), LanguageID)
            If (Not bParsed) Then LanguageID = 1
        End If

        If (Request.Form("frequency") <> "") Then
            Frequency = Request.Form("frequency")
        End If
        If DefaultToEnhancedCustomReport = 1 Then
            builder.Append(CreateMeijerDefaultReport(LanguageID, OptList))
        Else
            builder.Append(CreateReport(LanguageID, OptList))
        End If
        Response.Write(builder)
    End Sub

    ' Create enhanced report for downloading
    Function CreateReport(ByVal LanguageID As Integer, OptList As String()) As String
        Dim ReportStartDate As Date
        Dim ReportEndDate As Date
        Dim ReportWeeks As Integer
        Dim RowCount As Integer
        Dim CumulativeImpress As Integer
        Dim CumulativeRedeem As Integer
        Dim CumulativeAmtRedeem As Double
        Dim RedemptionRate As Double
        Dim AmtRedeem As Double
        Dim Redemptions As Integer
        Dim Impressions As Integer
        Dim i As Integer
        Dim OfferID As String = ""
        Dim dst As System.Data.DataTable
        Dim bParsed As Boolean
        Dim builder As StringBuilder = New StringBuilder()
        Dim frequency As String = ""

        Dim WhereClause As New StringBuilder("")
        ' buiding the report header
        builder.Append("OfferID")
        builder.Append(",")
        If (Request.Form("hiddenImpression") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenCumulativeImpressions") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.cumulativeimpressions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenRedemption") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.redemptions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenCumulativeRedemptions") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.cumulativeredemptions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenMarkdown") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.markdowns", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenCumulativeMarkdowns") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.cumulativemarkdowns", LanguageID))
            builder.Append(",")
        End If
        builder.Append("Reporting Date")
        builder.Append(vbNewLine)

        'If Request.Form("impressionChecked") = "true" Then 
        If (Request.Form("Impressions") <> "" AndAlso Request.Form("ImpressionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumImpressions" & ConvertToOperand(Request.Form("ImpressionType")) & Request.Form("Impressions"))
            'WhereClause.Append("NumImpressions" & Request.Form("ImpressionType") & Request.Form("Impressions"))
        End If
        'End If
        'If Request.Form("redemptionChecked") = "true" Then   
        If (Request.Form("Redemptions") <> "" AndAlso Request.Form("RedemptionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumRedemptions" & ConvertToOperand(Request.Form("RedemptionType")) & Request.Form("Redemptions"))
            'WhereClause.Append("NumRedemptions" & Request.Form("RedemptionType") & Request.Form("Redemptions"))
        End If
        'End If
        'If Request.Form("markdownChecked") = "true" Then    
        If (Request.Form("MarkDowns") <> "" AndAlso Request.Form("MarkDownType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("AmountRedeemed" & ConvertToOperand(Request.Form("MarkDownType")) & Request.Form("MarkDowns"))
            'WhereClause.Append("AmountRedeemed" & Request.Form("MarkDownType") & Request.Form("MarkDowns"))
        End If
        'End If  

        If WhereClause.Length > 0 Then
            WhereClause.Append(" and OfferID = ")
        Else
            WhereClause.Append("where OfferID = ")
        End If

        If (Request.Form("hiddenReportingType") = "5") Then
            bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate2"), ReportEndDate)
            If (Not bParsed) Then ReportEndDate = Now()
        ElseIf (Request.Form("hiddenReportingType") = "1") Then
            bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportEndDate)
            If (Not bParsed) Then
                ReportEndDate = Now()
            End If
            ReportStartDate = ReportEndDate.AddDays(-30)
        ElseIf (Request.Form("hiddenReportingType") = "0") Then
            bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportEndDate)
            If (Not bParsed) Then
                ReportEndDate = Now()
            End If
            ReportEndDate = ReportEndDate.Date.AddDays(-1)
            ReportStartDate = ReportEndDate.Date.AddDays(-30)
            ReportEndDate = ReportEndDate.AddTicks(-1)
        ElseIf (Request.Form("hiddenReportingType") = "3") Then
            bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            ReportEndDate = Now()
        ElseIf (Request.Form("hiddenReportingType") = "4") Then
            bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
            If (Not bParsed) Then
                ReportStartDate = Now()
            End If
            ReportStartDate = ReportStartDate.Date.AddDays(1)
            ReportEndDate = Now()
        ElseIf Request.Form("hiddenReportingType") = "2"
            bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            ReportStartDate = ReportStartDate.Date
            ReportEndDate = ReportStartDate.AddDays(1).AddTicks(-1)
        End If

        Common.Open_LogixWH()
        For count = 0 To OptList.Length - 1
            If (OptList(count)<> "") Then
                OfferID = Long.Parse(OptList(count))

                Common.QueryStr = "select NumImpressions, NumRedemptions, NumTransactions, AmountRedeemed, ReportingDate from OfferReporting with (nolock) " & _
                                WhereClause.ToString & OfferID & " " & _
                                "and ReportingDate between '" & ReportStartDate.ToString & "' and '" & ReportEndDate.ToString & "' " & _
                                "order by ReportingDate"
                dst = Common.LWH_Select
                If dst.Rows.Count > 0 Then
                    If (Request.Form("frequency") = "1") Then
                        dst = RollupReportWeek(dst, ReportStartDate, ReportEndDate)
                        frequency = "weekly"
                    ElseIf (Request.Form("frequency") = "2") Then
                        dst = FillInDays(dst, ReportStartDate, ReportEndDate)
                        frequency = "daily"
                    End If

                    'builder.Append(ExportNewCustomReport(dst, frequency, OfferID))
                    builder.Append(ExportCustomReport(dst, frequency, OfferID))
                End If
            End If
        Next
        Common.Close_LogixWH()

        Return builder.ToString
    End Function
    '----------------------------------------------------------------------------------
    ' AMSPS-2009 Create Meijer Default enhanced report for downloading
    Function CreateMeijerDefaultReport(ByVal LanguageID As Integer, OptList As String()) As String
        Dim ReportStartDate As Date
        Dim ReportEndDate As Date
        Dim ReportWeeks As Integer
        Dim RowCount As Integer
        Dim CumulativeImpress As Integer
        Dim CumulativeRedeem As Integer
        Dim CumulativeAmtRedeem As Double
        Dim RedemptionRate As Double
        Dim AmtRedeem As Double
        Dim Redemptions As Integer
        Dim Impressions As Integer
        Dim i As Integer
        Dim OfferID As String = ""
        Dim dst As System.Data.DataTable
        Dim bParsed As Boolean
        Dim builder As StringBuilder = New StringBuilder()
        Dim frequency As String = ""

        Dim WhereClause As New StringBuilder("")
        ' buiding the report header
        builder.Append("OfferID")
        builder.Append(",")
        If (Request.Form("hiddenImpression") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenCumulativeImpressions") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.cumulativeimpressions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenRedemption") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.redemptions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenCumulativeRedemptions") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.cumulativeredemptions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenTransaction") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.transactions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenCumulativeTransactions") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.cumulativetransactions", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenMarkdown") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.markdowns", LanguageID))
            builder.Append(",")
        End If
        If (Request.Form("hiddenCumulativeMarkdowns") = "true") Then
            builder.Append(Copient.PhraseLib.Lookup("term.cumulativemarkdowns", LanguageID))
            builder.Append(",")
        End If
        builder.Append("Reporting Date")
        builder.Append(vbNewLine)

        'If Request.Form("impressionChecked") = "true" Then 
        If (Request.Form("Impressions") <> "" AndAlso Request.Form("ImpressionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumImpressions" & ConvertToOperand(Request.Form("ImpressionType")) & Request.Form("Impressions"))
            'WhereClause.Append("NumImpressions" & Request.Form("ImpressionType") & Request.Form("Impressions"))
        End If
        'End If
        'If Request.Form("redemptionChecked") = "true" Then   
        If (Request.Form("Redemptions") <> "" AndAlso Request.Form("RedemptionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumRedemptions" & ConvertToOperand(Request.Form("RedemptionType")) & Request.Form("Redemptions"))
            'WhereClause.Append("NumRedemptions" & Request.Form("RedemptionType") & Request.Form("Redemptions"))
        End If
        'End If
        If (Request.Form("Transactions") <> "" AndAlso Request.Form("TransactionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumTransactions" & ConvertToOperand(Request.Form("TransactionType")) & Request.Form("Transactions"))
        End If
        'If Request.Form("markdownChecked") = "true" Then    
        If (Request.Form("MarkDowns") <> "" AndAlso Request.Form("MarkDownType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("AmountRedeemed" & ConvertToOperand(Request.Form("MarkDownType")) & Request.Form("MarkDowns"))
            'WhereClause.Append("AmountRedeemed" & Request.Form("MarkDownType") & Request.Form("MarkDowns"))
        End If
        'End If  

        If WhereClause.Length > 0 Then
            WhereClause.Append(" and OfferID = ")
        Else
            WhereClause.Append("where OfferID = ")
        End If

        If Request.Form("reportingDate1") = "" And Request.Form("reportingDate2") = "" Then
            bParsed = DateTime.TryParse( GetReportStartingDateByOfferList(OptList), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            ReportEndDate = Now()
        Else
            If (Request.Form("hiddenReportingType") = "5") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then ReportStartDate = Now()
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate2"), ReportEndDate)
                If (Not bParsed) Then ReportEndDate = Now()
                ReportEndDate = ReportEndDate.AddDays(1).AddTicks(-1)
            ElseIf (Request.Form("hiddenReportingType") = "1") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportEndDate)
                If (Not bParsed) Then
                    ReportEndDate = Now()
                End If
                ReportEndDate = ReportEndDate.AddDays(1).AddTicks(-1)
                ReportStartDate = ReportEndDate.AddDays(-30)
            ElseIf (Request.Form("hiddenReportingType") = "0") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportEndDate)
                If (Not bParsed) Then
                    ReportEndDate = Now()
                End If
                ReportEndDate = ReportEndDate.Date
                ReportStartDate = ReportEndDate.Date.AddDays(-30)
                ReportEndDate = ReportEndDate.AddTicks(-1)
            ElseIf (Request.Form("hiddenReportingType") = "3") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then ReportStartDate = Now()
                ReportEndDate = Now()
            ElseIf (Request.Form("hiddenReportingType") = "4") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then
                    ReportStartDate = Now()
                End If
                ReportStartDate = ReportStartDate.Date.AddDays(1)
                ReportEndDate = Now()
            Else   'Request.Form("hiddenReportingType") = "2"
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then ReportStartDate = Now()
                ReportStartDate = ReportStartDate.Date
                ReportEndDate = ReportStartDate.AddDays(1).AddTicks(-1)
            End If
        End If

        Common.Open_LogixWH()
        For count = 0 To OptList.Length - 1
            If (OptList(count)<> "") Then
                OfferID = Long.Parse(OptList(count))

                Common.QueryStr = "select NumImpressions, NumRedemptions, NumTransactions, AmountRedeemed, CAST(ReportingDate as DATE) as ReportingDate from OfferReporting with (nolock) " & _
                                WhereClause.ToString & OfferID & " " & _
                                "and ReportingDate between '" & ReportStartDate.ToString & "' and '" & ReportEndDate.ToString & "' " & _
                                "order by ReportingDate"
                dst = Common.LWH_Select
                If dst.Rows.Count > 0 Then
                    If (Request.Form("frequency") = "1") Then
                        dst = RollupReportWeek(dst, ReportStartDate, ReportEndDate)
                        frequency = "weekly"
                    ElseIf (Request.Form("frequency") = "2") Then
                        dst = FillInDays(dst, ReportStartDate, ReportEndDate)
                        frequency = "daily"
                    End If

                    builder.Append(ExportNewCustomReport(dst, frequency, OfferID))
                End If
            End If
        Next
        Common.Close_LogixWH()

        Return builder.ToString
    End Function

    '----------------------------------------------------------------------------------

    Function RollupReportWeek(ByVal dst As DataTable, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date) As DataTable
        Dim dstWeek As New DataTable
        Dim i, j As Integer
        Dim numRedeem As Integer
        Dim numImpression As Integer
        Dim numTransaction As Integer
        Dim amtRedeem As Double

        If (dst.Rows.Count > 0) Then
            Dim CurrentStart As Date
            Dim CurrentEnd As Date
            Dim ReportWeeks As Integer
            Dim row As DataRow
            Dim rowCt As Integer

            dstWeek = dst.Copy()
            dstWeek.Clear()

            CurrentStart = ReportStartDate
            CurrentEnd = ReportStartDate.AddDays(6)
            ReportWeeks = DateDiff(DateInterval.Day, ReportStartDate, ReportEndDate) / 7

            For i = 0 To ReportWeeks
                If (DateTime.Compare(ReportEndDate, CurrentStart) >= 0) Then
                    dst.DefaultView.RowFilter = "ReportingDate >= '" & CurrentStart.ToString() & "' and ReportingDate <= '" & CurrentEnd.ToString() & "'"
                    rowCt = dst.DefaultView.Count
                    If (rowCt > 0) Then
                        For j = 0 To rowCt - 1
                            numRedeem += dst.DefaultView(j).Item("NumRedemptions")
                            amtRedeem += dst.DefaultView(j).Item("AmountRedeemed")
                            numImpression += dst.DefaultView(j).Item("NumImpressions")
                            numTransaction += dst.DefaultView(j).Item("NumTransactions")
                            If (j = dst.DefaultView.Count - 1) Then
                                row = dst.DefaultView(j).Row
                                row.Item("ReportingDate") = CurrentStart
                                row.Item("NumRedemptions") = numRedeem
                                row.Item("AmountRedeemed") = amtRedeem
                                row.Item("NumImpressions") = numImpression
                                row.Item("NumTransactions") = numTransaction
                                dstWeek.ImportRow(row)
                            End If
                        Next
                    Else
                        row = dstWeek.NewRow()
                        row.Item("ReportingDate") = CurrentStart
                        row.Item("NumRedemptions") = 0
                        row.Item("AmountRedeemed") = 0.0
                        row.Item("NumImpressions") = 0
                        row.Item("NumTransactions") = 0
                        dstWeek.Rows.Add(row)
                    End If
                    numRedeem = 0
                    amtRedeem = 0.0
                    numImpression = 0
                    numTransaction = 0
                    CurrentStart = CurrentEnd.AddDays(1)
                    CurrentEnd = CurrentStart.AddDays(6)
                End If
            Next
        End If

        Return dstWeek
    End Function

    '----------------------------------------------------------------------------------

    Function FillInDays(ByVal dst As DataTable, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date) As DataTable
        Dim dstDay As New DataTable
        Dim CurrentDate As Date
        Dim RptDate As Date
        Dim row As DataRow

        dstDay = dst.Copy()
        dstDay.Clear()

        CurrentDate = ReportStartDate
        RptDate = ReportStartDate

        For Each row In dst.Rows
            RptDate = row.Item("ReportingDate")
            If (CurrentDate < RptDate) Then
                AddEmptyDays(dstDay, CurrentDate, RptDate)
                dstDay.ImportRow(row)
            Else
                dstDay.ImportRow(row)
            End If
            CurrentDate = RptDate.AddDays(1)
        Next

        If (ReportEndDate > RptDate) Then
            If (RptDate = ReportStartDate) Then
                AddEmptyDays(dstDay, RptDate, ReportEndDate.Date.AddDays(1).AddTicks(-1))
            Else
                AddEmptyDays(dstDay, RptDate.AddDays(1), ReportEndDate.Date.AddDays(1).AddTicks(-1))
            End If
        End If

        Return dstDay
    End Function

    '----------------------------------------------------------------------------------

    Sub AddEmptyDays(ByRef dst As DataTable, ByVal StartDate As Date, ByVal EndDate As Date)
        Dim CurrentDate As Date
        Dim row As DataRow
        CurrentDate = StartDate
        While (CurrentDate < EndDate)
            row = dst.NewRow()
            row.Item("ReportingDate") = CurrentDate
            row.Item("NumRedemptions") = 0
            row.Item("NumTransactions") = 0
            row.Item("NumImpressions") = 0
            row.Item("AmountRedeemed") = 0.0
            dst.Rows.Add(row)
            CurrentDate = CurrentDate.AddDays(1)
        End While
    End Sub

    '----------------------------------------------------------------------------------

    Function ExportCustomReport(ByVal dst As DataTable, ByVal frequency As String, ByVal OfferID As String) As String
        Dim builder As StringBuilder = New StringBuilder()

        If (Not dst Is Nothing) Then
            builder.Append("OfferID=")
            builder.Append(OfferID)
            builder.Append(",")
            builder.Append(WriteExportRow(dst, "ReportingDate", False))
            If (Request.Form("hiddenImpression") = "true") Then
                If (frequency.Contains("weekly")) Then
                    builder.Append("Impressions (weekly),")
                Else
                    builder.Append("Impressions (daily),")
                End If
                builder.Append(WriteExportRow(dst, "NumImpressions", False))
            End If
            If (Request.Form("hiddenCumulativeImpressions") = "true") Then
                builder.Append("Impressions (cumulative),")
                builder.Append(WriteExportRow(dst, "NumImpressions", True))
            End If
            If (Request.Form("hiddenRedemption") = "true") Then
                If (frequency.Contains("weekly")) Then
                    builder.Append("Redemptions (weekly),")
                Else
                    builder.Append("Redemptions (daily),")
                End If
                builder.Append(WriteExportRow(dst, "NumRedemptions", False))
            End If
            If (Request.Form("hiddenCumulativeredemptions") = "true") Then
                builder.Append("Redemptions (cumulative),")
                builder.Append(WriteExportRow(dst, "NumRedemptions", True))
            End If
            If (Request.Form("hiddenMarkdown") = "true") Then
                If (frequency.Contains("weekly")) Then
                    builder.Append("Mark Downs ($) (weekly),")
                Else
                    builder.Append("Mark Downs ($) (daily),")
                End If
                builder.Append(WriteExportRow(dst, "AmountRedeemed", False))
            End If
            If (Request.Form("hiddenCumulativeMarkdowns") = "true") Then
                builder.Append("Mark Downs ($) (cumulative),")
                builder.Append(WriteExportRow(dst, "AmountRedeemed", True))
            End If
            ' builder.Append("Redemption Rate,")
            ' builder.Append(WriteRedemptionRow(dst))
        End If

        Return builder.ToString
    End Function

    '----------------------------------------------------------------------------------

    Function WriteExportRow(ByVal dst As DataTable, ByVal field As String, ByVal bCumulative As Boolean) As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim RowCount, i As Integer
        Dim cumulative As Double
        Dim dt As Date

        RowCount = dst.Rows.Count

        For i = 0 To (RowCount - 1)
            If (field = "" OrElse IsDBNull(dst.Rows(i).Item(field))) Then
                If field = "AmountRedeemed" And bCumulative = True Then
                    builder.Append(cumulative)
                Else
                    builder.Append("0")
                End If
            Else
                If (bCumulative) Then
                    cumulative += dst.Rows(i).Item(field)
                    builder.Append(cumulative)
                Else
                    If (IsDate(dst.Rows(i).Item(field))) Then
                        dt = dst.Rows(i).Item(field)
                        builder.Append(dt.ToString("M/dd/yyyy"))
                    Else
                        builder.Append(dst.Rows(i).Item(field))
                    End If
                End If
            End If
            If (i = (RowCount - 1)) Then
                builder.Append(vbNewLine)
            Else
                builder.Append(",")
            End If
        Next
        Return builder.ToString()
    End Function
    '----------------------------------------------------------------------------------
    ' AMSPS-2009   
    Function ExportNewCustomReport(ByVal dst As DataTable, ByVal frequency As String, ByVal OfferID As String) As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim RowCount, i As Integer
        Dim cumulative As Double
        Dim dt As Date

        Dim cumuImpression As Integer
        Dim cumuRedeem As Integer
        Dim cumuTransaction As Integer
        Dim cumuAmtRedeem As Double

        RowCount = dst.Rows.Count

        ' building the report body
        For i = 0 To (RowCount - 1)
            builder.Append(OfferID)
            builder.Append(",")
            If (Request.Form("hiddenImpression") = "true") Then
                builder.Append(dst.Rows(i).Item("NumImpressions"))
                builder.Append(",")
            End If

            If (Request.Form("hiddenCumulativeImpressions") = "true") Then
                cumuImpression += dst.Rows(i).Item("NumImpressions")
                builder.Append(cumuImpression )
                builder.Append(",")
            End If

            If (Request.Form("hiddenRedemption") = "true") Then
                builder.Append(dst.Rows(i).Item("NumRedemptions"))
                builder.Append(",")
            End If

            If (Request.Form("hiddenCumulativeRedemptions") = "true") Then
                cumuRedeem += dst.Rows(i).Item("NumRedemptions")
                builder.Append(cumuRedeem)
                builder.Append(",")
            End If

            If (Request.Form("hiddenTransaction") = "true") Then
                builder.Append(dst.Rows(i).Item("NumTransactions"))
                builder.Append(",")
            End If

            If (Request.Form("hiddenCumulativeTransactions") = "true") Then
                cumuTransaction += dst.Rows(i).Item("NumTransactions")
                builder.Append(cumuTransaction)
                builder.Append(",")
            End If

            If (Request.Form("hiddenMarkdown") = "true") Then
                'builder.Append(dst.Rows(i).Item("AmountRedeemed"))
                If dst.Rows(i).Item("AmountRedeemed") <> 0.0 Then
                    builder.Append(Format(dst.Rows(i).Item("AmountRedeemed"), "0.00"))
                Else
                    builder.Append("0")
                End If
                builder.Append(",")
            End If

            If (Request.Form("hiddenCumulativeMarkdowns") = "true") Then
                cumuAmtRedeem += dst.Rows(i).Item("AmountRedeemed")
                'builder.Append(cumuAmtRedeem)
                If cumuAmtRedeem <> 0.0 Then
                    builder.Append(Format(cumuAmtRedeem, "0.00"))
                Else
                    builder.Append("0")
                End If
                builder.Append(",")
            End If

            dt = dst.Rows(i).Item("ReportingDate")
            builder.Append(dt.ToString("M/dd/yyyy"))
            builder.Append(vbNewLine)
        Next
        Return builder.ToString()
    End Function
    '----------------------------------------------------------------------------------

    Function WriteRedemptionRow(ByVal dst As DataTable) As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim RowCount, i As Integer
        Dim Impressions, Redemptions As Integer
        Dim RedemptionRate As Double

        RowCount = dst.Rows.Count

        For i = 0 To (RowCount - 1)
            Impressions = Common.NZ(dst.Rows(i).Item("NumImpressions"), 0)
            Redemptions = Common.NZ(dst.Rows(i).Item("NumRedemptions"), 0)
            If (Impressions > 0) Then
                RedemptionRate = Redemptions / Impressions
            Else
                RedemptionRate = 0.0
            End If
            builder.Append(RedemptionRate.ToString("0.####"))
            If (i = (RowCount - 1)) Then
                builder.Append(vbNewLine)
            Else
                builder.Append(",")
            End If
        Next

        Return builder.ToString()
    End Function
    ' From reports-detail above 
    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_Page_Links()
        Send("<style> ")
        Send("  #intro { ")
        Send("    top:65px;")
        Send("    height:10px;")
        Send("  }")
        Send("  * html #intro {")
        Send("    position: absolute;")
        Send("    height: 10px;")
        Send("    top: 65px;")
        Send("  }")
        Send("  #main { ")
        Send("    top:85px;")
        Send("  }")
        Send("  * html #main {")
        Send("    position: absolute;")
        Send("    top: 87px;")
        Send("    height: 107%;")
        Send("  }")
        Send("</style>")
    End Sub
    '-------------------------------------------------------------------------------------------------------------
    ' AMSPS-2009
    Function GetReportStartingDate(Byval OfferIDs As String) As String
        Dim RptStartDate As Date = Now()
        Dim dt As System.Data.DataTable
        Dim Count As Integer
        Dim OptList As String() = OfferIDs.Split(",")

        Common.Open_LogixRT()

        For count = 0 To OptList.Length - 1
            If (OptList(count) <> "") Then
                Common.QueryStr = "SELECT ProdStartDate FROM CM_ST_Offers WITH (NOLOCK) WHERE Deleted=0 AND IsTemplate=0 AND OfferID=" & OptList(count)
                dt = Common.LRT_Select
                If dt.Rows.Count > 0 Then
                    If RptStartDate > Common.NZ(dt.Rows(0).Item("ProdStartDate"), Now().ToShortDateString()) Then
                        RptStartDate = Common.NZ(dt.Rows(0).Item("ProdStartDate"), Now().ToShortDateString())
                    End If
                End If
            End If
        Next
        Common.Close_LogixRT()
        return RptStartDate.ToShortDateString()
    End Function
    '-------------------------------------------------------------------------------------------------------------
    ' AMSPS-2009
    Function GetReportStartingDateByOfferList(Byval OfferList As String()) As String
        Dim RptStartDate As Date = Now()
        Dim dt As System.Data.DataTable
        Dim Count As Integer

        Common.Open_LogixRT()
        For count = 0 To OfferList.Length - 1
            If (OfferList(count) <> "") Then
                Common.QueryStr = "Select ProdStartDate from CM_ST_Offers WITH (NOLOCK) Where Deleted=0 and IsTemplate=0 and OfferID=" & OfferList(count)
                dt = Common.LRT_Select
                If dt.Rows.Count > 0 Then
                    If RptStartDate > Common.NZ(dt.Rows(0).Item("ProdStartDate"), Now().ToShortDateString()) Then
                        RptStartDate = Common.NZ(dt.Rows(0).Item("ProdStartDate"), Now().ToShortDateString())
                    End If
                End If
            End If
        Next
        Common.Close_LogixRT()
        return RptStartDate.ToShortDateString()
    End Function
    '-------------------------------------------------------------------------------------------------------------  
    Public Sub Open_Criteria_Box(Optional ByVal BoxTitle As String = "", Optional ByVal ExtraTitleText As String = "", Optional ByVal BoxWidth As String = "")

        Dim dst As DataTable
        Dim BoxObjectName As String = ""
        Dim BoxOpen As Integer = 1
        Dim ValidBoxID As Boolean = False
        Dim WidthStr As String = ""
        Dim BoxID As Integer = 1

        If Not (BoxWidth = "") Then WidthStr = "style=""width: " & BoxWidth & ";"""
        Send("<div class=""box"" style=""margin-bottom: 0px; "" id=""" & BoxObjectName & "box"" " & WidthStr & ">")
        Send("  <div style=""position: relative;float: left;""><font size=""3"" color=""white""><b>" & BoxTitle & "</b></font>" & ExtraTitleText & "</div>")
        Send("  <div class=""resizer"" style=""position: relative;"">")
        Send("    <a href="""" onclick=""resizeDiv('" & BoxObjectName & "body','img" & BoxObjectName & "body','" & BoxTitle & "', '" & BoxID & "', '" & AdminUserID & "'); return false;"">")
        If BoxOpen = 1 Then
            Send("    <img id=""img" & BoxObjectName & "body"" src=""/images/arrowup-off.png"" alt=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ title=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ onmouseover = ""handleResizeHover(true,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" onmouseout=""handleResizeHover(false,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" />")
        Else
            Send("    <img id=""img" & BoxObjectName & "body"" src=""/images/arrowdown-off.png"" alt=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ title=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ onmouseover = ""handleResizeHover(true,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" onmouseout=""handleResizeHover(false,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" />")
        End If
        Send("    </a>")
        Send("  </div> <!-- resizer -->")
        Send("  <br clear=""all"" />")
        If BoxOpen = 1 Then
            Send("  <div id=""" & BoxObjectName & "body"">")
            Send_Criteria_Form()
        Else
            Send("  <div id=""" & BoxObjectName & "body""  style=""display: none;"">")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------

    Private Sub Send_Criteria_Form()
        Send("    <form action=""#"" name=""mainform"" id=""mainform"" method=""post"">")
        Send("    <table style= ""border:0px; border-collapse: collapse; border-spacing:5px"" width=""100%"" cellspacing=""0"" cellpadding=""0"" style=""width: 670px;"" summary=""")
        Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID))
        Send(""" >")
        Send("      <tr colspan=""all"" >&nbsp</tr>")
        Send("      <tr>")
        Send("<td style=""width:5%;"" />")
        Send("        <td style=""width:50%; padding:0px; border-width:5px; margin:0px;"">")
        Send("          <table style= ""border:0px; border-collapse: collapse; border-spacing:5px"" width=""100%"" cellspacing=""0"" celloadding=""0"" >")
        Send("            <tr>")
        Send("              <td style=""width:30%; padding:0px; border-width:5px; margin:0px;"">")
        Send("                <label for=""offerIds""><b>&nbsp;&nbsp;&nbsp;")
        Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID))
        Send(":  </b></label>")
        Send("              </td>")
        Send("              <td style=""width:20%; padding:0px; border-width:5px; margin:0px;"">")
        Send("              <select id=""offerIds"" name=""offerIds""  multiple=""multiple"" style=""height: 90px; width: 75px;"" size=""4"" >")
        If Request.Form("hdnOfferIds") <> "" Then
            Dim count As Integer
            Dim OptList As String() = Request.Form("hdnOfferIds").Split(",")
            For count = 0 To OptList.Length - 1
                If (OptList(count) <> "") Then
                    Send(" <option value=""")
                    Sendb(OptList(count).Trim())
                    Send(""">")
                    Sendb(OptList(count).Trim())
                    Send("</opption>")
                End If
            Next
        End If
        Sendb(OfferIdOptions)
        Send("              </select>")
        Send("              </td>")
        Send("              <td style=""width:40%"">")
        Send("                &nbsp;<input type=""button"" class=""regular"" style=""width:90px;"" id=""btnAdd"" name=""btnAdd"" value=""")
        Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID))
        Send(""" onclick=""addOffer(); updateCheckboxState() "" /><br />")
        Send("                <br />")
        Send("                &nbsp;<input type=""button"" class=""regular"" style=""width:90px;"" id=""btnRemove"" name=""btnRemove"" value=""")
        Sendb(Copient.PhraseLib.Lookup("term.remove", LanguageID))
        Send(""" onclick=""removeOffer();"" />")
        Send("              </td>")
        Send("            </tr>")

        Send("            <tr><td colspan=""3""/></tr>")
        Send("            <tr>")
        Send("              <td rowspan=""1"" style=""width: 140px; padding:0px; border-width:5px; margin:0px;"">&nbsp;&nbsp;")
        Send("                <input type=""checkbox"" id=""impression"" name=""impression"" ")
        If Request.Form("hiddenImpression") = "true" Then
            Send(" checked ")
        Else
            Send(" ")
        End If
        Send(" onclick=""updateCheckboxState()"" /><label for=""impressions"">")
        Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
        Send("</label>")
        Send("              </td>")
        Send("              <td rowspan=""1"" style=""width: 75px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                <select id=""impressionType"" name=""impressionType"" style=""width: 75px;"">")
        Send("                <option value=""1""")
        If Request.Form("hiddenImpressionType") = "0" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(1, ImpressionType))
        Send(">&lt;</option>")
        Send("                <option value=""2""")
        If Request.Form("hiddenImpressionType") = "1" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(2, ImpressionType))
        Send(">&lt;=</option>")
        Send("                <option value=""3""")
        If Request.Form("hiddenImpressionType") = "2" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(3, ImpressionType))
        Send(">=</option>")
        Send("                <option value=""4""")
        If Request.Form("hiddenImpressionType") = "3" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(4, ImpressionType))
        Send(">&gt;=</option>")
        Send("                <option value=""5""")
        If Request.Form("hiddenImpressionType") = "4" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(5, ImpressionType))
        Send(">&gt;</option>")
        Send("                </select>")
        Send("              </td>")
        Send("              <td rowspan=""1"" style=""width: 80px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                &nbsp;&nbsp;<input type=""text"" id=""impressions"" name=""impressions"" value=""")
        Impressions = Request.Form("hiddenImpressions")
        Sendb(Impressions)
        Send(""" style=""width:80px;"" />")
        Send("              </td>")
        Send("            </tr>")
        Send("            <tr>")
        Send("              <td rowspan=""1"" style=""width: 140px; padding:0px; border-width:5px; margin:0px;"">&nbsp;&nbsp;")
        Send("                <input type=""checkbox"" id=""redemption"" name=""redemption"" ")
        If DefaultToEnhancedCustomReport = 1 Then
            If Request.Form("hiddenRedemption") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
        Else
            If Request.Form("hiddenRedemption") = "true" Then
                Send(" checked ")
            Else
                Send(" ")
            End If
        End If
        Send(" onclick=""updateCheckboxState()"" /><label for=""redemptions"">")
        Sendb(Copient.PhraseLib.Lookup("term.redemptions", LanguageID))
        Send("</label>")
        Send("              </td>")
        Send("              <td rowspan=""1"" style=""width: 75px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                <select name=""redemptionType"" id=""redemptionType"" style=""width: 75px;"">")
        Send("                <option value=""1""")
        If Request.Form("hiddenRedemptionType") = "0" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(1, RedemptionType))
        Send(""">&lt;</option>")
        Send("                <option value=""2""")
        If Request.Form("hiddenRedemptionType") = "1" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(2, RedemptionType))
        Send(""">&lt;=</option>")
        Send("                <option value=""3""")
        If Request.Form("hiddenRedemptionType") = "2" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(3, RedemptionType))
        Send(""">=</option>")
        Send("                <option value=""4""")
        If Request.Form("hiddenRedemptionType") = "3" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(4, RedemptionType))
        Send(""">&gt;=</option>")
        Send("                <option value=""5""")
        If Request.Form("hiddenRedemptionType") = "4" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(5, RedemptionType))
        Send(""">&gt;</option>")
        Send("                </select>")
        Send("              </td>")
        Send("              <td rowspan=""1"" style=""width: 80px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                &nbsp;&nbsp;<input type=""text"" name=""redemptions"" id=""redemptions"" value=""")
        Redemptions = Request.Form("hiddenRedemptions")
        Sendb(Redemptions)
        Send(""" style=""width:80px;"" />")
        Send("              </td>")
        Send("            </tr>")
        If DefaultToEnhancedCustomReport = 1 Then
            Send("            <tr>")
            Send("              <td rowspan=""1"" style=""width: 140px; padding:0px; border-width:5px; margin:0px;"">&nbsp;&nbsp;")
            Send("                <input type=""checkbox"" id=""transaction"" name=""transaction"" ")
            If Request.Form("hiddenTransaction") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
            Send(" onclick=""updateCheckboxState()"" />")
            Send("<label for=""transactions"">")
            Sendb(Copient.PhraseLib.Lookup("term.transactions", LanguageID))
            Send("</label>")
            Send("              </td>")
            Send("              <td rowspan=""1"" style=""width: 75px; padding:0px; border-width:5px; margin:0px;"">")
            Send("                <select name=""transactionType"" id=""transactionType"" style=""width: 75px;"">")
            Send("                <option value=""1""")
            If Request.Form("hiddenTransactionType") = "0" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(1, TransactionType))
            Send(""">&lt;</option>")
            Send("                <option value=""2""")
            If Request.Form("hiddenTransactionType") = "1" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(2, TransactionType))
            Send(""">&lt;=</option>")
            Send("                <option value=""3""")
            If Request.Form("hiddenTransactionType") = "2" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(3, TransactionType))
            Send(""">=</option>")
            Send("                <option value=""4""")
            If Request.Form("hiddenTransactionType") = "3" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(4, TransactionType))
            Send(""">&gt;=</option>")
            Send("                <option value=""5""")
            If Request.Form("hiddenTransactionType") = "4" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(5, TransactionType))
            Send(""">&gt;</option>")
            Send("                </select>")
            Send("              </td>")
            Send("              <td rowspan=""1"" style=""width: 80px; padding:0px; border-width:5px; margin:0px;"">")
            Send("                &nbsp;&nbsp;<input type=""text"" name=""transactions"" id=""transactions"" value=""")
            Transactions = Request.Form("hiddenTransactions")
            Sendb(Transactions)
            Send(""" style=""width:80px;"" />")
            Send("              </td>")
            Send("            </tr>")
        Else
            ' Dummy versions that aren't displayed, needed so code can function normally
            Send("                <input type=""checkbox"" id=""transaction"" name=""transaction"" style=""display:none"" onclick=""updateCheckboxState()"" />")
            Send("                <select name=""transactionType"" id=""transactionType"" style=""display:none"" >")
            Send("                <option value=""1""")
            If Request.Form("hiddenTransactionType") = "0" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(1, TransactionType))
            Send(""">&lt;</option>")
            Send("                <option value=""2""")
            If Request.Form("hiddenTransactionType") = "1" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(2, TransactionType))
            Send(""">&lt;=</option>")
            Send("                <option value=""3""")
            If Request.Form("hiddenTransactionType") = "2" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(3, TransactionType))
            Send(""">=</option>")
            Send("                <option value=""4""")
            If Request.Form("hiddenTransactionType") = "3" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(4, TransactionType))
            Send(""">&gt;=</option>")
            Send("                <option value=""5""")
            If Request.Form("hiddenTransactionType") = "4" Then
                Send(" selected ")
            End If
            Sendb(SetAsSelected(5, TransactionType))
            Send(""">&gt;</option>")
            Send("                <input type=""text"" name=""transactions"" id=""transactions"" value="""" style=""display:none"" />")
        End If
        Send("            <tr>")
        Send("              <td rowspan=""1"" style=""width: 140px; padding:0px; border-width:5px; margin:0px;"">&nbsp;&nbsp;")
        Send("                <input type=""checkbox"" id=""markdown"" name=""markdown"" ")
        If DefaultToEnhancedCustomReport = 1 Then
            If Request.Form("hiddenMarkdown") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
        Else
            If Request.Form("hiddenMarkdown") = "true" Then
                Send(" checked ")
            Else
                Send(" ")
            End If
        End If
        Send(" onclick=""updateCheckboxState()"" /><label for=""markDowns"">")
        Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID))
        Send("</label>")
        Send("              </td>")
        Send("              <td rowspan=""1"" style=""width: 75px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                <select name=""markDownType"" id=""markDownType"" style=""width: 75px;"">")
        Send("                <option value=""1""")
        If Request.Form("hiddenMarkdownType") = "0" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(1, MarkDownType))
        Send(">&lt;</option>")
        Send("                <option value=""2""")
        If Request.Form("hiddenMarkdownType") = "1" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(2, MarkDownType))
        Send(">&lt;=</option>")
        Send("                <option value=""3""")
        If Request.Form("hiddenMarkdownType") = "2" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(3, MarkDownType))
        Send(">=</option>")
        Send("                <option value=""4""")
        If Request.Form("hiddenMarkdownType") = "3" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(4, MarkDownType))
        Send(">&gt;=</option>")
        Send("                <option value=""5""")
        If Request.Form("hiddenMarkdownType") = "4" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(5, MarkDownType))
        Send(">&gt;</option>")
        Send("                </select>")
        Send("              </td>")
        Send("              <td rowspan=""1"" style=""width: 80px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                &nbsp;&nbsp;<input type=""text"" name=""markDowns"" id=""markDowns"" value=""")
        MarkDowns = Request.Form("hiddenMarkdowns")
        Sendb(MarkDowns)
        Send(""" style=""width:80px;"" />")
        Send("              </td>")
        Send("            </tr>")
        Send("            <tr>")
        Send("              <td rowspan=""1"" style=""width: 140px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                <label for=""reportingDate1"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
        Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))
        Send("</label>")
        Send("              </td>")
        Send("              <td rowspan=""1"" style=""width: 75px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                <select name=""reportingType"" id=""reportingType"" onChange=""updateSelectboxState();"" style=""width: 75px;"" optionSelected=""Betwwn""")

        Send(""" onchange=""toggleDate2(this);"">")
        Send("                <option value=""1""")
        If Request.Form("hiddenReportingType") = "0" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(1, ReportingType))
        Send(">&lt;</option>")
        Send("                <option value=""2""")
        If Request.Form("hiddenReportingType") = "1" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(2, ReportingType))
        Send(">&lt;=</option>")
        Send("                <option value=""3""")
        If Request.Form("hiddenReportingType") = "2" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(3, ReportingType))
        Send(">=</option>")
        Send("                <option value=""4""")
        If Request.Form("hiddenReportingType") = "3" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(4, ReportingType))
        Send(">&gt;=</option>")
        Send("                <option value=""5""")
        If Request.Form("hiddenReportingType") = "4" Then
            Send(" selected ")
        End If
        Sendb(SetAsSelected(5, ReportingType))
        Send(">&gt;</option>")
        Send("                <option value=""6""")
        If Request.Form("hiddenReportingType") = "5" OR RptTypeFromFolder = "5" Then
            Send(" selected ")
            DisplayRptEnd = "display:inline;"
            RptTypeFromFolder = ""
        End If
        Sendb(SetAsSelected(6, ReportingType))
        Send(">")
        Sendb(Copient.PhraseLib.Lookup("term.between", LanguageID))
        Send("</option>")
        Send("                </select>")
        Send("              </td>")
        Send("                <td rowspan=""1"" style=""width: 80px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                  &nbsp;&nbsp;<input type=""text"" name=""reportingDate1"" id=""reportingDate1""")
        Send(" value=""")
        If RptStartDateFromFolder <> "" Then
            ReportingStartDate = RptStartDateFromFolder
            RptStartDateFromFolder = ""
        Else
            ReportingStartDate = Request.Form("hiddenReportingDate1")
        End If
        Sendb(ReportingStartDate)
        Send(""" style=""width:80px;""/>")
        Send("                </td>")
        Send("              </tr>")
        If (Request.Form("hiddenReportingType") = "0" OR Request.Form("hiddenReportingType") = "1") AND Request.Form("hiddenReportingDate1") <> "" AND Request.Form("hiddenEnhancedReport") = "true" then
            DisplayHint = "Reports start 30 days back."
            ColspanHint = "colspan='2'"
        Else
            DisplayHint = ""
            ColspanHint = ""
        End If
        Send("              <tr>")
        Send("                <td rowspan=""1"" style=""width: 140px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                </td>")
        Send("                <td rowspan=""1""")
        Sendb(ColspanHint)
        Send(" style=""width: 75px; padding:0px; border-width:5px; margin:0px;"">")
        Sendb(DisplayHint)
        Send("                </td>")
        Send("                <td rowspan=""1"" style=""width: 80px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                  &nbsp;&nbsp;<input type=""text"" name=""reportingDate2"" id=""reportingDate2"" style=""")
        Sendb(DisplayRptEnd)
        Send("; width:80px;"" value=""")
        If RptEndDateFromFolder <> "" Then
            ReportingEndDate = RptEndDateFromFolder
            RptEndDateFromFolder = ""
        Else
            ReportingEndDate = Request.Form("hiddenReportingDate2")
        End If
        Sendb(ReportingEndDate)
        Send(""" />")
        Send("                </td>")
        Send("              </tr>")
        Send("            </table>")
        Send("      </td>")

        Send("          <td style=""width:35%"">")
        Send("<br/>")
        Send("            <table style= ""border:1px solid black; border-collapse: collapse; border-spacing:10px"" cellspacing=""0"" celloadding=""0"">")
        Send("              <tr>")
        Send("                <td colspan=""3"" style=""padding:3px; border-width:5px; margin:0px; "">")
        Send("                  &nbsp;&nbsp;<input type=""checkbox"" id=""EnhancedReport"" name=""EnhancedReport"" type=""checkbox"" ")
        If DefaultToEnhancedCustomReport = 1 Then
            If Request.Form("hiddenEnhancedReport") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
        Else
            If Request.Form("hiddenEnhancedReport") = "true" Then
                Send(" checked ")
            Else
                Send(" ")
            End If
        End If
        Send(" onclick=""EnableEnhancedOptions()"" /><label for=""Enhanced Reporting"" id=""SelectEnhancedreporting"">")
        Sendb(Copient.PhraseLib.Lookup("term.enhancedreporting", LanguageID))
        Send("</label>")
        Send("                </td>")
        Send("              </tr>")
        Send("              <tr>")
        Send("                <td style=""width: 0px; padding:0px; border-width:5px; margin:0px;""></td>")
        Send("                <td style=""width: 100px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                  &nbsp;&nbsp;<input type=""radio"" id=""byweek"" name=""byweek"" value=""false"" onclick=""ByWeekClick();"" ")
        If Request.Form("hiddenByweek") = "true" Then
            Send(" checked ")
        Else
            Send(" ")
        End If
        Send(" > " & Copient.PhraseLib.Lookup("term.byWeek", LanguageID) & "")
        Send("                </td>")
        Send("                <td style=""width: 100px; padding:0px; border-width:5px; margin:0px;"">")
        Send("                  <input type=""radio"" id=""byday"" name=""byday"" value=""false"" onclick=""ByDayClick();"" ")
        If DefaultToEnhancedCustomReport = 1 Then
            If Request.Form("hiddenByday") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
        Else
            If Request.Form("hiddenByday") = "true" Then
                Send(" checked ")
            Else
                Send(" ")
            End If
        End If
        Send(" > " & Copient.PhraseLib.Lookup("term.byday", LanguageID) & "")
        Send("                </td>")
        Send("              </tr>")
        Send("              <tr>")
        Send("                <td style=""width: 0px; padding:0px; border-width:5px; margin:0px;""></td>")
        Send("                <td colspan=""2"" style=""padding:0px; border-width:5px; margin:0px;"">")
        Send("                  &nbsp;&nbsp;<input type=""checkbox"" id=""CumulativeImpressions"" name=""CumulativeImpressions"" type=""checkbox"" ")
        If Request.Form("hiddenCumulativeImpressions") = "true" Then
            Send(" checked ")
        Else
            Send(" ")
        End If
        Send(" onclick=""updateCheckboxState()"" /><label for=""CumulImpressions"" id=""CumulImpressions"">")
        Sendb(Copient.PhraseLib.Lookup("term.cumulativeimpressions", LanguageID))
        Send("</label>")
        Send("                </td>")
        Send("              </tr>")
        Send("              <tr>")
        Send("                <td style=""width: 0px; padding:0px; border-width:5px; margin:0px;""></td>")
        Send("                <td colspan=""2"" style=""padding:0px; border-width:5px; margin:0px;"">")
        Send("                  &nbsp;&nbsp;<input type=""checkbox"" id=""CumulativeRedemptions"" name=""CumulativeRedemptions"" type=""checkbox"" ")
        If DefaultToEnhancedCustomReport = 1 Then
            If Request.Form("hiddenCumulativeRedemptions") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
        Else
            If Request.Form("hiddenCumulativeRedemptions") = "true" Then
                Send(" checked ")
            Else
                Send(" ")
            End If
        End If
        Send(" onclick=""updateCheckboxState()"" /><label for=""CumulRedemptions"" id=""CumulRedemptions"">")
        Sendb(Copient.PhraseLib.Lookup("term.cumulativeredemptions", LanguageID))
        Send("</label>")
        Send("                </td>")
        Send("              </tr>")
        If DefaultToEnhancedCustomReport = 1 Then
            Send("              <tr>")
            Send("                <td style=""width: 0px; padding:0px; border-width:5px; margin:0px;""></td>")
            Send("                <td colspan=""2"" style=""padding:0px; border-width:5px; margin:0px;"">")
            Send("                  &nbsp;&nbsp;<input type=""checkbox"" id=""CumulativeTransactions"" name=""CumulativeTransactions"" type=""checkbox"" ")
            If Request.Form("hiddenCumulativeTransactions") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
            Send(" onclick=""updateCheckboxState()"" />")
            Send("<label for=""CumulTransactions"" id=""CumulTransactions"">")
            Sendb(Copient.PhraseLib.Lookup("term.cumulativetransactions", LanguageID))
            Send("</label>")
            Send("                </td>")
            Send("              </tr>")
        Else
            Send("<input type=""checkbox"" id=""CumulativeTransactions"" name=""CumulativeTransactions"" type=""checkbox"" style=""display:none"" onclick=""updateCheckboxState()"" />")
        End If
        Send("              <tr>")
        Send("                <td style=""width: 10px; padding:0px; border-width:5px; margin:0px;""></td>")
        Send("                <td colspan=""2"" style=""padding:0px; border-width:5px; margin:0px;"">")
        Send("                  &nbsp;&nbsp;<input type=""checkbox"" id=""CumulativeMarkdowns"" name=""CumulativeMarkdowns"" type=""checkbox"" ")
        If DefaultToEnhancedCustomReport = 1 Then
            If Request.Form("hiddenCumulativeMarkdowns") = "false" Then
                Send(" ")
            Else
                Send(" checked ")
            End If
        Else
            If Request.Form("hiddenCumulativeMarkdowns") = "true" Then
                Send(" checked ")
            Else
                Send(" ")
            End If
        End If
        Send(" onclick=""updateCheckboxState()"" /><label for=""CumulMarkdowns"" id=""CumulMarkdowns"">")
        Sendb(Copient.PhraseLib.Lookup("term.cumulativemarkdowns", LanguageID))
        Send("</label>")
        Send("              </tr>")
        Send("              <tr><td colspan=""3"">&nbsp;</td></tr>")
        Send("            </table>")

        Send("<br/>")
        Send("<br/>")
        Send("    <table>")
        Send("        <tr>")
        Send("          <td width=""15""></td>")
        Send("          <td colspan=""1"" rowspan=""1"" style=""height: 21px; vertical-align: bottom;"">")
        Send("            <input type=""button"" class=""regular"" id=""btnClear""  name=""btnClear"" value=""")
        Send(Copient.PhraseLib.Lookup("term.clear", LanguageID))
        Send("""onclick=""clearForm();"" />")
        Send("            <input type=""button"" class=""regular"" id=""generateReport"" name=""generateReport"" value=""")
        Sendb(Copient.PhraseLib.Lookup("term.GetResults", LanguageID))
        Send(""" onclick=""submitForm();""")
        If Request.Form("OfferIds") <> "" OR OfferIdOptions <> "" OR Request.Form("hdnOfferIds") <> "" Then
            Send(" ")
        Else
            Send(" disabled=""disabled"" ")
        End If
        Send(" />")
        Send("          </td>")
        Send("        </tr>")
        Send("    </table>")
        Send("  </td>")

        Send("<td style=""width:5%;"" />")
        Send("        </tr>")

        Send("  <div id=""datepicker"" class=""dpDiv"">")
        Send("  </div>")
        If Request.Browser.Type = "IE6" Then
            Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
        Send("        <input type=""hidden"" id=""Reports"" name=""Reports"" value=""1"" />")
        Send("        <input type=""hidden"" id=""lang"" name=""lang"" value=""")
        Sendb(LanguageID)
        Send(""" />")
        Send("        <input type=""hidden"" name=""frequency"" id=""frequency"" value="""" />")
        Send("        <input type=""hidden"" name=""offerID"" id=""offerID"" value=""55"" />")
        ' from reports-detail.aspx above   

        Send("        <input type=""hidden"" name=""exportRpt"" id=""exportRpt"" value=""0"" />")

        Send("        <input type=""hidden"" name=""sortText"" id=""sortText"" value=""")
        Sendb(SortText)
        Send(""" />")
        Send("        <input type=""hidden"" name=""sortDir"" id=""sortDir"" value=""")
        Sendb(SortDirection)
        Send(""" />")
        Send("        <input type=""hidden"" name=""hdnOfferIDs"" id=""hdnOfferIDs"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenImpression"" id=""hiddenImpression"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenImpressionType"" id=""hiddenImpressionType"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenImpressions"" id=""hiddenImpressions"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenRedemption"" id=""hiddenRedemption"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenRedemptionType"" id=""hiddenRedemptionType"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenRedemptions"" id=""hiddenRedemptions"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenTransaction"" id=""hiddenTransaction"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenTransactionType"" id=""hiddenTransactionType"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenTransactions"" id=""hiddenTransactions"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenMarkdown"" id=""hiddenMarkdown"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenMarkdownType"" id=""hiddenMarkdownType"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenMarkdowns"" id=""hiddenMarkdowns"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenReportingType"" id=""hiddenReportingType"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenReportingDate1"" id=""hiddenReportingDate1"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenReportingDate2"" id=""hiddenReportingDate2"" value="""" />")
        Send("        <input type=""hidden"" name=""hideDownloadButton"" id=""hideDownloadButton"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenEnhancedReport"" id=""hiddenEnhancedReport"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenByweek"" id=""hiddenByweek"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenByday"" id=""hiddenByday"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenCumulativeImpressions"" id=""hiddenCumulativeImpressions"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenCumulativeRedemptions"" id=""hiddenCumulativeRedemptions"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenCumulativeTransactions"" id=""hiddenCumulativeTransactions"" value="""" />")
        Send("        <input type=""hidden"" name=""hiddenCumulativeMarkdowns"" id=""hiddenCumulativeMarkdowns"" value="""" />")

        Send("      </table>")
        Send("  </form>")
    End Sub
    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_Report_Criteria_Box()

        Open_Criteria_Box("Criteria", "")

        Close_UI_Box()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_Page()

        Send("<div id=""intro"">")
        Send("  <h1 id=""title"">")
        Sendb(Copient.PhraseLib.Lookup("term.customreports", LanguageID))
        Send("  </h1>")
        Send("  <div id=""controls"">")
        Send("    <input type=""button"" class=""regular"" id=""download"" name=""download""  value=""")
        Sendb(Copient.PhraseLib.Lookup("term.download", LanguageID))
        Send(""" style="" margin-right:15px; ")
        If Request.Form("hiddenEnhancedReport") = "false" OR (Request.Form("hiddenEnhancedReport") = "true" AND DefaultToEnhancedCustomReport=1) Then
            StyleDownloadBtn = Request.Form("hideDownloadButton")
        End If
        Sendb(StyleDownloadBtn)
        Send(""" title=""")
        Sendb(Copient.PhraseLib.Lookup("report-custom-downloadreport", LanguageID))
        Send(""" onclick=""handleDownload();"" />")
        Send("  </div>")
        Send("</div>")
        Send("<br/>")
        Send("<br/>")
        Send("<br/>")
        Send("<br/>")
        Send("<div id=""main"">")
        If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
        Send("<div class=""gutter""></div>")

        Send("<div id=""columnfull"" style=""padding-left:0.5%; width:730px;"">")
        Send_Report_Criteria_Box()

        If Request.Form("hiddenEnhancedReport") = "true" Then
            If DefaultToEnhancedCustomReport = 1 Then
                Open_New_Enhanced_Result_Box()
            Else
                Open_Enhanced_Result_Box()
            End If
        Else
            Dim validation = ValidateForm()
            If (validation.Length = 0) Then
                Open_Result_Box()
            Else
                Send("<div style=""visibility: visible;color: #cc0000; width:720px; margin-top: 10px; "">")

                Send(" <center>" + validation + "</center></div>")
            End If
        End If

        Send("</div> <!-- columnfull -->")

        Send("</div> <!-- main -->")

    End Sub

    Function ValidateForm() As String
        Dim ParseReturn As Long = 0
        Dim ParseReturnDecimal As Decimal = 0
        Impressions = Request.Form("impressions")
        Redemptions = Request.Form("redemptions")
        Transactions = Request.Form("transactions")
        MarkDowns = Request.Form("markDowns")
        If (Impressions <> "") Then
            If (Long.TryParse(Impressions, ParseReturn) <> True) Then
                Return Copient.PhraseLib.Lookup("term.impressions", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-9223372036854775807)"
            End If
        End If


        If (Redemptions <> "") Then
            If (Long.TryParse(Redemptions, ParseReturn) <> True) Then
                Return Copient.PhraseLib.Lookup("term.redemptions", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-9223372036854775807)"
            End If
        End If

        If (Transactions <> "") Then
            If (Long.TryParse(Transactions, ParseReturn) <> True) Then
                Return Copient.PhraseLib.Lookup("term.transactions", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-9223372036854775807)"
            End If
        End If

        If (MarkDowns <> "") Then
            If (Decimal.TryParse(MarkDowns, ParseReturnDecimal) <> True) Then
                Return Copient.PhraseLib.Lookup("term.markdowns", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-99999999999999999999999999999999999999)"
            End If
        End If
        Return ""
    End Function
    '------------------------------------------------------------------------------------------------------------- 
    Function ConvertToOperand(ByVal code As String) As String
        Dim operand As String = ""
        Select Case code
            Case 1
                operand = " < "
            Case 2
                operand = " <= "
            Case 3
                operand = " = "
            Case 4
                operand = " >= "
            Case 5
                operand = " > "
            Case 6
                operand = " between "
            Case Else
                operand = " = "
        End Select
        Return operand
    End Function

    Function SetAsSelected(ByVal code As String, ByVal value As String) As String
        Dim selected As String = ""
        If (code = value) Then
            selected = " selected=""selected"""
        End If
        Return selected
    End Function

    '-----------------------------------------------------------------------------------------
    Sub Open_Enhanced_Result_Box()

        Dim dt As DataTable
        Dim ReportStartDate As String = ""
        Dim ReportEndDate As String = ""
        Dim ProdStartDate As String = ""
        Dim ProdEndDate As String = ""
        Dim OfferID As Long = -1
        Dim EngineID As Long = -1
        Dim OfferName As String = ""
        Dim Status As String = ""
        Dim ShowReport As Boolean = False
        Dim bParsed As Boolean = False
        Dim dtStart As Date
        Dim dtEnd As Date
        Dim OfferLink As String = "#"
        Dim TempDate As Date

        Dim OptList As String()
        Dim ToTop As integer = 24
        Dim Interval As integer = 20
        Dim DivBoxHeight As Integer = 159    '175
        Dim SpanHeight As Integer = 16
        Dim Shaded As String = ""

        ' calculating DivBox height below
        If Request.Form("hiddenImpression") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("hiddenRedemption") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("hiddenMarkdown") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("hiddenCumulativeImpressions") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("hiddenCumulativeRedemptions") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        If Request.Form("hiddenCumulativeMarkdowns") = "false" Then
            DivBoxHeight = DivBoxHeight - SpanHeight
        End If
        ' calculating DivBox height above

        OfferID = Common.Extract_Val(Request.QueryString("OfferID"))

        'AMSPS-2231
        '    OfferIds = Request.Form("hdnOfferIDs")
        OfferIds = SortOfferIdsInString(Request.Form("hdnOfferIDs"), True)
        'AMSPS-2231 above
        OfferIds = Replace(OfferIds, vbCrLf, "")
        OfferIds = Replace(OfferIds, vbCr, "")
        If OfferIds = "" Then
            Return
        End If
        Impressions = Request.Form("impressions")
        ImpressionType = Request.Form("impressionType")
        Redemptions = Request.Form("redemptions")
        RedemptionType = Request.Form("redemptionType")
        MarkDowns = Request.Form("markDowns")
        MarkDownType = Request.Form("markDownType")
        ReportingType = Request.Form("reportingType")

        ReportStartDate = Request.Form("reportingDate1")

        If Date.TryParse(ReportStartDate, Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
            ReportStartDate = Logix.ToShortDateString(TempDate, Common)
        End If

        ReportEndDate = Request.Form("reportingDate2")

        If Date.TryParse(ReportEndDate, Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, TempDate) Then
            ReportEndDate = Logix.ToShortDateString(TempDate, Common)
        End If

        Send("<div class=""box"" id=""enhancedResult"" style=""visibility: visible; width:720px; margin-top: 10px; "">")
        Send("  <h2>")
        Send("    <span id=""resultTitle"" style=""visibility: visible;"">")
        Sendb("Enhanced Result")
        Send("    </span>")
        Send("  </h2>")

        OptList = OfferIds.Split(",")

        Send("<div id='rptContainer' style=""overflow: hidden;float: left; display:none;"">")
        For count = 0 To OptList.Length - 1
            If (OptList(count) <> "") Then
                OfferID = Long.Parse(OptList(count))

                If OfferID > 0 Then
                    Common.QueryStr = "select IncentiveName as Name from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & _
                                      " union " & _
                                      "select Name from Offers with (NoLock) where OfferID=" & OfferID & ";"
                    dt = Common.LRT_Select
                    If dt.Rows.Count > 0 Then
                        OfferName = Common.NZ(dt.Rows(0).Item("Name"), "")
                    End If
                End If

                bParsed = DateTime.TryParse(ReportStartDate, Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, dtStart)
                If (bParsed) Then
                    ProdStartDate = Logix.ToShortDateString(dtStart, Common)
                Else
                    ProdStartDate = StrConv(Copient.PhraseLib.Lookup("term.never", LanguageID), VbStrConv.Lowercase)
                End If

                bParsed = DateTime.TryParse(ReportEndDate, Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, dtEnd)
                If (bParsed) Then
                    ProdEndDate = Logix.ToShortDateString(dtEnd, Common)
                Else
                    ProdEndDate = StrConv(Copient.PhraseLib.Lookup("term.never", LanguageID), VbStrConv.Lowercase)
                End If

                Send("<div class=""box "" id=""rptHeader"" style=""margin-top:0px; margin-left: 20px; height:")
                Sendb(DivBoxHeight.ToString())
                Send("px; width: 220px;"" >")
                Send("  <h2>")
                Send("    <span id=""reportTitle"" style=""visibility: visible;line-height:20px; font-size:12px;"">")
                Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))
                Send("&nbsp;for offer ID:&nbsp;")
                Sendb(OfferID)
                Send("    </span>")
                Send("  </h2>")


                If (Request.Form("hiddenImpression") = "true") Then
                    Send("  <span id=""rowHdr1"" style=""top: ")
                    Sendb(ToTop.toString())
                    Send("px;"" class=""reportRowHeader")
                    Sendb(Shaded)
                    If DivBoxHeight > 79 then
                        ToTop = ToTop + Interval
                        If(Shaded = "" ) Then
                            Shaded = " shaded"
                        Else
                            Shaded = ""
                        End If
                    Else
                        ToTop = 24
                        Shaded = ""
                    End If
                    Send(""" ondblclick=""javascript:toggleHighlight(1);"">")
                    Send("   <a href=""javascript:launchGraph(1, ")
                    Sendb(OptList(count))
                    Send(");"">")
                    Send("      <img src=""../images/graph.png"" alt=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" title=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" onmouseover=""this.src='../images/graph-on.png';"" onmouseout=""this.src='../images/graph.png';"" />")
                    Send("    </a>&nbsp;&nbsp;&nbsp;")
                    Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
                    Send("&nbsp;&nbsp;<span id=""row1"" >")
                    If (Request.Form("frequency") = "1") Then
                        Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID))
                    Else
                        Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID))
                    End If
                    Send("  </span>")
                    Send("  </span>")
                End If

                If (Request.Form("hiddenCumulativeImpressions") = "true") Then
                    Send("  <span id=""rowHdr2"" style=""top: ")
                    Sendb(ToTop.toString())
                    Send("px;"" class=""reportRowHeader")
                    Sendb(Shaded)
                    If ((DivBoxHeight = 95 and ToTop = 44) or (DivBoxHeight = 79 and ToTop = 24)) then
                        ToTop = 24
                        Shaded = ""
                    Else
                        ToTop = ToTop + Interval
                        If Shaded = "" Then
                            Shaded = " shaded"
                        Else
                            Shaded = ""
                        End If
                    End If
                    Send(""" ondblclick=""javascript:toggleHighlight(2);"">")
                    Send("    <a href=""javascript:launchGraph(2, ")
                    Sendb(OptList(count))
                    Send(");"">")
                    Send("      <img src=""../images/graph.png"" alt=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" title=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" onmouseover=""this.src='../images/graph-on.png';"" onmouseout=""this.src='../images/graph.png';"" />")
                    Send("    </a>&nbsp;&nbsp;&nbsp;")
                    Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")")
                    Send("  </span>")
                End If

                If (Request.Form("hiddenRedemption") = "true") Then
                    Send("  <span id=""rowHdr3"" style=""top: ")
                    Sendb(ToTop.toString())
                    Send("px;"" class=""reportRowHeader")
                    Sendb(Shaded)
                    If ((DivBoxHeight = 111 and ToTop = 64) or (DivBoxHeight = 95 and ToTop = 44) or (DivBoxHeight = 79 and ToTop = 24)) then
                        ToTop = 24
                        Shaded = ""
                    Else
                        ToTop = ToTop + Interval
                        If Shaded = "" Then
                            Shaded = " shaded"
                        Else
                            Shaded = ""
                        End If
                    End If
                    Send(""" ondblclick=""javascript:toggleHighlight(3);"">")
                    Send("    <a href=""javascript:launchGraph(3, ")
                    Sendb(OptList(count))
                    Send(");"">")
                    Send("      <img src=""../images/graph.png"" alt=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" title=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" onmouseover=""this.src='../images/graph-on.png';"" onmouseout=""this.src='../images/graph.png';"" />")
                    Send("</a>&nbsp;&nbsp;&nbsp;")
                    Sendb(Copient.PhraseLib.Lookup("term.redemption", LanguageID))
                    Send("&nbsp;&nbsp;<span id=""row2"" >")
                    If (Request.Form("frequency") = "1") Then
                        Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID))
                    Else
                        Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID))
                    End If
                    Send("  </span>")
                    Send("  </span>")
                End If

                If (Request.Form("hiddenCumulativeRedemptions") = "true") Then
                    Send("  <span id=""rowHdr4"" style=""top: ")
                    Sendb(ToTop.toString())
                    Send("px;"" class=""reportRowHeader")
                    Sendb(Shaded)
                    If ((DivBoxHeight = 127 and ToTop = 84) or (DivBoxHeight = 111 and ToTop = 64) or (DivBoxHeight = 95 and ToTop = 44) or (DivBoxHeight = 79 and ToTop = 24)) then
                        ToTop = 24
                        Shaded = ""
                    Else
                        ToTop = ToTop + Interval
                        If Shaded = "" Then
                            Shaded = " shaded"
                        Else
                            Shaded = ""
                        End If
                    End If
                    Send(""" ondblclick=""javascript:toggleHighlight(4);"">")
                    Send("    <a href=""javascript:launchGraph(4, ")
                    Sendb(OptList(count))
                    Send(");"">")
                    Send("      <img src=""../images/graph.png"" alt=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" title=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" onmouseover=""this.src='../images/graph-on.png';"" onmouseout=""this.src='../images/graph.png';"" />")
                    Send("    </a>&nbsp;&nbsp;&nbsp;")
                    Sendb(Copient.PhraseLib.Lookup("term.redemption", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")")
                    Send("  </span>")
                End If

                If (Request.Form("hiddenMarkdown") = "true") Then
                    Send("  <span id=""rowHdr5"" style=""top: ")
                    Sendb(ToTop.toString())
                    Send("px;"" class=""reportRowHeader")
                    Sendb(Shaded)
                    If ((DivBoxHeight = 143 and ToTop = 104) or (DivBoxHeight = 127 and ToTop = 84) or (DivBoxHeight = 111 and ToTop = 64) or (DivBoxHeight = 95 and ToTop = 44) or (DivBoxHeight = 79 and ToTop = 24)) then
                        ToTop = 24
                        Shaded = ""
                    Else
                        ToTop = ToTop + Interval
                        If Shaded = "" Then
                            Shaded = " shaded"
                        Else
                            Shaded = ""
                        End If
                    End If
                    Send(""" ondblclick=""javascript:toggleHighlight(5);"">")
                    Send("    <a href=""javascript:launchGraph(5, ")
                    Sendb(OptList(count))
                    Send(");"">")
                    Send("      <img src=""../images/graph.png"" alt=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" title=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" onmouseover=""this.src='../images/graph-on.png';"" onmouseout=""this.src='../images/graph.png';"" />")
                    Send("    </a>&nbsp;&nbsp;&nbsp;")
                    Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID))
                    Send("&nbsp;&nbsp;<span id=""row3"" >")
                    If (Request.Form("frequency") = "1") Then
                        Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID))
                    Else
                        Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID))
                    End If
                    Send("  </span>")
                    Send("  </span>")
                End If

                If (Request.Form("hiddenCumulativeMarkdowns") = "true") Then
                    Send("  <span id=""rowHdr6"" style=""top: ")
                    Sendb(ToTop.toString())

                    Send("px;"" class=""reportRowHeader")
                    Sendb(Shaded)
                    ToTop = 24
                    Shaded = ""
                    Send(""" ondblclick=""javascript:toggleHighlight(6);"">")
                    Send("    <a href=""javascript:launchGraph(6, OptList(count));"">")
                    Send("      <img src=""../images/graph.png"" alt=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" title=""")
                    Sendb(Copient.PhraseLib.Lookup("reports.viewgraph", LanguageID))
                    Send(""" onmouseover=""this.src='../images/graph-on.png';"" onmouseout=""this.src='../images/graph.png';"" />")
                    Send("    </a>&nbsp;&nbsp;&nbsp;")
                    Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")")
                    Send("  </span>")
                End If

                Send("</div>")

            End If

        Next
        Send("</div>")

        Send("<div id=""report"" style=""margin-top: 0px; margin-right: 20px; overflow: hidden;top:-5px;position:relative;left:-2px;"">")
        '    Send("<div id=""report"">")
        Send("</div>")
        Send("<div style=""clear:both""></div>")

        ' Send("<div class=""box"" id=""enhancedResult"" style=""visibility: hidden; "">") 
        Send_Scripts(New String() {"datePicker.js"})
        Send("<hr class=""hidden"" />")
        Send("<br />")
        Send("<div id=""nota"" style=""display: none;"">")
        Send("</div>")
        '    Send("<div class=""box"" id=""wait"" style=""visibility: hidden;"">")
        Send("<div class=""box"" id=""wait"" style=""display: none;"">")
        Send("</div>")


        Send("</div>")  '-- <div id="enhancedResult"
        Send("<br clear=""all"" />")
        Send("<!-- <img src=""reports-graph.aspx"" alt=""graph"" /> -->")
        '   Send("</div>")

        If (Request.Form("exportRpt") = "1") Then
            Response.ClearHeaders()
            Response.AddHeader("Content-Disposition", "attachment; filename=Offer" & OfferID & "_Rpt.csv")
            Response.ContentType = "application/octet-stream"
            Response.Clear()
            GenerateReport(Request.Form("Reports"), OptList)
            Response.Flush()
            Response.End()
        End If

    End Sub
    '-----------------------------------------------------------------------------------------
    ' AMSPS-2009 
    Sub Open_New_Enhanced_Result_Box()

        Dim dt As DataTable
        Dim OptList As String()

        Dim ReportStartDate As Date
        Dim ReportEndDate As Date

        Dim RowCount As Integer
        Dim CumulativeImpress As Integer
        Dim CumulativeRedeem As Integer
        Dim CumulativeTransact As Integer
        Dim CumulativeAmtRedeem As Double
        Dim RedemptionRate As Double
        Dim AmtRedeem As Double
        Dim Redemptions As Integer
        Dim Transactions As Integer
        Dim Impressions As Integer
        Dim i As Integer
        Dim OfferID As String = ""
        Dim dst As System.Data.DataTable
        Dim bParsed As Boolean
        Dim builder As StringBuilder = New StringBuilder()
        Dim frequency As String = ""
        Dim WhereClause As New StringBuilder("")

        'AMSPS-2231
        '    OfferIds = Request.Form("hdnOfferIDs")
        OfferIds = SortOfferIdsInString(Request.Form("hdnOfferIDs"), True)
        'AMSPS-2231 above
        OfferIds = Replace(OfferIds, vbCrLf, "")
        OfferIds = Replace(OfferIds, vbCr, "")
        If OfferIds = "" Then
            Return
        End If

        OptList = OfferIds.Split(",")

        If (Request.Form("Impressions") <> "" AndAlso Request.Form("ImpressionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumImpressions" & ConvertToOperand(Request.Form("ImpressionType")) & Request.Form("Impressions"))
            'WhereClause.Append("NumImpressions" & Request.Form("ImpressionType") & Request.Form("Impressions"))
        End If

        If (Request.Form("Redemptions") <> "" AndAlso Request.Form("RedemptionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumRedemptions" & ConvertToOperand(Request.Form("RedemptionType")) & Request.Form("Redemptions"))
            'WhereClause.Append("NumRedemptions" & Request.Form("RedemptionType") & Request.Form("Redemptions"))
        End If

        If (Request.Form("Transactions") <> "" AndAlso Request.Form("TransactionType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("NumTransactions" & ConvertToOperand(Request.Form("TransactionType")) & Request.Form("Transactions"))
        End If

        If (Request.Form("MarkDowns") <> "" AndAlso Request.Form("MarkDownType") <> "") Then
            WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
            WhereClause.Append("AmountRedeemed" & ConvertToOperand(Request.Form("MarkDownType")) & Request.Form("MarkDowns"))
            'WhereClause.Append("AmountRedeemed" & Request.Form("MarkDownType") & Request.Form("MarkDowns"))
        End If

        If WhereClause.Length > 0 Then
            WhereClause.Append(" and OfferID = ")
        Else
            WhereClause.Append("where OfferID = ")
        End If

        ' Getting starting and ending dates
        If Request.Form("reportingDate1").Trim() = "" And Request.Form("reportingDate2").Trim() = "" Then
            bParsed = DateTime.TryParse(GetReportStartingDateByOfferList(OptList), ReportStartDate)
            If (Not bParsed) Then ReportStartDate = Now()
            ReportEndDate = Now()
            If OptList.Length > 0 AND Session("RPTSTARTDATE") Is Nothing Then
                Session.Add("RPTSTARTDATE", ReportStartDate)
            Else
                Session.Remove("RPTSTARTDATE")
            End If
        Else
            If (Request.Form("hiddenReportingType") = "5") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then ReportStartDate = Now()
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate2"), ReportEndDate)
                If (Not bParsed) Then ReportEndDate = Now()
                ReportEndDate = ReportEndDate.AddDays(1).AddTicks(-1)
            ElseIf (Request.Form("hiddenReportingType") = "1") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportEndDate)
                If (Not bParsed) Then
                    ReportEndDate = Now()
                End If
                ReportEndDate = ReportEndDate.AddDays(1).AddTicks(-1)
                ReportStartDate = ReportEndDate.AddDays(-30)
            ElseIf (Request.Form("hiddenReportingType") = "0") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportEndDate)
                If (Not bParsed) Then
                    ReportEndDate = Now()
                End If
                ReportEndDate = ReportEndDate.Date
                ReportStartDate = ReportEndDate.Date.AddDays(-30)
                ReportEndDate = ReportEndDate.AddTicks(-1)
            ElseIf (Request.Form("hiddenReportingType") = "3") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then ReportStartDate = Now()
                ReportEndDate = Now()
            ElseIf (Request.Form("hiddenReportingType") = "4") Then
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then
                    ReportStartDate = Now()
                End If
                ReportStartDate = ReportStartDate.Date.AddDays(1)
                ReportEndDate = Now()
            Else   'Request.Form("hiddenReportingType") = "2"
                bParsed = DateTime.TryParse(Request.Form("hiddenReportingDate1"), ReportStartDate)
                If (Not bParsed) Then ReportStartDate = Now()
                ReportStartDate = ReportStartDate.Date
                ReportEndDate = ReportStartDate.AddDays(1).AddTicks(-1)
            End If
        End If

        Send("<div class=""box"" id=""enhancedResult"" style=""visibility: visible; width:720px; margin-top: 10px; "">")
        Send("  <h2>")
        Sendb("Enhanced Result")
        Send("  </h2>")

        If Request.Form("hideDownloadButton") = "visibility:visible;" Then
            StyleDownloadBtn = Request.Form("hideDownloadButton")

            Send("  <div id=""reportlist"" onclick=""return reportlist_onclick()"">")
            Send("  <table id=""list"" style=""width: 100%;"" summary=""")
            Sendb(Copient.PhraseLib.Lookup("term.reports", LanguageID))
            Send(""">")
            Send("    <thead>")
            Send("      <tr>")
            Send("        <th align=""left"" class=""th-id"" scope=""col"">")
            Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID))

            If Request.Form("hiddenImpression") = "true" Then
                Send("        </th>")
                Send("        <th align=""left"" id=""th1"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
            End If

            If Request.Form("hiddenCumulativeImpressions") = "true" Then
                Send("        </th>")
                Send("        <th align=""left"" id=""th11"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.cumulativeimpressions", LanguageID))
            End If

            If Request.Form("hiddenRedemption") = "true" Then
                Send("       </th>")
                Send("       <th align=""left"" id=""th2"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.redemptions", LanguageID))
            End If

            If Request.Form("hiddenCumulativeRedemptions") = "true" Then
                Send("       </th>")
                Send("       <th align=""left"" id=""th21"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.cumulativeredemptions", LanguageID))
            End If

            If Request.Form("hiddenTransaction") = "true" Then
                Send("       </th>")
                Send("       <th align=""left"" id=""th3"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.transactions", LanguageID))
            End If

            If Request.Form("hiddenCumulativeTransactions") = "true" Then
                Send("       </th>")
                Send("       <th align=""left"" id=""th31"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.cumulativetransactions", LanguageID))
            End If

            If Request.Form("hiddenMarkdown") = "true" Then
                Send("       </th>")
                Send("       <th align=""left"" id=""th4"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID))
            End If

            If Request.Form("hiddenCumulativeMarkdowns") = "true" Then
                Send("       </th>")
                Send("       <th align=""left"" id=""th41"" scope=""col"">")
                Sendb(Copient.PhraseLib.Lookup("term.cumulativemarkdowns", LanguageID))
            End If

            Send("       </th>")
            Send("       <th align=""left"" class=""th-datetime"" id=""th5"" scope=""col"">")
            Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))

            Send("       </th>")
            Send("     </tr>")
            Send("   </thead>")
            Send("   <tbody>")

            Send("<div id='rptContainer' style=""overflow: hidden;float: left; display:none;"">")
            Common.Open_LogixWH()
            For count = 0 To OptList.Length - 1
                If (OptList(count) <> "") Then
                    OfferID = Long.Parse(OptList(count))
                    If OfferID > 0 Then
                        Common.QueryStr = "select NumImpressions, NumRedemptions, NumTransactions, AmountRedeemed, CAST(ReportingDate as DATE) as ReportingDate from OfferReporting with (nolock) " & _
                                    WhereClause.ToString & OfferID & " " & _
                                    "and ReportingDate between '" & ReportStartDate.ToString & "' and '" & ReportEndDate.ToString & "' " & _
                                    "order by ReportingDate"
                        dt = Common.LWH_Select
                        If dt.Rows.Count > 0 Then
                            If (Request.Form("frequency") = "1") Then
                                dt = RollupReportWeek(dt, ReportStartDate, ReportEndDate)
                                frequency = "weekly"
                            ElseIf (Request.Form("frequency") = "2") Then
                                dt = FillInDays(dt, ReportStartDate, ReportEndDate)
                                frequency = "daily"
                            End If
                            WriteCustomReportRow(dt, frequency, offerID)
                        End If
                    End If
                End If
            Next
            Common.Close_LogixWH()
            Send("</div>")

            Send("    </tbody>")
            Send("  </table>")
            Send("</div>")

        ElseIf (Request.Form("generateReport") <> "" Or Request.Form("exportRpt") <> "") Then
            Send("<br /><center><i>" & Copient.PhraseLib.Lookup("term.noresults", LanguageID) & "</i></center>")
        End If

        If (ShowLimitNote) Then
            Send("<center><i>* " & Copient.PhraseLib.Lookup("report-custom.reachedlimit", LanguageID) & "</i></center>")
        End If

        Send("</div>")

        If (Request.Form("exportRpt") = "1") Then
            Response.ClearHeaders()
            Response.AddHeader("Content-Disposition", "attachment; filename=Offer" & OptList(0).ToString & "_Rpt.csv")
            Response.ContentType = "application/octet-stream"
            Response.Clear()
            GenerateReport(Request.Form("Reports"), OptList)
            Response.Flush()
            Response.End()
        End If

    End Sub
    '----------------------------------------------------------------------------------
    Function WriteCustomReportRow(ByVal dst As DataTable, ByVal frequency As String, ByVal OfferID As String) As String
        Dim builder As StringBuilder = New StringBuilder()
        Dim RowCount, i As Integer
        Dim cumulative As Double
        Dim dt As Date

        Dim cumuImpression As Integer
        Dim cumuRedeem As Integer
        Dim cumuTransact As Integer
        Dim strAmtRedeem As String = "0"
        Dim cumuAmtRedeem As Double
        Dim strCumuAmtRedeem As String = "0"

        ' building the report body       
        RowCount = dst.Rows.Count
        For i = 0 To (RowCount - 1)
            Send(" " & ControlChars.Tab & "    <tr class=""" & Shaded & """>")
            Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & OfferID  & "</td>")

            If (Request.Form("hiddenImpression") = "true") Then
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & dst.Rows(i).Item("NumImpressions") & "</td>")
            End If

            If (Request.Form("hiddenCumulativeImpressions") = "true") Then
                cumuImpression += dst.Rows(i).Item("NumImpressions")
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & cumuImpression & "</td>")
            End If

            If (Request.Form("hiddenRedemption") = "true") Then
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & dst.Rows(i).Item("NumRedemptions") & "</td>")
            End If

            If (Request.Form("hiddenCumulativeRedemptions") = "true") Then
                cumuRedeem += dst.Rows(i).Item("NumRedemptions")
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & cumuRedeem & "</td>")
            End If

            If (Request.Form("hiddenTransaction") = "true") Then
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & dst.Rows(i).Item("NumTransactions") & "</td>")
            End If

            If (Request.Form("hiddenCumulativeTransactions") = "true") Then
                cumuTransact += dst.Rows(i).Item("NumTransactions")
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & cumuTransact & "</td>")
            End If

            If (Request.Form("hiddenMarkdown") = "true") Then
                If dst.Rows(i).Item("AmountRedeemed") <> 0.0 Then
                    strAmtRedeem = Format(dst.Rows(i).Item("AmountRedeemed"), "0.00")
                Else
                    strAmtRedeem = "0"
                End If
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & strAmtRedeem & "</td>")
            End If

            If (Request.Form("hiddenCumulativeMarkdowns") = "true") Then
                cumuAmtRedeem += dst.Rows(i).Item("AmountRedeemed")
                If cumuAmtRedeem <> 0.0 Then
                    strCumuAmtRedeem = Format(cumuAmtRedeem, "0.00")
                Else
                    strCumuAmtRedeem = "0"
                End If
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & strCumuAmtRedeem & "</td>")
            End If

            dt = dst.Rows(i).Item("ReportingDate")
            Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & dt.ToString("M/dd/yyyy") & "</td>")
            Send(" " & ControlChars.Tab & "    </tr>")
            If Shaded = "shaded" Then
                Shaded = ""
            Else
                Shaded = "shaded"
            End If
        Next
        Return builder.ToString()
    End Function

    '-----------------------------------------------------------------------------------------
    Sub Open_Result_Box()

        ' Do not initially run the query and hide the results table
        If (Request.Form("generateReport") <> "" Or Request.Form("exportRpt") <> "") Then
            OfferIds = Request.Form("hdnOfferIDs")
            OfferIds = Replace(OfferIds, vbCrLf, "")
            OfferIds = Replace(OfferIds, vbCr, "")
            Impressions = Request.Form("impressions")
            ImpressionType = Request.Form("impressionType")
            Redemptions = Request.Form("redemptions")
            RedemptionType = Request.Form("redemptionType")
            Transactions = Request.Form("transactions")
            TransactionType = Request.Form("transactionType")
            MarkDowns = Request.Form("markDowns")
            MarkDownType = Request.Form("markDownType")
            ReportingStartDate = Request.Form("reportingDate1")
            ReportingType = Request.Form("reportingType")
            ReportingEndDate = Request.Form("reportingDate2")
            SortText = Request.Form("sortText")
            SortDirection = Request.Form("sortDir")

            If (SortText = "") Then
                SortText = "OfferID"
            End If

            If (SortDirection = "ASC") Then
                NextSortDirection = "DESC"
            Else
                NextSortDirection = "ASC"
            End If

            Dim WhereClause As New StringBuilder(" ")

            If (Impressions <> "" AndAlso ImpressionType <> "") Then
                WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
                WhereClause.Append("NumImpressions" & ConvertToOperand(ImpressionType) & Impressions)
            End If

            If (Redemptions <> "" AndAlso RedemptionType <> "") Then
                WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
                WhereClause.Append("NumRedemptions" & ConvertToOperand(RedemptionType) & Redemptions)
            End If

            If (Transactions <> "" AndAlso TransactionType <> "") Then
                WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
                WhereClause.Append("NumTransactions" & ConvertToOperand(TransactionType) & Transactions)
            End If

            If (MarkDowns <> "" AndAlso MarkDownType <> "") Then
                WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))
                WhereClause.Append("AmountRedeemed" & ConvertToOperand(MarkDownType) & MarkDowns)
            End If

            If (ReportingStartDate <> "" AndAlso ReportingType <> "") Then
                WhereClause.Append(IIf(WhereClause.Length > 0, " and ", " Where "))

                WhereClause.Append("ReportingDate" & ConvertToOperand(ReportingType) & "'" & Copient.commonShared.ConvertToSqlDate(ReportingStartDate, Common.GetAdminUser.Culture) & "'")

                If (ReportingEndDate <> "") Then
                    WhereClause.Append(" and '" & Copient.commonShared.ConvertToSqlDate(ReportingEndDate, Common.GetAdminUser.Culture) & "'")
                    DisplayRptEnd = "display:inline;"
                End If
            End If

            WhereClause.Append(" Order by " & SortText & " " & SortDirection)

            If (Not OfferIds Is Nothing AndAlso OfferIds.Length > 0) Then

                Dim dtOfferIDs As New DataTable
                dtOfferIDs.TableName = "OfferIDs"
                dtOfferIDs.Columns.Add("OfferID")

                Dim offers As New StringBuilder("")
                OfferIds = OfferIds.Trim(New Char() {","})

                For Each OfferID In OfferIds.Split(",")
                    offers.Append("<option value=""" & OfferID & """>" & OfferID)
                    dtOfferIDs.Rows.Add(OfferID)
                Next
                OfferIdOptions = offers.ToString()

                If infoMessage = "" AndAlso dtOfferIDs.Rows.Count > 0 Then
                    Common.QueryStr = "dbo.pa_GetReports"
                    Common.Open_LWHsp()
                    Common.LWHsp.Parameters.AddWithValue("@OffersT", dtOfferIDs)
                    Common.LWHsp.Parameters.AddWithValue("@WhereClause", WhereClause.ToString())
                    Common.LWHsp.Parameters.AddWithValue("@RecordLimit", RECORD_LIMIT)
                    dtResult = Common.LWHsp_select()
                    Common.Close_LWHsp()

                    sizeOfData = dtResult.Rows.Count
                End If
            End If
            'Write the data as a CSV file for download
            If (Request.Form("exportRpt") = "1") Then
                Response.ClearHeaders()
                Response.AddHeader("Content-Disposition", "attachment; filename=Rpt" & Now().ToShortDateString() & ".csv")
                Response.ContentType = "application/octet-stream"
                Response.Clear()
                'Write the column headers
                Sendb("OfferID,")
                Sendb("Impressions,")
                Sendb("Redemptions,")
                If DefaultToEnhancedCustomReport = 1 and Request.Form("hiddenTransaction") = "true" Then
                    Sendb("Transactions,")
                End If
                Sendb("Mark Downs,")
                Sendb("Reporting Date")
                Send("")
                'Write the data rows
                i = 0
                For Each row In dtResult.Rows
                    Sendb(Common.NZ(row.Item("OfferId"), ""))
                    Sendb(",")
                    Sendb(Common.NZ(row.Item("NumImpressions"), ""))
                    Sendb(",")
                    Sendb(Common.NZ(row.Item("NumRedemptions"), ""))
                    Sendb(",")
                    If DefaultToEnhancedCustomReport = 1 and Request.Form("hiddenTransaction") = "true" Then
                        Sendb(Common.NZ(row.Item("NumTransactions"), ""))
                        Sendb(",")
                    End If
                    Sendb(Common.NZ(row.Item("AmountRedeemed"), ""))
                    Sendb(",")
                    If Not IsDBNull(row.Item("ReportingDate")) Then
                        Sendb(Logix.ToShortDateString(row.Item("ReportingDate"), Common))
                    End If
                    Send("")
                    i = i + 1
                Next
                Response.Flush()
                Response.End()
            End If

            StyleDownloadBtn = "visibility:visible;"
            ShowReportList = (sizeOfData > 0)
            ShowLimitNote = (sizeOfData >= RECORD_LIMIT)
        End If
        ' Insert from original above	
        Send("<div class=""box"" id=""results"" style=""width: 720px; margin-top:10px;"">")
        Send("  <h2>")
        Sendb(Copient.PhraseLib.Lookup("term.results", LanguageID))
        Send("  </h2>")
        If (ShowReportList) Then
            Send("  <div id=""reportlist"" onclick=""return reportlist_onclick()"">")
            Send("  <table id=""list"" style=""width: 100%;"" summary=""")
            Sendb(Copient.PhraseLib.Lookup("term.reports", LanguageID))
            Send(""">")
            Send("    <thead>")
            Send("      <tr>")
            Send("        <th align=""left"" class=""th-id"" scope=""col"">")
            Send("          <a href=""javascript:doSort('OfferID', '")
            Sendb(NextSortDirection)
            Send("');"">")
            Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID))
            Send("          </a>")
            If SortText = "OfferID" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            End If

            Send("        </th>")
            Send("        <th align=""left"" id=""th1"" scope=""col"">")
            Send("          <a href=""javascript:doSort('NumImpressions', '")
            Sendb(NextSortDirection)
            Send("');"">")
            Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID))
            Send("   </a>")
            If SortText = "NumImpressions" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            End If

            Send("       </th>")
            Send("       <th align=""left"" id=""th2"" scope=""col"">")
            Send("         <a href=""javascript:doSort('NumRedemptions', '")
            Sendb(NextSortDirection)
            Send("');"">")
            Sendb(Copient.PhraseLib.Lookup("term.redemptions", LanguageID))
            Send("  </a>")
            If SortText = "NumRedemptions" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            End If

            Send("       </th>")

            If DefaultToEnhancedCustomReport = 1 and Request.Form("hiddenTransaction") = "true" Then
                Send("       <th align=""left"" id=""th2"" scope=""col"">")
                Send("         <a href=""javascript:doSort('NumTransactions', '")
                Sendb(NextSortDirection)
                Send("');"">")
                Sendb(Copient.PhraseLib.Lookup("term.transactions", LanguageID))
                Send("  </a>")
                If SortText = "NumTransactions" Then
                    If SortDirection = "ASC" Then
                        Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                        Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                End If
                Send("       </th>")
            End If

            Send("       <th align=""left"" id=""th3"" scope=""col"">")
            Send("         <a href=""javascript:doSort('AmountRedeemed', '")
            Sendb(NextSortDirection)
            Send("');"">")
            Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID))
            Send("    </a>")
            If SortText = "AmountRedeemed" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            End If

            Send("       </th>")
            Send("       <th align=""left"" class=""th-datetime"" id=""th4"" scope=""col"">")
            Send("         <a href=""javascript:doSort('ReportingDate', '")
            Sendb(NextSortDirection)
            Send("');"">")
            Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))
            Send("    </a>")
            If SortText = "ReportingDate" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            End If

            Send("       </th>")
            Send("     </tr>")
            Send("   </thead>")
            Send("   <tbody>")

            For Each row In dtResult.Rows

                Send(" " & ControlChars.Tab & "    <tr class=""" & Shaded & """>")
                Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & Common.NZ(row.Item("OfferId"), "") & "</td>")

                If Request.Form("hiddenImpression") = "true" Then
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & Common.NZ(row.Item("NumImpressions"), "") & "</td>")
                Else
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>--</td>")
                End If
                If Request.Form("hiddenRedemption") = "true" Then
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & Common.NZ(row.Item("NumRedemptions"), "") & "</td>")
                Else
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>--</td>")
                End If

                If DefaultToEnhancedCustomReport = 1 and Request.Form("hiddenTransaction") = "true" Then
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & Common.NZ(row.Item("NumTransactions"), "") & "</td>")
                End If

                If Request.Form("hiddenMarkdown") = "true" Then
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & Common.NZ(row.Item("AmountRedeemed"), "") & "</td>")
                Else
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>--</td>")
                End If
                If IsDBNull(row.Item("ReportingDate")) Then
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td></td>")
                Else
                    Send(" " & ControlChars.Tab & ControlChars.Tab & "<td>" & Logix.ToShortDateTimeString(row.Item("ReportingDate"), Common) & "</td>")
                End If
                Send(" " & ControlChars.Tab & "   </tr>")
                If Shaded = "shaded" Then
                    Shaded = ""
                Else
                    Shaded = "shaded"
                End If
                i = i + 1
            Next

            Send("    </tbody>")
            Send("  </table>")
            Send("</div>")

        ElseIf (Request.Form("generateReport") <> "" Or Request.Form("exportRpt") <> "") Then
            Send("<br /><center><i>" & Copient.PhraseLib.Lookup("term.noresults", LanguageID) & "</i></center>")
        End If
        If (ShowLimitNote) Then
            Send("<center><i>* " & Copient.PhraseLib.Lookup("report-custom.reachedlimit", LanguageID) & "</i></center>")
        End If

        Send("</div>")
    End Sub


    'AMSPS-2231
    Function SortOfferIdsInString(ByVal OfferIds As String, Ascending As Boolean) As String
        Dim SortedOfferIds As String = ""
        Dim bSuccess As Boolean = False
        Dim i As Long
        Dim str() As String = OfferIds.Split(",")
        Dim InArray(str.Length - 1) As Long
        Dim OutArray(str.Length - 1) As Long
        For i = 0 To str.Length - 1
            bSuccess = Long.TryParse(str(i), InArray(i) )
        Next

        OutArray = BubbleSrt(InArray, True)
        For i = LBound(OutArray) + 1 To UBound(OutArray)
            SortedOfferIds = SortedOfferIds + OutArray(i).ToString
            If (i < UBound(OutArray)) Then SortedOfferIds = SortedOfferIds + ","
        Next i

        SortOfferIdsInString = SortedOfferIds
    End Function


    Function BubbleSrt(ArrayIn As Long(), Ascending As Boolean) As Long()
        Dim SrtTemp As Object
        Dim i As Long
        Dim j As Long

        If Ascending = True Then
            For i = LBound(ArrayIn) To UBound(ArrayIn)
                For j = i + 1 To UBound(ArrayIn)
                    If ArrayIn(i) > ArrayIn(j) Then
                        SrtTemp = ArrayIn(j)
                        ArrayIn(j) = ArrayIn(i)
                        ArrayIn(i) = SrtTemp
                    End If
                Next j
            Next i
        Else
            For i = LBound(ArrayIn) To UBound(ArrayIn)
                For j = i + 1 To UBound(ArrayIn)
                    If ArrayIn(i) < ArrayIn(j) Then
                        SrtTemp = ArrayIn(j)
                        ArrayIn(j) = ArrayIn(i)
                        ArrayIn(i) = SrtTemp
                    End If
                Next j
            Next i
        End If

        BubbleSrt = ArrayIn

    End Function
    'AMSPS-2231 above


</script>
<script type="text/javascript" runat="server">

  Function Impressions_checked()
    'alert("Hello!");
    Return "javascript: document.getElementById('impressions').checked.toString();"
  End Function

  Function Redemptions_checked()
    Return "javascript: document.getElementById('redemptions').checked.toString();"
  End Function

  Function Transactions_checked()
    Return "javascript: document.getElementById('transactions').checked.toString();"
  End Function

  Function Markdowns_checked()
    Return "javascript: document.getElementById('markdowns').checked.toString();"
  End Function

</script>

<%
  
'-------------------------------------------------------------------------------------------------------------
' Main code - execution starts here ...   
   
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim LinkIds As String = Request.QueryString("LinkIds")
 
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  RecordLimitValue = Common.Fetch_SystemOption(176)
  If RecordLimitValue <> "" Then
    If IsNumeric(RecordLimitValue) Then
      If CInt(RecordLimitValue) > 0 Then
        RECORD_LIMIT = RecordLimitValue
      End If
    End If
  End If

  If Common.Fetch_SystemOption(274) <> "" And Common.Fetch_SystemOption(274) = "1" Then
    DefaultToEnhancedCustomReport = 1
  End If
  
  Response.Expires = 0
  Common.AppName = "reports-custom.aspx"
  On Error GoTo ErrorTrap
  Common.Open_LogixRT()
  Common.Open_LogixWH()

  AdminUserID = Verify_AdminUser(Common, Logix)
  
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  Send_HeadBegin("term.reports")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>

<script type="text/javascript">
 
  var datePickerDivID = "datepicker";
  
  var highlightedRow = -1;
  var highlightedCol = -1;
  
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
  
  <% Send_Calendar_Overrides(Common) %>

  function handlePageClick(e) {
    var calFrame = document.getElementById('calendariframe');
    var el=(typeof event!=='undefined')? event.srcElement : e.target        
    
    if (el != null) {
      var pickerDiv = document.getElementById(datePickerDivID);
      if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
        if (el.id!="prod-start-picker" && el.id!="prod-end-picker") {
          if (!isDatePickerControl(el.className)) {
            pickerDiv.style.visibility = "hidden";
            pickerDiv.style.display = "none";
            if (calFrame != null) {
              calFrame.style.visibility = 'hidden';
              calFrame.style.display = "none";
            }
          }
        } else  {
          pickerDiv.style.visibility = "visible";            
          pickerDiv.style.display = "block";            
          if (calFrame != null) {
            calFrame.style.visibility = 'visible';
            calFrame.style.display = "block";
          }
        }
      }
    }
  }
  
  function isDatePickerControl(ctrlClass) {
    var retVal = false;
    
    if (ctrlClass != null && ctrlClass.length >= 2) {
      if (ctrlClass.substring(0,2) == "dp") {
        retVal = true;
      }
    }
    
    return retVal;
  }

  function handleSearch() {
 
    if(document.getElementById("byweek").checked)
      document.getElementById("frequency").value = '1';
    else
      document.getElementById("frequency").value = '2';

    var elemfrequency = document.getElementById("frequency");
    var rptStart = document.getElementById("reportingDate1");
    var rptEnd = document.getElementById("reportingDate2");

    if(elemfrequency.value == '1'){
      if(document.getElementById("impression").checked)
        document.getElementById("row1").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID)) %>)';
      if(document.getElementById("redemption").checked)
        document.getElementById("row2").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID)) %>)';
      if(document.getElementById("markdown").checked)
        document.getElementById("row3").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.weekly", LanguageID)) %>)';
     
    }else {
      if(document.getElementById("impression").checked)
        document.getElementById("row1").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID)) %>)';
      if(document.getElementById("redemption").checked)
        document.getElementById("row2").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID)) %>)';
      if(document.getElementById("markdown").checked)
        document.getElementById("row3").innerHTML = '(<%Sendb(Copient.PhraseLib.Lookup("term.daily", LanguageID)) %>)';
    }
             
    generateReport();
  }
    
  function generateReport() {
 
    xmlhttpPost("XMLFeeds.aspx")

  }
  
  function xmlhttpPost(strURL) {
    var xmlHttpReq = false;
    var self = this;
 
    document.getElementById("nota").style.display = "none";

    document.getElementById("rptHeader").style.visibility = 'hidden';
    document.getElementById("reportTitle").style.visibility = "hidden";

    document.getElementById("report").style.visibility = "hidden";
 
   
    resetRow(highlightedRow);
    highlightedRow = -1;
    
    document.getElementById("wait").style.visibility = "visible";
    document.getElementById("wait").innerHTML = "<div class=\"loading\"><br \/><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %><\/div>';
    
    // Mozilla/Safari
    if (window.XMLHttpRequest) {
      self.xmlHttpReq = new XMLHttpRequest();
    }
    // IE
    else if (window.ActiveXObject) {
      self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }

    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function() {

      if (self.xmlHttpReq.readyState == 4) {
        updatepage(self.xmlHttpReq.responseText);
      }
    }
    //alert("Before self.xmlHttpReq.send()");
    self.xmlHttpReq.send(getPostData());
  }
  
  function getPostData() {
    var form = document.forms[1];
  
    var freq = 0;
    if(document.getElementById("byweek").checked){ 
        freq = 1;
    }
    else{
        freq = 2;
    }    
    var startDate = document.getElementById("reportingDate1").value;
    var endDate = document.getElementById("reportingDate2").value;
    var offerId = document.getElementById("offerID").value;

    var impressionChecked = document.getElementById("impression").checked.toString();
    var redemptionChecked =  document.getElementById('redemption').checked.toString();
    var transactionChecked =  document.getElementById('transaction').checked.toString();
    var markdownChecked =  document.getElementById('markdown').checked.toString();
    var CumuImpressionsChecked =  document.getElementById('CumulativeImpressions').checked.toString();
    var CumuRedemptionsChecked =  document.getElementById('CumulativeRedemptions').checked.toString();
    var CumuTransactionsChecked =  document.getElementById('CumulativeTransactions').checked.toString();
    var CumuMarkdownsChecked =  document.getElementById('CumulativeMarkdowns').checked.toString();
    var impressions =  document.getElementById("impressions").value;
    var imp =  document.getElementById("impressionType");
    var impressionType = imp.options[imp.selectedIndex].text;
    var redemptions =  document.getElementById("redemptions").value;
    var red =  document.getElementById("redemptionType");
    var redemptionType =  red.options[red.selectedIndex].text;
    var transactions =  document.getElementById("transactions").value;
    var trx =  document.getElementById("transactionType");
    var transactionType =  trx.options[trx.selectedIndex].text;
    var markdowns =  document.getElementById("markDowns").value;
    var mar = document.getElementById("markDownType");
    var markdownType = mar.options[mar.selectedIndex].text;
    var ofrIds = document.getElementById("hdnOfferIDs").value;
    var rpttype =  document.getElementById("reportingType");
    var reportingType =  rpttype.options[rpttype.selectedIndex].text;

    //alert("ofrIds = " + ofrIds);
    
    qstr = 'CustomReports=1&frequency=' + escape(freq); // NOTE: no '?' before querystring
    qstr += "&reportstart=" + escape(startDate);
    qstr += "&reportend=" + escape(endDate);  
    qstr += "&lang=<% Sendb(LanguageID) %>";
    qstr += "&offerID=" + escape(offerId);
    qstr += "&ofrIds=" + escape(ofrIds);
    qstr += "&impressionChecked=" + escape(impressionChecked);
    qstr += "&redemptionChecked=" + escape(redemptionChecked);
    qstr += "&transactionChecked=" + escape(transactionChecked);
    qstr += "&markdownChecked=" + escape(markdownChecked);
    qstr += "&CumuImpressionsChecked=" + escape(CumuImpressionsChecked);
    qstr += "&CumuRedemptionsChecked=" + escape(CumuRedemptionsChecked);
    qstr += "&CumuTransactionsChecked=" + escape(CumuTransactionsChecked);
    qstr += "&CumuMarkdownsChecked=" + escape(CumuMarkdownsChecked);
    qstr += "&Impressions=" + escape(impressions);
    qstr += "&ImpressionType=" + escape(impressionType);
    qstr += "&Redemptions=" + escape(redemptions);
    qstr += "&RedemptionType=" + escape(redemptionType);
    qstr += "&Transactions=" + escape(transactions);
    qstr += "&TransactionType=" + escape(transactionType);
    qstr += "&MarkDowns=" + escape(markdowns);
    qstr += "&MarkDownType=" + escape(markdownType);
    qstr += "&ReportingType=" + escape(reportingType);
    //alert(qstr.toString());    
    return qstr;
  }

  
  function updatepage(str){
    if (str != null && str.indexOf("No Data Found") == -1) {
      document.getElementById("report").innerHTML = str;
      document.getElementById("nota").style.display = "none";
      document.getElementById("download").style.visibility = "visible";
      document.getElementById("reportTitle").style.visibility = "visible";
      document.getElementById("report").style.visibility = "visible";
      document.getElementById("rptHeader").style.visibility = "visible";
    } else {
      document.getElementById("nota").innerHTML = str; 
      document.getElementById("nota").style.display = "inline"
    }
    
    document.getElementById("generateReport").disabled = false;
    document.getElementById("reportstart").disabled = false;
    document.getElementById("reportend").disabled = false;
    document.getElementById("download").disabled = false;
  }
  
  
  function launchGraph(type, offerId) {
    var strURL = "graph-display.aspx?type=" + type;
    var form = document.forms[1];
    var freq = document.getElementById("frequency").value;
    
    var startDate = document.getElementById("reportingDate1").value;
    var endDate = document.getElementById("reportingDate2").value;
   
    strURL += "&offerId=" + offerId + "&freq=" + freq + "&start=" + startDate + "&end=" + endDate;
    openReports(strURL);
  }
  
  function toggleHighlight(row) {
    var rowHdrElem = document.getElementById("rowHdr"+row);
    var rowBodyElem = document.getElementById("rowBody"+row);
    var bHighlight = false;
    var bSameRow = false;
    
    if (highlightedCol > -1) {
      toggleHighlight(highlightedCol);
      highlightedCol = -1;    
    }
    
    bSameRow = (row == highlightedRow);
    resetRow(highlightedRow);
    highlightedRow = -1;
    
    if (rowHdrElem != null && rowBodyElem != null && !bSameRow) {
      bHighlight = (rowHdrElem.className != "reportRowHeader rowHighlighted");
      if (bHighlight) {
        rowHdrElem.className = 'reportRowHeader rowHighlighted';
        rowBodyElem.className = 'rowHighlighted';
        highlightedRow = row;
      }
    }
  }
  
  function resetRow(row) {
    var rowHdrElem = document.getElementById("rowHdr"+row);
    var rowBodyElem = document.getElementById("rowBody"+row);
    
    if (rowHdrElem != null && rowBodyElem != null) {
      if (row % 2 == 1) {
        rowHdrElem.className = "reportRowHeader"
        rowBodyElem.className = "noclass"
      } else {
        rowHdrElem.className = "reportRowHeader shaded"
        rowBodyElem.className = "shaded"
      }
    }
  }


    function addOffer(){
       //alert('<%Sendb(OfferIdOptions)%>');

       var id = prompt('<% Sendb(Copient.PhraseLib.Lookup("report-custom-offerprompt", LanguageID)) %>', "");
       addOfferbyid(id);
       //alert("hdnOfferIDs = #" + document.getElementById('hdnOfferIDs').value + "#");

    }

    
    function addOfferbyid(id) {
        var slctElem = null;
        
        var ids = null;
        var offerID = null;

        if (id < 1 && id >= 0) {
            return;
        } else if (id < 0) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("report-custom-nonnumeric", LanguageID)) %>');
            return;
        } else {
            slctElem = document.getElementById("offerIds");

            var maxLength =  <%Sendb(MaxBigIntLength) %>            
            if (id != null && id != "" && slctElem != null) {     
                ids = id.split(",");
                for (var i=0; i < ids.length; i++) {
                    offerID = ids[i];
                    if (offerID != "" && offerID.length > 0 && !isNaN(offerID)) {                      
                          if(offerID.length > maxLength){
                            alert('<% Sendb(Copient.PhraseLib.Lookup("error.offerIDtoolong", LanguageID)) %>');
                          }
                          else{
                            slctElem.options.add( new Option(offerID, offerID));
                          }                        
                    }

                    
                } 
            }
        }
        selectAllOffers();
        if(document.getElementById("EnhancedReport").checked){
          document.getElementById("hideDownloadButton").value = "visibility:hidden;";
          document.getElementById("download").style.dispplay = "none";
          document.mainform.exportRpt.value = "";
          document.mainform.submit();
        }
        handleResultsButton();
        
    }

    function removeOffer() {
        var slctElem = document.getElementById("offerIds");
        
        if (slctElem != null) {
            for (var i=0; i < slctElem.options.length; i++) {
                if (slctElem.options[i].selected == true) {
                    slctElem.options[slctElem.options.selectedIndex] = null;
                    i--;
                }
            }
        }
        selectAllOffers();
        if(document.getElementById("EnhancedReport").checked){
            document.mainform.exportRpt.value = "";
            updateHiddenInputs();
            document.getElementById("hideDownloadButton").value = "visibility:hidden;";
            document.getElementById("download").style.dispplay = "none";
            document.mainform.exportRpt.value = "";
            document.mainform.submit();
        }
        handleResultsButton();

    }
    
    function handleResultsButton() {
        var slctElem = document.getElementById("offerIds");
        var btnElem = document.getElementById("generateReport");
    
        if (slctElem != null && btnElem != null) {
            btnElem.disabled = (slctElem.options.length == 0) ? true : false;
        }            
    }
    
    function clearForm() {
        var slctElem = document.getElementById("offerIds");
        
        if (slctElem != null) {
            while (slctElem.options.length > 0 ){
                slctElem.options[0] = null;
            }
        }
        document.mainform.impression.checked = false;
        document.mainform.impressions.value = "";
        document.mainform.impressionType.options[0].selected = true;
        document.mainform.redemption.checked = false;
        document.mainform.redemptions.value = "";
        document.mainform.redemptionType.options[0].selected = true;
        document.mainform.transaction.checked = false;
        document.mainform.transactions.value = "";
        document.mainform.transactionType.options[0].selected = true;
        document.mainform.markdown.checked = false;
        document.mainform.markDowns.value="";
        document.mainform.markDownType.options[0].selected = true;
        document.mainform.reportingType.options[0].selected = true;
        document.mainform.reportingDate1.value="";
        document.mainform.reportingDate2.value="";
        document.mainform.reportingDate2.style.display = "none";
        document.mainform.EnhancedReport.checked = false;
        document.mainform.byweek.checked = false;
        document.mainform.byday.checked = false;
        document.mainform.CumulativeImpressions.checked = false;
        document.mainform.CumulativeRedemptions.checked = false;
        document.mainform.CumulativeTransactions.checked = false;
        document.mainform.CumulativeMarkdowns.checked = false;

        document.mainform.exportRpt.value = "";
        handleResultsButton();
        updateHiddenInputs();
        document.getElementById("hideDownloadButton").value = "visibility:hidden;";

        document.mainform.submit();
    }
    
    function validateForm() {        
        if (document.mainform.impressions.value != "") {
            if (isNaN(document.mainform.impressions.value)) {
                alert('<% Sendb(Copient.PhraseLib.Lookup("report-custom-impressionerror", LanguageID)) %>');
                return false;
            }
            if (parseInt(document.mainform.impressions.value)>9223372036854775807 || parseInt(document.mainform.impressions.value) < 0) {                
                alert('<% Sendb(Copient.PhraseLib.Lookup("term.impressions", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-9223372036854775807)") %>');
                return false;
            }
        }
        
        if (document.mainform.redemptions.value != "") {
            if (isNaN(document.mainform.redemptions.value)) {
                alert('<% Sendb(Copient.PhraseLib.Lookup("report-custom-redemptionerror", LanguageID)) %>');
                return false;
            }
            if (parseInt(document.mainform.redemptions.value)>9223372036854775807 || parseInt(document.mainform.redemptions.value) < 0) {                
                alert('<% Sendb(Copient.PhraseLib.Lookup("term.redemptions", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-9223372036854775807)") %>');
                return false;
            }
        }
        
        if (document.mainform.transactions.value != "") {
            if (isNaN(document.mainform.transactions.value)) {
                alert('<% Sendb(Copient.PhraseLib.Lookup("report-custom-transactionerror", LanguageID)) %>');
                return false;
            }
            if (parseInt(document.mainform.transactions.value)>9223372036854775807 || parseInt(document.mainform.transactions.value) < 0) {                
                alert('<% Sendb(Copient.PhraseLib.Lookup("term.transactions", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-9223372036854775807)") %>');
                return false;
            }
        }
        
        if (document.mainform.markDowns.value != "") {
            if (isNaN(document.mainform.markDowns.value)) {
                alert('<% Sendb(Copient.PhraseLib.Lookup("report-custom-markdownerror", LanguageID)) %>');
                return false;
            }
            if (parseInt(document.mainform.markDowns.value)>99999999999999999999999999999999999999 || parseInt(document.mainform.markDowns.value) < 0) {                
                alert('<% Sendb(Copient.PhraseLib.Lookup("term.markdowns", LanguageID) + ": " + Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) + " (0-99999999999999999999999999999999999999)") %>');
                return false;
            }
        }
        
        selectAllOffers();
        return true;        
    }
    
    function selectAllOffers(){
    var obj1 = document.getElementById("hdnOfferIDs")
    var slctElem = document.getElementById("offerIds");
        var str = '';
        if (slctElem != null) {
        
            for (var i=0; i < slctElem.options.length; i++) {
            var offerID = slctElem.options[i].value;
             var maxLength =  <%Sendb(MaxBigIntLength) %>    
                if(offerID.length > maxLength){
                     alert(offerID+': <% Sendb(Copient.PhraseLib.Lookup("error.offerIDtoolong", LanguageID)) %>');
                }
                str = str + offerID + ',';
            }
        } 
        obj1.value = str;  
        //alert("hdnOfferIDs = " + document.getElementById('hdnOfferIDs').value);
    }


    function updateHiddenInputs()
    {
        selectAllOffers();
        //alert(" In updateHiddenInputs " + document.getElementById('hdnOfferIDs').value);
        document.getElementById('hiddenImpression').value = document.getElementById('impression').checked.toString();
        document.getElementById('hiddenImpressionType').value = document.getElementById('impressionType').options.selectedIndex;
        document.getElementById('hiddenImpressions').value = document.getElementById('impressions').value

        document.getElementById('hiddenRedemption').value = document.getElementById('redemption').checked.toString();
        document.getElementById('hiddenRedemptionType').value = document.getElementById('redemptionType').options.selectedIndex;
        document.getElementById('hiddenRedemptions').value = document.getElementById('redemptions').value

        document.getElementById('hiddenTransaction').value = document.getElementById('transaction').checked.toString();
        document.getElementById('hiddenTransactionType').value = document.getElementById('transactionType').options.selectedIndex;
        document.getElementById('hiddenTransactions').value = document.getElementById('transactions').value

        document.getElementById('hiddenMarkdown').value = document.getElementById('markdown').checked.toString();
        document.getElementById('hiddenMarkdownType').value = document.getElementById('markDownType').options.selectedIndex;
        document.getElementById('hiddenMarkdowns').value = document.getElementById('markDowns').value

        document.getElementById('hiddenReportingType').value = document.getElementById('reportingType').options.selectedIndex;
        document.getElementById('hiddenReportingDate1').value = document.getElementById('reportingDate1').value;
        if (document.getElementById('hiddenReportingType').value == "5" )
            document.getElementById('hiddenReportingDate2').value = document.getElementById('reportingDate2').value;
        else
            document.getElementById('hiddenReportingDate2').value = "";
        document.getElementById('hiddenEnhancedReport').value = document.getElementById('EnhancedReport').checked.toString();
        document.getElementById('hiddenByday').value = document.getElementById('byday').checked.toString();
        document.getElementById('hiddenByweek').value = document.getElementById('byweek').checked.toString();
        if(document.getElementById('EnhancedReport').checked && document.getElementById('byweek').checked)
            document.getElementById("frequency").value = '1';
        else
            document.getElementById("frequency").value = '2';
        document.getElementById('hiddenCumulativeImpressions').value = document.getElementById('CumulativeImpressions').checked.toString();
        document.getElementById('hiddenCumulativeRedemptions').value = document.getElementById('CumulativeRedemptions').checked.toString();
        document.getElementById('hiddenCumulativeTransactions').value = document.getElementById('CumulativeTransactions').checked.toString();
        document.getElementById('hiddenCumulativeMarkdowns').value = document.getElementById('CumulativeMarkdowns').checked.toString();
    }

     // report loading indication below
     function showloadingscreen() {
      
      var collisionsBox = document.getElementById("loading");
      if (collisionsBox != null) {
     
      collisionsBox.style.display = 'block';
     
      }
     }

     function closeloadingscreen(){
       var collisionsBox = document.getElementById("loading");
       if (collisionsBox != null) {
     
         collisionsBox.style.display = 'none';
     
       }
     }
     // report loading indication above

     function submitForm() {
      var form = document.mainform;
      var sel = document.getElementById("reportingType").options.selectedIndex;

      //alert(sel.toString());
      var DefaultToEnhancedCustomReport = <%Sendb(DefaultToEnhancedCustomReport)%>;
      //alert(DefaultToEnhancedCustomReport.toString());
   
      if((sel.toString() == "5") && (document.getElementById("reportingDate1").value.toString() == "" || document.getElementById("reportingDate2").value.toString() == ""))
      {
          alert("Report date range are incorrect!");
          return;
      }
        
      updateHiddenInputs();

      //if(document.getElementById("EnhancedReport").checked){
      if(document.getElementById("EnhancedReport").checked && DefaultToEnhancedCustomReport == 0){        
        if( document.getElementById('reportingDate1').value == "")
        {
          alert("Please enter a date for enhanced reporting.");
          return;
        }

        document.getElementById("rptContainer").style.display = "block";
        handleSearch();

        document.getElementById("hideDownloadButton").value = "visibility:hidden;";

        return true;
      }
      else{
        if(document.getElementById("EnhancedReport").checked && DefaultToEnhancedCustomReport == 1)
        {  
          showloadingscreen();
        }

        form.exportRpt.value = "0";

        if (validateForm()) {
          document.getElementById("download").style.display = 'display';
          document.getElementById("hideDownloadButton").value = "visibility:visible;";       
          form.submit();
        }
        handleResultsButton();
      }
    }

    
    function doSort(sortText, sortDir) {
        document.mainform.sortText.value = sortText;
        document.mainform.sortDir.value = sortDir;
        submitForm();
    }
    
   

    function toggleDate2(slct) {
        var elemDtTxt = document.getElementById("reportingDate2");
        
        if (slct != null) {
            if (slct.options.selectedIndex == 5) {
                elemDtTxt.style.display = "inline";
            } else {
                elemDtTxt.value = "";
                elemDtTxt.style.display = "none";
            }
        }
    }

    
    function handleDownload() {
      var form = document.mainform;
      //var DefaultToEnhancedCustomReport = <%Sendb(DefaultToEnhancedCustomReport)%>;

      updateHiddenInputs();
      if (validateForm()) {
          form.exportRpt.value = "1";
          form.action = "#";
          form.method = "post";
          form.submit();
      }
    }


    function EnableEnhancedOptions() {
        var status = document.getElementById('hiddenEnhancedReport');
        var box = document.getElementById("EnhancedReport");
        status.value = box.checked.toString();

        if (box.checked){

            document.getElementById('impression').checked = false;
            document.getElementById('redemption').checked = true;
            document.getElementById('transaction').checked = true;
            document.getElementById('markdown').checked = true;

            if(document.getElementById('byday').disabled)    document.getElementById('byday').disabled = false;
            document.getElementById('byday').checked = true;
            //document.getElementById("frequency").value = '1';
            if(document.getElementById('byweek').disabled)  document.getElementById('byweek').disabled = false;
            document.getElementById('byweek').checked = false;
            if(document.getElementById('CumulativeImpressions').disabled) document.getElementById('CumulativeImpressions').disabled = false;
            

            if(document.getElementById('CumulativeRedemptions').disabled)  document.getElementById('CumulativeRedemptions').disabled= false;
            document.getElementById('CumulativeRedemptions').checked = true;

            if(document.getElementById('CumulativeTransactions').disabled)  document.getElementById('CumulativeTransactions').disabled= false;
            document.getElementById('CumulativeTransactions').checked = true;

            if(document.getElementById('CumulativeMarkdowns').disabled)  document.getElementById('CumulativeMarkdowns').disabled = false;
            document.getElementById('CumulativeMarkdowns').checked = true;
            
        }
        else
        {
            document.getElementById('byweek').checked=false;
            document.getElementById('byweek').setAttribute('disabled', 'disabled');
            document.getElementById('byday').checked=false;
            document.getElementById('byday').setAttribute('disabled', 'disabled');
            document.getElementById('CumulativeImpressions').checked = false;
            document.getElementById('CumulativeImpressions').setAttribute('disabled', 'disabled');
            document.getElementById('CumulativeRedemptions').checked = false;
            document.getElementById('CumulativeRedemptions').setAttribute('disabled', 'disabled');
            document.getElementById('CumulativeTransactions').checked = false;
            document.getElementById('CumulativeTransactions').setAttribute('disabled', 'disabled');
            document.getElementById('CumulativeMarkdowns').checked = false;
            document.getElementById('CumulativeMarkdowns').setAttribute('disabled', 'disabled');
           
        }
        updateHiddenInputs();
        document.getElementById("hideDownloadButton").value = "visibility:hidden;";
        document.getElementById("download").style.dispplay = "none";
        document.mainform.exportRpt.value = "";
        document.mainform.submit();
        handleResultsButton();
    }


    function ByWeekClick() {
        var box = document.getElementById('byweek');

        if (box.checked && document.getElementById('byday').checked == false){
           return;
        }
        else{
           document.getElementById('byweek').checked = document.getElementById('byday').checked;
           document.getElementById('byday').checked = !document.getElementById('byweek').checked;
        }

        document.getElementById('hiddenByday').value = document.getElementById('byday').checked.toString();
        document.getElementById('hiddenByweek').value = document.getElementById('byweek').checked.toString();

        if(box.checked)
            document.getElementById("frequency").value = '1';
        else
            document.getElementById("frequency").value = '2';

        selectAllOffers();            
        updateHiddenInputs();
        document.mainform.exportRpt.value = "";
        document.mainform.submit();
        handleResultsButton();
    }


    function ByDayClick() {
        var box = document.getElementById('byday');
        //alert(box.checked.toString());
        if (box.checked && document.getElementById('byweek').checked == false){
           return;
        }
        else{
           document.getElementById('byday').checked = document.getElementById('byweek').checked;
           document.getElementById('byweek').checked = !document.getElementById('byday').checked;
        }
           
        document.getElementById('hiddenByday').value = document.getElementById('byday').checked.toString();
        document.getElementById('hiddenByweek').value = document.getElementById('byweek').checked.toString();

        if(box.checked)
            document.getElementById("frequency").value = '2';
        else
            document.getElementById("frequency").value = '1';

        selectAllOffers();            
        updateHiddenInputs();
        document.mainform.exportRpt.value = "";
        document.mainform.submit();
        handleResultsButton();
    }


    handleResultsButton();
    function reportlist_onclick() {

    }

 
    function updateCheckboxState(){
 
        updateHiddenInputs();
        document.mainform.exportRpt.value = "";
        document.mainform.submit();

        handleResultsButton();   
    }

    function updateSelectboxState(){
        document.getElementById("hideDownloadButton").value = "visibility:hidden;";

        updateHiddenInputs();
        document.mainform.exportRpt.value = "";
        document.mainform.submit();

        handleResultsButton();   
    }

    // For keeping data
    window.onunload = function(){
      //alert("In onbeforeunload");
      localStorage.setItem(frequency, $('#frequency').val());
      localStorage.setItem(hiddenImpression, $('#hiddenImpression').val());
      localStorage.setItem(hiddenRedemption, $('#hiddenRedemption').val());
      localStorage.setItem(hiddenTransaction, $('#hiddenTransaction').val());
      localStorage.setItem(hiddenMarkdown, $('#hiddenMarkdown').val());
      localStorage.setItem(reportingDate1, $('#reportingDate1').val());
      localStorage.setItem(reportingDate2, $('#reportingDate2').val());

      localStorage.setItem(hideDownloadButton, $('#hideDownloadButton').val());
      localStorage.setItem(hiddenEnhancedReport, $('#hiddenEnhancedReport').val());
    }

    window.onload = function(){
      //alert("In onload");
      var reportingDate1 = localStorage.getItem(reportingDate1);

      if (reportingDate1 !== null) $('#reportingDate1').val(reportingDate1);
      var reportingDate2 = localStorage.getItem(reportingDate2);
      if (reportingDate2 !== null) $('#reportingDate2').val(reportingDate2);
      var frequency = localStorage.getItem(frequency);
      if (frequency !== null) $('#frequency').val(frequency);
      var hiddenImpression = localStorage.getItem(hiddenImpression);
      if (hiddenImpression !== null) $('#hiddenImpression').val(hiddenImpression);    
      var hiddenRedemption = localStorage.getItem(hiddenRedemption);
      if (hiddenRedemption !== null) $('#hiddenRedemption').val(hiddenRedemption);
      var hiddenTransaction = localStorage.getItem(hiddenTransaction);
      if (hiddenTransaction !== null) $('#hiddenTransaction').val(hiddenTransaction);
      var hiddenMarkdown = localStorage.getItem(hiddenMarkdown);
      if (hiddenMarkdown !== null) $('#hiddenMarkdown').val(hiddenMarkdown);
    
      var hideDownloadBtn = localStorage.getItem(hideDownloadButton);
      if (hideDownloadBtn !== null) $('#hideDownloadButton').val(hideDownloadButton);
      var hiddenEnhancedReport = localStorage.getItem(hiddenEnhancedReport);
      //alert( hiddenEnhancedReport);
      if (hiddenEnhancedReport !== null) $('#hiddenEnhancedReport').val(hiddenEnhancedReport);
    }



</script>
<%
  If Session("OFFERIDS") IsNot Nothing Then
    OfferIdOptions = Session("OFFERIDS").ToString()

    RptTypeFromFolder = "5"
    RptStartDateFromFolder = GetReportStartingDate(Session("OFFERIDSONLY").ToString())
    Session.Remove("OFFERIDSONLY")
    RptEndDateFromFolder = Now().ToShortDateString()

    Session.Remove("OFFERIDS")
  Else If Session("RPTSTARTDATE") IsNot Nothing Then
    If ( Request.Form("hdnOfferIds") <> "" And Request.Form("hiddenReportingType") = "0" And Request.Form("hiddenReportingDate1") = "" And Request.Form("hiddenReportingDate2") = "") Then
      RptTypeFromFolder = "5"
      RptStartDateFromFolder = Session("RPTSTARTDATE")
      RptEndDateFromFolder = Now().ToShortDateString()
    End If
    Session.Remove("RPTSTARTDATE")
  Else
    RptTypeFromFolder = ""
    RptStartDateFromFolder = ""
    RptEndDateFromFolder = ""
  End If

  Send_Scripts()
  'Send_Scripts(New String() {"datePicker.js"})
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 8)


  If (Logix.UserRoles.AccessReports = False) Then
    Send_Denied(1, "perm.admin-reports")
    GoTo done
  End If

%>
<script type="text/javascript">
<% Send_Date_Picker_Terms() %>

</script>
<%	
  Send_Page()

GenerateReportBox()

done:
  Send_BodyEnd()
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()

  Common = Nothing

  Response.End()


ErrorTrap:
  Response.Write("<pre>" & Common.Error_Processor() & "</pre>")

  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
  
%>
