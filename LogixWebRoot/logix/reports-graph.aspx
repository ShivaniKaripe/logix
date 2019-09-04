<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%@ Import Namespace="System.IO" %>
<%
  Response.ContentType = "image/png"
  
  Dim CopientFileName As String = "reports-graph.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As System.Data.DataTable
  Dim OfferID As String
  Dim GraphField As String
  Dim GraphYTitle As String
  Dim ReportStartDate As Date
  Dim ReportEndDate As Date
  Dim objBitmap As New Bitmap(700, 450)
  Dim objGraphics As Graphics = Graphics.FromImage(objBitmap)
  Dim font As New Font("Arial", 12, FontStyle.Bold)
  Dim GraphType As Integer = 1
  Dim cx As Integer = 110
  Dim cy As Integer = 110
  Dim topYaxisVal As Double
  Dim redemptionRate As Double
  Dim dayCount As Integer = 0
  Dim RowCount As Integer = 0
  Dim i As Integer = 0
  Dim Frequency As Integer
  Dim bParsed As Boolean
  Dim bCumulative As Boolean
  Const YAxisIntervals As Integer = 6
  Const XAxisIntervals As Integer = 52
  
  On Error GoTo done
  MyCommon.Open_LogixWH()
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  MyCommon.AppName = "reports-graph.aspx"
  
  'load data
  bParsed = DateTime.TryParse(Request.QueryString("start"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ReportStartDate)
  If (Not bParsed) Then ReportStartDate = Now()
  bParsed = DateTime.TryParse(Request.QueryString("end"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, ReportEndDate)
  If (Not bParsed) Then ReportEndDate = Now
  dayCount = DateDiff(DateInterval.Day, ReportStartDate, ReportEndDate)
  OfferID = Request.QueryString("offerId")
  bParsed = Integer.TryParse(Request.QueryString("type"), GraphType)
  bParsed = Integer.TryParse(Request.QueryString("freq"), Frequency)
  
  Select Case GraphType
    Case 1
      GraphField = "NumImpressions"
      GraphYTitle = Copient.PhraseLib.Lookup("term.impressions", LanguageID)
      bCumulative = False
    Case 2
      GraphField = "NumImpressions"
      GraphYTitle = Copient.PhraseLib.Lookup("term.impressions", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")"
      bCumulative = True
    Case 3
      GraphField = "NumRedemptions"
      GraphYTitle = Copient.PhraseLib.Lookup("term.redemptions", LanguageID)
      bCumulative = False
    Case 4
      GraphField = "NumRedemptions"
      GraphYTitle = Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")"
      bCumulative = True
    Case 5
      GraphField = "AmountRedeemed"
      GraphYTitle = Copient.PhraseLib.Lookup("term.markdown", LanguageID) & "($)"
      bCumulative = False
    Case 6
      GraphField = "AmountRedeemed"
      GraphYTitle = Copient.PhraseLib.Lookup("term.markdown", LanguageID) & "($ " & Copient.PhraseLib.Lookup("term.cumulative", LanguageID) & ")"
      bCumulative = True
    Case 7
      GraphField = "RedemptionRate"
      GraphYTitle = Copient.PhraseLib.Lookup("term.redemptionrate", LanguageID)
      bCumulative = False
    Case Else
      GraphField = "NumImpressions"
      GraphYTitle = Copient.PhraseLib.Lookup("term.impressions", LanguageID)
      bCumulative = False
  End Select
  
  
  
  MyCommon.QueryStr = "select NumImpressions, NumRedemptions, AmountRedeemed, 0.0000 as 'RedemptionRate', ReportingDate from OfferReporting with (nolock) " & _
                      "where OfferID = " & OfferID & " " & _
                      "and ReportingDate between '" & Copient.commonShared.ConvertToSqlDate(ReportStartDate, MyCommon.GetAdminUser.Culture) & "' " & _
                      "and '" & Copient.commonShared.ConvertToSqlDate(ReportEndDate, MyCommon.GetAdminUser.Culture) & "' " & _
                      "order by ReportingDate"
  dst = MyCommon.LWH_Select
  RowCount = dst.Rows.Count
  objGraphics.Clear(Color.White)
  objGraphics.FillRectangle(Brushes.White, 0, 0, 700, 450)
  objGraphics.FillRectangle(New SolidBrush(Color.FromArgb(235, 235, 235)), 75, 50, 600, 300)
  
  If (RowCount > 0) Then
    ' adjust for weekly, if necessary
    If (Frequency = 1) Then
      dst = RollupReportWeek(dst, ReportStartDate, ReportEndDate)
      RowCount = dst.Rows.Count
    End If
    For i = 0 To (RowCount - 1)
      If (bCumulative) Then
        topYaxisVal += MyCommon.NZ(dst.Rows(i).Item(GraphField), 0)
      Else
        ' calculate the Redemption Rate, if necessary, and set it between 0 and 100
        If (GraphType = 7 And Frequency <> 1) Then
          dst.Rows(i).Item("RedemptionRate") = (MyCommon.NZ(dst.Rows(i).Item("NumRedemptions"), 0) / MyCommon.NZ(dst.Rows(i).Item("NumImpressions"), 1)) * 100
        End If
        If (topYaxisVal < MyCommon.NZ(dst.Rows(i).Item(GraphField), 0)) Then
          topYaxisVal = MyCommon.NZ(dst.Rows(i).Item(GraphField), 0)
        End If
      End If
    Next
  End If
  
  MyCommon.Close_LogixWH()
  
  'write graph header
  objGraphics.DrawString(Copient.PhraseLib.Lookup("reports-graph.redemptionreport", LanguageID), New Font("Arial", 16, FontStyle.Bold), Brushes.Black, 250, 5)
  
  'write horizontal graph title
  If (Frequency = 1) Then
    objGraphics.DrawString(Copient.PhraseLib.Lookup("reports.weekly", LanguageID), font, Brushes.Black, 270, 430)
  Else
    objGraphics.DrawString(Copient.PhraseLib.Lookup("reports.daily", LanguageID), font, Brushes.Black, 270, 430)
  End If
  
  'draw vertical graph title
  objGraphics.ResetTransform()
  objGraphics.TranslateTransform(-cx, -cy, Drawing2D.MatrixOrder.Append)
  objGraphics.RotateTransform(270, Drawing2D.MatrixOrder.Append)
  objGraphics.TranslateTransform(cx, cy, Drawing2D.MatrixOrder.Append)
  objGraphics.DrawString(GraphYTitle, font, Brushes.Black, 0, 0)
  objGraphics.ResetTransform()
  
  'draw x-axis
  objGraphics.DrawLine(Pens.Black, 75, 350, 675, 350)
  objGraphics.DrawLine(Pens.Black, 75, 351, 675, 351)
  
  'draw y-axis
  objGraphics.DrawLine(Pens.Black, 75, 50, 75, 350)
  objGraphics.DrawLine(Pens.Black, 74, 50, 74, 350)
  
  'draw intervals
  objGraphics.DrawLine(Pens.LightGray, 76, 300, 675, 300)
  objGraphics.DrawLine(Pens.LightGray, 76, 250, 675, 250)
  objGraphics.DrawLine(Pens.LightGray, 76, 200, 675, 200)
  objGraphics.DrawLine(Pens.LightGray, 76, 150, 675, 150)
  objGraphics.DrawLine(Pens.LightGray, 76, 100, 675, 100)
  objGraphics.DrawLine(Pens.LightGray, 76, 50, 675, 50)
  
  'draw the x- scale interval
  Dim xSpacing As Single = 10
  Dim tickFont As New Font("Arial", 9, FontStyle.Regular)
  Dim vertFormat As New StringFormat(StringFormatFlags.DirectionVertical)
  
  xSpacing = CInt(600 / (RowCount + 1))
  If (RowCount >= XAxisIntervals) Then
    xSpacing = 11.4
    For i = 0 To XAxisIntervals - 1
      objGraphics.DrawLine(Pens.Black, 85 + (i * xSpacing), 346, 85 + (i * xSpacing), 356)
      'avoid cluttering the dates by display ever other date
      If (i Mod 2 = 0) Then
        objGraphics.DrawString(MyCommon.NZ(dst.Rows(i).Item("ReportingDate"), ""), tickFont, Brushes.Black, 77 + (i * xSpacing), 360, vertFormat)
      End If
    Next
  Else
    xSpacing = CInt(xSpacing)
    For i = 1 To RowCount
      objGraphics.DrawLine(Pens.Black, 85 + (i * xSpacing), 346, 85 + (i * xSpacing), 356)
      If (RowCount <= 30) Then
        objGraphics.DrawString(MyCommon.NZ(dst.Rows(i - 1).Item("ReportingDate"), ""), tickFont, Brushes.Black, 77 + (i * xSpacing), 360, vertFormat)
      Else
        If ((i + 1) Mod 2 = 0) Then
          objGraphics.DrawString(MyCommon.NZ(dst.Rows(i - 1).Item("ReportingDate"), ""), tickFont, Brushes.Black, 77 + (i * xSpacing), 360, vertFormat)
        End If
      End If
    Next
  End If
  
  'draw the y-scale interval
  Dim ySpacing As Integer = 10
  Dim topYAxisValue As Integer
  Dim yIntervalLen As Integer
  topYaxisVal += 1
  ySpacing = CInt(topYaxisVal / YAxisIntervals)
  
  For i = 1 To YAxisIntervals
    objGraphics.DrawLine(Pens.Black, 71, 0 + (i * 50), 79, 0 + (i * 50))
    topYAxisValue = ySpacing * i
    yIntervalLen = (topYAxisValue.ToString.Length + 0.5) * tickFont.SizeInPoints
    objGraphics.DrawString(topYAxisValue, tickFont, Brushes.Black, 75 - yIntervalLen, 344 - (i * 50))
  Next
  
  'plot the points from the data on the graph
  Dim pixelsPerUnit As Double
  Dim yPt As Integer
  Dim points As ArrayList = New ArrayList(30)
  Dim pt, leftPt, rightPt As Point
  Dim cumulativeVal As Double
  
  If (topYAxisValue = 0) Then topYAxisValue = 1
  pixelsPerUnit = 300 / CInt(topYAxisValue)
  xSpacing = CInt(600 / (RowCount + 1))
  
  ' first add the origin
  pt = New Point(75, 350)
  points.Add(pt)
  
  For i = 1 To RowCount
    If (bCumulative) Then
      cumulativeVal += MyCommon.NZ(dst.Rows(i - 1).Item(GraphField), 0)
      yPt = pixelsPerUnit * cumulativeVal
    Else
      yPt = pixelsPerUnit * MyCommon.NZ(dst.Rows(i - 1).Item(GraphField), 0)
    End If
    pt = New Point(83 + (i * xSpacing), 350 - yPt)
    points.Add(pt)
    objGraphics.DrawRectangle(Pens.Black, pt.X, pt.Y, 3, 3)
  Next
  
  'connect the lines of the data
  For i = 0 To points.Count - 1
    leftPt = points(i)
    If (i < points.Count - 1) Then
      rightPt = points(i + 1)
    Else
      rightPt = leftPt
    End If
    If (i > 0) Then
      objGraphics.DrawLine(Pens.Red, leftPt.X, leftPt.Y + 2, rightPt.X, rightPt.Y + 2)
    Else
      objGraphics.DrawLine(Pens.Red, leftPt.X, leftPt.Y, rightPt.X, rightPt.Y + 2)
    End If
  Next
  
  'objBitmap.Save(Response.OutputStream, ImageFormat.Jpeg)
  Dim MemStream As New System.IO.MemoryStream()
  objBitmap.Save(MemStream, ImageFormat.Png)
  MemStream.WriteTo(Response.OutputStream)
  
  Response.End()
%>

<script runat="server">
  Function RollupReportWeek(ByVal dst As DataTable, ByVal ReportStartDate As Date, ByVal ReportEndDate As Date) As DataTable
    Dim dstWeek As New DataTable
    Dim i, j As Integer
    Dim numRedeem As Integer
    Dim numImpression As Integer
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
              If (j = dst.DefaultView.Count - 1) Then
                row = dst.DefaultView(j).Row
                row.Item("ReportingDate") = CurrentStart
                row.Item("NumRedemptions") = numRedeem
                row.Item("AmountRedeemed") = amtRedeem
                row.Item("NumImpressions") = numImpression
                If (numImpression = 0) Then
                  row.Item("RedemptionRate") = 0.0
                Else
                  row.Item("RedemptionRate") = (numRedeem / numImpression) * 100
                End If
                dstWeek.ImportRow(row)
              End If
            Next
          Else
            row = dstWeek.NewRow()
            row.Item("ReportingDate") = CurrentStart
            row.Item("NumRedemptions") = 0
            row.Item("AmountRedeemed") = 0.0
            row.Item("NumImpressions") = 0
            row.Item("RedemptionRate") = 0.0
            dstWeek.Rows.Add(row)
          End If
          numRedeem = 0
          amtRedeem = 0.0
          numImpression = 0
          CurrentStart = CurrentEnd.AddDays(1)
          CurrentEnd = CurrentStart.AddDays(6)
        End If
      Next
    End If
    Return dstWeek
  End Function
</script>

<%
done:
  objGraphics.Clear(Color.White)
  objGraphics.DrawString(Err.Source & vbNewLine & Err.Description, font, Brushes.Black, 0, 0)
  objBitmap.Save(Response.OutputStream, ImageFormat.Jpeg)
  ' clean up...
  Response.Flush()
  objGraphics.Dispose()
  objBitmap.Dispose()
%>
