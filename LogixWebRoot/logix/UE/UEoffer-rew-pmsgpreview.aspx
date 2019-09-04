<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew-pmsgpreview.aspx 
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
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim PrinterTypeID As Long
  Dim PrinterName As String
  Dim PageWidth As Integer
  Dim FixedWidthFont As Boolean
  Dim RawMessage As String
  Dim MESSAGE As String
  Dim TierLevel As String
  Dim x As Integer = 0
  Dim y As Integer = 0
  Dim i As Integer = 0
  Dim PrinterTag As String
  Dim PreviewText As String
  Dim PrintLine As String
  Dim Found As Boolean = False
  Dim ExtFound As Boolean = False
  Dim NewLine As String
  Dim NewBRLine As String
  Dim FormatLine As String
  Dim WrappedPrintLine As String
  Dim RawTextLine As String
  Dim FormatChar As String
  Dim ExtFormatChar As String
  Dim ENDMESSAGE As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  Dim TabWidth As Integer = 0
  Dim TabMatch As Match
  Dim AdjustlineWidth As Integer = 0
  Dim AdjustlineMatch As Match
  Dim FontType As Integer = -1
  Dim FontMatch As Match
  
  Dim SVScorecardText As String = ""
  Dim ScorecardText As String = ""
  Dim ScorecardTextBottom As String = ""
  Dim ScorecardTextTop As String = ""
  Dim DScorecardText As String = ""
  Dim DScorecardTextBottom As String = ""
  Dim DScorecardTextTop As String = ""
  Dim TempStyleStr As String
  Dim StyleStr As String = ""
  Dim AlignmentStr As String = ""
  Dim StartStylePos As Integer
  Dim EndStylePos As Integer
  Dim StyleTag As String
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  ' This 2D array will hold the replace and replacement for real-time engine tags, just add or modify here
  Dim SubTags(,) As String = { _
  {"|RewExpAmt|", "|POSNUM|", "|CUSTOMERID|", "|FIRSTNAME|", "|LASTNAME|", "|TSD|", "|LYTS|", "|HHSTD|", "|CURRDATE|", "|OFFERSTART|", "|OFFEREND|", "|TOTALPOINTS|", "|PTSASPEN|", "|ACCUMAMT|", "|REMAINAMT|", "|FTS|", "|FTSPG|", "|BACKSLASH|", "\n", "\2n", "\p", "\l", "\t", "\f", "\3j", "\0j", "\i", "\h", "\w", "\q", "\s", "\m", "\BX", "\BX.", "\1BX.", "\0BX."}, _
  {"####", "####", "###################", "John", "Smith", "000.00", "000.00", "000.00", "xx/xx/xxxx", "xx/xx/xxxx", "xx/xx/xxxx", "xx", "xx", "000.00", "000.00", "###.##", "###.##", "\", "<br />", "", "|NORMAL|", "********************************************", "\0t", "\0f", "|CENTER|", "|LEFT|", "|THIN|", "|HIGH|", "|WIDE|", "|QUAD|", "|SMALL|", "|MEDIUM|", "\1BX{,,}", "\1BX{,,}", "\1BX{,,}", "\0BX"}}
  
  Dim SubExpTags(,) As String = { _
  {"\|NET\#[[0-9]+]\|", "\|EARNED\#[[0-9]+]\|", "\|INITIAL#\[[0-9]+]\|", "\|REDEEMED#\[[0-9]+]\|", "\|NET\$\[[0-9]+]\|", "\|INITIAL\$\[[0-9]+]\|", "\|EARNED\$\[[0-9]+]\|", "\|REDEEMED\$\[[0-9]+]\|", "\|SVBAL\[[0-9]+\]\|", "\|SVVAL\[[0-9]+\]\|", "\|UPCA\[[0-9]+\]\|", "\|UPCB\[[0-9]+\]\|", "\|SVRATIO\[[0-9]+\,[0-9]+\,[0-9,.]+\]\|"}, _
  {"###", "###", "###", "###", "###.##", "###.##", "###.##", "###.##", "###", "###.##", "<img src='/images/upca.png' \/>", "<img src='/images/upcb.png' \/>", "###.##"}}
  
  Response.Expires = 0
  MyCommon.AppName = "offer-rew-pmsgpreview.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  Send_HeadBegin("term.preview")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  
  ' Get printer type from the URL
  PrinterTypeID = Request.QueryString("PrinterTypeID")
  RawMessage = Request.QueryString("message")
  
  ' replace all the < and > 
  ' before processing
  RawMessage = RawMessage.Replace("<", "&lt;")
  RawMessage = RawMessage.Replace(">", "&gt;")
  
  ' Pull out details (name, page width, etc.) for that printer
  MyCommon.QueryStr = "select top 1 PTy.PrinterTypeID ,PTy.PageWidth, PTy.FixedWidthFont, PTy.Name as PrinterName " & _
                      "from PrinterTypes as PTy with (NoLock) " & _
                      "where PrinterTypeID=" & PrinterTypeID & ";"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    PrinterName = row.Item("PrinterName")
    PageWidth = row.Item("PageWidth")
    FixedWidthFont = row.Item("FixedWidthFont")
  Next
  
  ' Calculate how many pixels wide the pmsgpreviewbody div should be, based on the page width
  Dim PageWidthPixels As Integer
  If (PageWidth = 0) Or (PageWidth > 50) Then
    PageWidthPixels = 400
  Else
    PageWidthPixels = PageWidth * 8
  End If
  
  ' Generate custom CSS to style the DIV
  Send("<style type=""text/css"">")
  Send("#pmsgpreview {")
  Send("  width: " & (PageWidthPixels + 2) & "px;")
  Send("}")
  Send("* html #pmsgpreview {")
  Send("  width: " & (PageWidthPixels + 40) & "px;")
  Send("  white-space: nowrap;")
  Send("}")
  Send("#pmsgpreviewbody {")
  Send("  border: 0;")
  Send("  padding: 0;")
  Send("  font-size: 13px;")
  If (FixedWidthFont = False) Then
    Send("  font-family: Verdana, Arial;")
  Else
    Send("  font-family: monospace;")
  End If
  If (PrinterTypeID = 4) Then
    Send("  overflow-x: auto;")
    Send("  white-space: normal;")
  Else
    Send("  overflow-x: hidden;")
  End If
  
  Send("  width: " & (PageWidthPixels + 2) & "px;")
  ' Note: the +2 added to PageWidthPixels shouldn't be necessary; it's only there
  ' because of a problem sizing medium text in IE that couldn't otherwise be solved
  Send("}")
  Send("</style>")
  
  ' Here we will format the message.
  ' The message is currently in RawMessage, so let's get it formatted correctly, the same as the local server
  ' First, look for and replace tags
  
  MESSAGE = MyCommon.NZ(RawMessage, " ")
  For x = 0 To SubTags.GetUpperBound(1)
    MESSAGE = Replace(MESSAGE, SubTags(0, x), SubTags(1, x))
  Next
  
  For x = 0 To SubExpTags.GetUpperBound(1)
    If MESSAGE = "" Then
      MESSAGE = " "
    End If
    MESSAGE = System.Text.RegularExpressions.Regex.Replace(MESSAGE, SubExpTags(0, x), SubExpTags(1, x))
  Next
  
  If PrinterTypeID = 4 Then
    ' ...it's HTML
    Dim strArray() As String = MESSAGE.Split(vbLf)
    ' ok we split lets blank out MESSAGE so we can refill it
    MESSAGE = ""
    ENDMESSAGE = ""
    ' get the tags
    MyCommon.QueryStr = "select PrinterTypeID, '|' + tag + '|' as Tag, isnull(pt.previewchars,'') as PreviewChars from MarkupTags as MT with (NoLock) " & _
                        "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID;"
    rst = MyCommon.LRT_Select
    For Each PrintLine In strArray
      ' now lets search on each printer tag to replace
      ' Pull out details (name, page width, etc.) for that printer
      RawTextLine = PrintLine
      For Each row In rst.Rows
        PrinterTag = row.Item("tag")
        PreviewText = row.Item("previewchars")
        Dim rowPrinterTypeID = MyCommon.NZ(row.Item("PrinterTypeID"), 0)
        If (rowPrinterTypeID = 4 AndAlso InStr(RawTextLine, PrinterTag)) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, PreviewText)
          ENDMESSAGE = "</span>"
        End If
      Next
      ' now get any remaining tags
      For Each row In rst.Rows
        PrinterTag = row.Item("tag")
        PreviewText = row.Item("previewchars")
        If (InStr(RawTextLine, PrinterTag)) Then
          ' MESSAGE = MESSAGE & "<span style=""color:#ff0000;"">no html equiv replacing with nothing</span><br />"
          RawTextLine = Replace(RawTextLine, PrinterTag, "")
        End If
      Next
      MESSAGE = MESSAGE & RawTextLine & ENDMESSAGE & "<br />"
    Next
  ElseIf PrinterTypeID = 6 Or PrinterTypeID = 8 Then
    ' ...its an ACS POS printer
    Dim strArray() As String = MESSAGE.Split(vbLf)
    ' ok we split lets blank out MESSAGE so we can refill it
    MESSAGE = ""
    ENDMESSAGE = ""
    ' get the tags
    MyCommon.QueryStr = "select PrinterTypeID, '|' + tag + '|' as Tag, isnull(pt.previewchars,'') as PreviewChars from MarkupTags as MT with (NoLock) " & _
                        "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID;"
    rst = MyCommon.LRT_Select
    For Each PrintLine In strArray
      ' now lets search on each printer tag to replace
      ' Pull out details (name, page width, etc.) for that printer
      RawTextLine = PrintLine
      RawTextLine = LineWrap(PageWidth, RawTextLine)
      For Each row In rst.Rows
        PrinterTag = row.Item("tag")
        PreviewText = row.Item("previewchars")
        If (InStr(RawTextLine, PrinterTag) And MyCommon.NZ(row.Item("PrinterTypeID"), 0) = 6) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, PreviewText)
          ENDMESSAGE = "</span>"
        End If
      Next
      ' now get any remaining tags
      For Each row In rst.Rows
        PrinterTag = row.Item("tag")
        PreviewText = row.Item("previewchars")
        If (InStr(RawTextLine, PrinterTag)) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, "")
        End If
      Next
      MESSAGE = MESSAGE & RawTextLine & ENDMESSAGE
    Next
  ElseIf PrinterTypeID = 3 Or PrinterTypeID = 7 Then
    ' ...its an IBM POS printer
    Dim strArray() As String = MESSAGE.Split(vbLf)
    ' ok we split lets blank out MESSAGE so we can refill it
    MESSAGE = ""
    ' get the tags
    MyCommon.QueryStr = "select '|' + tag + '|' as tag, isnull(pt.previewchars,'') as PreviewChars from MarkupTags as MT with (NoLock) " & _
                        "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID " & _
                        "where PrinterTypeID=" & PrinterTypeID & ";"
    rst = MyCommon.LRT_Select
    For Each PrintLine In strArray
      ' now lets search on each printer tag to replace
      ' Pull out details (name, page width, etc.) for that printer
      Found = False
      ExtFound = False
      ' check if were over 38 and wrap if we are
      ExtFormatChar = ""
      FormatChar = ""
      RawTextLine = PrintLine
      ' eat all the tags out of the line to wrap
      For Each row In rst.Rows
        PrinterTag = MyCommon.NZ(row.Item("tag"), "")
        PreviewText = MyCommon.NZ(row.Item("previewchars"), "")
        If (InStr(RawTextLine, PrinterTag)) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, "")
        End If
      Next
      ' RawTextLine now contains unformatted text of the line lets wrap it
      RawTextLine = LineWrap(PageWidth, RawTextLine)
      WrappedPrintLine = PrintLine
      For Each row In rst.Rows
        PrinterTag = MyCommon.NZ(row.Item("tag"), "")
        PreviewText = MyCommon.NZ(row.Item("previewchars"), "")
        If (InStr(WrappedPrintLine, PrinterTag) And Not Found And PrinterTag <> "|INV|" And PrinterTag <> "|V|" And PrinterTag <> "|U|") Then
          Found = True
          FormatChar = PreviewText
        ElseIf (InStr(WrappedPrintLine, PrinterTag) And Not ExtFound) Then
          ExtFound = True
          ExtFormatChar = PreviewText
        End If
      Next
      If Found Then RawTextLine = RawTextLine & "</span>"
      If ExtFound Then RawTextLine = RawTextLine & "</span>"
      MESSAGE = MESSAGE & FormatChar & ExtFormatChar & RawTextLine
    Next
  ElseIf PrinterTypeID = 9 Then
    ' Rite Aid printer
    Dim strArray() As String = MESSAGE.Split(vbLf)
    ' Message split; now blank out MESSAGE so we can refill it
    MESSAGE = ""
    ENDMESSAGE = ""
    ' Get all tags for the Rite Aid printer
    MyCommon.QueryStr = "select PrinterTypeID, Tag, isnull(PT.PreviewChars,'') as PreviewChars from MarkupTags as MT with (NoLock) " & _
                        "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID " & _
                        "where PrinterTypeID=9 and CentralRendered=1;"
    rst = MyCommon.LRT_Select
    MyCommon.QueryStr = "select PrinterTypeID, '|' + Tag + '|' as Tag, isnull(PT.PreviewChars,'') as PreviewChars from MarkupTags as MT with (NoLock) " & _
                        "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID " & _
                        "where PrinterTypeID=9 and CentralRendered=0;"
    rst2 = MyCommon.LRT_Select
    For Each PrintLine In strArray
      RawTextLine = PrintLine
      'RawTextLine = LineWrap(PageWidth, RawTextLine)
      'First, extract a few special tag value where the user-defined input needs to be preserved
      TabMatch = Regex.Match(RawTextLine, "\\\d+t")
      If TabMatch IsNot Nothing Then
        TabWidth = MyCommon.Extract_Val(TabMatch.Value)
      End If
      AdjustlineMatch = Regex.Match(RawTextLine, "\\\d+l")
      If AdjustlineMatch IsNot Nothing Then
        AdjustlineWidth = MyCommon.Extract_Val(AdjustlineMatch.Value)
      End If
      FontMatch = Regex.Match(RawTextLine, "\\\d+f")
      If FontMatch IsNot Nothing Then
        FontType = MyCommon.Extract_Val(FontMatch.Value)
      End If
      'For each line in the message, replace each centrally-rendered tag (which will likely contain
      'user-defined values) with the default form of the tag present in the MarkupTags.Tag field.
      'This allows us in the subsequent step to do a conventional search/replace.
      RawTextLine = Regex.Replace(RawTextLine, "\\\d+([Aflt])", "\^1$1")
      RawTextLine = Regex.Replace(RawTextLine, "\\A{.*}", "\A{^1}")
      RawTextLine = Regex.Replace(RawTextLine, "\\V{.*}", "\V{^1}")
      RawTextLine = Regex.Replace(RawTextLine, "\\Z{.*}", "\Z{^1}")
      RawTextLine = Regex.Replace(RawTextLine, "\\HX{[0-9A-Fa-f]+}", "\HX{^1}")
      RawTextLine = Regex.Replace(RawTextLine, "\\MP{.*}", "\MP{^1}")
      RawTextLine = Regex.Replace(RawTextLine, "\\MSG{.*}", "\MSG{^1}")
      RawTextLine = Regex.Replace(RawTextLine, "\\Z{.*,.*}", "\Z{^1,^2}")
      RawTextLine = Regex.Replace(RawTextLine, "\\CUT{.*,.*}", "\CUT{^1,^2}")
      RawTextLine = Regex.Replace(RawTextLine, "\\BA{.*,.*}", "\BA{^1,^2}")
      RawTextLine = Regex.Replace(RawTextLine, "\\BC{.*,.*,.*}", "\BC{^1,^2,^3}")
      RawTextLine = Regex.Replace(RawTextLine, "\\IF{.*,.*,.*}", "\IF{^1,^2,^3}")
      RawTextLine = Regex.Replace(RawTextLine, "\\1BX{.*,.*,.*}", "\1BX{^1,^2,^3}")
      RawTextLine = Regex.Replace(RawTextLine, "\\L{.*,.*,.*,.*}", "\L{^1,^2,^3,^4}")
      RawTextLine = Regex.Replace(RawTextLine, "\\BX{.*,.*,.*,.*}", "\BX{^1,^2,^3,^4}")
      
      'Now that the tags have been cleaned up, replace them with their preview equivalents
      For Each row In rst.Rows
        PrinterTag = row.Item("Tag")
        PreviewText = MyCommon.NZ(row.Item("PreviewChars"), "")
        If (InStr(RawTextLine, PrinterTag)) Then
          If (PrinterTag = "\^1t") Then
            PreviewText = ""
            For i = 1 To TabWidth
              PreviewText &= row.Item("PreviewChars")
            Next
          ElseIf (PrinterTag = "\^1l") Then
            PreviewText = ""
            For i = 1 To AdjustlineWidth
              PreviewText &= row.Item("PreviewChars")
            Next
          ElseIf (PrinterTag = "\^1f") Then
            If FontType = 0 Then
              PreviewText = "</span>"
            ElseIf FontType = 1 Then
              PreviewText = "<span style=""letter-spacing:12px;"">"
            ElseIf FontType = 2 Then
              PreviewText = "<span style=""line-height:200%;"">"
            ElseIf FontType = 3 Then
              PreviewText = "<span style=""font-size:30px;"">"
            ElseIf FontType = 4 Then
              PreviewText = "<span style=""font-size:9px;"">"
            ElseIf FontType = 5 Then
              PreviewText = "<span style=""font-size:9px;letter-spacing:12px;"">"
            ElseIf FontType = 6 Then
              PreviewText = "<span style=""font-size:9px;line-height:200%;"">"
            End If
          End If
          RawTextLine = Replace(RawTextLine, PrinterTag, PreviewText)
          ENDMESSAGE = "</span>"
        End If
      Next
      
      'Finally do substitution on the non-centrally rendered tags
      For Each row2 In rst2.Rows
        PrinterTag = row2.Item("Tag")
        PreviewText = MyCommon.NZ(row2.Item("PreviewChars"), "")
        If (InStr(RawTextLine, PrinterTag)) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, PreviewText)
        End If
      Next
      
      'The final result:
      MESSAGE = MESSAGE & RawTextLine & ENDMESSAGE
      
      'Reinitialize variables for next line
      TabWidth = 0
      AdjustlineWidth = 0
      FontType = -1
      PrinterTag = ""
      PreviewText = ""
      
    Next

  ElseIf PrinterTypeID = 10 Or PrinterTypeID = 11 Then
    Send("<!-- processing PrinterTypeID=10 -->")
    ' ...its an Advanced Store printer
    StyleStr = ""
    AlignmentStr = ""
    Dim strArray() As String = MESSAGE.Split(vbLf)
    ' ok we split lets blank out MESSAGE so we can refill it
    MESSAGE = ""
    ' get the tags
    MyCommon.QueryStr = "select isnull(PrinterTypeID, 0) as PrinterTypeID, '|' + tag + '|' as Tag, isnull(pt.previewchars,'') as PreviewChars, AlignmentTag from MarkupTags as MT with (NoLock) " & _
                        "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID;"
    rst = MyCommon.LRT_Select
    For Each PrintLine In strArray
      
      RawTextLine = PrintLine
      'RawTextLine = LineWrap(PageWidth, RawTextLine)
      RawTextLine = ProcessMsgLine(PrinterTypeID, RawTextLine, rst, StyleStr, AlignmentStr)
      ' now get any remaining tags
      For Each row In rst.Rows
        PrinterTag = row.Item("tag")
        PreviewText = row.Item("previewchars")
        If (InStr(RawTextLine, PrinterTag)) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, "")
        End If
      Next
      MESSAGE = MESSAGE & RawTextLine
      If UCase(Right(Trim(MESSAGE), 9)) = "</CENTER>" Or UCase(Right(Trim(MESSAGE), 4)) = "</P>" Then
        'leave the string alone, the center tag will cause a wrap to the next line
      Else
        MESSAGE = MESSAGE & "<br />"
      End If
    Next

    
  End If ' PrinterTypeID = 3 Or PrinterTypeID = 7
  
  'Special search/replace for scorecards, which have variable previews based on whether (and where) there are total lines
  MyCommon.QueryStr = "select ScorecardID, ScorecardTypeID, Description, Bold, PrintTotalLine, TotalLinePosition from Scorecards with (NoLock) " & _
                      "where Deleted=0 order by ScorecardID;"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    SVScorecardText = FormatScorecardLine("ueoffer-rew-pmsgpreview.SampleSVProgA", "3", PageWidth) & ControlChars.CrLf & FormatScorecardLine("ueoffer-rew-pmsgpreview.SampleSVProgB", "5", PageWidth)
    ScorecardText = FormatScorecardLine("ueoffer-rew-pmsgpreview.SamplePtProgA", "3", PageWidth) & ControlChars.CrLf & FormatScorecardLine("ueoffer-rew-pmsgpreview.SamplePtProgB", "5", PageWidth)
    ScorecardTextBottom = ScorecardText & ControlChars.CrLf & Space(38) & ControlChars.CrLf & "<span style=""font-weight:bold;"">" & FormatScorecardLine("term.total", "8", PageWidth) & "</span>"
    ScorecardTextTop = "<span style=""font-weight:bold;"">" & FormatScorecardLine("term.ScorecardTitle", "8", PageWidth) & "</span" & ControlChars.CrLf & Space(38) & ControlChars.CrLf & FormatScorecardLine("ueoffer-rew-pmsgpreview.SamplePtProgA", "3", PageWidth) & ControlChars.CrLf & FormatScorecardLine("ueoffer-rew-pmsgpreview.SamplePtProgB", "5", PageWidth) & "<br />"
    DScorecardText = "<br \>" & FormatScorecardLine("ueoffer-rew-pmsgpreview.DiscountA", "0.50", PageWidth) & ControlChars.CrLf & FormatScorecardLine("ueoffer-rew-pmsgpreview.DiscountB", "0.50", PageWidth) & "<br \>"
    DScorecardTextBottom = "<br />" & DScorecardText & ControlChars.CrLf & Space(38) & "<br /><span style=""font-weight:bold;"">" & FormatScorecardLine("term.total", "$  1.00", PageWidth) & "</span><br />"
    DScorecardTextTop = "<br /><span style=""font-weight:bold;"">" & FormatScorecardLine("term.ScorecardTitle", "$  1.00", PageWidth) & "</span>" & ControlChars.CrLf & Space(38) & ControlChars.CrLf & FormatScorecardLine("ueoffer-rew-pmsgpreview.DiscountA", "0.50", PageWidth) & ControlChars.CrLf & FormatScorecardLine("ueoffer-rew-pmsgpreview.DiscountB", "0.50", PageWidth) & "<br />"

    If (MyCommon.NZ(row.Item("ScorecardTypeID"), 0) = 1) Then
      ' Points scorecard
      If (MyCommon.NZ(row.Item("PrintTotalLine"), 0) = 0) Then
        MESSAGE = Regex.Replace(MESSAGE, "\|SCORECARD\[" & row.Item("ScorecardID") & "\]\|", IIf(MyCommon.NZ(row.Item("Bold"), 0) = 0, ScorecardText, "<b>" & ScorecardText & "</b>"))
      Else
        If (MyCommon.NZ(row.Item("TotalLinePosition"), 0) = 0) Then
          MESSAGE = Regex.Replace(MESSAGE, "\|SCORECARD\[" & row.Item("ScorecardID") & "\]\|", IIf(MyCommon.NZ(row.Item("Bold"), 0) = 0, ScorecardTextBottom, "<b>" & ScorecardTextBottom & "</b>"))
        ElseIf (MyCommon.NZ(row.Item("TotalLinePosition"), 0) = 1) Then
          ScorecardTextTop = Regex.Replace(ScorecardTextTop, "Scorecard Title               ", Left(MyCommon.NZ(row.Item("Description"), ""), 30).PadRight(30))
          MESSAGE = Regex.Replace(MESSAGE, "\|SCORECARD\[" & row.Item("ScorecardID") & "\]\|", IIf(MyCommon.NZ(row.Item("Bold"), 0) = 0, ScorecardTextTop, "<b>" & ScorecardTextTop & "</b>"))
        End If
      End If
    ElseIf (MyCommon.NZ(row.Item("ScorecardTypeID"), 0) = 2) Then
      ' Stored value scorecard
      MESSAGE = Regex.Replace(MESSAGE, "\|SVSCORECARD\[" & row.Item("ScorecardID") & "\]\|", IIf(MyCommon.NZ(row.Item("Bold"), 0) = 0, SVScorecardText, "<b>" & SVScorecardText & "</b>"))
    ElseIf (MyCommon.NZ(row.Item("ScorecardTypeID"), 0) = 3) Then
      ' Discount scorecard
      If (MyCommon.NZ(row.Item("PrintTotalLine"), 0) = 0) Then
        MESSAGE = Regex.Replace(MESSAGE, "\|DSCORECARD\[" & row.Item("ScorecardID") & "\]\|", IIf(MyCommon.NZ(row.Item("Bold"), 0) = 0, DScorecardText, "<b>" & DScorecardText & "</b>"))
      Else
        If (MyCommon.NZ(row.Item("TotalLinePosition"), 0) = 0) Then
          MESSAGE = Regex.Replace(MESSAGE, "\|DSCORECARD\[" & row.Item("ScorecardID") & "\]\|", IIf(MyCommon.NZ(row.Item("Bold"), 0) = 0, DScorecardTextBottom, "<b>" & DScorecardTextBottom & "</b>"))
        ElseIf (MyCommon.NZ(row.Item("TotalLinePosition"), 0) = 1) Then
          DScorecardTextTop = Regex.Replace(DScorecardTextTop, "Scorecard Title               ", Left(MyCommon.NZ(row.Item("Description"), ""), 30).PadRight(30))
          MESSAGE = Regex.Replace(MESSAGE, "\|DSCORECARD\[" & row.Item("ScorecardID") & "\]\|", IIf(MyCommon.NZ(row.Item("Bold"), 0) = 0, DScorecardTextTop, "<b>" & DScorecardTextTop & "</b>"))
        End If
      End If
    End If
  Next
  
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.AccessOffers = False) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.printedmessagepreview", LanguageID))%>
  </h1>
  <%--
  <div id="controls">
    <button id="refresh" name="refresh" type="button"><% Sendb(Copient.PhraseLib.Lookup("term.refresh", LanguageID))%></button>
  </div>
  --%>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <%
    TierLevel = MyCommon.NZ(Request.QueryString("TierLevel"), "")
    If TierLevel.Length > 0 Then
      'Sendb(TierLevel)
    End If
  %>
  <div id="column2x">
    <div id="pmsgpreview">
      <div id="pmsgpreviewbody">
<%
  Send("<pre>" & MESSAGE & "</pre>")
  %>        
  
      </div>
    </div>
  </div>
  <div id="gutter">
  </div>
  <div id="column1x">
    <div class="box" id="message" style="height: 425px;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.printer", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<b>" & PrinterName & "</b><br />")
        If (PageWidth = 0) Then
        Else
          If (FixedWidthFont = True) Then
            Send(PageWidth & " " & StrConv(Copient.PhraseLib.Lookup("term.fixedwidth", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.characters", LanguageID), VbStrConv.Lowercase) & "<br />")
          Else
            Send(PageWidth & " " & StrConv(Copient.PhraseLib.Lookup("term.characters", LanguageID), VbStrConv.Lowercase) & "<br />")
          End If
        End If
      %>
    </div>
  </div>
</div>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>

<script runat="server">
  Public DefaultLanguageID
  Public MyCommon As New Copient.CommonInc
  Function LineWrap(ByVal wrapLength As Integer, ByVal NewLine As String) As String
    Dim NewBRLine As String
    Dim FormatLine As String
    Dim y As Integer
    'If the line is longer then <wrapLength> characters, wrap it
    If wrapLength = 0 Then
      wrapLength = 50
    End If
    NewBRLine = NewLine
    FormatLine = ""
    If NewBRLine.Length > wrapLength Then
      ' safety check to prevent infinite loop
      y = 0
      While (NewBRLine.Length > 0 And y < 100)
        y = y + 1
        If (NewBRLine.Length > wrapLength) Then
          FormatLine = FormatLine & NewBRLine.Substring(0, wrapLength) & "<br />"
          NewBRLine = NewBRLine.Remove(0, wrapLength)
        Else
          FormatLine = FormatLine & NewBRLine & "<br />"
          NewBRLine = ""
        End If
      End While
    Else
      FormatLine = NewLine & "<br />"
    End If
    Return FormatLine
  End Function
  
  '----------------------------------------------------------------------------------------------
  
  Function ProcessMsgLine(ByVal PrinterTypeID As Integer, ByVal PrintLine As String, ByRef TagDST As DataTable, ByRef StyleStr As String, ByRef AlignmentStr As String) As String
    
    Dim ReturnVal As String = ""
    Dim Done As Boolean = False
    Dim ptr As Integer
    Dim TagStart As Integer
    Dim TagEnd As Integer
    Dim TempTag As String
    Dim PrinterTag As String
    Dim PreviewText As String
    Dim TagLength As Integer
    Dim row As DataRow
    Dim StartStylePos As Integer
    Dim EndStylePos As Integer
    Dim TempStyleStr As String
    Dim ParamPos As Integer
    Dim index As Integer = 0
    Dim TempStr As String
    Dim TagMatchSuccess As Boolean = False
    Dim tempValueOfString As String
    
    ptr = 1
    Send("<!-- PrintLine=" & PrintLine & " -->")
    While Not (Done) 'look for tags until we are finished with the line
      index = index + 1
      'get the next tag from the PrintLine
      TagStart = InStr(ptr, PrintLine, "|", CompareMethod.Binary)
      If TagStart > 0 Then
        TagEnd = InStr(TagStart + 1, PrintLine, "|", CompareMethod.Binary)
        If TagEnd > 0 Then
          TagEnd = TagEnd + 1
          TempTag = Mid(PrintLine, TagStart, TagEnd - TagStart)
          'handling for tags with parameters
          ParamPos = InStr(TempTag, "[", CompareMethod.Binary)
          If ParamPos > 0 Then
            TempTag = Left(TempTag, ParamPos - 1) & "|"
          End If
          Send("  <!-- Found possible tag: " & TempTag & " -->")
          'we have found the starting and ending points of a possible tag
          'nows search through our list and see if there is a match
          TagMatchSuccess = False
          For Each row In TagDST.Rows
            PrinterTag = row.Item("tag")
            TagLength = Len(PrinterTag)
            PreviewText = row.Item("previewchars")
            'Send("<!-- TempTag=" & TempTag & "  PrinterTag=" & PrinterTag & " -->")
            If TempTag = PrinterTag And row.Item("PrinterTypeID") = PrinterTypeID Then
              'we found a tag match
              TagMatchSuccess = True
              If InStr(PreviewText, "style=", CompareMethod.Binary) > 0 Then
                'we are starting a new style (bold, underline, etc.) or alignment
                'get the style command from the PreviewText and add it to the StyleStr
                Send("  <!-- START STYLE TAG - previewtext=" & PreviewText & " -->")
                StartStylePos = InStr(PreviewText, """", CompareMethod.Binary)
                If StartStylePos > 0 Then
                  StartStylePos = StartStylePos + 1 'skip the quotation mark
                  EndStylePos = InStr(StartStylePos, PreviewText, """", CompareMethod.Binary)
                  If EndStylePos > 0 Then
                    TempStyleStr = Mid(PreviewText, StartStylePos, EndStylePos - StartStylePos)
                    PreviewText = ""
                    If Len(StyleStr) > 0 Then
                      PreviewText = "</span>"
                    End If
                    If row.Item("AlignmentTag") = True Then
                      AlignmentStr = TempStyleStr
                      Send("  <!-- Updated AlignmentStr=" & AlignmentStr & " -->")
                      PreviewText = PreviewText & "<span style=""" & AlignmentStr & """>"
                    Else
                      StyleStr = StyleStr & TempStyleStr
                      Send("  <!-- Updated StyleStr=" & StyleStr & " -->")
                      PreviewText = PreviewText & "<span style=""" & StyleStr & """>"
                    End If
                  End If 'EndStylePos>0
                End If 'StartStylePos>0
                Send("  <!-- FINAL previewtext=" & PreviewText & " -->")
                'Now replace the tag in the PrintLine with the PreviewText
                PrintLine = Mid(PrintLine, 1, TagStart - 1) & PreviewText & Mid(PrintLine, TagEnd)
                Send("  <!-- Updated PrintLine=" & PrintLine & " -->")
              ElseIf Len(PreviewText) >= 2 AndAlso Left(PreviewText, 1) = "/" Then
                'we are ending a style (/bold, /underline, etc.)
                Send("  <!-- END STYLE TAG - previewtext=" & PreviewText & " -->")
                TempStyleStr = Mid(PreviewText, 2)
                Send("  <!-- TempStyleStr='" & TempStyleStr & "'   StyleStr='" & StyleStr & "' -->")
                'remove the style from the string
                If row.Item("AlignmentTag") = True Then
                  AlignmentStr = "" 'you can only have one alignment at a time, so we can just blank out the alignment string
                  Send("  <!-- Updated AlignmentStr=" & AlignmentStr & " -->")
                Else
                  StyleStr = Replace(StyleStr, TempStyleStr, "", 1, 1)
                  Send("  <!-- Updated StyleStr=" & StyleStr & " -->")
                End If
                PreviewText = "</span>"
                If Not (StyleStr = "") Then PreviewText = PreviewText & "<span style=""" & StyleStr & """>"
                Send("  <!-- FINAL previewtext=" & PreviewText & " -->")
                PrintLine = Mid(PrintLine, 1, TagStart - 1) & PreviewText & Mid(PrintLine, TagEnd)
                Send("  <!-- Updated PrintLine=" & PrintLine & " -->")
              ElseIf PreviewText = "</span>" Then
                'we are clearing all the styles ... this would be from a 'normal' tag
                Send("  <!-- Clear all StyleStr -->")
                StyleStr = ""
                PrintLine = Mid(PrintLine, 1, TagStart - 1) & PreviewText & Mid(PrintLine, TagEnd)
                Send("  <!-- Updated PrintLine=" & PrintLine & " -->")
              Else
                'if this is an alignment tag, then close the style and re-open it after the alignment tag
                Send("  <!-- Non style tag -->")
                If row.Item("AlignmentTag") = True And Not (Mid(PreviewText, 2, 1) = "/") Then
                  Send("  <!-- Opening AlignmentTag=" & PreviewText & " -->")
                  If Not (AlignmentStr = "") Then
                    PreviewText = "</" & Mid(AlignmentStr, 2) & PreviewText
                  End If
                  AlignmentStr = PreviewText
                  Send("  <!-- Updated AlignmentStr='" & AlignmentStr & "' -->")
                End If
                tempValueOfString = (PreviewText & Mid(PrintLine, TagEnd)).Replace("Discount", "<br />Discount")
                PrintLine = Mid(PrintLine, 1, TagStart - 1) & tempValueOfString
                Send("  <!-- Updated PrintLine=" & PrintLine & " -->")
              End If
            End If
          Next
          If TagMatchSuccess Then
            ptr = TagEnd  'starting looking for the next tag with the first charater AFTER the last tag end
        Else
            'what we thought might have been a valid tag, wasn't a match in the database - maybe it was just pipe characters in the message text?
            ptr = TagEnd - 1 'start looking for the next tag using the invalid tag's ending pipe character
        End If
          
      Else
        Done = True
      End If
      Else
        Done = True
      End If
      'enable for runaway prevention
      If index > 100 Then
        Done = True
        Send("<!-- loop runaway in ProcessMsgLine -->")
      End If
    End While
    Send("<!-- index=" & index & " -->")

    Insert_Line_Breaks(PrintLine)
    
    Return PrintLine
    
  End Function
  
  
  Private Sub Insert_Line_Breaks(ByRef PrintLine As String)
    Dim MsgCt As Integer = 0
    Dim InsideTag As Boolean
    Dim i As Integer
    Dim ShouldCtChar As Boolean
    Dim TagName As String = ""

    If PrintLine IsNot Nothing Then
      For i = 0 To PrintLine.Length - 1
        ShouldCtChar = True

        Select Case PrintLine.Chars(i)
          Case "<"
            ' peek ahead one char to see if this is a tag
            If i + 1 < PrintLine.Length - 1 Then
              InsideTag = Char.IsLetter(PrintLine.Chars(i + 1)) OrElse PrintLine.Chars(i + 1) = "/"
              ShouldCtChar = IIf(InsideTag, False, True)

              ' check if this is a line break tag; if so, reset the counter
              If i + 4 < PrintLine.Length - 1 Then
                TagName = PrintLine.Substring(i, 4).ToLower
                If TagName = "<br " OrElse TagName = "<br>" Then
                  MsgCt = 0
                End If
              End If

            End If
          Case ">"
            If InsideTag Then
              InsideTag = False
              ShouldCtChar = False
            End If
          Case Else
            If InsideTag Then ShouldCtChar = False
        End Select

        If ShouldCtChar Then MsgCt += 1

        If MsgCt > 44 Then
          PrintLine = PrintLine.Substring(0, i) & "<br />" & PrintLine.Substring(i)
          MsgCt = 0
        End If
      Next
    End If

  End Sub
  
  Function FormatScorecardLine(ByVal LabelPhrase As String, ByVal Value As String, ByVal PageWidth As Integer) As String
    Dim Col1Width As Integer = 30
    Dim Col2Width As Integer = 14
    
    If PageWidth > 0 Then
      Col1Width = CInt(PageWidth * 0.8)
      If Col1Width <= 0 Then Col1Width = 1

      Col2Width = PageWidth - Col1Width
      If Col2Width <= 0 Then Col1Width = 1
    End If
    
    Return String.Format("{0," & Col1Width & "}{1," & Col2Width & "}", Left(Copient.PhraseLib.Lookup(LabelPhrase, LanguageID), Col1Width), Left(Value, Col2Width))
  End Function

</script>

