<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: UEoffer-rew-cmsgpreview.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2011.  All rights reserved by:
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
  Dim PrinterName As String
  Dim PageWidth As Integer
  Dim FixedWidthFont As Boolean
  Dim RawMessage As String
  Dim MESSAGE As String
  Dim TierLevel As String
  Dim x As Integer
  Dim PrinterTag As String
  Dim PreviewText As String
  Dim PrintLine As String
  Dim Found As Boolean = False
  Dim ExtFound As Boolean = False
  Dim NewLine As String
  Dim NewBRLine As String
  Dim y As Integer = 0
  Dim FormatLine As String
  Dim WrappedPrintLine As String
  Dim RawTextLine As String
  Dim FormatChar As String
  Dim ExtFormatChar As String
  Dim ENDMESSAGE As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  ' This 2D array will hold the replace and replacement for real-time engine tags, just add or modify here
  Dim SubTags(,) As String = { _
  {"|CUSTOMERID|", "|TSD|", "|LYTS|", "|CURRDATE|", "|OFFERSTART|", "|OFFEREND|", "|TOTALPOIINTS|", "|ACCUMANT|", "|REMAINAMT|"}, _
  {"###################", "000.00", "000.00", "xx/xx/xxxx", "xx/xx/xxxx", "xx/xx/xxxx", "xx", "000.00", "000.00"}}
  
  Dim SubExpTags(,) As String = { _
  {"\|NET\#[[0-9]+]\|", "\|EARNED\#[[0-9]+]\|", "\|INITIAL#\[[0-9]+]\|", "\|REDEEMED#\[[0-9]+]\|", "\|NET\$\[[0-9]+]\|", "\|INITIAL\$\[[0-9]+]\|", "\|EARNED\$\[[0-9]+]\|", "\|REDEEMED\$\[[0-9]+]\|", "\|SVBAL\[[0-9]+\]\|", "\|SVVAL\[[0-9]+\]\|", "\|UPCA\[[0-9]+\]\|", "\|UPCB\[[0-9]+\]\|", "\|SCORECARD\[[0-9]+\]\|", "\|SVRATIO\[[0-9]+\,[0-9]+\,[0-9,.]+\]\|", "\|PTBAL\[[0-9]+\]\|"}, _
  {"###", "###", "###", "###", "###.##", "###.##", "###.##", "###.##", "####", "###.##", "<img src='../images/upca.png'>", "<img src='../images/upcb.png'>", "**************SCORECARD**************", "###.##", "####"}}
  
  Response.Expires = 0
  MyCommon.AppName = "UEoffer-rew-cmsgpreview.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  Send_HeadBegin("term.preview")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  
  Dim lp as Integer
  for lp = 1 to 9 
    if(Request.QueryString("Line" & lp) <> "") then
		RawMessage = RawMessage & Request.QueryString("Line" & lp) + vbLf + vbLf
	end if
  next
  RawMessage = RawMessage & Request.QueryString("Line" & lp)
  
  ' replace all the < and > 
  ' before processing
  RawMessage = RawMessage.Replace("<", "&lt;")
  RawMessage = RawMessage.Replace(">", "&gt;")
  
  PrinterName = "Cashier"
  PageWidth = 20
  FixedWidthFont = True
  
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
  Send("  overflow-x: hidden;")
  
  Send("  width: " & (PageWidthPixels + 2) & "px;")
  ' Note: the +2 added to PageWidthPixels shouldn't be necessary; it's only there
  ' because of a problem sizing medium text in IE that couldn't otherwise be solved
  Send("}")
  Send("</style>")
  
  ' here we will format the message, the message is currently in RawMessage 
  ' so lets get it formatted correctly same same
  ' as the local server
  ' first look for and replace  CUSTOMERID  TSD  LYTS  CURRDATE  OFFERSTART  OFFEREND  TOTALPOIINTS  ACCUMANT  REMAINAMT 
  
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
  
  Dim strArray() As String = MESSAGE.Split(vbLf)
  ' ok we split lets blank out MESSAGE so we can refill it
  MESSAGE = ""
  ' get the tags
  MyCommon.QueryStr = "select '|' + tag + '|' as tag,isnull(pt.previewchars,'') as PreviewChars from MarkupTags as MT with (NoLock) " & _
                      "left join PrinterTranslation as PT with (NoLock) on MT.MarkupID=PT.MarkupID " & _
                      "where PrinterTypeID=0;"
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
      PrinterTag = row.Item("tag")
      PreviewText = row.Item("previewchars")
      If (InStr(RawTextLine, PrinterTag)) Then
        RawTextLine = Replace(RawTextLine, PrinterTag, "")
      End If
    Next
    ' RawTextLine now containts unformatted text of the line lets wrap it
    RawTextLine = LineWrap(PageWidth, RawTextLine)
    WrappedPrintLine = PrintLine
    For Each row In rst.Rows
      PrinterTag = row.Item("tag")
      PreviewText = row.Item("previewchars")
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
    <% Sendb(Copient.PhraseLib.Lookup("term.cashiermessagepreview", LanguageID))%>
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
      Sendb(TierLevel)
    End If
  %>
  <div id="column2x">
    <div id="pmsgpreview">
      <div id="pmsgpreviewbody">
        <% Sendb(MESSAGE)%>
      </div>
    </div>
  </div>
  
  <div id="gutter">
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
  Public DefaultLanguageID As Integer
  Public MyCommon As New Copient.CommonInc
  Function LineWrap(ByVal wrapLength As Integer, ByVal NewLine As String) As String
    Dim NewBRLine As String
    Dim FormatLine As String
    Dim y As Integer
    '' if the line is longer then 38 characters then wrap it
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
</script>
