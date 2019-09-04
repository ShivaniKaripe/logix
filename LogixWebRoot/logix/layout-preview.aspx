<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: layout-preview.aspx 
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
  Dim dt As DataTable = Nothing
  Dim row As DataRow = Nothing
  Dim LayoutID As Integer
  Dim Index As Integer = 0
  Dim ColorName As String = ""
  Dim FontColor As String = ""
  Dim BackgroundImg As String = ""
  Dim GraphicPath As String = ""
  Dim imgExt As String = ""
  Dim ImgType As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Const DEFAULT_GRAPHIC_PATH As String = "C:\"
  
  Response.Expires = 0
  MyCommon.AppName = "layout-preview.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  LayoutID = MyCommon.Extract_Val(Request.QueryString("LayoutID"))
  
  ' Build the graphic file path string
  GraphicPath = MyCommon.Fetch_SystemOption(47)
  If (GraphicPath.Trim().Length = 0) Then
    GraphicPath = DEFAULT_GRAPHIC_PATH
  End If
  If Not (Right(GraphicPath, 1) = "\") Then
    GraphicPath = GraphicPath & "\"
  End If
  If (ImgType = "1") Then
    imgExt = "jpg"
  ElseIf (ImgType = "2") Then
    imgExt = "gif"
  End If
  
  'GraphicPath = GraphicPath & CStr(AdID) & "img." & imgExt
  
  MyCommon.QueryStr = "select Name from ScreenLayouts with (NoLock) where LayoutID=" & LayoutID & " and deleted=0;"
  dt = MyCommon.LRT_Select
  If dt.Rows.Count > 0 Then
    Dim LayoutName As String = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
  End If
  
  Send_HeadBegin("term.layoutpreview")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Scripts()
  Send_HeadEnd()
  Send("<body style=""background-color:#ffffff;"">")
  
  If (Logix.UserRoles.AccessLayouts = False) Then
    Send_Denied(2, "perm.layouts-access")
    GoTo done
  End If
  
  MyCommon.QueryStr = "select CellID, Name, ContentsID, X, Y, Width, Height, BackgroundImg from ScreenCells with (NoLock) where LayoutID=" & LayoutID & " and deleted=0;"
  dt = MyCommon.LRT_Select
  If (dt.Rows.Count > 0) Then
    Index = 0
    For Each row In dt.Rows
      Index = Index + 1
      If Index > 4 Then Index = 1
      Select Case Index
        Case 1
          ColorName = "yellow"
          FontColor = "black"
        Case 2
          ColorName = "red"
          FontColor = "white"
        Case 3
          ColorName = "green"
          FontColor = "white"
        Case 4
          ColorName = "blue"
          FontColor = "white"
      End Select
      BackgroundImg = MyCommon.NZ(row.Item("BackgroundImg"), 0)
      If BackgroundImg > 0 Then
        MyCommon.QueryStr = "select ImageType from OnScreenAds with (NoLock) where OnScreenAdID=" & BackgroundImg & " and deleted=0;"
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
          ImgType = MyCommon.NZ(dt.Rows(0).Item("ImageType"), 1)
        End If
        
        ' Build the graphic file path string
        GraphicPath = MyCommon.Fetch_SystemOption(47)
        If (GraphicPath.Trim().Length = 0) Then
          GraphicPath = DEFAULT_GRAPHIC_PATH
        End If
        If Not (Right(GraphicPath, 1) = "\") Then
          GraphicPath = GraphicPath & "\"
        End If
        If (ImgType = "1") Then
          imgExt = "jpg"
        ElseIf (ImgType = "2") Then
          imgExt = "gif"
        End If
        GraphicPath = GraphicPath & CStr(BackgroundImg) & "img." & imgExt
        
        Send("<div style=""position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) & "px; width:" & MyCommon.NZ(row.Item("Width"), 1) & "px; height:" & MyCommon.NZ(row.Item("Height"), 1) & "px; background-color:" & ColorName & ";"">")
        Send("<img src=""graphic-display-img.aspx?path=" & Server.UrlEncode(GraphicPath) & "&mode=ad&adid=" & BackgroundImg & """ width=""" & MyCommon.NZ(row.Item("Width"), 1) & """ height=""" & MyCommon.NZ(row.Item("Height"), 1) & """ alt="""" />")
        Send("</div>")
        Send("<div style=""position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) & "px; width:" & MyCommon.NZ(row.Item("Width"), 1) & "px; height:" & MyCommon.NZ(row.Item("Height"), 1) & "px;"">")
      Else
        Send("<div style=""position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) & "px; width:" & MyCommon.NZ(row.Item("Width"), 1) & "px; height:" & MyCommon.NZ(row.Item("Height"), 1) & "px; background-color:" & ColorName & ";"">")
      End If
      Send(" <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""height:100%;"" summary=""" & Copient.PhraseLib.Lookup("term.cell", LanguageID) & """>")
      Send("  <tr>")
      Send("   <td valign=""middle"">")
      Send("    <center>")
      Send("     <table border=""0"" cellpadding=""2"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.cell", LanguageID) & """>")
      Send("      <tr>")
      Send("       <td class=""cell"">")
      Send("        " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
      Send("       </td>")
      Send("      </tr>")
      Send("     </table>")
      Send("    </center>")
      Send("   </td>")
      Send("  </tr>")
      Send(" </table>")
      Send("</div>")
    Next
  End If
  Send("</body>")
  Send("</html>")
  
done:
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
