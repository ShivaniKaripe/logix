<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: graphic-preview.aspx 
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
  Dim AdID As Long
  Dim TouchAreaID As Integer = 0
  Dim TouchAreaCount As Integer = 0
  Dim AreaCounter As Integer = 0
  Dim OfferID As Integer = 0
  Dim ShowDeliverables As Integer = 0
  Dim AddressPath As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.AppName = "graphic-preview.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = MyCommon.Extract_Val(Request.QueryString("offerID"))
  AdID = MyCommon.Extract_Val(Request.QueryString("adId"))
  ShowDeliverables = MyCommon.Extract_Val(Request.QueryString("show"))
  TouchAreaID = MyCommon.Extract_Val(Request.QueryString("areaId"))
  AddressPath = Request.QueryString("path")
  
  'ShowDeliverables = 1 ' Remove after testing
  'OfferID = 341 'Remove after testing
  
  Send_HeadBegin("term.graphicpreview")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Scripts()
  Send_HeadEnd()
  Send("<body style=""background-color:#ffffff;"">")
  
  If (Logix.UserRoles.AccessGraphics = False) Then
    Send_Denied(2, "perm.graphics-access")
    GoTo done
  End If
  
  If (TouchAreaID > 0 And OfferID > 0) Then
    ' look up the deliverables associated with this touch area
    MyCommon.QueryStr = "select DeliverableID, RewardOptionID, DeliverableTypeID, OutputID " & _
                        "from CPE_Deliverables with (NoLock) where rewardoptionid in " & _
                        "(select RewardOptionID from CPE_deliverableroids with (NoLock) where incentiveid = " & OfferID & " and areaid = " & TouchAreaID & " and deleted=0) " & _
                        "order by DeliverableTypeID;"
    dt = MyCommon.LRT_Select()
    If (dt.Rows.Count > 0) Then
      For Each row In dt.Rows
        If (row.Item("DeliverableTypeID") = 1) Then 'Graphic
          Send(GenerateGraphicCode(MyCommon, AddressPath, MyCommon.NZ(row.Item("OutputID"), 0), ShowDeliverables))
        Else
          Send("<div>")
          Send(Copient.PhraseLib.Lookup("term.deliverable", LanguageID) & ": " & MyCommon.NZ(row.Item("DeliverableTypeID"), 0))
          Send("<br />")
          Send(Copient.PhraseLib.Lookup("term.output", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & MyCommon.NZ(row.Item("OutputID"), 0))
          Send("</div>")
        End If
        Send("<br />")
      Next
    End If
  Else
    Send(GenerateGraphicCode(MyCommon, AddressPath, AdID, ShowDeliverables))
  End If
%>

<script runat="server">
  Function GenerateGraphicCode(ByRef MyCommon As Copient.CommonInc, ByRef AddressPath As String, ByVal AdID As Long, ByVal ShowDeliverables As Integer) As String
    Dim htmlBuf As New StringBuilder
    Dim ImgWidth As Integer
    Dim ImgHeight As Integer
    Dim ImgType As Integer
    Dim ImgName As String = ""
    Dim LabelsStr As String
    Dim dt As DataTable = Nothing
    Dim row As DataRow = Nothing
    Dim GraphicPath As String = ""
    Dim imgExt As String = ""
    Dim top As Integer = 0
    
    Const DEFAULT_GRAPHIC_PATH As String = "C:\"
    
    MyCommon.QueryStr = "select Name, Width, Height, ImageType from OnScreenAds with (NoLock) where OnScreenAdID=" & AdID & " and deleted=0;"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      ImgWidth = MyCommon.NZ(dt.Rows(0).Item("Width"), 1)
      ImgHeight = MyCommon.NZ(dt.Rows(0).Item("Height"), 1)
      ImgType = MyCommon.NZ(dt.Rows(0).Item("ImageType"), 1)
      ImgName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
      AddressPath += ImgName
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
    GraphicPath = GraphicPath & CStr(AdID) & "img." & imgExt
    top = IIf(ShowDeliverables = 1, 20, 0)
    htmlBuf.Append("<div style=""position:absolute; top:" & top & "px; left:0px; width:" & ImgWidth & "px; height:" & ImgHeight & "px; background-color:#ffffff;"">" & vbCrLf)
    htmlBuf.Append(" <img src=""graphic-display-img.aspx?path=" & Server.UrlEncode(GraphicPath) & "&amp;mode=ad&amp;lang=" & LanguageID & "&amp;adid=" & AdID & """ width=""" & ImgWidth & """ height=""" & ImgHeight & """ alt="""" />" & vbCrLf)
    LabelsStr = ""
    MyCommon.QueryStr = "select AreaID, Name, X, Y, Width, Height from TouchAreas with (NoLock) where OnScreenAdID=" & AdID & " and Deleted=0;"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      For Each row In dt.Rows
        'send the horizontal lines
        htmlBuf.Append(" <img src=""/images/blackdot.png"" style=""position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) & "px;"" width=""" & MyCommon.NZ(row.Item("Width"), 1) & """ height=""1"" alt="""" />" & vbCrLf)
        htmlBuf.Append(" <img src=""/images/blackdot.png"" style=""position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) + MyCommon.NZ(row.Item("Height"), 1) & "px;LEFT:" & MyCommon.NZ(row.Item("X"), 0) & "px;"" width=""" & MyCommon.NZ(row.Item("Width"), 1) & """ height=""1"" alt="""" />" & vbCrLf)
        'send the vertical lines
        htmlBuf.Append(" <img src=""/images/blackdot.png"" style=""position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) & "px;"" width=""1"" height=""" & MyCommon.NZ(row.Item("Height"), 1) & """ alt="""" />" & vbCrLf)
        htmlBuf.Append(" <img src=""/images/blackdot.png"" style=""position:absolute; top:" & MyCommon.NZ(row.Item("Y"), 0) & "px; left:" & MyCommon.NZ(row.Item("X"), 0) + MyCommon.NZ(row.Item("Width"), 1) & "px;"" width=""1"" height=""" & MyCommon.NZ(row.Item("Height"), 1) & """ alt="""" />" & vbCrLf)
        If (ShowDeliverables > 0) Then
          LabelsStr = LabelsStr & "<div onclick=""javascript:showNextDeliverable(" & MyCommon.NZ(row.Item("AreaID"), "") & ", '" & MyCommon.NZ(row.Item("Name"), "") & "');"" style=""position: absolute; top: " & top + MyCommon.NZ(row.Item("Y"), 0) & "px; left: " & MyCommon.NZ(row.Item("X"), 0) & "px; width: " & MyCommon.NZ(row.Item("Width"), 1) & "px; height: " & MyCommon.NZ(row.Item("Height"), 1) & "px;"">" & vbCrLf
        Else
          LabelsStr = LabelsStr & "<div style=""position: absolute; top: " & 30 + MyCommon.NZ(row.Item("Y"), 0) & "px; left: " & MyCommon.NZ(row.Item("X"), 0) & "px; width: " & MyCommon.NZ(row.Item("Width"), 1) & "px; height: " & MyCommon.NZ(row.Item("Height"), 1) & "px;"">" & vbCrLf
        End If
        LabelsStr = LabelsStr & " <table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""height:100%;width:100%;"" summary=""" & MyCommon.NZ(row.Item("Name"), "") & """>" & vbCrLf
        LabelsStr = LabelsStr & "  <tr>" & vbCrLf
        LabelsStr = LabelsStr & "   <td valign=""middle"">" & vbCrLf
        LabelsStr = LabelsStr & "    <center>" & vbCrLf
        LabelsStr = LabelsStr & "    <table border=""0"" cellpadding=""2"" cellspacing=""0"" summary="""">" & vbCrLf
        LabelsStr = LabelsStr & "     <tr>" & vbCrLf
        LabelsStr = LabelsStr & "      <td valign=""middle"" bgcolor=""white"">" & vbCrLf
        LabelsStr = LabelsStr & "       <center><span style=""font-family:sans-serif;font-size:12px;"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</span></center>" & vbCrLf
        LabelsStr = LabelsStr & "      </td>" & vbCrLf
        LabelsStr = LabelsStr & "     </tr>" & vbCrLf
        LabelsStr = LabelsStr & "    </table>" & vbCrLf
        LabelsStr = LabelsStr & "    </center>" & vbCrLf
        LabelsStr = LabelsStr & "   </td>" & vbCrLf
        LabelsStr = LabelsStr & "  </tr>" & vbCrLf
        LabelsStr = LabelsStr & " </table>" & vbCrLf
        LabelsStr = LabelsStr & "</div>" & vbCrLf
      Next
    End If
    htmlBuf.Append("</div>" & vbCrLf)
    htmlBuf.Append(LabelsStr)
    Return htmlBuf.ToString
  End Function
</script>
<% 
  If (ShowDeliverables = 1) Then
    Send("<div id=""addressBar"" style=""position:absolute;top:0px;left:0px;width:100%;height:20px;"">")
    Send(AddressPath)
    Send("</div>")
  End If
%>
<script type="text/javascript">
    function showNextDeliverable(areaId, areaName) {
        var path = "<% Sendb(AddressPath.Replace("""", "\""")) %>" + "(" + areaName+ ") - ";
        location.href = "graphic-preview.aspx?adId=<%Sendb(AdID)%>&areaId=" + areaId + "&path=" + path + "&show=<%Sendb(ShowDeliverables) %>";
    }
</script>
<%
  Send("</body>")
  Send("</html>")
done:
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
