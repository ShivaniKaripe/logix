<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" ValidateRequest="false" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>

<%
  ' *****************************************************************************
  ' * FILENAME: email-not-pmsgpreview.aspx 
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
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim row As DataRow
  Dim OfferID As Integer = 0
  Dim ImgType As Integer = 1
  Dim i As Integer = 0
  Dim OutputID As Integer = 0
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
    
  Response.Expires = 0
  MyCommon.AppName = "email-not-pmsgpreview.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  Send_HeadBegin("term.preview")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
    
  
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
    <% Sendb(Copient.PhraseLib.Lookup("term.emailpreview", LanguageID))%>
  </h1>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="emailpreview">
    <%
      OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
      
      ' Pull out details (name, page width, etc.) for that printer
      MyCommon.QueryStr = "select DeliverableTypeID, OutputID from CPE_Deliverables DEL with (NoLock)  " & _
                          "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DEL.RewardOptionID and RO.Deleted=0 and DEL.Deleted=0 " & _
                          "where RO.IncentiveID=" & OfferID & " and RewardOptionPhase=1 " & _
                          "order by DEL.Priority;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then        
        For Each row In rst.Rows
          OutputID = MyCommon.NZ(row.Item("OutputID"), 0)
          If MyCommon.NZ(row.Item("DeliverableTypeID"), 4) = 4 Then
            ' write the printed message
            Send(GetPMsgText(OutputID, MyCommon))
          ElseIf MyCommon.NZ(row.Item("DeliverableTypeID"), 4) = 1 Then
            ' get the image type
            MyCommon.QueryStr = "select ImageType from OnScreenAds with (NoLock) where OnScreenAdID=" & OutputID
            rst2 = MyCommon.LRT_Select
            If (rst2.Rows.Count > 0) Then
              ImgType = MyCommon.NZ(rst2.Rows(0).Item("ImageType"), 1)
            End If
            
            ' write the graphic
            Send("<img src=""" & GetGraphicsPath(OutputID, ImgType, MyCommon) & """ />")
          End If
          Send("<br />")
        Next
      End If


     %>
    </div>
</div>
<script runat="server">
  
  Function GetPMsgText(ByVal MessageID As Integer, ByRef MyCommon As Copient.CommonInc) As String
    Dim PMsgText As String = ""
    Dim rst As DataTable = Nothing
    
    If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()

    MyCommon.QueryStr = "select BodyText from PrintedMessageTiers where MessageID=" & MessageID & " and TierLevel=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      PMsgText = MyCommon.NZ(rst.Rows(0).Item("BodyText"), "")
      PMsgText = PMsgText.Replace(vbCrLf, "<br />")
    End If
    
    Return PMsgText
  End Function
  
  Function GetGraphicsPath(ByVal AdID As Integer, ByVal ImgType As Integer, ByRef MyCommon As Copient.CommonInc) As String
    Dim GraphicPath As String = ""
    Dim ImgExt As String = "jpg"
    
    If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()

    ' Build the graphic file path string
    GraphicPath = MyCommon.Fetch_SystemOption(47)
    If (GraphicPath.Trim().Length = 0) Then
      GraphicPath = "C:\"
    End If
    If Not (Right(GraphicPath, 1) = "\") Then
      GraphicPath = GraphicPath & "\"
    End If
    If (ImgType = 1) Then
      ImgExt = "jpg"
    ElseIf (ImgType = 2) Then
      ImgExt = "gif"
    End If
    GraphicPath = GraphicPath & CStr(AdID) & "img." & ImgExt
    
    Return GraphicPath
    
  End Function
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
