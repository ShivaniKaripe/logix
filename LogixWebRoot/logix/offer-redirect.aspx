<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-redirect.aspx 
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
  Dim OfferID As Integer
  Dim EngineID As Integer = 0
  Dim RedirectPage As String = ""
  Dim Popup As Boolean = False
  
  Response.Expires = 0
  MyCommon.AppName = "offer-redirect.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
  EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  
  ' find the EngineID for the given OfferID 
  If (EngineID <= 0) Then
    MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
    End If
  End If
  
  ' see if the page needs to be in popup form
  If (Request.QueryString("Popup") <> "") Then
    Popup = True
  End If
  
  Select Case EngineID
    Case 0, 1
      RedirectPage = "offer-sum.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
    Case 2
      RedirectPage = "CPEoffer-sum.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
    Case 3
      RedirectPage = "web-offer-sum.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
      'Case 4
      '  RedirectPage = "DP-offer-sum.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
    Case 5
      RedirectPage = "email-offer-sum.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
    Case 6
      RedirectPage = "CAM/CAM-offer-sum.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
    'Case 7
    '  RedirectPage = "desktop/PDE/PDEoffer-gen.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
    Case 9
      RedirectPage = "/logix/UE/UEoffer-sum.aspx?OfferID=" & OfferID & IIf(Popup, "&Popup=1", "")
    Case Else
      RedirectPage = "error-notfound.aspx"
  End Select
  
  Response.Status = "301 Moved Permanently"
  Response.AddHeader("location", RedirectPage)
  
done:
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
