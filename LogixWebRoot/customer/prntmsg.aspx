<%@ Page Language="vb" Debug="true" CodeFile="cwCB.vb" Inherits="cwCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>


<%
  dim MyCommon As New Copient.CommonInc
  dim Logix As New Copient.LogixInc
  dim Localization As Copient.Localization
  dim CustomerPK As Long
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  Dim LanguageID As Integer = 1
  Dim infoMessage As String = ""
  Dim Popup As Boolean = True
  Dim Frames As Boolean = False
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.AppName = "prntmsg.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then MyCommon.Open_PrefManRT()
  Localization = New Copient.Localization(MyCommon)
  If Not (Long.TryParse(GetCgiValue("identifier"), CustomerPK)) Then CustomerPK = 0
  CustLanguageID = Localization.GetCustLanguageID(CustomerPK)
  
  Send_HeadBegin(Handheld, "Adam's Natural Foods: Details")
  Send_Metas(Handheld)
  Send("<style type=""text/css"">")
  Send("body { font-size:12px;font-family:arial;text-align:center; }")
  Send("</style>")
  Send_HeadEnd(Handheld)
  Send_BodyBegin(Handheld, False)
  Send_Pmsg(Handheld)
  Send_BodyEnd(Handheld)
  
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then MyCommon.Close_PrefManRT()
  MyCommon = Nothing
%>
