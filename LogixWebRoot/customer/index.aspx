<%@ Page Language="vb" Debug="true" CodeFile="cwCB.vb" Inherits="cwCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim LanguageID As Integer = 1
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  Dim dst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim Identifier As String = ""
  Dim infoMessage As String = ""
  Dim Popup As Boolean = False
  Dim Framed As Boolean = False
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Identifier = IIf(Request.Form("identifier") <> "", Request.Form("identifier"), "")
  If (Identifier = "") Then
    Identifier = IIf(Request.QueryString("identifier") <> "", Request.QueryString("identifier"), "")
  End If
  
  MyCommon.AppName = "index.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  
  Send_HeadBegin(Handheld, "Customer Website: Login")
  Send_Metas(Handheld)
  Send_Links(Handheld)
  Send_HeadEnd(Handheld)
  Send_BodyBegin(Handheld, Popup)
  If Not Handheld Then
    Send_InnerWrapBegin(Handheld)
  End If
  Send_Logo(Handheld)
  Send_Menu(Handheld)
  Send_Submenu(Handheld)
  Send_SidebarBegin(Handheld)
  Send_Login(Handheld)
  Send_SidebarEnd(Handheld)
  If Not Handheld Then
    Send_Gutter(Handheld)
  End If
  Send_MainBegin(Handheld)
%>
  <h1>• Welcome! •</h1>
  <hr />
  <br />
  <p>
    <img src="images/card.png" alt="NCR Program Card" title="NCR Program Card" align="left" />
    Thanks for visiting NCR Store! To get the most out of your visits to our stores and our website, be sure to get your free NCR Program Card. The card entitles you to special savings opportunities, plus you can use it to log in to this website and check your points balances, accumulations and offers. Visit customer service to sign up!
  </p>
  <p>
    If you have a card and would like to see the personalized offers you can get at NCR, log in using the form to the left.  Once you're logged in, you'll be able to opt into or out of offers, as well as see and edit your information.
  </p>
<%
  Send_MainEnd(Handheld)
  If Not Handheld Then
    Send_Footer(Handheld)
    Send_InnerWrapEnd(Handheld)
    Send_Legal(Handheld)
  End If
  Send_BodyEnd(Handheld)
  
done:
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
%>
