<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-general.aspx 
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
  
%>
  
  
<script runat="server">

    Dim Common As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim Cust As New Copient.Customer
    Dim MyLookup As New Copient.CustomerLookup()
    Dim Handheld As Boolean = False
    Dim Debug As Boolean = False


    '---------------------------------------------------------------------


    Sub Get_Customer_Info(ByVal CustomerPK As Long, ByRef CardPK As Long, ByRef ExtCardID As String, ByRef IsHousehold As Boolean)

        Dim dst As DataTable

        CardPK = 0
        ExtCardID = ""
        IsHousehold = False

        If CustomerPK > 0 Then
            CardPK = Common.Extract_Val(GetCgiValue("CardPK"))

            If CardPK > 0 Then
                ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
            End If
            Common.QueryStr = "select C.CustomerTypeID  " & _
                                  "from Customers C with (NoLock) " & _
                                  "where C.CustomerPK=" & CustomerPK
            dst = Common.LXS_Select
            If dst.Rows.Count > 0 Then
                If dst.Rows(0).Item("CustomerTypeID") = 1 Then IsHousehold = True
            End If
        End If

    End Sub


    '---------------------------------------------------------------------

    Sub Send_Page_Title(ByVal IsHousehold As Boolean, ByVal ExtCardID As String)

        If ExtCardID = "" Then
            If (IsHousehold) Then
                Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
            Else
                Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
            End If
        Else
            If (IsHousehold) Then
                Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & ExtCardID)
            Else
                Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
            End If
        End If

    End Sub

    '---------------------------------------------------------------------  

    Sub Send_Prefs_Box(ByVal CustomerPK As Long)

        Dim FormData As String
        Dim RawRequest As String
        Dim RawURI As String = ""
        Dim HostURI As String = ""
        Dim TargetAddress As String
        Dim dst As DataTable
        Dim ConnInc As New Copient.ConnectorInc()

        'The code for the contents of this box lives in customer.prefs.editbox.aspx.  

        RawRequest = Get_Raw_RequestData(Request.InputStream)
        If Debug Then
            Send("<!-- Raw data:")
            Send(RawRequest)
            Send("-->")
        End If

        Common.QueryStr = "select isnull(HTTP_RootURI, '') as HTTP_RootURI from Integrations where IntegrationID=1;"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            HostURI = dst.Rows(0).Item("HTTP_RootURI")
        End If
        dst = Nothing

        HostURI = Trim(HostURI)
        If HostURI = "" Then
            Send(Copient.PhraseLib.Lookup("customer-languages.URINotSet", LanguageID))
        Else
            If Not (Right(HostURI, 1) = "/") Then HostURI = HostURI & "/"
            Send("<!-- HostURI=" & HostURI & " -->")
            TargetAddress = HostURI & "UI/customer.prefs.editbox.aspx"
            Send("<! -- TargetAddress=" & TargetAddress & " -->")
            'Open_UI_Box(4, AdminUserID, Common, "")
            Send("<div class=""box"" id=""prefeditbox"">")
            Send("<h2>")
            Send("  <span>")
            Sendb(Copient.PhraseLib.Lookup("term.preferences", LanguageID))
            Send("  </span>")
            Send("</h2>")

            FormData = "AuthToken=" & HttpUtility.UrlEncode(Request.Cookies("AuthToken").Value) & "&ParentURI=customer-prefs.aspx&ThemeURIHost=" & HostURI & "UI/" & "&CustomerPK=" & CustomerPK & "&" & RawRequest
            Send(ConnInc.Retrieve_HttpResponse(TargetAddress, FormData))
            'Close_UI_Box()
        End If
        Send("</div><!-- prefedibox -->")


    End Sub

    '---------------------------------------------------------------------


    Sub Send_Page(ByVal CustomerPK As Long)

        Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
        Dim CopientFileVersion As String = "7.3.1.138972"
        Dim CopientProject As String = "Copient Logix"
        Dim CopientNotes As String = ""
        Dim restrictLinks As Boolean = False
        Dim ExtraLink As String = ""
        Dim CardPK As Long = 0
        Dim ExtCardID As String = ""
        Dim IsHousehold As Boolean = False

        MyLookup.SetAdminUserID(AdminUserID)
        MyLookup.SetLanguageID(LanguageID)
        restrictLinks = MyLookup.IsRestrictedUser(AdminUserID)

        Get_Customer_Info(CustomerPK, CardPK, ExtCardID, IsHousehold)

        If CardPK > 0 Then
            Send_HeadBegin("term.customer", "term.preferences", Common.Extract_Val(ExtCardID))
        Else
            Send_HeadBegin("term.customer", "term.general")
        End If
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 3)
        If (Not restrictLinks) Then
            Send_Tabs(Logix, 3)
            Send_Subtabs(Logix, 32, 7, LanguageID, CustomerPK, , CardPK)
        Else
            Send_Subtabs(Logix, 91, 4, LanguageID, CustomerPK, ExtraLink, cardpk)
        End If

        Send("<div id=""intro"">")
        Sendb(" <h1 id=""title"">")
        Send_Page_Title(IsHousehold, ExtCardID)
        Sendb("</h1>")
        'If (Logix.UserRoles.EditSystemConfiguration = True) Then
        Send(" <div id=""controls"">")
        Send(" </div> <!-- controls -->")
        'End If
        Send("</div>")
        Send("")
        Send("<div id=""main"">")
        'If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
        Send("<div id=""column"">")
        Send("<div id=""statusbar"" class=""red-background"" style=""display: none;""></div>")

        Send_Prefs_Box(CustomerPK)

        Send("</div><br clear=""all"" /> <!-- column -->")

        Send_BodyEnd("mainform", "runfreq")

    End Sub

    '---------------------------------------------------------------------


</script>


<%
    Dim Mode As String
    Dim AppID As Long
    Dim CustomerPK As Long = 0

    Response.Expires = 0
    Common.AppName = "customer-prefs.aspx"
    On Error GoTo ErrorTrap

    Common.Open_LogixRT()
    Common.Open_LogixXS()
    AdminUserID = Verify_AdminUser(Common, Logix)
    CMS.AMS.CurrentRequest.Resolver.AppName = Common.AppName

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    CustomerPK = Common.Extract_Val(GetCgiValue("CustPK"))
    Send_Page(CustomerPK)

    GoTo Finish

ErrorTrap:
    Send("<pre>" & Common.Error_Processor() & "</pre>")

Finish:
    Common.Close_LogixRT()
    Common.Close_LogixXS()
    Common = Nothing
    Logix = Nothing
    Response.End()

%>
