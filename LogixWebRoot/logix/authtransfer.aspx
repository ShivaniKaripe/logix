<%@ Page Language="vb" Debug="true" CodeFile="logixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: login.aspx 
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
    Public MyCommon As New Copient.CommonInc
    Public Logix As New Copient.LogixInc

    '-------------------------------------------------------------------------------------------------------------

    Sub Send_Page()
        Dim SendToURI As String
        Dim TransferKey As String
        Dim TransferParam As String
        Dim Debug As Boolean = False
        Dim isError As Boolean = False
        Dim LogixHostURI As String
        Dim EPMHostURI As String
        Dim SendUrl As String

        If Debug Then MyCommon.Write_Log("auth.txt", "In AuthTranfer.aspx - AdminUserID=" & AdminUserID)
        SendToURI = System.Web.HttpUtility.UrlDecode(GetCgiValue("sendtouri"))
        If SendToURI = "" Then
            Send(Copient.PhraseLib.Lookup("authtransfer.SendToURIError", LanguageID))
        Else
            'get Logix Host URI
            LogixHostURI = HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Authority)
            'get EPM Host Integration URI
            Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
            If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)
                EPMHostURI = IntegrationVals.HTTP_RootURI
            End If

            If Not SendToURI.StartsWith(LogixHostURI) And (string.IsNullOrEmpty(EPMHostURI) Or Not SendToURI.StartsWith(EPMHostURI)) Then
                If Not ValidateRelativePathURL(SendToURI) Then
                    Send(Copient.PhraseLib.Lookup("error.notfound", LanguageID))
                    isError = True
                End If
            End If
            If Not isError
                'create the TransferKey
                TransferKey = MyCommon.Generate_GUID
                TransferParam = "transferkey=" & TransferKey
                If InStr(SendToURI, "?") > 0 Then
                    TransferParam = "&" & TransferParam
                Else
                    TransferParam = "?" & TransferParam
                End If
                SendUrl = SendToURI & TransferParam
            Else
                SendUrl = "error-forbidden.aspx"
            End If

            
            'store the TransferKey in the AdminUsers table
            MyCommon.QueryStr = "Update AdminUsers set TransferKey='" & TransferKey & "' where AdminUserID=" & AdminUserID & ";"
            MyCommon.LRT_Execute()
            If Debug Then MyCommon.Write_Log("auth.txt", "In AuthTranfer.aspx - Storing the transfer key: " & MyCommon.QueryStr)
            'Bounce the user to the target URI with their transfer key
            Send("<html xmlns=""http://www.w3.org/1999/xhtml"">")
            Send("<head>")
            Send("<title>" & Copient.PhraseLib.Lookup("term.logix", LanguageID) & "</title>")
            Send("<meta http-equiv=""Refresh"" content=""0; URL=" & SendUrl & """>")
            Send("</head>")
            Send("<body bgcolor=""#ffffff"">")
            Send("<!-- bouncing -->")
            Send("</body>")
            Send("</html>")

        End If

    End Sub

    Private Function ValidateRelativePathURL(ByVal url As String) As Boolean
        Dim validatedUri As Uri = Nothing
        'Url passed is invalid if its absolute and does not start with Logix Website
        If Uri.IsWellFormedUriString(url, UriKind.Absolute) Then
            Return False
        ElseIf Uri.IsWellFormedUriString(url, UriKind.Relative) Then
            Return True
        End If

        Return False
    End Function


</script>

<%
  Dim Mode As String
  MyCommon.AppName = "login.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Send_Page()
  
  If Not (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Close_LogixRT()
%>

<%
  Response.End()
ErrorTrap:
  Response.Write("<pre>" & MyCommon.Error_Processor() & "</pre>")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
%>
