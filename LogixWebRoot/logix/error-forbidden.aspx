<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: error-application.aspx 
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
    Dim Handheld As Boolean = False
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "error-forbidden.aspx"
    MyCommon.Open_LogixRT()
    LanguageID = MyCommon.Fetch_SystemOption(1)
    Send_HeadBegin("term.error")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Logos(LanguageID)
    Send_Tabs(Logix, 0)
    'Send("<div class=""tabs"" id=""tabs"">&nbsp;")
    'Send("<hr class=""hidden"" />")
    'Send("</div>")
    'Send("")
%>
<div id="intro">
    <h1 id="title">
        <% Sendb(Copient.PhraseLib.Lookup("term.error", LanguageID))%>
    </h1>
</div>
<div id="main">
    <% 
        If Request.QueryString("error") <> "" Then
            Sendb(Request.QueryString("error"))
        Else
            Sendb(Copient.PhraseLib.Lookup("error.unauthorized-activity-error-message", LanguageID))
        End If
    %>
</div>
<%
done:
    Send_BodyEnd()
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    Logix = Nothing
%>
