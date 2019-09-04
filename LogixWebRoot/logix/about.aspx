<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%
    ' *****************************************************************************
    ' * FILENAME: about.aspx
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
    Dim dst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim hDate As New DateTime
    Dim hDateString As String
    Dim shaded As Boolean = True
    Dim i As Integer = 0
    Dim LanguageCode As String = ""
    Dim LogixHost As String = System.Net.Dns.GetHostName
    Dim LogixIPs As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(LogixHost)
    Dim EXInstalled As Boolean = False
    Dim DocPath As String = ""
    Dim DocSize As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    If (System.Environment.GetEnvironmentVariable("LEXSERVER") <> "") AndAlso (System.Environment.GetEnvironmentVariable("LEXDATABASE") <> "") Then
        EXInstalled = True
    End If

    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    MyCommon.Open_LogixWH()
    If EXInstalled Then
        MyCommon.Open_LogixEX()
    End If

    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    Response.Expires = 0
    MyCommon.AppName = "about.aspx"
    Send_HeadBegin("term.about")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 1)
    Send_Subtabs(Logix, 1, 2)

    MyCommon.QueryStr = "select MSNetCode from Languages where LanguageID=" & LanguageID
    dst = MyCommon.LRT_Select
    If dst.Rows.Count > 0 Then
        LanguageCode = dst.Rows(0).Item("MSNetCode")
    End If
%>
<div id="intro">
    <div id="version">
        <%
            MyCommon.QueryStr = "select top 1 VersionID, MajorVersion, MinorVersion, Build, Revision, InstallDate from InstalledVersions with (NoLock) order by InstallDate Desc;"
            dst = MyCommon.LRT_Select
            If dst.Rows.Count > 0 Then
                hDate = dst.Rows(0).Item("InstallDate")
                hDateString = Logix.ToShortDateString(hDate, MyCommon)
                Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID) & " " & dst.Rows(0).Item("MajorVersion") & "." & dst.Rows(0).Item("MinorVersion") & " ")
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.build", LanguageID), VbStrConv.Lowercase) & " " & dst.Rows(0).Item("Build") & " ")
                Sendb(Left(StrConv(Copient.PhraseLib.Lookup("term.revision", LanguageID), VbStrConv.Lowercase), 3) & " " & dst.Rows(0).Item("Revision") & ", ")
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.installed", LanguageID), VbStrConv.Lowercase) & " " & hDateString)
            End If
        %>
    </div>
</div>
<div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
        <div class="box" id="installation">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.installation", LanguageID))%>
                </span>
            </h2>
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.installation", LanguageID))%>">
                <%
                    ' Installation location
                    Send("<tr>")
                    Send("  <td style=""min-width:80px;"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</td>")
                    Send("  <td>" & MyCommon.Fetch_SystemOption(2) & "</td>")
                    Send("</tr>")
                    Send("<tr>")
                    Send("  <td>" & Copient.PhraseLib.Lookup("term.host", LanguageID) & "</td>")
                    Send("  <td>" & LogixIPs.HostName & "</td>")
                    Send("</tr>")
                    Send("<tr>")
                    Send("  <td>" & Copient.PhraseLib.Lookup("term.url", LanguageID) & "</td>")
                    Send("  <td>" & Left(Request.Url.ToString, Request.Url.ToString.LastIndexOf("/") + 1) & "</td>")
                    Send("</tr>")
                    Send("<tr>")
                    Send("  <td>" & Copient.PhraseLib.Lookup("term.IP", LanguageID) & "</td>")
                    Sendb("  <td>")
                    i = 0
                    For Each IP As System.Net.IPAddress In LogixIPs.AddressList
                        If i > 0 Then
                            Sendb(", ")
                        End If
                        Send(IP.ToString)
                        i = i + 1
                    Next
                    Send("</td>")
                    Send("</tr>")

                    ' RT database details
                    Send("<tr>")
                    Send("  <td>RT " & Copient.PhraseLib.Lookup("term.database", LanguageID) & "</td>")
                    Try
                        MyCommon.QueryStr = "select database_id as DatabaseID, name as DatabaseName, create_date as CreateDate, state as State, state_desc as StateDesc " &
                                            "from master.sys.databases where database_id=db_id();"
                        dst = MyCommon.LRT_Select
                        If dst.Rows.Count > 0 Then
                            Sendb("  <td>" & MyCommon.NZ(dst.Rows(0).Item("DatabaseName"), "&nbsp;") & " ")
                            If StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) <> "Online" Then
                                Sendb("(<span class=""red"">" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & "</span>)")
                            Else
                                Sendb("(" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & ")")
                            End If
                            Send("</td>")
                        Else
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                        End If
                    Catch ex As Exception
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    End Try
                    Send("</tr>")

                    ' XS database details
                    Send("<tr>")
                    Send("  <td>XS " & Copient.PhraseLib.Lookup("term.database", LanguageID) & "</td>")
                    Try
                        MyCommon.QueryStr = "select database_id as DatabaseID, name as DatabaseName, create_date as CreateDate, state as State, state_desc as StateDesc " &
                                            "from master.sys.databases where database_id=db_id();"
                        dst = MyCommon.LXS_Select
                        If dst.Rows.Count > 0 Then
                            Sendb("  <td>" & MyCommon.NZ(dst.Rows(0).Item("DatabaseName"), "&nbsp;") & " ")
                            If StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) <> "Online" Then
                                Sendb("(<span class=""red"">" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & "</span>)")
                            Else
                                Sendb("(" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & ")")
                            End If
                            Send("</td>")
                        Else
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                        End If
                    Catch ex As Exception
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    End Try
                    Send("</tr>")

                    ' WH database details
                    Send("<tr>")
                    Send("  <td>WH " & Copient.PhraseLib.Lookup("term.database", LanguageID) & "</td>")
                    Try
                        MyCommon.QueryStr = "select database_id as DatabaseID, name as DatabaseName, create_date as CreateDate, state as State, state_desc as StateDesc " &
                                            "from master.sys.databases where database_id=db_id();"
                        dst = MyCommon.LWH_Select
                        If dst.Rows.Count > 0 Then
                            Sendb("  <td>" & MyCommon.NZ(dst.Rows(0).Item("DatabaseName"), "&nbsp;") & " ")
                            If StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) <> "Online" Then
                                Sendb("(<span class=""red"">" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & "</span>)")
                            Else
                                Sendb("(" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & ")")
                            End If
                            Send("</td>")
                        Else
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                        End If
                    Catch ex As Exception
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    End Try
                    Send("</tr>")

                    ' EX database details
                    If EXInstalled Then
                        Send("<tr>")
                        Send("  <td>EX " & Copient.PhraseLib.Lookup("term.database", LanguageID) & "</td>")
                        Try
                            MyCommon.QueryStr = "select database_id as DatabaseID, name as DatabaseName, create_date as CreateDate, state as State, state_desc as StateDesc " &
                                                "from master.sys.databases where database_id=db_id();"
                            dst = MyCommon.LEX_Select
                            If dst.Rows.Count > 0 Then
                                Sendb("  <td>" & MyCommon.NZ(dst.Rows(0).Item("DatabaseName"), "&nbsp;") & " ")
                                If StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) <> "Online" Then
                                    Sendb("(<span class=""red"">" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & "</span>)")
                                Else
                                    Sendb("(" & StrConv(MyCommon.NZ(dst.Rows(0).Item("StateDesc"), ""), VbStrConv.ProperCase) & ")")
                                End If
                                Send("</td>")
                            Else
                                Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                            End If
                        Catch ex As Exception
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                        End Try
                        Send("</tr>")
                    End If

                    ' Server details
                    Send("<tr>")
                    Send("  <td>" & Copient.PhraseLib.Lookup("term.server", LanguageID) & "</td>")
                    Try
                        MyCommon.QueryStr = "select server_id as ServerID, name as ServerName, product as Product, provider as Provider " &
                                            "from master.sys.servers where server_id=0;"
                        dst = MyCommon.LRT_Select
                        If dst.Rows.Count > 0 Then
                            Sendb("  <td>" & MyCommon.NZ(dst.Rows(0).Item("ServerName"), "&nbsp;") & "</td>")
                        Else
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                        End If
                    Catch ex As Exception
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    End Try
                    Send("</tr>")

                    ' Phrase library details
                    MyCommon.QueryStr = "select top(1) LastUpdate from PhraseText order by LastUpdate desc;"
                    dst = MyCommon.LRT_Select
                    Send("<tr>")
                    Send("  <td>" & Copient.PhraseLib.Lookup("term.phraselib", LanguageID) & "</td>")
                    If dst.Rows.Count > 0 Then
                        hDate = MyCommon.NZ(dst.Rows(0).Item("LastUpdate"), New Date(1900, 1, 1))
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.updated", LanguageID) & " " & Logix.ToShortDateTimeString(hDate, MyCommon) & "</td>")
                    Else
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    End If
                    Send("</tr>")
                %>
            </table>
        </div>
        <div class="box" id="badges">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.about", LanguageID))%>
                </span>
            </h2>
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.content", LanguageID))%>">
                <tr>
                    <td valign="top">
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.content", LanguageID))%></b>
                        <br />
                        <a href="http://developer.mozilla.org/en/JavaScript" target="_blank">
                            <img src="../images/badges/javascript.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.javascript", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.javascript", LanguageID)) %>" /></a><br />
                        <a href="http://www.cookiecentral.com/faq/" target="_blank">
                            <img src="../images/badges/cookies.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.cookies", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.cookies", LanguageID)) %>" /></a><br />
                        <a href="http://www.w3.org/Style/CSS/" target="_blank">
                            <img src="../images/badges/css.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.css", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.css", LanguageID)) %>" /></a><br />
                        <a href="http://www.w3.org/MarkUp/" target="_blank">
                            <img src="../images/badges/xhtml.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.xhtml", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.xhtml", LanguageID)) %>" /></a><br />
                        <a href="http://www.unicode.org/" target="_blank">
                            <img src="../images/badges/utf8.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.utf8", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.utf8", LanguageID)) %>" /></a><br />
                        <a href="http://www.libpng.org/pub/png/" target="_blank">
                            <img src="../images/badges/png.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.png", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.png", LanguageID)) %>" /></a><br />
                        <a href="http://get.adobe.com/reader/" target="_blank">
                            <img src="../images/badges/pdf.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.pdf", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.pdf", LanguageID)) %>" /></a><br />
                        <br />
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.powered", LanguageID))%></b>
                        <br />
                        <a href="http://www.asp.net/" target="_blank">
                            <img src="../images/badges/aspnet4.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.aspnet4", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.aspnet4", LanguageID)) %>" /></a><br />
                        <a href="http://www.microsoft.com/sql/" target="_blank">
                            <img src="../images/badges/sql2008.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.sql", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.sql", LanguageID)) %>" /></a><br />
                        <a href="http://www.microsoft.com/sql/" target="_blank">
                            <img src="../images/badges/sql2012.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.sql12", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.sql12", LanguageID)) %>" /></a><br />
                        <a href="http://www.microsoft.com/sql/" target="_blank">
                            <img src="../images/badges/sql2016.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.sql16", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.sql16", LanguageID)) %>" /></a><br />
                        <br />
                    </td>
                    <td valign="top">
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.compatibility", LanguageID))%></b>
                        <br />
                        <a href="http://www.microsoft.com/windows/internet-explorer/" target="_blank">
                            <img src="../images/badges/explorer.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.internetexplorer", LanguageID))%> 9.0, 10.0 <% Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> 11.0"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.internetexplorer", LanguageID))%> 9.0, 10.0 <% Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> 11.0" /></a><br />
                        <a href="https://www.google.com/chrome/" target="_blank">
                            <img src="../images/badges/chrome.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.chrome", LanguageID))%>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.chrome", LanguageID))%>" /></a><br />
                        <!--
            AL-1506 AMS only officially supports IE 8 and 9 per 5.19 SRD
            <a href="http://www.mozilla.com/<%Sendb(LanguageCode) %>/firefox/" target="_blank"><img src="../images/badges/firefox.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.firefox", LanguageID))%> 12+" title="<% Sendb(Copient.PhraseLib.Lookup("term.firefox", LanguageID))%> 12+" /></a><br />
            <a href="http://www.google.com/chrome/" target="_blank"><img src="../images/badges/chrome.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.chrome", LanguageID))%> 19+" title="<% Sendb(Copient.PhraseLib.Lookup("term.chrome", LanguageID))%> 19+" /></a><br />
            -->
                        <!--
            <a href="http://www.apple.com/safari/" target="_blank"><img src="../images/badges/safari.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.safari", LanguageID))%> 5+" title="<% Sendb(Copient.PhraseLib.Lookup("term.safari", LanguageID))%> 5+" /></a><br />
            <a href="http://www.opera.com/" target="_blank"><img src="../images/badges/opera.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.opera", LanguageID))%> 12+" title="<% Sendb(Copient.PhraseLib.Lookup("term.opera", LanguageID))%> 12+" /></a><br />
            <a href="http://caminobrowser.org/" target="_blank"><img src="../images/badges/camino.png" class="badge" alt="Camino" title="Camino" /></a><br />
            <a href="http://gnome.org/projects/epiphany/" target="_blank"><img src="../images/badges/epiphany.png" class="badge" alt="Epiphany" title="Epiphany" /></a><br />
            <a href="http://www.flock.com/" target="_blank"><img src="../images/badges/flock.png" class="badge" alt="Flock" title="Flock" /></a><br />
            <a href="http://galeon.sourceforge.net/" target="_blank"><img src="../images/badges/galeon.png" class="badge" alt="Galeon" title="Galeon" /></a><br />
            <a href="http://www.icab.de/" target="_blank"><img src="../images/badges/icab.png" class="badge" alt="iCab" title="iCab" /></a><br />
            <a href="http://kmeleon.sf.net/" target="_blank"><img src="../images/badges/k-meleon.png" class="badge" alt="K-Meleon" title="K-Meleon" /></a><br />
            <a href="http://www.konqueror.org/" target="_blank"><img src="../images/badges/konqueror.png" class="badge" alt="Konqueror" title="Konqueror" /></a><br />
            <a href="http://lynx.isc.org/" target="_blank"><img src="../images/badges/lynx.png" class="badge" alt="Lynx" title="Lynx" /></a><br />
            <a href="http://www.mozilla.org/" target="_blank"><img src="../images/badges/mozilla.png" class="badge" alt="Mozilla" title="Mozilla" /></a><br />
            <a href="http://browser.netscape.com" target="_blank"><img src="../images/badges/netscape.png" class="badge" alt="Netscape" title="Netscape" /></a><br />
            <a href="http://www.omnigroup.com/applications/omniweb/" target="_blank"><img src="../images/badges/omniweb.png" class="badge" alt="OmniWeb" title="OmniWeb" /></a>
            <a href="http://www.seamonkey-project.org/" target="_blank"><img src="../images/badges/seamonkey.png" class="badge" alt="SeaMonkey" title="SeaMonkey" /></a><br />
            -->
                        <br />
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.conformance", LanguageID))%></b>
                        <br />
                        <a href="http://www.w3.org/TR/WCAG10/" target="_blank">
                            <img src="../images/badges/wcaga.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.wcag", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.wcag", LanguageID)) %>" /></a><br />
                        <a href="http://www.section508.gov/" target="_blank">
                            <img src="../images/badges/section508.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.section508", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.section508", LanguageID)) %>" /></a><br />
                        <!--
            <a href="http://www.anybrowser.org/campaign/" target="_blank"><img src="../images/badges/anybrowser.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.vwab", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.vwab", LanguageID)) %>" /></a><br />
            -->
                        <a href="http://webstandardsgroup.org/" target="_blank">
                            <img src="../images/badges/wsg.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.wsg", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.wsg", LanguageID)) %>" /></a><br />
                        <a href="http://www.hwg.org/" target="_blank">
                            <img src="../images/badges/hwg.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.hwg", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.hwg", LanguageID)) %>" /></a><br />
                        <br />
                    </td>
                    <td valign="top">
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.corporate", LanguageID))%></b>
                        <br />
                        <a href="http://www.ncr.com/company" target="_blank">
                            <img src="../images/badges/ncr.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.ncr", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.ncr", LanguageID)) %>" /></a><br />
                        <a href="http://www.ncr.com/retail/petroleum-convenience/customer-engagement/loyalty-offer-management/advanced-management-solution-ams"
                            target="_blank">
                            <img src="../images/badges/ncrams.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.ncrams", LanguageID)) %>"
                                title="<% Sendb(Copient.PhraseLib.Lookup("term.ncrams", LanguageID)) %>" /></a><br />
                        <!--
            <a href="http://www.copienttech.com/" target="_blank"><img src="../images/badges/copient.png" class="badge" alt="<% Sendb(Copient.PhraseLib.Lookup("term.copienttechnologies", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.copienttechnologies", LanguageID)) %>" /></a><br />
            -->
                        <br />
                    </td>
                </tr>
            </table>
            <hr class="hidden" />
        </div>
        <div class="box" id="contact">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contact", LanguageID))%>
                </span>
            </h2>
            <% Send(Copient.PhraseLib.Lookup("about.copientcontact", LanguageID))%>
            <% Send("<a href=""mailto:" & MyCommon.Fetch_SystemOption(40) & """>" & MyCommon.Fetch_SystemOption(40) & "</a><br />")%>
        </div>
        <div class="box" id="legal">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.legal", LanguageID))%>
                </span>
            </h2>
            <%
                Send(Copient.PhraseLib.Lookup("about.copyright", LanguageID) & "  " & Copient.PhraseLib.Lookup("about.access", LanguageID) & "<br />")
            %>
        </div>
    </div>
    <div class="gutter">
    </div>
    <div id="column2">
        <div class="box" id="documentation">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.documentation", LanguageID))%>
                </span>
            </h2>
            <%
                DocPath = "../documentation/Logix User Manual.pdf"
                DocSize = GetFileSize(DocPath)
                Send("<a href=""" & DocPath & """ target=""manual"">")
                Send("  <img src=""../images/pdf.png"" /> ")
                Send("  " & Copient.PhraseLib.Lookup("term.logix", LanguageID) & " " & Copient.PhraseLib.Lookup("term.usermanual", LanguageID))
                Send("</a> " & If(DocSize <> "", " <small>(" & DocSize & ")</small>", ""))
                If MyCommon.IsEngineInstalled(0) Then
                    DocPath = "../documentation/Logix User Manual - CM Offer Builder.pdf"
                    DocSize = GetFileSize(DocPath)
                    Send("<br />")
                    Send("<a href=""" & DocPath & """ target=""manual"">")
                    Send("  <img src=""../images/pdf.png"" /> ")
                    Send("  " & Copient.PhraseLib.Lookup("term.logix", LanguageID) & " " & Copient.PhraseLib.Lookup("term.usermanual", LanguageID) & " - " & Copient.PhraseLib.Lookup("term.cm", LanguageID) & " " & Copient.PhraseLib.Lookup("term.offerbuilder", LanguageID))
                    Send("</a> " & If(DocSize <> "", " <small>(" & DocSize & ")</small>", ""))
                End If
                If MyCommon.IsEngineInstalled(2) Then
                    DocPath = "../documentation/Logix User Manual - CPE Offer Builder.pdf"
                    DocSize = GetFileSize(DocPath)
                    Send("<br />")
                    Send("<a href=""" & DocPath & """ target=""manual"">")
                    Send("  <img src=""../images/pdf.png"" /> ")
                    Send("  " & Copient.PhraseLib.Lookup("term.logix", LanguageID) & " " & Copient.PhraseLib.Lookup("term.usermanual", LanguageID) & " - " & Copient.PhraseLib.Lookup("term.cpe", LanguageID) & " " & Copient.PhraseLib.Lookup("term.offerbuilder", LanguageID))
                    Send("</a> " & If(DocSize <> "", " <small>(" & DocSize & ")</small>", ""))
                End If
                If MyCommon.IsEngineInstalled(9) Then
                    DocPath = "../documentation/Logix User Guide for the Universal Engine (UE).pdf"
                    DocSize = GetFileSize(DocPath)
                    Send("<br />")
                    Send("<a href=""" & DocPath & """ target=""manual"">")
                    Send("  <img src=""../images/pdf.png"" /> ")
                    Send("  " & Copient.PhraseLib.Lookup("term.uedoc", LanguageID))
                    Send("</a> " & If(DocSize <> "", " <small>(" & DocSize & ")</small>", ""))
                End If
            %>
        </div>
        <div class="box" id="versions">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.versionhistory", LanguageID))%>
                </span>
            </h2>
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.versionhistory", LanguageID))%>">
                <thead>
                    <tr>
                        <th class="th-version">
                            <% Sendb(Copient.PhraseLib.Lookup("term.version", LanguageID))%>
                        </th>
                        <th>
                            <% Sendb(Copient.PhraseLib.Lookup("term.installed", LanguageID))%>
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        MyCommon.QueryStr = "select VersionID, MajorVersion, MinorVersion, Build, Revision, InstallDate from InstalledVersions with (NoLock) order by InstallDate DESC;"
                        dst = MyCommon.LRT_Select
                        i = 0
                        For Each row In dst.Rows
                            hDate = MyCommon.NZ(row.Item("InstallDate"), New Date(1900, 1, 1))
                            Sendb("<tr")
                            If shaded = True Then Sendb(" class=""shaded""")
                            Send(">")
                            Sendb("  <td title=""" & MyCommon.NZ(row.Item("VersionID"), 0) & """>")
                            Sendb("    " & MyCommon.NZ(row.Item("MajorVersion"), 0) & "." & MyCommon.NZ(row.Item("MinorVersion"), 0) & "." & MyCommon.NZ(row.Item("Build"), 0) & "." & MyCommon.NZ(row.Item("Revision"), 0))
                            Send("</td>")
                            Send("  <td>")
                            Send("    " & Logix.ToShortDateTimeString(hDate, MyCommon))
                            Send("  </td>")
                            Send("</tr>")
                            If shaded = True Then
                                shaded = False
                            Else
                                shaded = True
                            End If
                            i = i + 1
                        Next
                    %>
                </tbody>
            </table>
            <hr class="hidden" />
        </div>
    </div>
</div>
<div style="position: absolute; bottom: 2px; right: 2px;">
    <a href="#" accesskey="\">
        <img src="../images/blackdot.png" alt="" title="" /></a>
</div>
<script runat="server">
    Private Function GetFileSize(ByVal DocPath As String) As String
        Dim Doc As New System.IO.FileInfo(Server.MapPath(DocPath))
        Dim Bytes As Long = 0
        Dim DocSize As String = ""

        If Doc.Exists Then
            Bytes = Doc.Length
            DocSize = String.Format("{0:0,0}", Int(Bytes / 1000)) & " KB"
        End If

        Return DocSize
    End Function
</script>
<%
done:
    Send_BodyEnd()
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    MyCommon.Close_LogixWH()
    MyCommon.Close_LogixEX()
    MyCommon = Nothing
    Logix = Nothing
%>