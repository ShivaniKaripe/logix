'
' version:7.3.1.138972.Official Build (SUSDAY10202)
' *****************************************************************************
' * FILENAME: ue-cb.vb
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
' * MODULE  : Logixinc
' *
' * PURPOSE :
' *
' * NOTES   :
' *
' * Version : 7.3.1.138972
' *
' *****************************************************************************
'

Imports System.Data

Public Class UECB
    Inherits System.Web.UI.Page
    'Public Logix As New Copient.LogixInc
    'Public Common As New Copient.CommonInc
    Public AdminName As String
    Public LanguageID As Integer
    Public AdminUserID As Integer
    Public OfferLockedforCollisionDetection As Boolean = False

    Dim Group_Record_Limit As Integer

    Public ReadOnly Property GroupRecordLimit() As Integer
        Get
            Dim MyCommon As New Copient.CommonInc
            MyCommon.Open_LogixRT()
            Group_Record_Limit = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(126)) '500
            MyCommon.Close_LogixRT()
            MyCommon = Nothing
            Return Group_Record_Limit
        End Get
        'Set(ByVal value)

        'End Set
    End Property

    Public Function Verify_AdminUser(ByRef Common As Copient.CommonInc, ByRef MyLogix As Object) As Long
        Dim AdminUserID As Long
        Dim Authtoken As String
        Dim MyURI As String
        Authtoken = ""
        If Not (Request.Cookies("AuthToken") Is Nothing) Then
            Authtoken = Request.Cookies("AuthToken").Value
        End If
        AdminUserID = 0
        AdminUserID = MyLogix.Auth_Token_Verify(Common, Authtoken, AdminName, LanguageID)
        If AdminUserID = 0 Then
            MyURI = System.Web.HttpUtility.UrlEncode(Request.Url.AbsoluteUri)
            Send("<!DOCTYPE html ")
            Send("     PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN""")
            Send("     ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
            Send("<html xmlns=""http://www.w3.org/1999/xhtml"">")
            Send("<head>")
            Send("<meta http-equiv=""refresh"" content=""0; url=/logix/login.aspx?mode=invalid&amp;bounceback=" & MyURI & """ />")
            Send("<title>Logix</title>")
            Send("</head>")
            Send("<body bgcolor=""#ffffff"">")
            Send("<!-- Bouncing -->")
            Send("</body>")
            Send("</html>")
            Response.End()
        End If
        Verify_AdminUser = AdminUserID
    End Function

    Public Sub Send(ByVal WebText As String)
        Response.Write(WebText & vbCrLf)
    End Sub

    Public Sub Sendb(ByVal WebText As String)
        Response.Write(WebText)
    End Sub

    Public Function CleanString(ByVal InString As String, Optional ByVal AdditionalValidCharacters As String = "") As String
        Dim tmpString As String = ""
        Dim z As Integer
        Dim supportstring As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.#$*-&%@!?/: "
        If AdditionalValidCharacters <> "" Then
            supportstring = supportstring & "_;+"
        End If
        If InString IsNot Nothing Then
            For z = 0 To InString.Length - 1
                If (InStr(supportstring & AdditionalValidCharacters, InString(z))) Then
                    tmpString = tmpString & InString(z)
                End If
            Next
        End If

        CleanString = tmpString
    End Function

    Public Function CleanUPC(ByVal InString As String) As String
        Dim z As Integer
        Dim IsClean As Boolean

        If InString IsNot Nothing Then
            For z = 0 To InString.Length - 1
                If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", InString(z))) Then
                    IsClean = True
                Else
                    IsClean = False
                    GoTo breakout
                End If
            Next
        End If

breakout:
        CleanUPC = IsClean
    End Function

    Public Sub Send_HeadBegin(Optional ByVal PageTitle As String = "", Optional ByVal PageSubTitle As String = "", Optional ByVal PageID As Decimal = 0, Optional ByVal DefaultLanguageID As Integer = 0)
        Dim TempLanguageID As Integer
        Dim MyCommon As New Copient.CommonInc
        Dim dst As System.Data.DataTable
        Dim row As System.Data.DataRow

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        If DefaultLanguageID > 0 Then
            TempLanguageID = DefaultLanguageID
        Else
            TempLanguageID = LanguageID
        End If
        Send("<!-- IE6 quirks mode -->")
        Send("<!DOCTYPE html ")
        Send("     PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN""")
        Send("     ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
        MyCommon.QueryStr = "select top 1 MSNetCode from Languages as L with (NoLock) where L.LanguageID=" & TempLanguageID
        dst = MyCommon.LRT_Select
        For Each row In dst.Rows
            Send("<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""" & row.Item("MSNetCode") & """ xml:lang=""" & row.Item("MSNetCode") & """>")
        Next
        Send("<head>")
        Sendb("<title>" & Copient.PhraseLib.Lookup("term.logix", TempLanguageID))
        If (PageTitle = "") Then
        Else
            Sendb(" > " & Copient.PhraseLib.Lookup(PageTitle, TempLanguageID))
        End If
        If (PageID <= 0) Then
        Else
            Sendb(" " & PageID)
        End If
        If (PageSubTitle = "") Then
        Else
            Sendb(" > " & Copient.PhraseLib.Lookup(PageSubTitle, TempLanguageID))
        End If
        Send("</title>")
        MyCommon.Close_LogixRT()
        MyCommon = Nothing
    End Sub

    Public Sub Send_Comments(Optional ByVal CopientProject As String = "", Optional ByVal CopientFileName As String = "", Optional ByVal CopientFileVersion As String = "", Optional ByVal CopientNotes As String = "")
        Send("<!-- ")
        Send("PROJECT:  " & IIf(CopientProject = "", "...", CopientProject))
        Send("FILENAME: " & IIf(CopientFileName = "", "...", CopientFileName))
        Send("VERSION:  " & IIf(CopientFileVersion = "", "...", CopientFileVersion))
        Send("NOTES:    " & IIf(CopientNotes = "", "...", CopientNotes))
        Send("-->")
    End Sub

    Public Sub Send_Metas(Optional ByVal DefaultLanguageID As Integer = 0)
        Dim TempLanguageID As Integer
        Dim MyCommon As New Copient.CommonInc
        Dim dst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim longDate As New DateTime
        Dim longDateString As String

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()

        If DefaultLanguageID > 0 Then
            TempLanguageID = DefaultLanguageID
        Else
            TempLanguageID = LanguageID
        End If
        MyCommon.QueryStr = "select top 1 VersionID, MajorVersion, MinorVersion, Build, Revision, InstallDate from InstalledVersions with (NoLock) order by InstallDate Desc;"
        dst = MyCommon.LRT_Select
        For Each row In dst.Rows
            Sendb("<meta name=""version"" content=""" & row.Item("MajorVersion") & "." & row.Item("MinorVersion") & " " & Copient.PhraseLib.Lookup("term.build", TempLanguageID) & " " & row.Item("Build") & "." & row.Item("Revision"))
            longDate = row.Item("InstallDate")
            longDateString = longDate.ToString("MMMM d, yyyy")
            Send(" (" & longDateString & ")"" />")
        Next
        Send("<meta name=""author"" content=""" & Copient.PhraseLib.Lookup("about.copientaddress", TempLanguageID) & """ />")
        Send("<meta name=""copyright"" content=""" & Copient.PhraseLib.Lookup("about.copyright", TempLanguageID) & """ />")
        Send("<meta name=""description"" content=""" & Copient.PhraseLib.Lookup("about.description", TempLanguageID) & """ />")
        Send("<meta name=""content-type"" content=""text/html; charset=utf-8"" />")
        Send("<meta name=""robots"" content=""noindex, nofollow"" />")
        Send("<meta name=""viewport"" content=""width=782"" />")
        Send("<meta http-equiv=""cache-control"" content=""no-cache"" />")
        Send("<meta http-equiv=""pragma"" content=""no-cache"" />")
        Send("<meta http-equiv=""X-UA-Compatible"" content=""IE=9"" />")
        MyCommon.Close_LogixRT()
        MyCommon = Nothing
    End Sub

    Private Function GetStyleFileNames() As String()
        Dim FileName As String() = {"logix-screen.css", ""}
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim dst As System.Data.DataTable
        Dim dst2 As System.Data.DataTable
        Dim AdminUserID As Long

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)

        MyCommon.QueryStr = "select UIS.FileName, UIS.BaseStyle from UIStyles as UIS " &
                            "inner join AdminUsers as AU on AU.StyleID=UIS.StyleID " &
                            "where AdminUserID=" & AdminUserID & ";"
        dst = MyCommon.LRT_Select
        If dst.Rows.Count > 0 Then
            'User-specific style defined.
            If Not MyCommon.NZ(dst.Rows(0).Item("BaseStyle"), False) Then
                FileName(0) = "logix-screen.css"
                FileName(1) = MyCommon.NZ(dst.Rows(0).Item("FileName"), "logix-screen.css")
            Else
                ReDim FileName(0)
                FileName(0) = MyCommon.NZ(dst.Rows(0).Item("FileName"), "logix-screen.css")
            End If
        Else
            'No style defined, so fall back on the system default
            MyCommon.QueryStr = "select FileName, BaseStyle from UIStyles where DefaultStyle=1 and StyleID>1;"
            dst2 = MyCommon.LRT_Select
            If dst2.Rows.Count > 0 Then
                If dst2.Rows(0).Item("BaseStyle") Then
                    FileName(0) = MyCommon.NZ(dst2.Rows(0).Item("FileName"), "logix-screen.css")
                Else
                    FileName(0) = "logix-screen.css"
                    FileName(1) = MyCommon.NZ(dst2.Rows(0).Item("FileName"), "logix-screen.css")
                End If
            End If
        End If

        MyCommon.Close_LogixRT()
        MyCommon = Nothing

        Return FileName
    End Function

    Public Sub Send_Links(Optional ByVal Handheld As Boolean = False, Optional ByVal Restricted As Boolean = False)
        Dim MyCommon As New Copient.CommonInc
        Dim dt As System.Data.DataTable
        Dim myUrl As String = ""
        Dim FileNames As String() = Nothing
        Dim i As Integer = 0

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()

        Send("<link rel=""icon"" href=""/images/favicon.ico"" type=""image/x-icon"" />")
        Send("<link rel=""shortcut icon"" href=""/images/favicon.ico"" type=""image/x-icon"" />")
        Send("<link rel=""apple-touch-icon"" href=""/images/touchicon.png"" />")
        myUrl = Request.CurrentExecutionFilePath
        If (myUrl = "/logix/login.aspx" OrElse myUrl = "/logix/requirements.aspx") Then
            If Not (Request.Cookies("Style") Is Nothing) Then
                If Request.Cookies("Style").Value <> "" Then
                    MyCommon.QueryStr = "select FileName, BaseStyle from UIStyles with (NoLock) where StyleID=" & Request.Cookies("Style").Value & ";"
                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count > 0 Then
                        If Not MyCommon.NZ(dt.Rows(0).Item("BaseStyle"), False) Then
                            Send("<link rel=""stylesheet"" href=""/css/logix-screen.css"" type=""text/css"" media=""screen"" />")
                        End If
                        Send("<link rel=""stylesheet"" href=""/css/" & dt.Rows(0).Item("FileName") & """ type=""text/css"" media=""screen"" />")
                    Else
                        Send("<link rel=""stylesheet"" href=""/css/logix-screen.css"" type=""text/css"" media=""screen"" />")
                        Request.Cookies("Style").Expires = DateTime.Now.AddDays(-1D)
                    End If
                Else
                    Send("<link rel=""stylesheet"" href=""/css/logix-screen.css"" type=""text/css"" media=""screen"" />")
                    Request.Cookies("Style").Expires = DateTime.Now.AddDays(-1D)
                End If
            Else
                MyCommon.QueryStr = "select FileName, BaseStyle from UIStyles where DefaultStyle=1 and StyleID>1;"
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    If dt.Rows(0).Item("BaseStyle") Then
                        Send("<link rel=""stylesheet"" href=""/css/" & dt.Rows(0).Item("FileName") & """ type=""text/css"" media=""screen"" />")
                    Else
                        Send("<link rel=""stylesheet"" href=""/css/logix-screen.css"" type=""text/css"" media=""screen"" />")
                        Send("<link rel=""stylesheet"" href=""/css/" & dt.Rows(0).Item("FileName") & """ type=""text/css"" media=""screen"" />")
                    End If
                Else
                    Send("<link rel=""stylesheet"" href=""/css/logix-screen.css"" type=""text/css"" media=""screen"" />")
                End If
            End If
        Else
            FileNames = GetStyleFileNames()
            For i = 0 To FileNames.GetUpperBound(0)
                Send("<link rel=""stylesheet"" href=""/css/" & FileNames(i) & """ type=""text/css"" media=""screen"" />")
            Next
        End If
        If Handheld Then
            Send("<link rel=""stylesheet"" href=""/css/logix-handheld.css"" type=""text/css"" media=""screen, handheld"" />")
        End If
        Send("<link rel=""stylesheet"" href=""/css/logix-aural.css"" type=""text/css"" media=""aural"" />")
        Send("<link rel=""stylesheet"" href=""/css/logix-print.css"" type=""text/css"" media=""braille, embossed, print, projection, tty"" />")
        If Restricted Then
            Send("<link rel=""stylesheet"" href=""/css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
        End If

        If Request.Browser.Browser = "IE" Or Request.Browser.Browser = "Opera" Then
            'IE-specific multilanguage-input tweak
            Send("<style type=""text/css"">")
            Send("  .ml { width: 88% !important; }")
            Send("</style>")
        End If

        MyCommon.Close_LogixRT()
        MyCommon = Nothing

    End Sub

    Public Sub Write_StyleCookie(ByVal StyleID As Integer)
        Dim StyleCookie As New HttpCookie("Style")
        StyleCookie.Value = StyleID
        StyleCookie.Expires = DateTime.Now.AddDays(365)
        Response.Cookies.Add(StyleCookie)
    End Sub

    Public Sub Send_Scripts(Optional ByVal ScriptNames As String() = Nothing)
        Dim i As Integer = 0

        Send("<script src=""/javascript/logix.js"" type=""text/javascript""></script>")
        Send("<script type=""text/javascript"" src=""/javascript/jquery.min.js""></script>")

        If (ScriptNames IsNot Nothing) Then
            For i = 0 To ScriptNames.GetUpperBound(0)
                Send("<script src=""/javascript/" & ScriptNames(i) & """ type=""text/javascript""></script>")
            Next
        End If

        Send_JavaScript_Terms()
    End Sub

    Public Sub Send_JavaScript_Terms()

        Send("<script type=""text/javascript"">")
        Send("  // replaces existing text found in javascript files with translations of that text into the user's language.")
        Send("  termSelectPrinter = '" & Copient.PhraseLib.Lookup("logix-js.SelectPrinter", LanguageID) & "';")
        Send("  termBrowser = '" & Copient.PhraseLib.Lookup("term.browser", LanguageID) & "';")
        Send("  termPlatform = '" & Copient.PhraseLib.Lookup("term.platform", LanguageID) & "';")
        Send("  termEndLogix = '" & Copient.PhraseLib.Lookup("logix-js.EndSession", LanguageID) & "';")
        Send("  termDateFormat = '" & Copient.PhraseLib.Lookup("logix-js.EnterDate", LanguageID) & "';")
        Send("  termValidMonth = '" & Copient.PhraseLib.Lookup("logix-js.EnterMonth", LanguageID) & "';")
        Send("  termValidDay = '" & Copient.PhraseLib.Lookup("logix-js.EnterDay", LanguageID) & "';")
        Send("  termValidYear = '" & Copient.PhraseLib.Lookup("logix-js.EnterYear", LanguageID) & "';")
        Send("  termValidDate = '" & Copient.PhraseLib.Lookup("logix-js.EnterValidDate", LanguageID) & "';")
        Send("  termEnterName = '" & Copient.PhraseLib.Lookup("logix-js.EnterOfferName", LanguageID) & "';")
        Send("  termValidStartHour = '" & Copient.PhraseLib.Lookup("logix-js.EnterStartHour", LanguageID) & "';")
        Send("  termValidEndHour = '" & Copient.PhraseLib.Lookup("logix-js.EnterEndHour", LanguageID) & "';")
        Send("  termStartMinute = '" & Copient.PhraseLib.Lookup("logix-js.EnterStartMinute", LanguageID) & "';")
        Send("  termEndMinute = '" & Copient.PhraseLib.Lookup("logix-js.EnterEndMinute", LanguageID) & "';")
        Send("  termPromptForSave = '" & Copient.PhraseLib.Lookup("sv-edit.ChangesMade", LanguageID) & "';")
        Send("  termSave = '" & Copient.PhraseLib.Lookup("term.save", LanguageID) & "';")
        Send("  termMarkupTagWarning = '" & Copient.PhraseLib.Lookup("logix-js.NoMarkupTags", LanguageID) & "';")
        Send("  termValueSelectOperation   = '" & Copient.PhraseLib.Lookup("prefentry.ValueSelectOperation", LanguageID) & "';")
        Send("  termValueNoUnused          = '" & Copient.PhraseLib.Lookup("prefentry.ValueNoUnused", LanguageID) & "';")
        Send("  termValueSelect            = '" & Copient.PhraseLib.Lookup("prefentry.ValueSelect", LanguageID) & "';")
        Send("  termValueEnter             = '" & Copient.PhraseLib.Lookup("prefentry.ValueEnter", LanguageID) & "';")
        Send("  termValueAlreadySelected   = '" & Copient.PhraseLib.Lookup("prefentry.ValueAlreadySelected", LanguageID) & "';")
        Send("  termValueOutsideRange      = '" & Copient.PhraseLib.Lookup("prefentry.ValueOutsideRange", LanguageID) & "';")
        Send("  termCurrentDate            = '" & Copient.PhraseLib.Lookup("term.currentdate", LanguageID) & "';")
        Send("  termEnterDateFormat        = '" & Copient.PhraseLib.Lookup("prefentry.EnterDateFormat", LanguageID) & "';")
        Send("  termEnterAnniversaryFormat = '" & Copient.PhraseLib.Lookup("prefentry.EnterAnniversaryFormat", LanguageID) & "';")
        Send("  termEnterValidMonth        = '" & Copient.PhraseLib.Lookup("prefentry.EnterValidMonth", LanguageID) & "';")
        Send("  termEnterValidDay          = '" & Copient.PhraseLib.Lookup("prefentry.EnterValidDay", LanguageID) & "';")
        Send("  termEnterValidYear         = '" & Copient.PhraseLib.Lookup("prefentry.EnterValidYear", LanguageID) & "';")
        Send("  termEnterValidValue        = '" & Copient.PhraseLib.Lookup("prefentry.EnterValidValue", LanguageID) & "';")
        Send("  termInvalidDays            = '" & Copient.PhraseLib.Lookup("prefentry.InvalidDays", LanguageID) & "';")
        Send("</sc" & "ript>")

    End Sub

    Public Sub Send_HeadEnd()
        Send("</head>")
    End Sub

    Public Sub Send_BodyBegin(ByVal BodyType As Integer)
        Send_PageBegin(BodyType)
        Send_WrapBegin()
    End Sub

    Public Sub Send_PageBegin(ByVal BodyType As Integer)
        Sendb("<body")
        If (BodyType = 1) Then
            Send(">")
        ElseIf (BodyType = 2) Then
            Send(" class=""popup"" onunload=""ChangeParentDocument()"">")
        ElseIf (BodyType = 3) Then
            Send(" class=""popup"">")
        ElseIf (BodyType = 4) Then
            Send(" onunload=""updateCookie()"">")
        ElseIf (BodyType = 5) Then
            Send(" class=""minipopup"" onunload=""ChangeParentDocument()"">")
        ElseIf (BodyType = 6) Then
            Send(" class=""minipopup"">")
        ElseIf (BodyType = 11) Then
            Send(" class=""template"">")
        ElseIf (BodyType = 12) Then
            Send(" class=""popup template"" onunload=""ChangeParentDocument()"">")
        ElseIf (BodyType = 13) Then
            Send(" class=""popup template"">")
        ElseIf (BodyType = 14) Then
            Send(" class=""template"" onunload=""updateCookie()"">")
        End If
    End Sub

    Public Sub Send_WrapBegin()
        Send("<div id=""custom1""></div>")
        Send("<div id=""wrap"">")
        Send("<div id=""custom2""></div>")
        Send("<a id=""top"" name=""top""></a>")
        Send("")
    End Sub

    Public Sub Send_Bar(Optional ByVal Handheld As Boolean = False)
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim AdminUserID As Long

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        MyCommon.SetAdminUser(Logix.Fetch_AdminUser(MyCommon, AdminUserID))

        Send("<div id=""bar"">")
        Send("  <span id=""skip""><a href=""#skiptabs"">▼</a></span>")
        Send("  <span id=""time"" title=""" & DateTime.Now.ToString("HH:mm:ss, G\MT zzz") & """>" & DateTime.Now.ToString("HH:mm") & " | </span>")
        If (Handheld = True) Then
            Send("  <span id=""date"">" & Logix.ToShortDateString(DateTime.Now, MyCommon) & " | </span>")
        Else
            Send("  <span id=""date"">" & Logix.ToLongDateString(DateTime.Now, MyCommon) & " | </span>")
        End If
        Send("  <span id=""user""><a href=""/logix/user-edit.aspx?UserID=" & AdminUserID & """>" & AdminName & "</a><span class=""noprint""> | </span></span>")
        Send("  <a href=""/logix/login.aspx?mode=logout"" id=""logout""><b>" & Copient.PhraseLib.Lookup("term.logout", LanguageID) & "</b></a>")
        Send("</div>")
        Send("")

        MyCommon.Close_LogixRT()
        MyCommon = Nothing
    End Sub

    Public Sub Send_Help(Optional ByVal FileName As String = "")
        'Send("<div id=""gethelp"">")
        'Send("  <a href=""javascript:openPopup('help.aspx?Popup=1" & IIf(FileName <> "", "&amp;FileName=" & FileName, "") & "')"">")
        'Send("    <img src=""/images/clear.png"" id=""helpbutton"" name=""helpbutton"" alt=""" & Copient.PhraseLib.Lookup("term.help", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.help", LanguageID) & """ />")
        'Send("  </a>")
        'Send("</div>")
        'Send("")
    End Sub

    Public Sub Send_Logos(Optional ByVal DefaultLanguageID As Integer = 0)
        Dim TempLanguageID As Integer
        If DefaultLanguageID > 0 Then
            TempLanguageID = DefaultLanguageID
        Else
            TempLanguageID = LanguageID
        End If
        Send("<div id=""logos"">")
        Send("  <div id=""logix"" title=""" & Copient.PhraseLib.Lookup("term.logix", TempLanguageID) & """></div>")
        Send("  <div id=""licensee"" title=""" & Copient.PhraseLib.Lookup("term.licensee", TempLanguageID) & """></div>")
        Send("  <br clear=""all"" />")
        Send("</div>")
        Send("")
    End Sub

    Public Sub Send_Tabs(ByRef MyLogix As Object, ByVal Tabset As Integer)
        Dim MyCommon As New Copient.CommonInc
        Dim dst As System.Data.DataTable
        Dim TabOns() As String = {"", "", "", "", "", "", "", "", "", ""}
        Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
        Dim EPMInstalled As Boolean
        Dim TabStyleOverride As String = ""
        Dim EPMPage As String = ""
        Dim EPMHostURI As String = ""
        Dim StyleCookie As Integer = 0

        TabOns(Tabset) = "class=""on"" "

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()

        EPMInstalled = MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)
        Send("<div id=""tabs"">")
        If EPMInstalled Then
            Send("<style type=""text/css"">")
            If Not Request.Cookies("Style") Is Nothing Then
                Integer.TryParse(Request.Cookies("Style").Value, StyleCookie)
            End If
            Select Case StyleCookie
                Case 1  'override the Logix Blue style
                    Send("#tabs a, #tabs a.on {background: url('../../images/tab_narrow1.png') no-repeat scroll 0 0 transparent; left: 7px; width: 82px;}")
                    Send("#tabs a:hover {background: url('../../images/tab-hover_narrow1.png') no-repeat;}")
                    Send("#tabs a.on {background: url('../../images/tab-on_narrow1.png') no-repeat;}")
                    Send("#tabs a.on:hover {background: url('../../images/tab-on_narrow1.png') no-repeat;}")
                Case 2  'override the NCR Blue style
                    Send("#tabs a, #tabs a.on {left: 7px; width: 82px;}")
                    Send("#tabs a:hover {background: url('../../images/ncr/tab-hover_narrow1.png') no-repeat;}")
                    Send("#tabs a.on {background: url('../../images/ncr/tab-on_narrow1.png') no-repeat; font-weight: bold; height: 25px;}")
                    Send("#tabs a.on:hover {background: url('../../images/ncr/tab-on_narrow1.png') no-repeat;}")
                Case 3  'override the GOLD style
                    Send(" ") 'nothing to change here
                Case Else 'override the NCR Green style (styleID=4)
                    Send("#tabs a, #tabs a.on {left: 7px; width: 82px;}")
                    Send("#tabs a:hover {background: url('../../images/ncrgreen/tab-hover_narrow1.png') no-repeat;}")
                    Send("#tabs a.on {background: url('../../images/ncrgreen/tab-on_narrow1.png') no-repeat; font-weight: bold; height: 25px;}")
                    Send("#tabs a.on:hover {background: url('../../images/ncrgreen/tab-on_narrow1.png') no-repeat;}")
            End Select
            Send("</style>")
        End If

        Send("  <a href=""/logix/status.aspx"" accesskey=""1"" " & TabOns(1) & TabStyleOverride & "id=""tab1"" title=""" & Copient.PhraseLib.Lookup("term.systemoverview", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.logix", LanguageID) & "</a>")
        Send("  <a href=""/logix/offer-list.aspx"" accesskey=""2"" " & TabOns(2) & TabStyleOverride & "id=""tab2"" title=""" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & "</a>")
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""3"" " & TabOns(3) & TabStyleOverride & "id=""tab3"" title=""" & Copient.PhraseLib.Lookup("term.customers", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.customers", LanguageID) & "</a>")
        End If
        Send("  <a href=""/logix/pgroup-list.aspx"" accesskey=""4"" " & TabOns(4) & TabStyleOverride & "id=""tab4"" title=""" & Copient.PhraseLib.Lookup("term.products", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.products", LanguageID) & "</a>")
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then
            Send("  <a href=""/logix/point-list.aspx"" accesskey=""5"" " & TabOns(5) & TabStyleOverride & "id=""tab5"" title=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & "</a>")
        End If
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then
            Send("  <a href=""/logix/graphic-list.aspx"" accesskey=""6"" " & TabOns(6) & TabStyleOverride & "id=""tab6"" title=""" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & "</a>")
        End If
        Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""7"" " & TabOns(7) & "id=""tab7"" title=""" & Copient.PhraseLib.Lookup("term.stores", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.locations", LanguageID) & "</a>")

        If (MyLogix.UserRoles.AccessStoreHealth = True) Then
            MyCommon.QueryStr = "select EngineID from PromoEngines where Installed=1 and DefaultEngine=1 and EngineID in (0,2,9);"
            dst = MyCommon.LRT_Select
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(91) = "1" AndAlso MyCommon.NZ(dst.Rows(0).Item("EngineID"), -1) = 9) Then
                Send("  <a href=""/logix/UE/UEServerHealthSummary.aspx"" accesskey=""8"" " & TabOns(8) & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
            Else
                MyCommon.QueryStr = "select EngineID from PromoEngines where Installed=1 and DefaultEngine=1 and EngineID in (0,2,9);"
                dst = MyCommon.LRT_Select
                If (dst.Rows.Count > 0) Then
                    Dim Tmp_EngineID As Integer = MyCommon.NZ(dst.Rows(0).Item("EngineID"), -1)
                    Select Case Tmp_EngineID
                        Case 0 ' link to CM store health
                            Send("  <a href=""/logix/store-health-cm.aspx?filterhealth=2"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                        Case 2 ' link to CPE store health
                            Send("  <a href=""/logix/store-health-cpe.aspx?filterhealth=2"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                        Case 9 ' link to UE store health
                            Send("  <a href=""/logix/UE/store-health-ue.aspx?filterhealth=2"" accesskey=""8"" " & TabOns(8) & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                        Case Else
                            Send("  <a href=""/logix/user-list.aspx"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                    End Select
                Else
                    Send("  <a href=""/logix/user-list.aspx"" accesskey=""8"" " & TabOns(8) & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                End If
            End If
        Else
            Send("  <a href=""/logix/user-list.aspx"" accesskey=""8"" " & TabOns(8) & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
        End If

        If (EPMInstalled) Then
            Try
                MyCommon.QueryStr = "select isnull(PageName, '') as PageName " &
                                "from PM_AdminUserStartPages as AUSP with (NoLock) Inner Join AdminUsers as AU with (NoLock) on AU.EPMStartPageID=AUSP.StartPageID " &
                                "where AU.AdminUserID=" & AdminUserID & ";"
                dst = MyCommon.LRT_Select
                If dst.Rows.Count > 0 Then
                    EPMPage = dst.Rows(0).Item("PageName")
                End If
            Catch ex As Exception
                '
            Finally
                If EPMPage = "" Then EPMPage = IntegrationVals.StartupPath
                EPMHostURI = IntegrationVals.HTTP_RootURI
                If Not (Right(EPMHostURI, 1) = "/") Then
                    EPMHostURI = EPMHostURI & "/"
                End If
                If Left(EPMPage, 1) = "/" Then
                    EPMPage = Mid(EPMPage, 2) 'get rid of the first character of the path (page) variable
                End If
                EPMHostURI = EPMHostURI & "UI/" & EPMPage
                Send("  <a href=""/logix/authtransfer.aspx?SendToURI=" & EPMHostURI & """ accesskey=""$"" " & TabOns(9) & TabStyleOverride & "id=""tab9"" title=""" & Copient.PhraseLib.Lookup(IntegrationVals.PhraseTerm, LanguageID) & """>" & Copient.PhraseLib.Lookup(IntegrationVals.PhraseTerm, LanguageID) & "</a>")
            End Try
        End If

        Send("  <br clear=""all"" />")
        Send("</div>")
        Send("")

        MyCommon.Close_LogixRT()
        MyCommon = Nothing
    End Sub

    Public Sub Send_Subtabs(ByRef MyLogix As Object, ByVal Subtabset As Integer, ByVal SubtabHighlight As Integer, Optional ByVal DefaultLanguageID As Integer = 0, Optional ByVal ID As String = "", Optional ByVal subtaburl As String = "", Optional ByVal SecondaryID As String = "")
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim dst As System.Data.DataTable
        Dim dt As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim TempLanguageID As Integer
        Dim CustPK As String = ""
        Dim CardPK As String = ""
        Dim DisableSubTabs As Boolean = False
        Dim IsNew As Boolean = False
        Dim OfferID As Integer
        Dim AdminUserID As Long = 0
        Dim CPEInstalled As Boolean = False
        Dim CMInstalled As Boolean = False
        Dim TempDateTime As New DateTime
        Dim BannersEnabled As Boolean = False

        MyCommon.AppName = "ue-cb"
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        MyCommon.SetAdminUser(Logix.Fetch_AdminUser(MyCommon, AdminUserID))
        MyCommon.Close_LogixRT()

        If DefaultLanguageID > 0 Then
            TempLanguageID = DefaultLanguageID
        Else
            TempLanguageID = LanguageID
        End If

        ' determine which engines are installed
        MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select EngineID from PromoEngines with (NoLock) where Installed=1;"
        dst = MyCommon.LRT_Select()
        For Each row In dst.Rows
            If row.Item("EngineID") = 0 Then CMInstalled = True
            If row.Item("EngineID") = 2 Then CPEInstalled = True
        Next

        'Are banners enabled
        MyCommon.QueryStr = "Select OptionValue from SystemOptions with (NoLock) where OptionID=66;"
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count = 1 Then
            If dt.Rows(0).Item("OptionValue") = 0 Then BannersEnabled = False
            If dt.Rows(0).Item("OptionValue") = 1 Then BannersEnabled = True
        End If

        MyCommon.Close_LogixRT()

        Dim SubtabOns() As String = {"", "", "", "", "", "", "", "", "", "", "", "", ""}
        SubtabOns(SubtabHighlight) = "class=""on"" "
        Send("<div id=""subtabs""" & IIf(TempLanguageID > 1, " class=""subtabs st" & Subtabset & """", "") & ">")

        If (Subtabset = 0) Then
            Send("  <a href=""/logix/login.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.login", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/requirements.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.requirements", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 1) Then
            Send("  <a href=""/logix/status.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.status", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/about.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.about", TempLanguageID) & "</a>")
            'Send("  <a href=""/logix/help.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.manual", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 20) Then
            ' Offers section
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.AccessFolders) Then
                Send("  <a href=""/logix/folders.aspx"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.folders", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/extoffer-list.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.external", TempLanguageID) & "</a>")
            If MyCommon.IsEngineInstalled(6) Then
                Send("  <a href=""/logix/CAM/CAM-offer-list.aspx"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.cam", TempLanguageID) & "</a>")
            End If
            If BannersEnabled Then
                Send("  <a href=""/logix/banneroffer-list.aspx"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"">" & Copient.PhraseLib.Lookup("term.banners", TempLanguageID) & "</a>")
            End If
            If MyCommon.IsEngineInstalled(9) Then
                ' Collision Section
                Send(" <a href=""/logix/CollidingOffers-list.aspx"" accesskey=""^"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.Collision", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 21) Then
            ' Offers section -- regular (CM) offers
            OfferID = ID
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "Select StatusFlag, DeployDeferred from Offers with (NoLock) where OfferID=" & OfferID
            dst = MyCommon.LRT_Select()
            For Each row In dst.Rows
                If (MyCommon.NZ(row.Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(row.Item("StatusFlag"), -1) > 10) Then
                    DisableSubTabs = True
                End If
                If (Not DisableSubTabs) Then
                    DisableSubTabs = (MyCommon.NZ(row.Item("DeployDeferred"), False) = True)
                End If
            Next
            MyCommon.Close_LogixRT()
            If (DisableSubTabs) Then
                Send("  <a href=""#"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</span></a>")
                Send("  <a href=""/logix/offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/offer-loc.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 22) Then
            ' Offers section -- regular offers templates
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/offer-loc.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/offer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/offer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/offer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 23) Then
            ' Offers section -- deleted offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 24) Then
            ' Offers section -- UE offers
            OfferID = ID
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""("" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "Select StatusFlag, DeployDeferred, EndDate, UpdateLevel, ExpireLocked from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
            dst = MyCommon.LRT_Select()
            For Each row In dst.Rows
                If (MyCommon.NZ(row.Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 11 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 12) Then
                    DisableSubTabs = True
                End If
                If (Not DisableSubTabs) Then
                    DisableSubTabs = (MyCommon.NZ(row.Item("DeployDeferred"), -1) = True)
                End If
                If (MyCommon.Fetch_UE_SystemOption(80) = "1") Then
                    TempDateTime = Logix.ToShortDateString(MyCommon.NZ(row.Item("EndDate"), New Date(1900, 1, 1)), MyCommon)
                    If TempDateTime < Today() AndAlso MyCommon.Extract_Val(MyCommon.NZ(row.Item("UpdateLevel"), -1)) > 0 Then
                        DisableSubTabs = True
                    End If
                End If
                If (MyCommon.NZ(row.Item("ExpireLocked"), 0) = 1) Then
                    DisableSubTabs = True
                End If
            Next
            MyCommon.Close_LogixRT()
            If (DisableSubTabs OrElse OfferLockedforCollisionDetection) Then
                Send("  <a href=""#"" accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</span></a>")
                'BZ2079: UE-feature-removal
                'Send("  <a href=""#"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.notifications", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</span></a>")
            Else
                Send("  <a href=""/logix/UE/UEoffer-loc.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/UE/UEoffer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/UE/UEoffer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/UE/UEoffer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/UE/UEoffer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 25) Then
            ' Offers section -- UE offer templates
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""("" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/UE/UEoffer-loc.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/UE/UEoffer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/UE/UEoffer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/UE/UEoffer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/UE/UEoffer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 26) Then
            ' Offers section -- CPE deleted offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/UE/UEoffer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 27) Then
            ' Offers section -- Web offers
            OfferID = ID
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "Select StatusFlag, DeployDeferred, EndDate, UpdateLevel, ExpireLocked from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
            dst = MyCommon.LRT_Select()
            For Each row In dst.Rows
                If (MyCommon.NZ(row.Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(row.Item("StatusFlag"), -1) > 10) Then
                    DisableSubTabs = True
                End If
                If (Not DisableSubTabs) Then
                    DisableSubTabs = (MyCommon.NZ(row.Item("DeployDeferred"), -1) = True)
                End If
                If (MyCommon.Fetch_UE_SystemOption(80) = "1") Then
                    TempDateTime = Logix.ToShortDateString(MyCommon.NZ(row.Item("EndDate"), New Date(1900, 1, 1)), MyCommon)
                    If TempDateTime < Today() AndAlso MyCommon.Extract_Val(MyCommon.NZ(row.Item("UpdateLevel"), -1)) > 0 Then
                        DisableSubTabs = True
                    End If
                End If
                If (MyCommon.NZ(row.Item("ExpireLocked"), 0) = 1) Then
                    DisableSubTabs = True
                End If
            Next
            MyCommon.Close_LogixRT()
            If (DisableSubTabs) Then
                Send("  <a href=""#"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.notifications", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</span></a>")
            Else
                Send("  <a href=""/logix/web-offer-not.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.notifications", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/web-offer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/web-offer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/web-offer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/web-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 29) Then
            ' Offers section -- Web deleted offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/web-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 203) Then
            ' Offers section -- Email offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/email-offer-not.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.notifications", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/email-offer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/email-offer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/email-offer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/email-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 204) Then
            ' Offers section -- Email deleted offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/email-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 205) Then
            ' Offers section -- CAM offers
            OfferID = ID
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""("" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "Select StatusFlag, DeployDeferred, EndDate, UpdateLevel, ExpireLocked from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
            dst = MyCommon.LRT_Select()
            For Each row In dst.Rows
                If (MyCommon.NZ(row.Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(row.Item("StatusFlag"), -1) > 10) Then
                    DisableSubTabs = True
                End If
                If (Not DisableSubTabs) Then
                    DisableSubTabs = (MyCommon.NZ(row.Item("DeployDeferred"), -1) = True)
                End If
                If (MyCommon.Fetch_UE_SystemOption(80) = "1") Then
                    TempDateTime = Logix.ToShortDateString(MyCommon.NZ(row.Item("EndDate"), New Date(1900, 1, 1)), MyCommon)
                    If TempDateTime < Today() AndAlso MyCommon.Extract_Val(MyCommon.NZ(row.Item("UpdateLevel"), -1)) > 0 Then
                        DisableSubTabs = True
                    End If
                End If
                If (MyCommon.NZ(row.Item("ExpireLocked"), 0) = 1) Then
                    DisableSubTabs = True
                End If
            Next
            MyCommon.Close_LogixRT()
            If (DisableSubTabs) Then
                Send("  <a href=""#"" accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</span></a>")
                Send("  <a href=""/logix/CAM/CAM-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/offer-loc.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CAM/CAM-offer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CAM/CAM-offer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CAM/CAM-offer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CAM/CAM-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 206) Then
            ' Offers section -- CAM offer templates
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""("" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/offer-loc.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CAM/CAM-offer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CAM/CAM-offer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CAM/CAM-offer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CAM/CAM-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 207) Then
            ' Offers section -- CAM deleted offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/CAM/CAM-offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 30) Then
            ' Customers section
            ID = "CustPK=" & ID
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx?" & ID & """ accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyLogix.UserRoles.ViewHistory = True) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1") Then
                Send("  <a href=""/logix/CM-cashier-report.aspx" & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.CashierHistory", LanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 31) Then
            ' Customers section -- groups
            If ID = 0 Then IsNew = True
            ID = "CustomerGroupID=" & ID
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/cgroup-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/cgroup-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 32) Then
            ' Customer section -- inquiry
            CustPK = "CustPK=" & ID
            If SecondaryID <> "" Then
                CardPK = "&amp;CardPK=" & MyCommon.Extract_Val(SecondaryID)
            End If
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyLogix.UserRoles.ViewHistory = True) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1" AndAlso ID = 0) Then
                Send("  <a href=""/logix/CM-cashier-report.aspx" & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & "Cashier History" & "</a>")
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso ID > 0) Then
                Send("  <a href=""/logix/customer-hist.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            If (ID > 0) Then
                'Send("  <a href=""/logix/customer-manual.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.manual", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/customer-transactions.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.transactions", TempLanguageID) & "</a>")
                If (MyLogix.UserRoles.AccessAdjustmentsPage = True) Then
                    Send("  <a href=""/logix/customer-adjustments.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.adjustments", TempLanguageID) & "</a>")
                End If
                Send("  <a href=""/logix/customer-offers.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/customer-general.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 33) Then
            ' CAM Customer section -- inquiry
            CustPK = "CustPK=" & ID
            If SecondaryID <> "" Then
                CardPK = "&amp;CardPK=" & MyCommon.Extract_Val(SecondaryID)
            End If
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso ID > 0) Then
                Send("  <a href=""/logix/customer-hist.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & "&amp;CAM=1"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            If (ID > 0) Then
                Send("  <a href=""/logix/CAM/CAM-customer-transactions.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.transactions", TempLanguageID) & "</a>")
                If (MyLogix.UserRoles.AccessAdjustmentsPage = True) Then
                    Send("  <a href=""/logix/CAM/CAM-customer-adjustments.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.adjustments", TempLanguageID) & "</a>")
                End If
                Send("  <a href=""/logix/CAM/CAM-customer-offers.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CAM/CAM-customer-general.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 40) Then
            ' Products section
            Send("  <a href=""/logix/pgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            MyCommon.Open_LogixRT()
            Send("  <a href=""/logix/product-inquiry.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.productinquiry", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/pgroup-inquiry.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.groupinquiry", TempLanguageID) & "</a>")
            MyCommon.Close_LogixRT()
        ElseIf (Subtabset = 41) Then
            ' Products section -- groups
            If ID = 0 Then IsNew = True
            ID = "ProductGroupID=" & ID
            Send("  <a href=""/logix/pgroup-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.productinquiry", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/pgroup-inquiry.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.groupinquiry", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/pgroup-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/pgroup-edit.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 42) Then
            ' Products section -- product inquiry
            ID = "ProductGroupID=" & ID
            Send("  <a href=""/logix/pgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.productinquiry", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/pgroup-inquiry.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.groupinquiry", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 43) Then
            ' Products section -- group inquiry
            ID = "ProductGroupID=" & ID
            Send("  <a href=""/logix/pgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.productinquiry", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/pgroup-inquiry.aspx"" accesskey=""#"" class=""on"" id=""subtab3"">" & Copient.PhraseLib.Lookup("term.groupinquiry", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 50) Then
            ' Programs section
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 51) Then
            ' Programs section -- points
            If ID = 0 Then IsNew = True
            ID = "ProgramGroupID=" & ID
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/point-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/point-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(5) & "id=""subtab5""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 52) Then
            ' Programs section -- stored value
            If ID = 0 Then IsNew = True
            ID = "ProgramGroupID=" & ID
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/sv-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/sv-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(5) & "id=""subtab5""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 53) Then
            ' Programs section -- promovars (In DP-only environments, this is under the PromoVars tab)
            If ID = 0 Then IsNew = True
            ID = "PromoVarID=" & ID
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/promovar-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/promovar-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(5) & "id=""subtab5""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 54) Then
            ' Programs section -- Advanced Limits
            If ID = 0 Then IsNew = True
            ID = "LimitID=" & ID
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled Or CMInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" class=""on"" id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/CM-advlimit-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/CM-advlimit-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(5) & "id=""subtab5""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 60) Then
            ' Graphics section
            Send("  <a href=""/logix/graphic-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.graphics", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/layout-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.layouts", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 61) Then
            ' Graphic section -- graphics
            If ID = 0 Then IsNew = True
            ID = "OnScreenAdId=" & ID
            Send("  <a href=""/logix/graphic-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.graphics", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/layout-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.layouts", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/graphic-hist.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/graphic-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 62) Then
            ' Graphics section -- layouts
            If ID = "0" Then IsNew = True
            ID = "LayoutID=" & ID
            Send("  <a href=""/logix/graphic-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.graphics", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/layout-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.layouts", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/layout-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/layout-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 70) Then
            ' Locations section
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 71) Then
            ' Locations section -- groups
            If ID = 0 Then IsNew = True
            ID = "LocationGroupID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/lgroup-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/lgroup-edit.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 72) Then
            ' Locations section -- stores
            If ID = 0 Then IsNew = True
            ID = "LocationID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/store-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/store-edit.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 73) Then
            ' Locations section -- servers (i.e., stores of LocationTypeID 2)
            If ID = 0 Then IsNew = True
            ID = "LocationID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" class=""on"" id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/store-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/store-edit.aspx?" & ID & IIf(IsNew, "&LocationTypeID=2", "") & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 8) Then
            ' Administration section
            Dim PageUrl As String = "/logix/UE/store-health-UE.aspx"
            MyCommon.Open_LogixRT()
            If (MyLogix.UserRoles.AccessSystemHealth) Then
                Send("  <a href=""/logix/agent-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.agents", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.AccessBanners AndAlso MyCommon.Fetch_SystemOption(66) = "1") Then
                Send("  <a href=""/logix/banner-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.banners", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.AccessConnectors) Then
                Send("  <a href=""/logix/connector-list.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.connectors", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.EditSystemConfiguration) Then
                Send("  <a href=""/logix/configuration.aspx"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.configuration", TempLanguageID) & "</a>")
            End If
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) AndAlso MyLogix.UserRoles.AccessLMGRejections) Then
                Send("  <a href=""/logix/CAM/LMG-rejections.aspx"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"">" & Copient.PhraseLib.Lookup("term.rejections", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.ViewOfferHealth) Then
                If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                    Send("  <a href=""/logix/offer-health.aspx"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.offerhealth", TempLanguageID) & "</a>")
                End If
            End If
            If (MyLogix.UserRoles.AccessStoreHealth) Then
                If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(91) = "1") Then
                    PageUrl = "/logix/UE/UEServerHealthSummary.aspx"
                Else
                    If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                        MyCommon.QueryStr = "select EngineID from PromoEngines where Installed=1 and DefaultEngine=1 and EngineID in (0,2,9);"
                        dst = MyCommon.LRT_Select
                        If (dst.Rows.Count > 0) Then
                            Dim Tmp_EngineID As Integer = MyCommon.NZ(dst.Rows(0).Item("EngineID"), -1)
                            Select Case Tmp_EngineID
                                Case 0 ' link to CM store health
                                    PageUrl = "/logix/store-health-cm.aspx?filterhealth=2"
                                Case 2 ' link to CPE store health
                                    PageUrl = "/logix/store-health-cpe.aspx?filterhealth=2"
                                Case 9 ' link to UE store health
                                    PageUrl = "/logix/UE/store-health-UE.aspx?filterhealth=2"
                            End Select
                        End If
                    End If
                End If

                If (PageUrl.Contains("UEServerHealth")) Then
                    Send("  <a href=""" & PageUrl & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.serverhealth", TempLanguageID) & "</a>")
                Else
                    Send("  <a href=""" & PageUrl & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.storehealth", TempLanguageID) & "</a>")
                End If
            End If
            If (MyLogix.UserRoles.AccessReports) Then
                If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                    Send("  <a href=""/logix/reports-list.aspx"" accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"">" & Copient.PhraseLib.Lookup("term.reports", TempLanguageID) & "</a>")
                End If
            End If
            Send("  <a href=""/logix/user-list.aspx"" accesskey=""("" " & SubtabOns(9) & "id=""subtab9"">" & Copient.PhraseLib.Lookup("term.users", TempLanguageID) & "</a>")
            MyCommon.Close_LogixRT()

        ElseIf (Subtabset = 9) Then
            ' Special tabset for the help desk, with only customer inquiry and product inquiry access
            AdminUserID = Verify_AdminUser(MyCommon, Logix)
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.customers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.products", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/user-edit.aspx?UserID=" & AdminUserID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.settings", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 91) Then
            AdminUserID = Verify_AdminUser(MyCommon, Logix)
            CustPK = "CustPK=" & ID
            If SecondaryID <> "" Then
                CardPK = "&amp;CardPK=" & MyCommon.Extract_Val(SecondaryID)
            End If
            Send("  <a href=""/logix/customer-inquiry.aspx?token=nothing" & subtaburl & """ accesskey=""!"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.customers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx?check=nothing" & subtaburl & """ accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.products", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/user-edit.aspx?UserID=" & AdminUserID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.settings", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso ID > 0) Then
                Send("  <a href=""/logix/customer-hist.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            If (ID > 0) Then
                Send("  <a href=""/logix/customer-transactions.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.transactions", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/customer-adjustments.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.adjustments", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/customer-offers.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/customer-general.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 92) Then
            AdminUserID = Verify_AdminUser(MyCommon, Logix)
            ID = "ProductGroupID=" & ID
            Send("  <a href=""/logix/customer-inquiry.aspx?check=nothing" & subtaburl & """ accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.customers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx?" & ID & subtaburl & """ accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.products", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/user-edit.aspx?UserID=" & AdminUserID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.settings", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 93) Then
            AdminUserID = Verify_AdminUser(MyCommon, Logix)
            ID = "ProductGroupID=" & ID
            Send("  <a href=""/logix/customer-inquiry.aspx?check=nothing" & subtaburl & """ accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.customers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx?" & ID & subtaburl & """ accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.products", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/user-edit.aspx?UserID=" & AdminUserID & """ accesskey=""#"" class=""on"" id=""subtab3"">" & Copient.PhraseLib.Lookup("term.settings", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 94) Then
            AdminUserID = Verify_AdminUser(MyCommon, Logix)
            CustPK = "CustPK=" & ID
            Send("  <a href=""/logix/customer-inquiry.aspx?token=nothing" & subtaburl & """ accesskey=""!"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.customers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/product-inquiry.aspx?check=nothing" & subtaburl & """ accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.products", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/user-edit.aspx?UserID=" & AdminUserID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.settings", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CAM/CAM-customer-manual.aspx?" & CustPK & """ accesskey=""("" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.manual", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True AndAlso ID > 0) Then
                Send("  <a href=""/logix/customer-hist.aspx?" & CustPK & "&amp;CAM=1"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            If (ID > 0) Then
                Send("  <a href=""/logix/CAM/CAM-customer-transactions.aspx?" & CustPK & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.transactions", TempLanguageID) & "</a>")
                If (MyLogix.UserRoles.AccessAdjustmentsPage = True) Then
                    Send("  <a href=""/logix/CAM/CAM-customer-adjustments.aspx?" & CustPK & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.adjustments", TempLanguageID) & "</a>")
                End If
                Send("  <a href=""/logix/CAM/CAM-customer-offers.aspx?" & CustPK & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CAM/CAM-customer-general.aspx?" & CustPK & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            End If
        End If

        Send("  <br clear=""all"" />")
        Send("  <hr class=""hidden"" />")
        Send("  <a id=""skiptabs"" name=""skiptabs""></a>")
        Send("</div>")
    End Sub

    Public Sub Send_Save(Optional ByVal Attributes As String = "")
        ' Find out if this is an offer were trying to send
        ' onclick=""if(confirm('" & Copient.PhraseLib.Lookup("offer.verifysave", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & "!""
        Send("<input type=""submit"" accesskey=""s"" class=""regular"" id=""save"" name=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Delete(Optional ByVal Attributes As String = "")
        Dim OnClickAttrib As String = ""

        If (Attributes <> "") Then
            OnClickAttrib = ParseAttribute(Attributes, "onclick", False, True)
        End If

        Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""" & OnClickAttrib & "if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')){}else{return false}"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Edit(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" accesskey=""e"" class=""regular"" id=""edit"" name=""edit"" title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_New(Optional ByVal Attributes As String = "")
        Sendb("<input type=""submit"" accesskey=""n"" class=""regular"" id=""new"" name=""new"" title=""" & Copient.PhraseLib.Lookup("term.new", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.new", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Import(Optional ByVal Attributes As String = "")
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim AdminUserID As Long
        Dim dt As System.Data.DataTable = Nothing
        Dim row As System.Data.DataRow = Nothing
        Dim BannersEnabled As Boolean = False
        Dim AllowMultipleBanners As Boolean = False
        Dim AllBannersCheckBox As String = ""
        Dim AllBannersCount As Integer = 1

        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)

        ' store banner options
        BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
        AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")

        Sendb("  <input type=""button"" class=""regular"" id=""importo"" name=""importo"" value=""" & Copient.PhraseLib.Lookup("term.import", LanguageID) & """ ")
        If Request.Browser.Type = "IE6" Then
            Send("onclick=""javascript:{document.getElementById('importer').style.display='block';document.getElementById('importeriframe').style.display='block';}"" />")
        Else
            Send("onclick=""javascript:{document.getElementById('importer').style.display='block';}"" />")
        End If
        Send("  <div id=""importer"" style=""display:none;"">")
        Send("    <form action=""/logix/offer-list.aspx"" method=""post"" enctype=""multipart/form-data"">")
        Send("      <div id=""importmain"">")
        Sendb("        <div id=""importclose"" title=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """><a href=""#"" ")
        If Request.Browser.Type = "IE6" Then
            Sendb("onclick=""javascript:document.getElementById('importer').style.display='none';javascript:document.getElementById('importeriframe').style.display='none';"">")
        Else
            Sendb("onclick=""javascript:document.getElementById('importer').style.display='none';"">")
        End If
        Send("x</a></div><br />")
        Send("        <div id=""importbody"">")
        Send("          " & Copient.PhraseLib.Lookup("term.importpath", LanguageID) & ":<br />")
        Send("          <input type=""file"" id=""browse"" name=""browse"" value=""" & Copient.PhraseLib.Lookup("term.browse", LanguageID) & """ />")
        Send("          <input type=""submit"" class=""regular"" id=""upload"" name=""upload"" value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """" & Attributes & " /><br />")

        If (BannersEnabled) Then
            MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name, BAN.AllBanners from Banners BAN with (NoLock) " &
                                "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " &
                                "where BAN.Deleted=0 and AdminUserID = " & AdminUserID & " " &
                                "order by BAN.Name;"
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
                Send("<label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banners", LanguageID) & ":</label><br />")
                Send("  <select class=""longest"" name=""banner"" id=""banner"" " & IIf(AllowMultipleBanners, "size=""5"" multiple=""multiple""", "") & ">")
                For Each row In dt.Rows
                    If (AllowMultipleBanners AndAlso MyCommon.NZ(row.Item("AllBanners"), False)) Then
                        ' exclude this all banners from the list box and store the option for later display
                        AllBannersCheckBox &= "<input type=""checkbox"" name=""allbannersid"" id=""allbannersid" & AllBannersCount & """ value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """ onclick=""updateBanners(this.checked," & AllBannersCount & ");"" />"
                        AllBannersCheckBox &= "<label for=""allbannersid" & AllBannersCount & """>" & MyCommon.NZ(row.Item("Name"), "") & "</label><br />"
                        AllBannersCount += 1
                    Else
                        Send("    <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                    End If
                Next
                Send("  </select><br />")
                Send("  <br class=""half"" />")
                If (AllBannersCheckBox <> "") Then
                    Send(AllBannersCheckBox)
                    If AllBannersCount > 2 Then
                        Send("<br /><span style=""color:red;"">" & Copient.PhraseLib.Lookup("offer-list.all-banner-note", LanguageID) & "</span>")
                    End If
                End If
            End If
        End If

        Send("        </div>")
        Send("      </div>")
        Send("    </form>")
        Send("  </div>")
        If Request.Browser.Type = "IE6" Then
            Send("  <iframe src=""javascript:'';"" id=""importeriframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""display:none;""></iframe>")
        End If

        MyCommon.Close_LogixRT()
        MyCommon = Nothing
    End Sub

    Public Sub Send_AddOffer(Optional ByVal CustomerPK As Integer = 0, Optional ByVal CardPK As Integer = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""addoffer"" name=""addoffer"" title=""" & Copient.PhraseLib.Lookup("term.addoffer", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.addoffer", LanguageID) & "..."" onclick=""openPopup('/logix/customer-addoffer.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');""" & Attributes & " />")
    End Sub

    Public Sub Send_AssignFolders(Optional ByVal OfferID As Long = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""assignfolders"" name=""assignfolders"" title=""" & Copient.PhraseLib.Lookup("term.assignfolders", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.assignfolders", LanguageID) & "..."" onclick=""openPopup('/logix/folder-browse.aspx?OfferID=" & OfferID & "');""" & Attributes & " />")
    End Sub

    Public Sub Send_CAMAddOffer(Optional ByVal CustomerPK As Integer = 0, Optional ByVal CardPK As Integer = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""addoffer"" name=""addoffer"" title=""" & Copient.PhraseLib.Lookup("term.addoffer", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.addoffer", LanguageID) & "..."" onclick=""openPopup('/logix/CAM/CAM-customer-addoffer.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');""" & Attributes & " />")
    End Sub

    Public Sub Send_CancelDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""canceldeploy"" name=""canceldeploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.canceldeploy", LanguageID) & "')){}else{return false}"" title=""" & Copient.PhraseLib.Lookup("term.canceldeploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.canceldeploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_CancelCollisionDetection(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""cancelcollisiondetection"" name=""cancelcollisiondetection"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.cancelcollisiondetection", LanguageID) & "')){}else{return false}"" style=""width: auto;"" title=""" & Copient.PhraseLib.Lookup("term.cancelcollisiondetection", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.cancelcollisiondetection", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_CancelGetApproval(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""cancelgetapproval"" name=""cancelgetapproval"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.cancelgetapproval", LanguageID) & "')){}else{return false}"" style=""width: auto;"" title=""" & Copient.PhraseLib.Lookup("term.cancelapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.cancelapproval", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Close(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""close"" name=""close"" title=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_CopyOffer(Optional ByVal IsTemplate As Boolean = False, Optional ByVal Attributes As String = "")
        Sendb("<input type=""submit"" class=""regular"" id=""copyoffer"" name=""copyoffer"" ")
        If IsTemplate Then
            Send("title=""" & Copient.PhraseLib.Lookup("term.copytemplate", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.copytemplate", LanguageID) & """" & Attributes & " />")
        Else
            Send("title=""" & Copient.PhraseLib.Lookup("term.copyoffer", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.copyoffer", LanguageID) & """" & Attributes & " />")
        End If
    End Sub

    Public Sub Send_CopyExpiredOffer(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""copyoffer"" name=""copyoffer""  title=""" & Copient.PhraseLib.Lookup("term.copyoffer", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.copyoffer", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_CopyGroup(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""copygroup"" name=""copygroup"" title=""" & Copient.PhraseLib.Lookup("term.copygroup", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.copygroup", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_CustomerNotes(Optional ByVal CustomerPK As Integer = 0, Optional ByVal CardPK As Integer = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""addnote"" name=""addnote"" title=""" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & "..."" onclick=""openPopup('/logix/customer-notes.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');"" />")
    End Sub

    Public Sub Send_DeferDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""deferdeploy"" name=""deferdeploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.deferdeploy", LanguageID) & "')){}else{return false}"" title=""" & Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_DeferDeployCollision(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""deferdeploycollision"" name=""deferdeploycollision"" onclick=""confirmDeployCollision(true);"" title=""" & Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Deploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""deploy"" name=""deploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.deploy", LanguageID) & "')){}else{return false}"" title=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_DeployCollision(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""deploycollision"" name=""deploycollision"" onclick=""confirmDeployCollision(false);"" title=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_RunCollisionDetection(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""runcollision"" name=""runcollision"" onclick=""runDetectionFromSummaryPage();"" title=""" & Copient.PhraseLib.Lookup("term.runCD", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.runCD", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_DisabledRunCollisionDetection(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" disabled=""disable"" id=""runcollision"" name=""runcollision""  title=""" & Copient.PhraseLib.Lookup("term.runCD", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.runCD", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_enableviewreport(Optional ByVal OfferID As Long = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""enableviewreport"" name=""enableviewreport"" onclick=""javascript: redirtToOfferReportPage(" & OfferID & ");"" title=""" & Copient.PhraseLib.Lookup("term.viewcollisionreport", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.viewcollisionreport", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_disableviewreport(Optional ByVal OfferID As Long = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""disableviewreport"" name=""disableviewreport"" disabled title=""" & Copient.PhraseLib.Lookup("term.viewcollisionreport", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.viewcollisionreport", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Download(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""download"" name=""download"" title=""" & Copient.PhraseLib.Lookup("term.download", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.download", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_ExportToEDW(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""exportedw"" name=""exportedw"" title=""" & Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Export(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""export"" name=""export"" title=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_ExportCME(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""exportCME"" name=""exportCME"" title=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ")""" & Attributes & " />")
    End Sub

    Public Sub Send_GenerateIPL(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""generateIPL"" name=""generateIPL"" title=""" & Copient.PhraseLib.Lookup("term.generateipl", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.generateipl", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_OfferFromTemp(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""OfferFromTemp"" name=""OfferFromTemp"" onclick=""javascript: return assignNoofDuplicateOffers(true);"" title=""" & Copient.PhraseLib.Lookup("term.newfromtemp", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.newfromtemp", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_ReDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""redeploy"" name=""redeploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.redeploy", LanguageID) & "')){}else{return false}"" title=""" & Copient.PhraseLib.Lookup("term.redeploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.redeploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_RevalidateAll(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""revalidateall"" name=""revalidateall"" title=""" & Copient.PhraseLib.Lookup("term.revalidate-all", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.revalidate-all", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_GetRecommendations(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""getRecommendations"" onclick=""if(confirm('Do you want to get recommendations?')){}else{return false}""  name=""getRecommendations"" title=""" & Copient.PhraseLib.Lookup("term.getrecommendations", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.getrecom", LanguageID) & """ />")
    End Sub

    Public Sub Send_GetRecommendationsAndDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""getRecommendationsAndDeploy""  onclick=""if(confirm('Do you want to get recommendations and deploy?')){}else{return false}""  name=""getRecommendationsAndDeploy""  title=""" & Copient.PhraseLib.Lookup("term.getrecommendations", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.getrecom", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """/> ")
    End Sub

    Public Sub Send_RequestApproval(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""reqApproval"" onclick=""if(confirm('Do you want to request approval?')){}else{return false}""  name=""reqApproval"" title=""" & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) & """ />")
    End Sub

    Public Sub Send_RequestApprovalCollision(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""reqApproval"" onclick=""javascript:requestApproval(13);""  name=""reqApproval"" title=""" & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) & """ />")
    End Sub

    Public Sub Send_RequestApprovalWithDeployment(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""reqApprovalWithDeployment"" onclick=""if(confirm('Do you want to request approval and deploy?')){}else{return false}""  name=""reqApprovalWithDeployment"" title=""" & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) & """ />")
    End Sub

    Public Sub Send_RequestApprovalWithDeploymentCollision(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""reqApprovalWithDeployment"" onclick=""javascript:requestApproval(14);""  name=""reqApprovalWithDeployment"" title=""" & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) & """ />")
    End Sub

    Public Sub Send_RequestApprovalWithDeferDeployment(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""reqApprovalWithDeferDeployment"" onclick=""if(confirm('Do you want to request approval and defer deploy?')){}else{return false}""  name=""reqApprovalWithDeferDeployment"" title=""" & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) & """ />")
    End Sub

    Public Sub Send_RequestApprovalWithDeferDeploymentCollision(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""reqApprovalWithDeferDeployment"" onclick=""javascript:requestApproval(15);""  name=""reqApprovalWithDeferDeployment"" title=""" & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.reqapprovedeferdeploy", LanguageID) & """ />")
    End Sub

    Public Sub Send_ApproveOffer(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""approveOffer""  onclick=""javascript:approveOfferBackground();""  name=""approveOffer"" title=""" & Copient.PhraseLib.Lookup("term.approveoffer", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.approveoffer", LanguageID) & """ />")
    End Sub

    Public Sub Send_RejectOffer(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""rejectOffer"" name=""rejectOffer"" title=""" & Copient.PhraseLib.Lookup("term.rejectoffer", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.rejectoffer", LanguageID) & """ onclick=""javascript:showRejectConfirmation();""/>")
    End Sub

    Public Sub Send_Saveastemp(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""saveastemp"" name=""saveastemp"" title=""" & Copient.PhraseLib.Lookup("term.saveastemp", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.saveastemp", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_SendOutbound(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""sendoutbound"" name=""sendoutbound"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.sendoutbound", LanguageID) & "')){}else{return false}"" title=""" & Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Upload(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""upload"" name=""upload"" title=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """" & Attributes & " onclick=""javascript:{document.getElementById('uploader').style.display='block';}"" />")
    End Sub

    Public Sub Send_Listbar(Optional ByVal linesPerPage As Integer = 0, Optional ByVal sizeOfData As Integer = 0, Optional ByVal PageNum As Integer = 0, Optional ByVal searchString As String = "", Optional ByVal SortText As String = "", Optional ByVal ShowExpired As String = "", Optional ByVal QueryString As String = "", Optional ByVal HideSearcher As Boolean = False, Optional ByVal AdminUserID As Integer = 0)
        Dim myUrl As String
        Dim startVal As Integer
        Dim endVal As Integer
        Dim expiredString As String
        Dim filterString As String
        Dim filterhealth As String
        Dim filterOffer As String = ""
        Dim CustomerInquiry As String = ""
        Dim MyCommon As New Copient.CommonInc
        Dim rst As System.Data.DataTable
        Dim dt As System.Data.DataRow

        MyCommon.Open_LogixRT()

        expiredString = "&amp;ShowExpired=" & ShowExpired
        filterString = "&amp;filterOffer=" & Request.QueryString("filterOffer")
        startVal = linesPerPage * PageNum
        endVal = linesPerPage * PageNum + linesPerPage
        If startVal = 0 Then
            startVal = 1
        Else
            startVal += 1
        End If

        If endVal > sizeOfData Then endVal = sizeOfData

        If (Request.QueryString("CustomerInquiry") <> "") Then
            CustomerInquiry = "&amp;CustomerInquiry=1"
        End If

        myUrl = Request.CurrentExecutionFilePath
        Send("<!-- MyURL=" & myUrl & " -->")
        If (UCase(myUrl) = "/LOGIX/UE/STORE-HEALTH-UE.ASPX") Then
            filterString = "&amp;filterhealth=" & Request.QueryString("filterhealth")
        End If
        Send("<!-- filterString=" & filterString & " -->")

        Send("<div id=""listbar"">")
        Send(" <form id=""searchform"" name=""searchform"" action=""#"">")
        If (myUrl = "/logix/scorecard-list.aspx") Then
            Send(" <input type=""hidden"" id=""ScorecardTypeID"" name=""ScorecardTypeID"" value=""" & MyCommon.Extract_Val(Request.QueryString("ScorecardTypeID")) & """ />")
        ElseIf (myUrl = "/logix/store-list.aspx") Then
            Send(" <input type=""hidden"" id=""LocationTypeID"" name=""LocationTypeID"" value=""" & MyCommon.Extract_Val(Request.QueryString("LocationTypeID")) & """ />")
        End If
        Sendb("  <div id=""searcher"" title=""" & Copient.PhraseLib.Lookup("term.searchterms", LanguageID) & """>")
        If Not HideSearcher Then
            Send("   <input type=""text"" id=""searchterms"" name=""searchterms"" maxlength=""100"" value=""" & searchString & """ />")
            If (myUrl = "/logix/offer-list.aspx") Then
                Send("   <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
                Send("   <input type=""button"" id=""advsearch"" name=""advsearch"" value=""..."" alt=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ onclick=""launchAdvSearch();"" /><br />")
            ElseIf (myUrl = "/logix/CM-cashier-report.aspx") Then
                Send("   <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
                Send("   <input type=""button"" id=""advsearch"" name=""advsearch"" value=""..."" alt=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ onclick=""launchAdvSearch();"" /><br />")
            Else
                Send("   <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ /><br />")
            End If
        End If
        Send("  </div>")
        Send("  <div id=""paginator"">")
        If (searchString = "") Then
        Else
            Sendb("   <span class=""printonly"">")
            Sendb(Copient.PhraseLib.Lookup("term.searchterms", LanguageID) & ": """ & searchString & """<br />")
            Send("</span>")
        End If
        If (PageNum > 0) Then
            If (QueryString <> "") Then
                Send("   <span id=""first""><a id=""firstPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=0&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a></span>&nbsp;")
                Send("   <span id=""previous""><a id=""previousPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=" & PageNum - 1 & "&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a></span>")
            Else
                Send("   <span id=""first""><a id=""firstPageLink"" href=""" & myUrl & "?pagenum=0&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a></span>&nbsp;")
                Send("   <span id=""previous""><a id=""previousPageLink"" href=""" & myUrl & "?pagenum=" & PageNum - 1 & "&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a></span>")
            End If
        Else
            Send("   <span id=""first""><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</span>&nbsp;")
            Send("   <span id=""previous"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</span>")
        End If
        If sizeOfData = 0 Then
            Send("   &nbsp;[ " & Copient.PhraseLib.Lookup("term.noresults", LanguageID) & " ]&nbsp;")
        Else
            Send("   &nbsp;[ <b>" & startVal & "</b> - <b>" & endVal & "</b> " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " <b>" & sizeOfData & "</b> ]&nbsp;")
        End If
        If (sizeOfData > endVal) Then
            If (QueryString <> "") Then
                Send("   <span id=""next""><a id=""nextPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=" & PageNum + 1 & "&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a></span>&nbsp;")
                Send("   <span id=""last""><a id=""lastPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a></span><br />")
            Else
                Send("   <span id=""next""><a id=""nextPageLink"" href=""" & myUrl & "?pagenum=" & PageNum + 1 & "&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a></span>&nbsp;")
                Send("   <span id=""last""><a id=""lastPageLink"" href=""" & myUrl & "?pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&amp;searchterms=" & searchString & SortText & expiredString & filterString & CustomerInquiry & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a></span><br />")
            End If
        Else
            Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</span>&nbsp;")
            Send("   <span id=""last"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></span><br />")
        End If
        Send("  </div>")
        Send("  <div id=""filter"" title=""" & Copient.PhraseLib.Lookup("term.filter", LanguageID) & """>")
        If (myUrl = "/logix/offer-list.aspx") Then
            Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
            filterOffer = Request.QueryString("filterOffer")
            If filterOffer = "" Then filterOffer = "1"
            Send("    <option value=""0""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            Send("    <option value=""1""" & IIf(filterOffer = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
            Send("    <option value=""2""" & IIf(filterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
            If (CustomerInquiry = "") Then
                Send("    <option value=""3""" & IIf(filterOffer = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showdiscrepancies", LanguageID) & "</option>")
            End If
            Send("  </select>")
        ElseIf (myUrl = "/logix/extoffer-list.aspx") Then
            Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
            filterOffer = Request.QueryString("filterOffer")
            If filterOffer = "" Then filterOffer = "0"
            MyCommon.QueryStr = "select ExtInterfaceID, Name from ExtCRMInterfaces with (NoLock) where ExtInterfaceID>0;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                Send("  <option value=""-1""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.alloffers", LanguageID) & "</option>")
                For Each dt In rst.Rows
                    Send("  <option value=""" & dt.Item("ExtInterfaceID") & """" & IIf(filterOffer = dt.Item("ExtInterfaceID"), " selected=""selected""", "") & ">" & dt.Item("Name") & "</option>")
                Next
                Send(" </select>")
            End If
        ElseIf (myUrl = "/logix/CAM/CAM-offer-list.aspx") Then 'Filter for the CAM offer page
            Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
            filterOffer = Request.QueryString("filterOffer")
            If filterOffer = "" Then filterOffer = "1"
            Send("    <option value=""0""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            Send("    <option value=""1""" & IIf(filterOffer = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
            Send("    <option value=""2""" & IIf(filterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
            If (CustomerInquiry = "") Then
                Send("    <option value=""3""" & IIf(filterOffer = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showdiscrepancies", LanguageID) & "</option>")
            End If
            Send("    <option value=""4""" & IIf(filterOffer = "4", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.showmutuallyexclusive", LanguageID) & "</option>")
            Send("  </select>")
        ElseIf (myUrl = "/logix/banneroffer-list.aspx") Then
            Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
            filterOffer = Request.QueryString("filterOffer")
            If filterOffer = "" Then filterOffer = "0"
            If AdminUserID = 0 Then
                MyCommon.QueryStr = "select BannerID,Name from Banners where BannerID > 0 order by Name"
            Else
                MyCommon.QueryStr = "select BAN.BannerID as BannerID, BAN.Name from Banners BAN with (NoLock) " &
                              "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " &
                              "WHERE(BAN.Deleted = 0 And AdminUserID = " & AdminUserID & ") order by BAN.Name"
            End If
            Send("  <option value=""-1""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                For Each dt In rst.Rows
                    Send("  <option value=""" & dt.Item("BannerID") & """" & IIf(filterOffer = dt.Item("BannerID"), " selected=""selected""", "") & ">" & dt.Item("Name") & "</option>")
                Next
            End If
            Send(" </select>")
        ElseIf (myUrl = "/logix/reports-list.aspx") Then
            Send("   <select id=""ShowExpired"" name=""ShowExpired"" onchange=searchform.submit()>")
            If (ShowExpired = "TRUE") Then
                Send("    <option value=""TRUE"" selected=""selected"">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
                Send("    <option value=""FALSE"" >" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
            Else
                Send("    <option value=""TRUE"" >" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
                Send("    <option value=""FALSE"" selected=""selected"">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
            End If
            Send("  </select>")
        ElseIf (myUrl = "/logix/store-health-cpe.aspx" Or myUrl = "/logix/store-health-cm.aspx" Or UCase(myUrl) = "/LOGIX/UE/STORE-HEALTH-UE.ASPX") Then
            filterhealth = Request.QueryString("filterhealth")
            Send("   <select id=""filterhealth"" name=""filterhealth"" onchange=searchform.submit()>")
            Send("    <option value=""0""" & IIf(filterhealth = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            Send("    <option value=""1""" & IIf(filterhealth = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.communications", LanguageID) & " " & Copient.PhraseLib.Lookup("term.ok", LanguageID) & "</option>")
            Send("    <option value=""2""" & IIf(filterhealth = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.allerrors", LanguageID) & "</option>")
            If (myUrl = "/logix/store-health-cpe.aspx") Then
                Send("    <option value=""3""" & IIf(filterhealth = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.centralerrors", LanguageID) & "</option>")
                Send("    <option value=""4""" & IIf(filterhealth = "4", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.localerrors", LanguageID) & "</option>")
            End If
            Send("    <option value=""5""" & IIf(filterhealth = "5", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.allactivefailovers", LanguageID) & "</option>")
            Send("    <option value=""6""" & IIf(filterhealth = "6", " selected=""selected""", "") & ">" & IIf(UCase(myUrl) = "/LOGIX/UE/STORE-HEALTH-UE.ASPX", Copient.PhraseLib.Lookup("term.failoverhistory", LanguageID), Copient.PhraseLib.Lookup("term.allfailovers", LanguageID)) & "</option>")
            Send("    <option value=""7""" & IIf(filterhealth = "7", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.iplneeded", LanguageID) & " " & Copient.PhraseLib.Lookup("term.yes", LanguageID) & "</option>")
            Send("    <option value=""8""" & IIf(filterhealth = "8", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.showactivelocations", LanguageID) & "</option>")
            Send("   </select>")
        ElseIf (myUrl = "/logix/offer-health.aspx" OrElse myUrl = "/logix/CRMoffer-list.aspx") Then
            filterhealth = Request.QueryString("filterhealth")
            Send("   <select id=""filterhealth"" name=""filterhealth"" onchange=searchform.submit()>")
            Send("    <option value=""0""" & IIf(filterhealth = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            Send("    <option value=""1""" & IIf(filterhealth = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.showvalid", LanguageID) & "</option>")
            Send("    <option value=""2""" & IIf(filterhealth = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.showinvalid", LanguageID) & "</option>")
            If MyCommon.Fetch_SystemOption(25) <> "0" Then
                Send("    <option value=""3""" & IIf(filterhealth = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.showcrm", LanguageID) & "</option>")
            End If
            Send("   </select>")
        ElseIf (myUrl = "/logix/customer-offers.aspx") Then
            filterhealth = Request.QueryString("filterhealth")
            Send("   <select id=""filterhealth"" name=""filterhealth"" onchange=searchform.submit()>")
            Send("    <option value=""0""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            Send("    <option value=""2""" & IIf(filterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
            Send("   </select>")
        ElseIf (myUrl = "/logix/PLU-list.aspx") Then
            Send("   <select id=""filterPLU"" name=""filterPLU"" onchange=""searchform.submit();"">")
            filterString = Request.QueryString("filterPLU")
            If filterString = "" Then filterString = "0"
            Send("    <option value=""0""" & IIf(filterString = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.ShowAllActivePLUs", LanguageID) & "</option>")
            Send("    <option value=""1""" & IIf(filterString = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.ShowActivePLURange", LanguageID) & "</option>")
            Send("    <option value=""2""" & IIf(filterString = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.ShowOutOfRangePLU", LanguageID) & "</option>")
            Send("  </select>")
        ElseIf (myUrl = "/logix/adjustmentUPC-list.aspx") Then
            Send("   <select id=""filterUPC"" name=""filterUPC"" onchange=""searchform.submit();"">")
            filterString = Request.QueryString("filterUPC")
            If filterString = "" Then filterString = "0"
            Send("    <option value=""0""" & IIf(filterString = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.AllActiveUPC", LanguageID) & "</option>")
            Send("    <option value=""1""" & IIf(filterString = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.ShowUPCInRange", LanguageID) & "</option>")
            Send("    <option value=""2""" & IIf(filterString = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.ShowUPCOutOfRange", LanguageID) & "</option>")
            Send("  </select>")
        Else
            Send("   <select id=""filterselect"" name=""filterselect"" style=""display: none;"">")
            Send("    <option value=""1"">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            Send("   </select>")
        End If
        Send("  </div>")
        If CustomerInquiry <> "" Then
            Send("  <input type=""hidden"" id=""CustomerInquiry"" name=""CustomerInquiry"" value=""1"" />")
        End If
        Send(" </form>")
        Send(" <hr class=""hidden"" />")
        Send("</div>")

        MyCommon.Close_LogixRT()
        MyCommon = Nothing
    End Sub

    Public Sub Send_Status(ByVal OfferID As Integer, Optional ByVal EngineID As Integer = 1)
        Dim MyCommon As New Copient.CommonInc
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim status As Integer = -1
        Dim preStatus As Integer = -1

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        ' Query for and set the offer's status value
        If (EngineID = 2 OrElse EngineID = 9) Then
            MyCommon.QueryStr = "Select StatusFlag,isnull(EndDate,0) as prodEndDate, DeployDeferred from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
        Else
            MyCommon.QueryStr = "select StatusFlag,isnull(prodEndDate,0) as prodEndDate, DeployDeferred from Offers with (NoLock) where OfferID=" & OfferID
        End If
        rst = MyCommon.LRT_Select()
        For Each row In rst.Rows
            status = row.Item("StatusFlag")
            preStatus = status
        Next
        For Each row In rst.Rows
            If ((MyCommon.Fetch_UE_SystemOption(200) = "1") AndAlso (MyCommon.NZ(row.Item("prodEndDate"), DateTime.Now) < DateTime.Now)) OrElse ((MyCommon.NZ(row.Item("prodEndDate"), Today) < Today)) Then
                status = 3
            End If

            If (MyCommon.NZ(row.Item("DeployDeferred"), False) = True) Then
                status = 4
            End If
        Next
        If (OfferID <= 0) Then
        Else
            'This piece of code will only work for deployed offer if status changed due to some other condition.
            If (preStatus = 2 AndAlso status <> 2) Then
                Sendb("<div id=""infobar"" class=""green-background"">" & Copient.PhraseLib.Lookup("offer.status2msg", LanguageID) & "</div>")
            End If


            Sendb("<div id=""statusbar""")
            If (OfferLockedforCollisionDetection = True) Then
                Sendb(" class=""green-background"">" & Copient.PhraseLib.Lookup("term.collisiondetectioninprogress", LanguageID))
            ElseIf (status = 0) Then
                Sendb(" class=""green-background"">" & Copient.PhraseLib.Lookup("offer.status0msg", LanguageID))
            ElseIf (status = 1) Then
                Sendb(" class=""orange-background"">" & Copient.PhraseLib.Lookup("offer.status1msg", LanguageID))
            ElseIf (status = 2) Then
                Sendb(" class=""green-background"">" & Copient.PhraseLib.Lookup("offer.status2msg", LanguageID))
            ElseIf (status = 3) Then
                Sendb(" class=""grey-background"">" & Copient.PhraseLib.Lookup("offer.status3msg", LanguageID))
            ElseIf (status = 4) Then
                Sendb(" class=""green-background"">" & Copient.PhraseLib.Lookup("offer.status4msg", LanguageID))
            ElseIf (status = 11 Or status = 12) Then
                Sendb(" class=""green-background"">" & Copient.PhraseLib.Lookup("alert.awaitingrecommendation", LanguageID))
            ElseIf (status = 13 Or status = 14 Or status = 15) Then
                Sendb(" class=""green-background"">" & Copient.PhraseLib.Lookup("alert.awaitingapproval", LanguageID))
            Else
                Sendb(">")
            End If
            Send("</div>")
        End If
        MyCommon.Close_LogixRT()
        MyCommon = Nothing
    End Sub

    'This is the old version of this function.  Stop using this and start calling the new one (below).
    Public Sub Send_Denied(ByVal BodyType As Integer, Optional ByVal PermPhraseName As String = "")
        If (BodyType = 0) Then
            Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
            If (PermPhraseName <> "") Then
                Send("<br />")
                Send(Copient.PhraseLib.Lookup("term.requiredpermission", LanguageID) & ": ")
                Send(Copient.PhraseLib.Lookup(PermPhraseName, LanguageID))
            End If
        ElseIf (BodyType = 1) OrElse (BodyType = 2) Then
            Send("<div id=""intro"">")
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.accessdenied", LanguageID) & "</h1>")
            If (BodyType = 1) Then
                Send("</div>")
                Send("<div id=""main"">")
                Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
                If (PermPhraseName <> "") Then
                    Send("<br />")
                    Send(Copient.PhraseLib.Lookup("term.requiredpermission", LanguageID) & ": ")
                    Send(Copient.PhraseLib.Lookup(PermPhraseName, LanguageID))
                End If
            ElseIf (BodyType = 2) Then
                Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
                If (PermPhraseName <> "") Then
                    Send("<br />")
                    Send(Copient.PhraseLib.Lookup("term.requiredpermission", LanguageID) & ": ")
                    Send(Copient.PhraseLib.Lookup(PermPhraseName, LanguageID))
                End If
            End If
            Send("</div>")
        End If
    End Sub

    'PermissionIDList = a comma-separated list of Permisions.PermissionID's that are required to access the denied resource
    'PermissionNameList = a comma-separated list of Permissions.Name's that are required to access the denied resource
    'PermissionIDList has priority over PermissionNameList
    Private Sub Send_Denied_Msg(ByRef Common As Copient.CommonInc, ByVal PermissionIDList As String, ByVal PermissionNameList As String)

        Dim dst As DataTable
        Dim row As DataRow
        Dim PermissionNameArray() As String
        Dim TempNameStr As String
        Dim index As Integer

        If Not (PermissionIDList = "") Then
            Common.QueryStr = "select isnull(PhraseID, 0) as PhraseID, isnull(Description, 'Unknown') as Description from Permissions where PermissionID in (" & PermissionIDList & ");"
        ElseIf Not (PermissionNameList = "") Then
            TempNameStr = ""
            PermissionNameArray = Split(PermissionNameList, ",")
            For index = 0 To UBound(PermissionNameArray)
                If Not (TempNameStr = "") Then TempNameStr = TempNameStr & ", "
                TempNameStr = TempNameStr & "'" & Trim(PermissionNameArray(index)) & "'"
            Next
            Common.QueryStr = "select isnull(PhraseID, 0) as PhraseID, isnull(Description, 'Unknown') as Description from Permissions where Description in (" & TempNameStr & ");"
        End If
        Send("<br />")
        Send(Copient.PhraseLib.Lookup("term.requiredpermission", LanguageID) & ": ")
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            Send("<UL>")
            For Each row In dst.Rows
                Send("<LI>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID, row.Item("Description")) & "</LI>")
            Next
            Send("</UL>")
        End If

    End Sub

    '******* This is the new Send_Denied function and should be used going forward *********
    'PermissionIDList = a comma-separated list of Permisions.PermissionID's that are required to access the denied resource
    'PermissionNameList = a comma-separated list of Permissions.Name's that are required to access the denied resource
    'PermissionIDList has priority over PermissionNameList
    Public Sub Send_Denied(ByVal BodyType As Integer, ByRef Common As Copient.CommonInc, Optional ByVal PermissionIDList As String = "", Optional ByVal PermissionNameList As String = "")
        If (BodyType = 0) Then
            Send("<B>" & Copient.PhraseLib.Lookup("error.forbidden", LanguageID) & "</B>")
            Send_Denied_Msg(Common, PermissionIDList, PermissionNameList)
        ElseIf (BodyType = 1) OrElse (BodyType = 2) Then
            Send("<div id=""intro"">")
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.accessdenied", LanguageID) & "</h1>")
            If (BodyType = 1) Then
                Send("</div>")
                Send("<div id=""main"">")
                Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
                Send_Denied_Msg(Common, PermissionIDList, PermissionNameList)
            ElseIf (BodyType = 2) Then
                Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
                Send_Denied_Msg(Common, PermissionIDList, PermissionNameList)
            End If
            Send("</div>")
        End If
    End Sub

    Public Sub Send_BoxResizer(ByVal DivName As String, ByVal ImgName As String, ByVal AltTitle As String, ByVal IsExpanded As Boolean)
        Dim ArrowDirection As String
        Dim ImgAlt As String
        ArrowDirection = IIf(IsExpanded, "up", "down")
        ImgAlt = IIf(IsExpanded, "Hide", "Show")
        ImgAlt += " " & AltTitle
        Send("<div class=""resizer"">")
        Send("  <a href=""javascript:resizeDiv('" & DivName & "','" & ImgName & "','" & AltTitle & "');"">")
        Send("    <img id=""" & ImgName & """ src=""/images/arrow" & ArrowDirection & "-off.png"" alt=""" & ImgAlt & """ title=""" & ImgAlt & """ ")
        Send("      onmouseover=""javascript:handleResizeHover(true,'" & DivName & "','" & ImgName & "');"" ")
        Send("      onmouseout=""javascript:handleResizeHover(false,'" & DivName & "','" & ImgName & "');"" />")
        Send("  </a>")
        Send("</div>")
        Sendb("<br clear=""all"" />")
    End Sub

    Public Sub Send_ProductGroupSelector(ByRef Logix As Object, ByRef TransactionLevelSelected As Object, ByRef FromTemplate As Object,
                                         ByRef Disallow_Edit As Object, ByRef selecteditem As Object, ByRef ExcludedItem As Object,
                                         ByRef RewardID As Object, ByRef EngineID As Integer,
                                         Optional ByVal IsTemplate As Boolean = False, Optional ByVal bDisallowEditPg As Boolean = False)
        Dim MyCommon As New Copient.CommonInc
        Dim row As System.Data.DataRow
        Dim rst As System.Data.DataTable
        Dim Limiter As String = ""

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        Send("        <div class=""box"" id=""groups"">")
        Send("            <h2><span>" & Copient.PhraseLib.Lookup("term.productcondition", LanguageID) & "</span>")

        If (IsTemplate) Then
            Send("<span class=""tempRequire"">")
            If (bDisallowEditPg) Then
                Send("<input type=""checkbox"" class=""tempcheck"" id=""DisallowEditPg1"" name=""DisallowEditPg"" checked=""checked"" />")
            Else
                Send("<input type=""checkbox"" class=""tempcheck"" id=""DisallowEditPg1"" name=""DisallowEditPg"" />")
            End If
            Send("<label for=""temp-Tiers"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
        ElseIf (FromTemplate And bDisallowEditPg) Then
            Send("<span class=""tempRequire"">")
            Send("<input type=""checkbox"" class=""tempcheck"" id=""DisallowEditPg2"" name=""DisallowEditPg"" disabled=""disabled"" checked=""checked"" />")
            Send("<label for=""temp-Tiers"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
        End If

        Send("  </h2>")
        Send("            <div style=""float:left;position:relative;"">")
        Send("                <label for=""pgroup-select"">" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</label><br clear=""all"" />")

        Send("                <select class=""longer"" id=""pgroup-select"" name=""pgroup-select"" size=""2"">")
        MyCommon.QueryStr = "select OFR.ProductGroupID,ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C on OFR.ProductGroupID=C.ProductGroupID and RewardID=" & RewardID & ";"

        rst = MyCommon.LRT_Select
        If (rst.Rows.Count = 0) Then
            Send("<option>" & Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "</option>")
            TransactionLevelSelected = True
        End If
        For Each row In rst.Rows
            Send("<option value=""" & row.Item("ProductGroupID") & """>" & row.Item("Name") & "</option>")
            selecteditem = row.Item("ProductGroupID")
        Next
        Send("                </select><br />")
        Send("                <br style=""line-height: 11px;"" />")

        Send("                <label for=""pgroup-exclude"">" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ":</label><br clear=""all"" />")
        Send("                <select class=""longer"" id=""pgroup-exclude"" name=""pgroup-exclude"" size=""2"">")

        MyCommon.QueryStr = "select OFR.ProductGroupID,OFR.ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C on OFR.ExcludedProdGroupID=C.ProductGroupID  where not(ExcludedProdGroupID=0) and RewardID=" & RewardID & ";"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            Send("<option value=""" & row.Item("ExcludedProdGroupID") & """>" & row.Item("Name") & "</option>")
            ExcludedItem = row.Item("ExcludedProdGroupID")
        Next
        Send("                </select>")

        Send("            </div>")
        Send("")
        Send("            <div style=""float:left;position:relative;padding: 14px 1px 0px 1px;"">")
        Sendb("                <input ")
        If Not (IsTemplate) Then
            If Not (Logix.userroles.editoffer And Not (FromTemplate And Disallow_Edit)) Then
                Sendb("disabled=""disabled"" ")
            End If
        Else
            If Not (Logix.userroles.edittemplates) Then
                Sendb("disabled=""disabled"" ")
            End If
        End If
        Send("type=""submit"" class=""arrowrem"" id=""pgroup-rem1"" name=""pgroup-rem1"" title=""" & Copient.PhraseLib.Lookup("term.unselect", LanguageID) & """ value=""&#187;"" />")
        Sendb("                <input ")
        If Not (IsTemplate) Then
            If Not (Logix.userroles.editoffer And Not (FromTemplate And Disallow_Edit)) Then
                Sendb("disabled=""disabled"" ")
            End If
        Else
            If Not (Logix.userroles.edittemplates) Then
                Sendb("disabled=""disabled"" ")
            End If
        End If
        Send("type=""submit"" class=""arrowadd"" id=""pgroup-add1"" name=""pgroup-add1"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ value=""&#171;"" /><br />")

        Send("                <br style=""line-height: 24px;"" />")
        Sendb("                <input ")
        If Not (IsTemplate) Then
            If Not (Logix.userroles.editoffer And Not (FromTemplate And Disallow_Edit)) Then
                Sendb("disabled=""disabled"" ")
            End If
        Else
            If Not (Logix.userroles.edittemplates) Then
                Sendb("disabled=""disabled"" ")
            End If
        End If
        Send("type=""submit"" class=""arrowrem"" id=""pgroup-rem2"" name=""pgroup-rem2"" title=""" & Copient.PhraseLib.Lookup("term.unexclude", LanguageID) & """ value=""&#187;"" />")
        Sendb("                <input ")
        If Not (IsTemplate) Then
            If Not (Logix.userroles.editoffer And Not (FromTemplate And Disallow_Edit)) Then
                Sendb("disabled=""disabled"" ")
            End If
        Else
            If Not (Logix.userroles.edittemplates) Then
                Sendb("disabled=""disabled"" ")
            End If
        End If
        Send("type=""submit"" class=""arrowadd"" id=""pgroup-add2"" name=""pgroup-add2"" title=""" & Copient.PhraseLib.Lookup("term.exclude", LanguageID) & """ value=""&#171;"" />")

        Send("            </div>")
        Send("")
        Send("            <div style=""float:left;position:relative;"">")
        Send("                <label for=""pgroup-avail"">" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</label><br clear=""all"" />")
        Send("                <select class=""longer"" id=""pgroup-avail"" name=""pgroup-avail"" size=""6"">")

        If (ExcludedItem) Then Limiter = "and ProductGroupID <> " & ExcludedItem
        If (selecteditem) Then Limiter = Limiter & " and ProductGroupID <> " & selecteditem

        MyCommon.QueryStr = "select ProductGroupID,CreatedDate,Name,LastUpdate,AnyProduct from ProductGroups with (NoLock) where Deleted=0" & Limiter & " order by AnyProduct desc, Name"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            Send("<option value=""" & row.Item("ProductGroupID") & """>" & row.Item("Name") & "</option>")
        Next
        Send("                </select>")
        Send("            </div>")
        Send("            <br clear=""left"" /><br class=""zero"" />")
        Send("            <hr class=""hidden"" />")
        Send("          </div>")
    End Sub

    Public Sub Send_ProductConditionSelector(ByRef Logix As Object, ByRef TransactionLevelSelected As Object, ByRef FromTemplate As Object,
                                             ByRef Disallow_Edit As Object, ByRef selecteditem As Object, ByRef ExcludedItem As Object,
                                             ByRef RewardID As Object, ByRef EngineID As Integer,
                                             Optional ByVal IsTemplate As Boolean = False, Optional ByVal bDisallowEditPg As Boolean = False)
        Dim MyCommon As New Copient.CommonInc
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim disabledattribute As String = ""

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        Send("<div class=""box"" id=""selector"">")
        Send("  <h2>")
        Send("    <span>" & Copient.PhraseLib.Lookup("term.productcondition", LanguageID) & "</span>")

        If (IsTemplate) Then
            Send("<span class=""tempRequire"">")
            If (bDisallowEditPg) Then
                Send("<input type=""checkbox"" class=""tempcheck"" id=""DisallowEditPg1"" name=""DisallowEditPg"" checked=""checked"" />")
            Else
                Send("<input type=""checkbox"" class=""tempcheck"" id=""DisallowEditPg1"" name=""DisallowEditPg"" />")
            End If
            Send("<label for=""temp-Tiers"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
        ElseIf (FromTemplate And bDisallowEditPg) Then
            Send("<span class=""tempRequire"">")
            Send("<input type=""checkbox"" class=""tempcheck"" id=""DisallowEditPg2"" name=""DisallowEditPg"" disabled=""disabled"" checked=""checked"" />")
            Send("<label for=""temp-Tiers"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
        End If

        Send("  </h2>")
        If Not (IsTemplate) Then
            If Not (Logix.userroles.editoffer And Not (FromTemplate And Disallow_Edit)) Then
                disabledattribute = " disabled=""disabled"""
            End If
        Else
            If Not (Logix.userroles.edittemplates) Then
                disabledattribute = " disabled=""disabled"""
            End If
        End If

        Send("<input type=""radio"" id=""functionradio1"" name=""functionradio"" " & IIf(MyCommon.Fetch_SystemOption(175) = "1", "checked=""checked""", "") & " " & disabledattribute & " /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
        Send("<input type=""radio"" id=""functionradio2"" name=""functionradio"" " & IIf(MyCommon.Fetch_SystemOption(175) = "2", "checked=""checked""", "") & " " & disabledattribute & " /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
        Send("<input type=""text"" class=""medium"" onkeyup=""handleKeyUp(200);"" id=""functioninput"" name=""functioninput"" maxlength=""100"" value=""" & disabledattribute & """ /><br />")
        Send("<div id=""pgList"">")
        Send("<select class=""longer"" id=""functionselect"" name=""functionselect"" size=""10""" & disabledattribute & ">")
        Dim Limiter As String = ""
        If (ExcludedItem) Then Limiter = "and ProductGroupID <> " & ExcludedItem
        If (selecteditem) Then Limiter = Limiter & " and ProductGroupID <> " & selecteditem

        MyCommon.QueryStr = "select ProductGroupID,CreatedDate,Name,LastUpdate,AnyProduct from ProductGroups with (NoLock) where Deleted=0" & Limiter & " order by AnyProduct desc, Name"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                Send("<option style=""font-weight:bold;color:brown;"" value=""" & row.Item("ProductGroupID") & """>" & row.Item("Name") & "</option>")
            Else
                Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
            End If
        Next
        Send("</select>")
        Send("</div>")
        Send("<br />")
        Send("<br class=""half"" />")
        Send("<b><label for=""selected"">" & Copient.PhraseLib.Lookup("term.selectedproducts", LanguageID) & ":</label></b><br />")
        Send("<input type=""button"" class=""regular select"" id=""select1"" name=""select1"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""handleSelectClick('select1');""" & disabledattribute & " />&nbsp;")
        Send("<input type=""button"" class=""regular deselect"" id=""deselect1"" name=""deselect1"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""handleSelectClick('deselect1');"" disabled=""disabled"" /><br />")
        Send("<br class=""half"" />")
        Send("<select class=""longer"" id=""selected"" name=""selected"" size=""2""" & disabledattribute & ">")

        MyCommon.QueryStr = "select OFR.ProductGroupID,ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C with (NoLock) on OFR.ProductGroupID=C.ProductGroupID and RewardID=" & RewardID & ";"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count = 0) Then
            'Send("  <option>" & Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "</option>")
            TransactionLevelSelected = True
        End If
        For Each row In rst.Rows
            If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                Send("  <option style=""font-weight:bold;color:brown;"" value=""" & row.Item("ProductGroupID") & """>" & row.Item("Name") & "</option>")
            Else
                Send("  <option value=""" & row.Item("ProductGroupID") & """>" & row.Item("Name") & "</option>")
            End If
            selecteditem = row.Item("ProductGroupID")
        Next
        Send("</select>")
        Send("<br />")

        Send("<b><label for=""excluded"">" & Copient.PhraseLib.Lookup("term.excludedproducts", LanguageID) & ":</label></b><br />")
        Send("<input type=""button"" class=""regular select"" id=""select2"" name=""select2"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""handleSelectClick('select2');"" disabled=""disabled""" & disabledattribute & " />&nbsp;")
        Send("<input type=""button"" class=""regular deselect"" id=""deselect2"" name=""deselect2"" value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;"" onclick=""handleSelectClick('deselect2');"" disabled=""disabled""" & disabledattribute & " /><br />")
        Send("<br class=""half"" />")
        Send("<select class=""longer"" id=""excluded"" name=""excluded"" size=""2""" & disabledattribute & ">")

        MyCommon.QueryStr = "select OFR.ProductGroupID,OFR.ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C with (NoLock) on OFR.ExcludedProdGroupID=C.ProductGroupID  where not(ExcludedProdGroupID=0) and RewardID=" & RewardID & ";"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            Send("<option value=""" & row.Item("ExcludedProdGroupID") & """>" & row.Item("Name") & "</option>")
            ExcludedItem = row.Item("ExcludedProdGroupID")
        Next
        Send("</select>")
        Send("<br />")

        Send("&nbsp;")
        Send("<br class=""half"" />")
        Send("<hr class=""hidden"" />")
        Send("</div>")
    End Sub

    Public Sub Send_NotesButton(Optional ByVal NoteTypeID As Integer = 0, Optional ByVal LinkID As Integer = 0, Optional ByVal AdminUserID As Integer = 0)
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim HasVisibleNotes As Boolean = False
        Dim HasNewNotes As Boolean = False
        Dim HasImportantNotes As Boolean = False
        Dim ActivityTypeID As Integer = 0

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        MyCommon.Open_LogixXS()

        AdminUserID = Verify_AdminUser(MyCommon, Logix)

        MyCommon.QueryStr = "select N.NoteTypeID, N.ActivityTypeID, A.PhraseID from NoteTypes as N with (NoLock) " &
                            "left join ActivityTypes as A with (NoLock) on A.ActivityTypeID=N.ActivityTypeID " &
                            "where NoteTypeID=" & NoteTypeID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            ActivityTypeID = MyCommon.NZ(rst.Rows(0).Item("ActivityTypeID"), 0)
        Else
            ActivityTypeID = 0
        End If

        If NoteTypeID = 4 Then
            MyCommon.QueryStr = "select N.NoteID, 4 as NoteTypeID, N.CustomerPK as LinkID, N.AdminUserID, N.CreatedDate, N.Private, " &
                                "N.Important, N.Deleted, 25 as ActivityTypeID from CustomerNotes as N with (NoLock) " &
                                "where CustomerPK=" & LinkID & " and Deleted=0;"
            rst = MyCommon.LXS_Select
        Else
            MyCommon.QueryStr = "select N.NoteID, N.NoteTypeID, N.LinkID, N.AdminUserID, N.CreatedDate, N.Private, " &
                                "N.Important, N.Deleted, NT.ActivityTypeID from Notes as N with (NoLock) " &
                                "inner join NoteTypes as NT on NT.NoteTypeID=N.NoteTypeID " &
                                "where N.NoteTypeID=" & NoteTypeID & " and LinkID=" & LinkID & " and Deleted=0;"
            rst = MyCommon.LRT_Select
        End If

        For Each row In rst.Rows
            If ((row.Item("Private") = False) OrElse ((row.Item("Private") = True AndAlso row.Item("AdminUserID") = AdminUserID))) Then
                HasVisibleNotes = True
            End If
            If (((row.Item("Private") = False) OrElse ((row.Item("Private") = True AndAlso row.Item("AdminUserID") = AdminUserID))) And row.Item("Important")) Then
                HasImportantNotes = True
            End If
            If DateDiff(DateInterval.Day, row.Item("CreatedDate"), DateTime.Today) = 0 Then
                HasNewNotes = True
            End If
        Next
        Send("")
        Send("<div id=""stickynote"">")
        Send("  <a href=""javascript:toggleNotes()"">")
        If HasVisibleNotes Then
            If HasImportantNotes Then
                If HasNewNotes Then
                    Sendb("    <img src=""/images/notes-newimportant.png""")
                Else
                    Sendb("    <img src=""/images/notes-someimportant.png""")
                End If
            Else
                If HasNewNotes Then
                    Sendb("    <img src=""/images/notes-new.png""")
                Else
                    Sendb("    <img src=""/images/notes-some.png""")
                End If
            End If
            Sendb(" id=""notesbutton"" name=""notesbutton""")
            If rst.Rows.Count = 1 Then
                Send(" alt=""" & rst.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.note", LanguageID), VbStrConv.Lowercase) & """ title=""" & rst.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.note", LanguageID), VbStrConv.Lowercase) & """ />")
            Else
                Send(" alt=""" & rst.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.notes", LanguageID), VbStrConv.Lowercase) & """ title=""" & rst.Rows.Count & " " & StrConv(Copient.PhraseLib.Lookup("term.notes", LanguageID), VbStrConv.Lowercase) & """ />")
            End If
        Else
            Send("    <img src=""/images/notes-none.png"" id=""notesbutton"" name=""notesbutton"" alt=""" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & """ />")
        End If
        Send("  </a>")
        Send("</div>")
        MyCommon.Close_LogixRT()
        MyCommon.Close_LogixXS()
        MyCommon = Nothing
    End Sub

    Public Sub Send_Notes(Optional ByVal NoteTypeID As Integer = 0, Optional ByVal LinkID As Integer = 0, Optional ByVal AdminUserID As Integer = 0)
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim rst2 As System.Data.DataTable
        Dim IsPrivate As Boolean = False
        Dim IsImportant As Boolean = False
        Dim ActivityTypeID As Integer = 0
        Dim NoteID As Integer = 0
        Dim FirstName As String = ""
        Dim LastName As String = ""
        Dim HistoryString As String = ""
        Dim NoteSortID As Integer = 0
        Dim NoteSortText As String = ""
        Dim NoteSortDirection As String = ""
        Dim NoteMode As Boolean = False

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        MyCommon.Open_LogixXS()

        AdminUserID = Verify_AdminUser(MyCommon, Logix)

        MyCommon.QueryStr = "select FirstName, LastName from AdminUsers with (NoLock) where AdminUserID=" & AdminUserID & ";"
        rst = MyCommon.LRT_Select
        FirstName = MyCommon.NZ(rst.Rows(0).Item("FirstName"), "")
        LastName = MyCommon.NZ(rst.Rows(0).Item("LastName"), "")

        MyCommon.QueryStr = "select N.NoteTypeID, N.ActivityTypeID, A.PhraseID from NoteTypes as N with (NoLock) " &
                            "left join ActivityTypes as A with (NoLock) on A.ActivityTypeID=N.ActivityTypeID " &
                            "where NoteTypeID=" & NoteTypeID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            ActivityTypeID = MyCommon.NZ(rst.Rows(0).Item("ActivityTypeID"), 0)
        Else
            ActivityTypeID = 0
        End If

        NoteMode = (Request.Form("NoteMode") = "on")

        If (Request.QueryString("NoteSortText") <> "") Then
            NoteSortText = Request.QueryString("NoteSortText")
        Else
            NoteSortText = "CreatedDate"
        End If
        If (Request.QueryString("NoteSortDirection") <> "") Then
            NoteSortDirection = Request.QueryString("NoteSortDirection")
        Else
            NoteSortDirection = "DESC"
        End If

        NoteSortID = MyCommon.Extract_Val(Request.Form("notesort"))
        If NoteSortID > 0 Then
            NoteMode = True
        End If
        Select Case NoteSortID
            Case 1
                NoteSortText = "CreatedDate"
                NoteSortDirection = "DESC"
            Case 2
                NoteSortText = "CreatedDate"
                NoteSortDirection = "ASC"
            Case 3
                NoteSortText = "LastName"
                NoteSortDirection = "DESC"
            Case 4
                NoteSortText = "LastName"
                NoteSortDirection = "ASC"
        End Select

        If Request.Form("notesave") <> "" Then
            If Request.Form("notetext") <> "" Then
                IsPrivate = IIf(Request.Form("private") = "1", 1, 0)
                IsImportant = IIf(Request.Form("important") = "1", 1, 0)
                If NoteTypeID = 4 Then
                    MyCommon.QueryStr = "dbo.pt_CustomerNotes_Insert"
                    MyCommon.Open_LXSsp()
                    MyCommon.LXSsp.Parameters.Add("@NoteTypeID", System.Data.SqlDbType.Int).Value = NoteTypeID
                    MyCommon.LXSsp.Parameters.Add("@LinkID", System.Data.SqlDbType.Int).Value = LinkID
                    MyCommon.LXSsp.Parameters.Add("@AdminUserID", System.Data.SqlDbType.Int).Value = AdminUserID
                    MyCommon.LXSsp.Parameters.Add("@FirstName", System.Data.SqlDbType.NVarChar, 50).Value = FirstName
                    MyCommon.LXSsp.Parameters.Add("@LastName", System.Data.SqlDbType.NVarChar, 50).Value = LastName
                    MyCommon.LXSsp.Parameters.Add("@Note", System.Data.SqlDbType.NVarChar, 1000).Value = Server.HtmlEncode(Request.Form("notetext"))
                    MyCommon.LXSsp.Parameters.Add("@Private", System.Data.SqlDbType.Bit).Value = IIf(IsPrivate, 1, 0)
                    MyCommon.LXSsp.Parameters.Add("@Important", System.Data.SqlDbType.Bit).Value = IIf(IsImportant, 1, 0)
                    MyCommon.LXSsp.Parameters.Add("@Deleted", System.Data.SqlDbType.Bit).Value = 0
                    MyCommon.LXSsp.Parameters.Add("@LanguageID", System.Data.SqlDbType.Int).Value = LanguageID
                    MyCommon.LXSsp.Parameters.Add("@NoteID", System.Data.SqlDbType.Int).Direction = System.Data.ParameterDirection.Output
                    MyCommon.LXSsp.ExecuteNonQuery()
                    NoteID = MyCommon.LXSsp.Parameters("@NoteID").Value
                Else
                    MyCommon.QueryStr = "dbo.pt_Notes_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@NoteTypeID", System.Data.SqlDbType.Int).Value = NoteTypeID
                    MyCommon.LRTsp.Parameters.Add("@LinkID", System.Data.SqlDbType.Int).Value = LinkID
                    MyCommon.LRTsp.Parameters.Add("@AdminUserID", System.Data.SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@FirstName", System.Data.SqlDbType.NVarChar, 50).Value = FirstName
                    MyCommon.LRTsp.Parameters.Add("@LastName", System.Data.SqlDbType.NVarChar, 50).Value = LastName
                    MyCommon.LRTsp.Parameters.Add("@Note", System.Data.SqlDbType.NVarChar, 1000).Value = Server.HtmlEncode(Request.Form("notetext"))
                    MyCommon.LRTsp.Parameters.Add("@Private", System.Data.SqlDbType.Bit).Value = IIf(IsPrivate, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@Important", System.Data.SqlDbType.Bit).Value = IIf(IsImportant, 1, 0)
                    MyCommon.LRTsp.Parameters.Add("@Deleted", System.Data.SqlDbType.Bit).Value = 0
                    MyCommon.LRTsp.Parameters.Add("@LanguageID", System.Data.SqlDbType.Int).Value = LanguageID
                    MyCommon.LRTsp.Parameters.Add("@NoteID", System.Data.SqlDbType.Int).Direction = System.Data.ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    NoteID = MyCommon.LRTsp.Parameters("@NoteID").Value
                End If

                If Not IsPrivate Then
                    HistoryString = Copient.PhraseLib.Lookup("history.note-add", LanguageID)
                    If Not IsDBNull(rst.Rows(0).Item("PhraseID")) Then
                        HistoryString = HistoryString & " " & StrConv(Copient.PhraseLib.Lookup("term.to", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup(rst.Rows(0).Item("PhraseID"), LanguageID), VbStrConv.Lowercase)
                    End If
                    If LinkID = 0 Then
                        Select Case NoteTypeID
                            Case 3, 5, 7, 8, 9, 10, 11, 12, 13, 14, 16, 18, 19, 20, 22, 26
                                HistoryString = HistoryString & " " & StrConv(Copient.PhraseLib.Lookup("term.list", LanguageID), VbStrConv.Lowercase)
                        End Select
                    End If
                    MyCommon.Activity_Log(ActivityTypeID, LinkID, AdminUserID, HistoryString)
                End If
            End If
        End If

        If Request.Form("notedelete") <> "" Then
            If NoteTypeID = 4 Then
                MyCommon.QueryStr = "dbo.pt_CustomerNotes_Delete"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@NoteID", System.Data.SqlDbType.Int).Value = Request.Form("NoteID")
                MyCommon.LXSsp.ExecuteNonQuery()
            Else
                MyCommon.QueryStr = "dbo.pt_Notes_Delete"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@NoteID", System.Data.SqlDbType.Int).Value = Request.Form("NoteID")
                MyCommon.LRTsp.ExecuteNonQuery()
            End If
            If Not IsPrivate Then
                HistoryString = Copient.PhraseLib.Lookup("history.note-delete", LanguageID)
                If Not IsDBNull(rst.Rows(0).Item("PhraseID")) Then
                    HistoryString = HistoryString & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup(rst.Rows(0).Item("PhraseID"), LanguageID), VbStrConv.Lowercase)
                End If
                If LinkID = 0 Then
                    Select Case NoteTypeID
                        Case 3, 5, 7, 8, 9, 10, 11, 12, 13, 14, 16, 18, 19, 20, 22, 26
                            HistoryString = HistoryString & " " & StrConv(Copient.PhraseLib.Lookup("term.list", LanguageID), VbStrConv.Lowercase)
                    End Select
                End If
                MyCommon.Activity_Log(ActivityTypeID, LinkID, AdminUserID, HistoryString)
            End If
        End If

        If NoteTypeID = 4 Then
            MyCommon.QueryStr = "select N.NoteID, 4 as NoteTypeID, N.CustomerPK as LinkID, N.AdminUserID, N.FirstName, N.LastName, " &
                                "N.CreatedDate, N.Note, N.Private, N.Important, N.Deleted, 25 as ActivityTypeID from CustomerNotes as N with (NoLock) " &
                                "where CustomerPK=" & LinkID & " and Deleted=0 order by " & NoteSortText & " " & NoteSortDirection & ";"
            rst = MyCommon.LXS_Select
        Else
            MyCommon.QueryStr = "select N.NoteID, N.NoteTypeID, N.LinkID, N.AdminUserID, N.FirstName, N.LastName, N.CreatedDate, N.Note, N.Private, " &
                                "N.Important, N.Deleted, NT.ActivityTypeID from Notes as N with (NoLock) " &
                                "inner join NoteTypes as NT on NT.NoteTypeID=N.NoteTypeID " &
                                "where N.NoteTypeID=" & NoteTypeID & " and LinkID=" & LinkID & " and Deleted=0 order by " & NoteSortText & " " & NoteSortDirection & ";"
            rst = MyCommon.LRT_Select
        End If

        Send("<div id=""notes""" & IIf(NoteMode, "", " style=""display:none;""") & ">")
        Send("  <div id=""notesbody"">")
        Send("    <form id=""notesform"" name=""notesform"" action=""#"" method=""post"">")
        Send("      <input type=""hidden"" name=""notedelete"" />")
        Send("      <input type=""hidden"" name=""noteID"" />")
        Send("      <input type=""hidden"" name=""notemode"" />")
        Send("      <span id=""notestitle"">" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & "</span>")
        Send("      <select id=""notesort"" name=""notesort"" onchange=""document.notesform.submit()"">")
        Send("        <option value=""1""" & IIf(NoteSortID = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("notes.SortByDateDesc", LanguageID) & "</option>")
        Send("        <option value=""2""" & IIf(NoteSortID = 2, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("notes.SortByDateAsc", LanguageID) & "</option>")
        Send("        <option value=""3""" & IIf(NoteSortID = 3, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("notes.SortByUserDesc", LanguageID) & "</option>")
        Send("        <option value=""4""" & IIf(NoteSortID = 4, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("notes.SortByUserAsc", LanguageID) & "</option>")
        Send("      </select>")
        Send("      <span id=""notesclose"" title=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """><a href=""javascript:toggleNotes()"">x</a></span>")
        Send("")
        Send("      <div id=""notesdisplay"">")
        Send("        <div id=""notesscroll"">")
        If (rst.Rows.Count > 0) Then
            For Each row In rst.Rows
                If ((row.Item("Private") = False) OrElse ((row.Item("Private") = True AndAlso row.Item("AdminUserID") = AdminUserID))) Then
                    MyCommon.QueryStr = "select UserName, FirstName, LastName, Employer from AdminUsers where AdminUserID=" & row.Item("AdminUserID") & ";"
                    rst2 = MyCommon.LRT_Select
                    Sendb("          <div class=""note")
                    If row.Item("Private") Then Sendb(" private")
                    If row.Item("Important") Then Sendb(" important")
                    Send(""" id=""note" & row.Item("NoteID") & """>")
                    Send("            <a name=""n" & row.Item("NoteID") & """></a>")
                    Send("            <span class=""notedate"">" & Logix.ToShortDateTimeString(MyCommon.NZ(row.Item("CreatedDate"), ""), MyCommon) & "</span>")
                    If rst2.Rows.Count = 0 Then
                        Sendb("            <span class=""noteuser grey"">")
                    Else
                        Sendb("            <span class=""noteuser"">")
                    End If
                    If NoteTypeID = 4 Then
                        Send(MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst2.Rows(0).Item("LastName"), ""), 25) & "</span>")
                    Else
                        Send(MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("LastName"), ""), 25) & "</span>")
                    End If
                    If (Logix.UserRoles.DeleteNotes) Then
                        Send("            <span class=""notedelete"">[<a href=""javascript:deleteNote('" & row.Item("NoteID") & "')"" title=""" & Copient.PhraseLib.Lookup("notes.deletethisnote", LanguageID) & """>X</a>]</span>")
                    End If
                    Send("            <br />")
                    Send("            " & MyCommon.SplitNonSpacedString(row.Item("Note").Replace(vbCrLf, "<br />"), 25))
                    Send("          </div>")
                End If
            Next
        Else
            Send("          <div class=""note"" id=""note"">")
            Send("            <span style=""color:#808050;"">" & Copient.PhraseLib.Lookup("notes.none", LanguageID) & "</span><br />")
            Send("          </div>")
        End If
        Send("        </div>")
        If (Logix.UserRoles.CreateNotes) Then
            Send("        <div id=""noteadddiv"" style=""text-align:center;"">")
            Send("          <input type=""button"" class=""regular"" id=""noteadd"" name=""noteadd"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""toggleNotesInput();"" /><br />")
            Send("        </div>")
        End If
        Send("      </div>")
        If (Logix.UserRoles.CreateNotes) Then
            Send("")
            Send("      <div id=""notesinput"" style=""display:none;"">")
            Send("        <textarea id=""notetext"" name=""notetext""></textarea>")
            Send("        <br />")
            Send("        <input type=""checkbox"" id=""private"" name=""private"" value=""1"" /><label for=""private"">" & Copient.PhraseLib.Lookup("term.private", LanguageID) & "</label>")
            Send("        <div id=""hider"" style=""display:none;"">")
            Send("          <input type=""checkbox"" id=""important"" name=""important"" value=""1"" /><label for=""important"">" & Copient.PhraseLib.Lookup("term.important", LanguageID) & "</label>")
            Send("        </div>")
            Send("        <br />")
            Send("        <br class=""half"" />")
            Send("        <input type=""submit"" class=""regular"" id=""notesave"" name=""notesave"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
            Send("        <input type=""button"" class=""regular"" id=""notecancel"" name=""notecancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""toggleNotesInput();"" /><br />")
            Send("      </div>")
        End If
        Send("    </form>")
        Send("  </div>")
        Send("  <div id=""notesshadow"">")
        Send("    <img src=""/images/notesshadow.png"" alt="""" />")
        Send("  </div>")
        If Request.Browser.Type = "IE6" Then
            Send("<iframe src=""javascript:'';"" id=""notesiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no""></iframe>")
        End If
        Send("</div>")
        Send("")
        MyCommon.Close_LogixRT()
        MyCommon.Close_LogixXS()
        MyCommon = Nothing
    End Sub

    Public Sub Send_BodyEnd(Optional ByVal FormName As String = "", Optional ByVal ElementName As String = "")
        Send_FocusScript(FormName, ElementName)
        Send_WrapEnd()
        Send_PageEnd()
    End Sub

    Public Sub Send_FocusScript(Optional ByVal FormName As String = "", Optional ByVal ElementName As String = "")
        If (FormName <> "" AndAlso ElementName <> "") Then
            Send("<script type=""text/javascript"" language=""javascript"">")
            Send("   var frmElem = document.forms['" & FormName & "'];")
            Send("   if (frmElem != null) { ")
            Send("     var elem = frmElem." & ElementName & ";")
            Send("     if (elem != null && elem.disabled == false) {")
            Send("       if(window.dialogArguments != null) {") 'In Case of Modal Dialog
            Send("         setTimeout (""document.forms['" & FormName & "']." & ElementName & ".focus();"", 50 );")
            Send("       }")
            Send("       else {")
            Send("        elem.focus();")
            Send("       }")
            Send("     }")
            Send("   }")
            Send("</script>")
        End If
    End Sub

    Public Sub Send_WrapEnd()
        Send("<a id=""bottom"" name=""bottom""></a>")
        Send("<div id=""footer"">")
        Send("  " & Copient.PhraseLib.Lookup("about.copyright", LanguageID))
        Send("</div>")
        Send("<div id=""custom3""></div>")
        Send("</div> <!-- End wrap -->")
        Send("<div id=""custom4""></div>")
    End Sub

    Public Sub Send_PageEnd()
        Send("</body>")
        Send("</html>")
    End Sub

    Public Sub Send_SV_Propagate(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""SaveProp"" name=""SaveProp"" value=""" & Copient.PhraseLib.Lookup("term.saveprop-sv", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_SV_Deploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""SavePropDeploy"" name=""SavePropDeploy"" value=""" & Copient.PhraseLib.Lookup("term.savepropdeploy-sv", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Date_Picker_Terms()
        Send("dayArray = new Array('" & Left(Copient.PhraseLib.Lookup("term.sunday", LanguageID), 2) & "', " &
             "'" & Left(Copient.PhraseLib.Lookup("term.monday", LanguageID), 2) & "', " &
             "'" & Left(Copient.PhraseLib.Lookup("term.tuesday", LanguageID), 2) & "', " &
             "'" & Left(Copient.PhraseLib.Lookup("term.wednesday", LanguageID), 2) & "', " &
             "'" & Left(Copient.PhraseLib.Lookup("term.thursday", LanguageID), 2) & "',  " &
             "'" & Left(Copient.PhraseLib.Lookup("term.friday", LanguageID), 2) & "', " &
             "'" & Left(Copient.PhraseLib.Lookup("term.saturday", LanguageID), 2) & "'); ")
        Send("monthArray = new Array('" & Copient.PhraseLib.Lookup("term.january", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.february", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.march", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.april", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.may", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.june", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.july", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.august", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.september", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.october", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.november", LanguageID) & "', " &
             "'" & Copient.PhraseLib.Lookup("term.december", LanguageID) & "');")

        Send("calendarPhrase = '" & Copient.PhraseLib.Lookup("term.calendar", LanguageID) & "';")
        Send("todayPhrase = '" & Copient.PhraseLib.Lookup("term.today", LanguageID) & "';")
    End Sub

    Public Function DetectHandheld(ByRef MobileDevice As Boolean, ByRef Platform As String, ByRef UserAgent As String) As Boolean
        Dim Handheld As Boolean = False

        If MobileDevice = True Then
            Handheld = True
        ElseIf Platform.IndexOf("WinCE") > -1 Or Platform.IndexOf("Palm") > -1 Or Platform.IndexOf("Pocket") > -1 Then
            Handheld = True
        ElseIf UserAgent.IndexOf("iPhone") > -1 Then
            Handheld = True
        Else
            Handheld = False
        End If

        Return Handheld
    End Function

    ' Takes in the attribute string and looks for the presence of AttribToParse token (e.g. onclick)
    ' then the function will cut out the name and only return the value of the token putting a semicolons
    ' where necessary as defined by the SemicolonStart and SemicolonEnd parameters

    Private Function ParseAttribute(ByRef Attributes As String, ByVal AttribToParse As String, ByVal SemicolonStart As Boolean, ByVal SemicolonEnd As Boolean) As String
        Dim OnClickPos As Integer = -1
        Dim Attrib As String = ""
        Dim QuoteStart, QuoteEnd As Integer

        If (Attributes <> "") Then
            OnClickPos = Attributes.IndexOf(AttribToParse, 0, StringComparison.OrdinalIgnoreCase)
            If (OnClickPos > -1) Then
                QuoteStart = Attributes.IndexOf("""", OnClickPos)
                QuoteEnd = Attributes.IndexOf("""", QuoteStart + 1)
                If (QuoteEnd - (QuoteStart + 1) > 0) Then
                    If (SemicolonStart AndAlso Left(Attrib, 1) <> ";") Then Attrib &= ";"
                    Attrib = Attributes.Substring(QuoteStart + 1, QuoteEnd - (QuoteStart + 1))
                    If (SemicolonEnd AndAlso Right(Attrib, 1) <> ";") Then Attrib &= ";"

                    ' now remove it from the attributes string
                    If ((QuoteEnd + 1) - OnClickPos > 0) Then
                        Attributes = Attributes.Remove(OnClickPos, (QuoteEnd + 1) - OnClickPos).Trim
                    End If
                End If
            End If
        End If

        Return Attrib
    End Function
    Public Function TruncateWordAppendEllipsis(ByVal input As String, ByVal length As Integer) As String
        If input Is Nothing OrElse input.Length < length Then
            Return input
        End If

        Dim iNextSpace = input.LastIndexOf(" ", length)
        Return String.Format("{0}...", input.Substring(0, (IIf((iNextSpace > 0), iNextSpace, length))).Trim())

    End Function
    Sub Send_ExportToExcel(Optional ByVal Attributes As String = "")
        Dim sExcelUrl As String
        Dim i As Integer
        Dim sId As String
        Dim sValue As String

        If Request.QueryString.Count > 0 Then
            If Request.QueryString("excel") = "" Then
                sExcelUrl = Request.RawUrl & "&excel=Excel"
            Else
                sExcelUrl = Request.RawUrl
            End If
        Else
            sExcelUrl = Request.RawUrl & "?excel=Excel"
        End If
        Send("<form id=""excelform"" name=""excelform"" action=""#"">")
        ' preserve the filtering and sorting of current database query
        Send("<input type=""hidden"" id=""ExcelUrl"" name=""ExcelUrl"" value=""" & sExcelUrl & """ />")
        For i = 0 To Request.Form.Keys.Count - 1
            sId = Request.Form.Keys.Item(i)
            sValue = Request.Form(sId)
            Send("<input type=""hidden"" id=""" & sId & """ name=""" & sId & """ value=""" & sValue & """ />")
        Next
        Sendb("<input type=""button"" class=""regular"" id=""excel"" name=""excel"" value=""" & Copient.PhraseLib.Lookup("offer-list.export", LanguageID) & """ onclick=""handleExcel();"" " & Attributes & " />")
        Send("</form>")
    End Sub

    Public Function GetCgiValue(ByVal VarName As String) As String
        Dim TempVal As String
        TempVal = ""
        TempVal = Request.QueryString(VarName)
        If TempVal = "" Then TempVal = Request.Form(VarName)
        GetCgiValue = TempVal
    End Function

    Public Function IsBrowserIE() As Boolean
        Return (Request.Browser.Browser = "IE")
    End Function

    Public Sub Close_UI_Box()

        Send("</div> <!-- closing UI box body -->")
        If IsBrowserIE() Then Send("<br clear=""all"" />")
        Send("</div> <!-- closing UI box -->")
        If IsBrowserIE() Then Send("<br clear=""all"" />")

    End Sub

    Public Sub Open_UI_Box(ByVal BoxID As Integer, ByVal AdminUserID As Integer, ByVal Common As Copient.CommonInc, Optional ByVal ExtraTitleText As String = "", Optional ByVal BoxWidth As String = "")

        Dim dst As DataTable
        Dim BoxTitle As String = ""
        Dim BoxObjectName As String = ""
        Dim BoxOpen As Integer = 1
        Dim ValidBoxID As Boolean = False
        Dim WidthStr As String = ""
        Dim ServerHealth As Boolean = False
        Common.QueryStr = "select OptionValue from dbo.UE_SystemOptions where OptionID=91"
        dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
        ServerHealth = (Common.Fetch_UE_SystemOption(91) = "1")
        Common.QueryStr = "select isnull(BoxTitle, '') as BoxTitle, isnull(BoxTitlePhraseTerm, '') as BoxTitlePhraseTerm, isnull(BoxObjectName, '') as BoxObjectName from UIBoxes where BoxID=" & BoxID & ";"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            BoxTitle = Copient.PhraseLib.Lookup(dst.Rows(0).Item("BoxTitlePhraseTerm"), LanguageID)
            BoxObjectName = dst.Rows(0).Item("BoxObjectName")
            ValidBoxID = True
        End If
        If ValidBoxID Then
            Common.QueryStr = "select isnull(BoxOpen, 1) as BoxOpen from AdminUserBoxStates where BoxID=" & BoxID & " and AdminUserID=" & AdminUserID & ";"
            dst = Common.LRT_Select
            If dst.Rows.Count > 0 Then
                If Not (dst.Rows(0).Item("BoxOpen")) Then
                    BoxOpen = 0
                End If
            End If
            If Not (BoxWidth = "") Then WidthStr = "style=""width: " & BoxWidth & ";"""
            Send("<div class=""box"" id=""" & BoxObjectName & "box"" " & WidthStr & ">")
            Send("  <div style=""position: relative;float: left;""><font size=""3""><b>" & IIf(ServerHealth, IIf(BoxTitle = Copient.PhraseLib.Lookup("term.storehealth", LanguageID), Copient.PhraseLib.Lookup("term.serverhealth", LanguageID), BoxTitle), BoxTitle) & "</b></font>" & ExtraTitleText & "</div>")
            Send("  <div class=""resizer"" style=""position: relative;"">")
            Send("    <a href=""#"" onclick=""resizeBox('" & BoxObjectName & "body','img" & BoxObjectName & "body','" & BoxTitle & "', '" & BoxID & "', '" & AdminUserID & "'); return false;"">")
            If BoxOpen = 1 Then
                Send("    <img id=""img" & BoxObjectName & "body"" src=""/images/arrowup-off.png"" alt=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ title=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ onmouseover = ""handleResizeHover(true,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" onmouseout=""handleResizeHover(false,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" />")
            Else
                Send("    <img id=""img" & BoxObjectName & "body"" src=""/images/arrowdown-off.png"" alt=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ title=""" & Copient.PhraseLib.Lookup("term.hide", LanguageID) & " " & BoxTitle & """ onmouseover = ""handleResizeHover(true,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" onmouseout=""handleResizeHover(false,'" & BoxObjectName & "body','img" & BoxObjectName & "body');"" />")
            End If
            Send("    </a>")
            Send("  </div> <!-- resizer -->")
            Send("  <br clear=""all"" />")
            If BoxOpen = 1 Then
                Send("  <div id=""" & BoxObjectName & "body"">")
            Else
                Send("  <div id=""" & BoxObjectName & "body"" style=""display: none;"">")
            End If
        End If  'ValidBoxID

    End Sub

    Public Function UEOffer_Has_AnyCustomer(ByVal Common As Copient.CommonInc, ByVal OfferID As Long) As Boolean
        Dim RetVal As Boolean = False
    Dim IsClosedRT As Boolean = False

    ' ensure everything we need is opened
    If Common.LRTadoConn.State = ConnectionState.Closed Then
      IsClosedRT = True
      Common.Open_LogixRT()
    End If

        Common.QueryStr = "dbo.pa_Check_AnyCustomer_In_Offer"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@AnyCustomerInOffer", SqlDbType.Int).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        If Common.LRTsp.Parameters("@AnyCustomerInOffer").Value = 1 Then RetVal = True
        Common.Close_LRTsp()

    If IsClosedRT Then
      If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
    End If
        Return RetVal

    End Function

    Public Sub Send_Calendar_Overrides(ByRef Common As Copient.CommonInc)
        Dim i As Integer
        Dim dt As DataTable
        Dim UserCulture As System.Globalization.CultureInfo = Nothing
        Dim FirstDayOfWeek As Integer = 0
        Dim days As String() = {"sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"}
        Dim months As String() = {"january", "february", "march", "april", "may", "june", "july", "august",
                                  "september", "october", "november", "december"}

        ' find the last day of the week in the user's language (region)
        Common.QueryStr = "select MSNetCode from Languages with (NoLock) where LanguageID=" & LanguageID
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            UserCulture = New System.Globalization.CultureInfo(Common.NZ(dt.Rows(0).Item("MSNetCode"), "en-US").ToString)
            FirstDayOfWeek = CInt(UserCulture.DateTimeFormat.FirstDayOfWeek)
            Send("firstDayOfWeek=" & FirstDayOfWeek & ";")
            Send("lastDayOfWeek = " & ((FirstDayOfWeek + 6) Mod 7) & ";")

            Sendb("dateArray = new Array(")
            For i = 0 To 31
                If i > 0 Then Sendb(",")
                Sendb("'" & Copient.commonShared.TranslateDigits(i, UserCulture) & "'")
            Next
            Send(")")
        End If

        ' send the days of the week in the user's language. start at the first day and wrap around the array to get any days that might remain
        Sendb("dayArray = new Array(")
        For i = 0 To days.GetUpperBound(0)
            If i > 0 Then Sendb(",")
            Sendb("'" & Copient.PhraseLib.Lookup("calendar." & days(i) & "abbreviation", LanguageID) & "'")
        Next
        Send(")")

        ' send the name of the months in the user's language
        Sendb("monthArray = new Array(")
        For i = 0 To months.GetUpperBound(0)
            If i > 0 Then Sendb(",")
            Sendb("'" & Copient.PhraseLib.Lookup("term." & months(i), LanguageID) & "'")
        Next
        Send(")")

        If UserCulture IsNot Nothing Then
            Send("  defaultDateSeparator = '" & UserCulture.DateTimeFormat.DateSeparator & "';")
            Send("  defaultDateFormat = '" & GetLocalizedDateFormat(UserCulture) & "';")
        End If

        Send("calendarPhrase = '" & Copient.PhraseLib.Lookup("term.calendar", LanguageID) & "';")
        Send("todayPhrase = '" & Copient.PhraseLib.Lookup("term.today", LanguageID) & "';")

    End Sub

    Private Function GetLocalizedDateFormat(ByVal UserCulture As System.Globalization.CultureInfo) As String
        Dim DateParts() As String
        Dim DateFormat As String = ""

        If UserCulture IsNot Nothing Then
            DateParts = UserCulture.DateTimeFormat.ShortDatePattern.Split(UserCulture.DateTimeFormat.DateSeparator)
            For Each Part As String In DateParts
                If Part IsNot Nothing AndAlso Part.Length > 0 Then
                    DateFormat &= Left(Part, 1).ToLower
                End If
            Next
            If DateFormat.Length <> 3 Then DateFormat = "mdy"
        End If

        Return DateFormat
    End Function

    Public Function CheckIfValidOffer(ByRef Common As Copient.CommonInc, ByVal OfferId As Long) As Boolean
	If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        Common.QueryStr = "select StatusFlag from CPE_Incentives with (NoLock) where IncentiveId=" & OfferId
        Dim dt As DataTable = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            Dim statusFlag As Int32 = Convert.ToInt32(dt.Rows(0)(0))
            If statusFlag = 2 Or statusFlag = 11 Or statusFlag = 12 Then
                Response.Redirect("/logix/UE/UEoffer-sum.aspx?OfferID=" & OfferId)
            End If
        End If
        Return True
    End Function

    Public Function IsOfferWaitingForApproval(ByVal OfferId As Long) As Boolean
        CMS.AMS.CurrentRequest.Resolver.AppName = "UE-CB.vb"
        Dim m_OAWService As CMS.AMS.Contract.IOfferApprovalWorkflowService = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOfferApprovalWorkflowService)()
        Dim m_IsOfferAwaitingApproval = m_OAWService.CheckIfOfferIsAwaitingApproval(OfferId).Result
        Return m_IsOfferAwaitingApproval
    End Function

    Public Sub ResetOfferApprovalStatus(ByVal OfferId As Long)
        CMS.AMS.CurrentRequest.Resolver.AppName = "UE-CB.vb"
        Dim m_OAWService As CMS.AMS.Contract.IOfferApprovalWorkflowService = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOfferApprovalWorkflowService)()
        m_OAWService.ResetOfferApprovalStatus(OfferId)
    End Sub
End Class