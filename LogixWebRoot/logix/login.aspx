<%@ Page Language="vb" Debug="true" CodeFile="logixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>

<%@ Import Namespace="CMS.AMS" %>
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
    Public DefaultLanguageID
    Public bExternalLoginEnabled As Boolean = False
    Public MyCommon As New Copient.CommonInc
    Public Logix As New Copient.LogixInc
    Public Const icExternalsourceId As Integer = 1
    'Private TempAdminUserID As Long = -1
    Public Common As New Copient.CommonInc
    Dim MyCryptlib As New CMS.CryptLib
    Dim HashLib As New CMS.HashLib.CryptLib
    Dim UserName As String
    Dim Password As String
    Dim userID As Long
    Dim emailId As String
    Dim isOTPValid As Boolean = True
    Dim isOTPResend As Boolean = False
    Dim isEmailValid As Boolean = True
    Dim strOTPErrormsg As String = ""
    Dim SecondFactorEnabled As Boolean = IIf(Common.Fetch_SystemOption(298) = "0", False, True)
    Dim SecondFactorBypassEnabled As Boolean = IIf(Common.Fetch_SystemOption(301) = "0", False, True)
    Dim BypassAuthorizedForCurrentUser As Boolean = False
    Dim bypass2FactorCookieName As String = String.Empty
    Dim otpHelper As IOTPHelper
    Sub Verify_User()
        Dim AuthToken As String = System.Web.HttpContext.Current.Session.SessionID
        Dim BounceBack As String
        Dim dt As System.Data.DataTable
        BounceBack = HttpUtility.UrlDecode(Request.Form("bounceback"))
        GetUserCredentials()
        Dim IsPasswordInvalid As Integer = 1
        Dim ErrorType As Integer
        Dim errorCode As String = ""

        CurrentRequest.Resolver.AppName = "login.aspx"
        Dim m_adminUserDataService As IAdminUserData = CurrentRequest.Resolver.Resolve(Of IAdminUserData)()

        If UserName <> "" And Password <> "" And (m_adminUserDataService.ValidatePassword(Password, UserName).ResultType = CMS.AMS.Models.AMSResultType.Success) Then
            Logix.Verify_AdminUser(MyCommon, UserName, Password, userID, AuthToken, emailId, errorCode)
            If userID = 0 Then 'Username/password are invalid - let them try again
                IsPasswordInvalid = 0
            ElseIf (m_adminUserDataService.ValidatePasswordExpiry(UserName).ResultType = CMS.AMS.Models.AMSResultType.Success) Then
                If (Not bExternalLoginEnabled AndAlso SecondFactorEnabled) Then
                    Session("UserID") = userID
                    Session("UserName") = UserName
                    Session("AuthToken") = AuthToken
                    Session("MailAddress") = emailId
                End If
                IsPasswordInvalid = 0
            Else
                IsPasswordInvalid = 1 'Password is expire
                ErrorType = PasswordValidation.PasswordExpire
            End If

        Else
            dt = m_adminUserDataService.GetAdminUserIDbyUserName(UserName, Password)
            If (UserName = "" Or Password = "" Or dt.Rows.Count = 0) Then
                IsPasswordInvalid = 0  
                'Send_Login_Page(Copient.PhraseLib.Lookup("login.invalidlogin", DefaultLanguageID))  //Commenting as this is being sent in the below code when AdminUserID = 0
            Else
                ErrorType = PasswordValidation.InvalidPassword
            End If
        End If

        'This Code is executed when password expire or invalid password
        If (IsPasswordInvalid = 1) Then
            Dim url = "ChangePassword.aspx"

            Response.Clear()
            Dim sb = New System.Text.StringBuilder()
            sb.Append("<html>")
            sb.AppendFormat("<body onload='document.forms[0].submit()'>")
            sb.AppendFormat("<form action='{0}' method='post'>", url)
            sb.AppendFormat("<input type='hidden' name='userName' enableviewstate='true' value='{0}'>", UserName)
            sb.AppendFormat("<input type='hidden' name='message' enableviewstate='true' value='{0}'>", ErrorType)
            sb.Append("</form>")
            sb.Append("</body>")
            sb.Append("</html>")
            Response.Write(sb.ToString())
            Response.[End]()
        End If


        Dim infoMessage As String = ""

        If userID = -1 Then
            'Username/password are expired (specific to SunOne)
            Dim redirectURL = MyCommon.Fetch_SystemOption(87)
            If Not redirectURL.Equals("") Then
                LogAttempt(UserName, Now(), False)
                Response.Redirect(redirectURL)
            Else
                Send_Login_Page(Copient.PhraseLib.Lookup(3790, DefaultLanguageID, Copient.PhraseLib.Lookup("logix.expiredpassword", LanguageID)))
                LogAttempt(UserName, Now(), False)
            End If
        ElseIf userID = 0 Then
            'Username/password are invalid - let them try again
            Dim aErrorCode As String() = errorCode.Split(";")
            If (aErrorCode.Length = 1) Then
                Send_Login_Page(Copient.PhraseLib.Lookup(errorCode, DefaultLanguageID, errorCode))
            ElseIf (aErrorCode.Length = 2) Then
                Dim errPhrase = Copient.PhraseLib.Lookup(aErrorCode(0), DefaultLanguageID, aErrorCode(0))
                Send_Login_Page(errPhrase.Replace("{0}", aErrorCode(1)))
            End If
            LogAttempt(UserName, Now(), False)
        Else
            bypass2FactorCookieName = String.Concat("SecondFactorAuthentication", userID)
            If Not bExternalLoginEnabled AndAlso SecondFactorEnabled AndAlso SecondFactorBypassEnabled = True AndAlso Request.Cookies(bypass2FactorCookieName) IsNot Nothing Then
                BypassAuthorizedForCurrentUser = (MyCryptlib.SQL_StringDecrypt(Request.Cookies(bypass2FactorCookieName)("user")) = UserName)
            End If
            If Not bExternalLoginEnabled AndAlso SecondFactorEnabled AndAlso (SecondFactorBypassEnabled = False OrElse (SecondFactorBypassEnabled = True AndAlso BypassAuthorizedForCurrentUser = False)) Then
                'If external login is enabled, then expectation is some other application will be providing authentication for logix user so excluding that
                If (String.IsNullOrEmpty(emailId)) Then
                    Send_Login_Page(String.Empty, True, True)
                Else
                    Send_Login_Page(String.Empty, True, False, True)
                End If
            Else
                AllowAccess(userID, BounceBack, AuthToken)
            End If
        End If
    End Sub
    Sub GetUserCredentials()
        UserName = Request.Form("username")
        Password = Request.Form("password")
    End Sub
    Sub AllowAccess(ByVal AdminUserID As Integer, BounceBack As String, AuthToken As String)
        Dim TargetURL As String
        Dim dst As System.Data.DataTable

        MyCommon.QueryStr = "update AdminUsers with (RowLock) set LastLoginExternal=0,LastLogin=getdate() where AdminUserID = @AdminUserID"
        MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        LogAttempt(UserName, Now(), True)

        'User passed verification - send them to where ever they need to go
        If BounceBack = "" Then
            MyCommon.Activity_Log(1, AdminUserID, AdminUserID, Copient.PhraseLib.Lookup("term.loggedin", DefaultLanguageID))
            MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName from AdminUsers as AU with (NoLock) " & _
                       "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                       "where AU.AdminUserID= @AdminUserID"
            MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
            dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            TargetURL = dst.Rows(0).Item("PageName")
        Else
            Dim HostURI As String = HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Authority)
            If Not BounceBack.StartsWith(HostURI) Then
                If Not ValidateRelativePathURL(BounceBack) Then
                    TargetURL = "error-forbidden.aspx"
                Else
                    TargetURL = GetBounceBack(BounceBack)
                End If
            Else                
                TargetURL = GetBounceBack(BounceBack)
            End If
        End If
        Response.CacheControl = "no-cache"
        Response.Cookies("AuthToken").Value = AuthToken
        LogBrowser()
        Send("<html xmlns=""http://www.w3.org/1999/xhtml"">")
        Send("<head>")
        Send("<title>" & Copient.PhraseLib.Lookup("term.logix", DefaultLanguageID) & "</title>")
        Send("<meta http-equiv=""Refresh"" content=""0; URL=" & TargetURL & """>")
        Send("</head>")
        Send("<body bgcolor=""#ffffff"">")
        Send("<!-- bouncing -->")
        Send("</body>")
        Send("</html>")
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

    Dim LogText As String = String.Format("{0:" + Common.DateFormat + "}", DateTime.Now)

    'term.otpemailmessage
    'term.otpsentmessage,  term.resendotp, term.otp
    Sub Send_Login_Page(Optional ByVal infoMessage As String = "", Optional ByVal EnableOTP As Boolean = False, Optional ByVal EnableEmailPopUp As Boolean = False, Optional ByVal EnableOTPPopup As Boolean = False)
        Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
        Dim CopientFileVersion As String = "7.3.1.138972"
        Dim CopientProject As String = "Copient Logix"
        Dim CopientNotes As String = ""
        Dim BounceBack As String
        Dim Mode As String
        Dim Logix As New Copient.LogixInc
        Dim Handheld As Boolean = False
        Dim maskedText As String = String.Empty

        If Not String.IsNullOrEmpty(Password) Then
            maskedText = Regex.Replace(Password, ".", "•")
        End If

        If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
            Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
        End If

        Mode = Request.QueryString("mode")
        If Mode = "invalid" Then infoMessage = Copient.PhraseLib.Lookup("login.invalidtimeout", DefaultLanguageID)

        BounceBack = Request.Form("bounceback")
        If String.IsNullOrWhiteSpace(BounceBack) Then BounceBack = Request.QueryString("bounceback")
        BounceBack = Trim(BounceBack)
        If Not (String.IsNullOrWhiteSpace(BounceBack)) Then BounceBack = GetBounceBack(BounceBack)
        'Response.Cookies("CopientLogix")("AuthToken") = ""
        Response.CacheControl = "no-cache"
        If (Not EnableOTPPopup AndAlso Not EnableEmailPopUp) Then
            DeleteAuthTokenCookie()
        End If
        'Send("<!-- DL=" & DefaultLanguageID & " -->")
        Send_HeadBegin(, , , DefaultLanguageID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas(1)
        Send_Links(Handheld)
        If LanguageID = 0 Then LanguageID = DefaultLanguageID 'Send_Scripts() which calls Send_JavaScript_Terms() needs LanguageID set since it does not use DefaultLanguageID.  LanguageID was not getting set yet since the application is still in start-up.  Setting it to DefaultLanguageID is simplest since Send_Scripts() is called from everywhere and LanguageID is properly set every other time.
        Send_Scripts()
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Logos(DefaultLanguageID)
        Send("<div class=""tabs"" id=""tabs"">&nbsp;")
        Send("<hr class=""hidden"" />")
        Send("</div>")
        Send("")
        Send_Subtabs(Logix, 0, 1, DefaultLanguageID)
        Send("")
        Send("<div id=""intro""></div>")
        Send("")
        Send("<div id=""main"">")
        If Not (infoMessage = "") Then Response.Write("  <div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
        Send("")
        Send("<div id=""column"">")
        Send("<br />")
        Send("")
        Send("<div style=""float:left; width:125px;"">")
        Send("<form action=""login.aspx"" id=""mainform"" name=""mainform"" method=""post"">")
        Send("  <input type=""hidden"" id=""mode"" name=""mode"" value=""verify"" />")
        Send("  <input type=""hidden"" id=""bounceback"" name=""bounceback"" value=""" & GetBounceBack(BounceBack) & """ />")
        Send("  <label for=""username"">" & Copient.PhraseLib.Lookup("term.username", DefaultLanguageID) & "&nbsp;</label><br />")
        Send("  <input type=""text"" id=""username"" name=""username"" style=""width:108px;"" tabindex=""1"" maxlength=""50"" value=""" & UserName & """ /><br />")
        Send("  <label for=""password"">" & Copient.PhraseLib.Lookup("term.password", DefaultLanguageID) & "&nbsp;</label><br />")
        Send("  <input type=""password"" id=""password"" name=""password"" style=""width:108px;"" tabindex=""2"" maxlength=""50"" value=""" & maskedText & """ /><br />")
        Send("  <br class=""half"" />")
        Send("  <input type=""submit"" class=""regular"" id=""btnSubmit"" name=""btnSubmit"" tabindex=""3"" value=""" & Copient.PhraseLib.Lookup("term.login", DefaultLanguageID) & """ /><br />")
        If (EnableOTP) Then
            SendOTPHtml()
        End If
        Send("</form>")
        Send("</div>")
        Send("")
        Send("<div style=""float: left; width: 360px;"">")
        Send(Copient.PhraseLib.Lookup("login.note1", DefaultLanguageID) & " ")
        Send("<a href=""mailto:" & MyCommon.Fetch_SystemOption(40) & """>" & MyCommon.Fetch_SystemOption(40) & "</a>.<br />")
        Send("<br class=""half"" />")
        Send(Copient.PhraseLib.Lookup("login.note2", DefaultLanguageID) & "<br />")
        Send("<br class=""half"" />")
        Send(Copient.PhraseLib.Lookup("login.note3", DefaultLanguageID) & "<br />")
        Send("</div>")
        Send("")
        Send("</div><br clear=""all"" />")
        Send("</div>")
        Send("")
        Send("<script type=""text/javascript"">")
        Send(" document.mainform.username.focus();")
        Send(" window.name = ""main"";")
        If (EnableOTP) Then
            SendOTPScript(EnableEmailPopUp, EnableOTPPopup)
        End If
        Send("</" & "script>")
        Send_BodyEnd()
        Logix = Nothing
    End Sub
    Sub SendOTPScript(EnableEmailPopUp As Boolean, EnableOTPPopup As Boolean)
        Send("var okText=""" & Copient.PhraseLib.Lookup("term.okay", DefaultLanguageID) & """;")
        Send("var cancelText=""" & Copient.PhraseLib.Lookup("term.cancel", DefaultLanguageID) & """;")
        Send("var continueText=""" & Copient.PhraseLib.Lookup("term.continue", DefaultLanguageID) & """;")

        Send("function HideEmailMessage(){")
        Send("$('#spanEmailValidation').css(""display"", ""none"");")
        Send("}")
        Send("function HideOTPMessage(){")
        Send("$('#spanOTPValidation').css(""display"", ""none"");")
        Send("}")
        Send("function ShowEmailPopUp(){")
        Send("$('#divEmailDialog').css(""display"",""block"");")
        Send("var dialogButtons = {};")
        Send("dialogButtons[okText]=OnEmailPopUpOk;")
        Send("dialogButtons[cancelText]=function(){  $(this).dialog(""close""); }")
        Send("$('#divEmailDialog').dialog({ appendTo:""#mainform"", width: 400, modal:true, close: ClearCredentials, buttons:dialogButtons,draggable: false,resizable: false});")
        Send("$('#intro').css(""z-index"", ""0""); $('#subtabs').css(""z-index"", ""0"");")
        Send("}")
        Send("function OnEmailPopUpOk(){")
        Send("$('#mode').val('OTPEmail'); $(""#mainform"").submit();")
        Send("}")
        Send("function ShowOTPPopUp(){")
        Send("$('#divOTPDialog').css(""display"",""block"");")
        Send("var dialogButtons = {};")
        Send("dialogButtons[continueText]=OnOTPPopUpOk")
        Send("dialogButtons[cancelText]=function(){  $(this).dialog(""close""); }")
        Send("$('#divOTPDialog').dialog({ appendTo:""#mainform"", width: 400, close: ClearCredentials, modal:true, buttons:dialogButtons,draggable: false,resizable: false});")
        Send("$('#intro').css(""z-index"", ""0""); $('#subtabs').css(""z-index"", ""0"");")
        Send("}")
        Send("function OnOTPPopUpOk(){")
        Send("$('#mode').val('OTP'); $(""#mainform"").submit();")
        Send("}")
        Send("function OnResendOTP(){")
        Send("$('#mode').val('ResendOTP'); $(""#mainform"").submit();")
        Send("}")
        Send("function ClearCredentials(event, ui){$('#mode').val(""verify""); $('#username').val(""""); $('#password').val("""");}")
        If (EnableEmailPopUp) Then
            Send("$(document).ready(function(){")
            Send("$('#divEmailDialog').keypress(function(e){")
            Send("if(e.keyCode == $.ui.keyCode.ENTER){")
            Send("OnEmailPopUpOk(); return false;")
            Send("}});")
            Send("ShowEmailPopUp();")
            If Not isEmailValid Then
                Send("$('#spanEmailValidation').css(""display"", ""block""); $('#spanEmailValidation').html(""" & Copient.PhraseLib.Lookup("emailValidation", DefaultLanguageID) & """);")
            End If
            Send("});")
        End If
        If (EnableOTPPopup) Then
            Send("$(document).ready(function(){")
            Send("$('#divOTPDialog').keypress(function(e){")
            Send("if(e.keyCode == $.ui.keyCode.ENTER){")
            Send("OnOTPPopUpOk();")
            Send("}});")
            Send("ShowOTPPopUp();")
            If Not isOTPValid Then
                Send("$('#spanOTPValidation').css(""display"", ""block""); $('#spanOTPValidation').html(""" & strOTPErrormsg & """);")
            End If
            If isOTPResend Then
                Send("$('#spanOTPValidation').css(""display"", ""block"");$('#spanOTPValidation').css(""color"", ""green""); $('#spanOTPValidation').html(""" & Copient.PhraseLib.Lookup("term.otpresend", DefaultLanguageID) & """);")
            End If
            Send("});")
        End If

    End Sub
    Sub SendOTPHtml()
        Send("<div id=""divOTPDialog"" name=""divOTPDialog"" title=""" & Copient.PhraseLib.Lookup("term.secondstepverification", DefaultLanguageID) & """ style=""display:none;"">")
        Send("<span>" & Copient.PhraseLib.Lookup("term.otpsentmessage", DefaultLanguageID) & "&nbsp;</span><br /><br />")
        Send("<label for=""otpBox"">" & Copient.PhraseLib.Lookup("term.otp", DefaultLanguageID) & ":&nbsp;</label>")
        Send("<input type=""text"" id=""otp"" name=""otp"" style=""width:108px;"" maxlength=""6"" value="""" onkeydown=""HideOTPMessage();""/><br />")
        Send("<span id=""spanOTPValidation"" style=""display:none;color:red;""></span><br />")
        If (SecondFactorBypassEnabled = True) Then
            Send("<input type='checkbox' id='bypassAuthentication' style='vertical-align:middle;' name='bypassAuthentication'>")
            Send("<label for=""bypassAuthentication"">" & Copient.PhraseLib.Lookup("term.dontaskagain", DefaultLanguageID) & "</label><br/><br/>")
        End If
        Send("<a onclick=""OnResendOTP()"" style=""color:blue"">" & Copient.PhraseLib.Lookup("term.resendotp", DefaultLanguageID) & "</a>")
        Send("</div>")
        Send("<div id=""divEmailDialog"" name=""divEmailDialog"" title=""" & Copient.PhraseLib.Lookup("term.emailaddress", DefaultLanguageID) & """ style=""display:none;"">")
        Send("<span>" & Copient.PhraseLib.Lookup("term.otpemailmessage", DefaultLanguageID) & "&nbsp;</span><br /><br />")
        Send("<label for=""otpBox"">" & Copient.PhraseLib.Lookup("term.email", DefaultLanguageID) & ":&nbsp;</label>")
        Send("<input type=""text"" id=""tbMailId"" name=""tbMailId"" style=""width:250px;"" maxlength=""50"" value="""" onkeydown=""HideEmailMessage();"" /><br />")
        Send("<span id=""spanEmailValidation"" style=""display:none;color:red;""></span><br />")
        Send("</div>")
    End Sub
    Function GenerateOTP() As Integer
        otpHelper = CurrentRequest.Resolver.Resolve(Of IOTPHelper)()
        Return otpHelper.GenerateCode(UserName)
    End Function
    Function SendOTP(otp As Integer, userName As String, emailAddress As String) As Boolean
        otpHelper = CurrentRequest.Resolver.Resolve(Of IOTPHelper)()
        Dim result As AMSResult(Of Boolean) = otpHelper.SendOTPNotification(otp, userName, emailAddress, DefaultLanguageID)

        If result.ResultType = AMSResultType.Exception Then
            Send_Login_Page(result.PhraseString, True, False, False)
            Return False
        End If

        Return True
    End Function
    Function ValidateOTP() As Boolean
        Dim otp As Integer = -1
        Integer.TryParse(Request.Form("otp"), otp)
        CurrentRequest.Resolver.AppName = "Login.aspx"
        Dim otpHelper As IOTPHelper = CurrentRequest.Resolver.Resolve(Of IOTPHelper)()
        Dim result As AMSResult(Of Boolean) = otpHelper.ValidateCode(otp, Session("UserName"))
        If result.ResultType = AMSResultType.Success AndAlso otp > 0 AndAlso result.Result = True Then
            isOTPValid = result.Result
            strOTPErrormsg = result.MessageString
            Return True
        Else
            isOTPValid = result.Result
            strOTPErrormsg = result.MessageString
            Return False
        End If
    End Function

    Function ValidateEmail() As Boolean
        If (MyCommon.EmailAddressCheck(Request.Form("tbMailId")) = False) Then
            Return False
        Else
            Return True
        End If
    End Function

    Sub LogAttempt(ByVal UserName As String, ByVal AccessDate As DateTime, ByVal Successful As Boolean, Optional ByVal sMessage As String = "")
        Dim LogFile As String = ""
        Dim IPAddress As String = ""
        Dim Result As String = ""

        LogFile = "LoginLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        IPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If IPAddress = "" Then
            IPAddress = Request.ServerVariables("REMOTE_ADDR")
        End If
        If Successful Then
            Result = "Successful"
        Else
            Result = "Failed"
        End If

        MyCommon.Load_System_Info()
        If sMessage.Length > 0 Then
            MyCommon.Write_Log(LogFile, LogText & " " & TimeOfDay & vbTab & UserName & vbTab & IPAddress & vbTab & Result & vbTab & sMessage)
        Else
            MyCommon.Write_Log(LogFile, LogText & " " & TimeOfDay & vbTab & UserName & vbTab & IPAddress & vbTab & Result)
        End If

    End Sub

    Sub LogBrowser()
        Dim LogFile As String = ""
        Dim Browser As String = Request.Browser.Browser
        Dim Version As String = Request.Browser.Version
        Dim Platform As String = Request.Browser.Platform
        Dim UserAgent As String = Request.ServerVariables("HTTP_USER_AGENT")

        LogFile = "BrowserLog." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

        MyCommon.Load_System_Info()
        MyCommon.Write_Log(LogFile, LogText & " " & TimeOfDay & vbTab & Browser & vbTab & Version & vbTab & Platform & vbTab & UserAgent)

    End Sub

    Function GetBounceBack(ByVal BounceBack As String) As String
        Try
            'if the bounceback 
            If Not (BounceBack = "") Then
                If Not System.Uri.IsWellFormedUriString(BounceBack, UriKind.RelativeOrAbsolute) Then
                    BounceBack = "/logix/status.aspx"
                End If
            End If
        Catch ex As Exception
            BounceBack = "/logix/status.aspx"
        End Try

        Return BounceBack
    End Function

    Sub DeleteAuthTokenCookie()
        'Delete Session and cookies for session-id and authtoken on logout
        Session.Abandon()
        Session.RemoveAll()

        If Request.Cookies("ASP.NET_SessionId") IsNot Nothing Then
            Response.Cookies("ASP.NET_SessionId").Value = String.Empty
            Response.Cookies("ASP.NET_SessionId").Expires = DateTime.Now.AddMonths(-20)
        End If

        If Request.Cookies("AuthToken") IsNot Nothing Then
            DeleteAuthToken(Request.Cookies("AuthToken").Value)
            Response.Cookies("AuthToken").Value = String.Empty
            Response.Cookies("AuthToken").Expires = DateTime.Now.AddMonths(-20)
        End If
    End Sub

    '*************************************  Begin External Login Section *****************************************

    Function GetExternalRoles() As String
        Dim saExtRoleNames As String()
        Dim sExtRoleNames As String
        Dim i As Integer

        sExtRoleNames = Request.ServerVariables("HTTP_IV_GROUPS")
        'If sExtRoleNames = "" Then
        '  ' Test Mode
        '  sExtRoleNames = System.Web.HttpUtility.UrlDecode(Request.Form("extrolenames"))
        '  If sExtRoleNames = "" Then
        '    sExtRoleNames = System.Web.HttpUtility.UrlDecode(Request.QueryString("extrolenames"))
        '  End If
        'End If

        If sExtRoleNames <> "" Then
            ' remove any single or double quotes
            sExtRoleNames = sExtRoleNames.Replace("'", "")
            sExtRoleNames = sExtRoleNames.Replace("""", "")
        End If

        If sExtRoleNames <> "" Then
            ' build string for subsequent query
            saExtRoleNames = sExtRoleNames.Split(",")
            sExtRoleNames = ""
            For i = 0 To saExtRoleNames.Length - 1
                saExtRoleNames(i) = saExtRoleNames(i).Trim()
                If saExtRoleNames(i) <> "" Then
                    If i = 0 Then
                        sExtRoleNames += "'" & saExtRoleNames(i) & "'"
                    Else
                        sExtRoleNames += ",'" & saExtRoleNames(i) & "'"
                    End If
                End If
            Next
        End If

        Return sExtRoleNames
    End Function

    Function GetExternalUserName() As String
        Dim sUserName As String = ""

        sUserName = Request.ServerVariables("HTTP_IV_USER")
        'If sUserName = "" Then
        '  ' Test Mode
        '  sUserName = System.Web.HttpUtility.UrlDecode(Request.Form("extusername"))
        '  If sUserName = "" Then
        '    sUserName = System.Web.HttpUtility.UrlDecode(Request.QueryString("extusername"))
        '  End If
        'End If

        Return sUserName
    End Function

    Function GetAdminIdForExternalUser(ByVal sUsername As String, ByVal iExternalSourceId As Integer, ByRef sAuthToken As String) As Long
        Dim iUserId As Integer = 0
        Dim iUserExtSourceId As Integer
        Dim sPassword As String
        Dim sExtUserRoles As String
        Dim sDefaultExtUserRoles As String = ""
        Dim dstUser As DataTable
        Dim dstRoles As DataTable = Nothing
        Dim Crypt As New CMS.CryptLib

        MyCommon.QueryStr = "select AdminUserID, ExternalSourceId from AdminUsers where UserName=@UserName"
        MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar).Value = sUsername
        dstUser = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        If dstUser.Rows.Count > 0 Then
            iUserId = MyCommon.NZ(dstUser.Rows(0).Item("AdminUserID"), 0)
            iUserExtSourceId = MyCommon.NZ(dstUser.Rows(0).Item("ExternalSourceId"), 0)
            If (iUserId = 1) Then
                ' Admin user can't login via TAM
                iUserId = -1
            Else
                sExtUserRoles = GetExternalRoles()
                If sExtUserRoles <> "" Then
                    MyCommon.QueryStr = "select RoleID from AdminRoles where ExtRoleName in ( @ExtUserRoles )"
                    MyCommon.DBParameters.Add("@ExtUserRoles", SqlDbType.NVarChar).Value = sExtUserRoles
                    dstRoles = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dstRoles.Rows.Count > 0 Then
                        ' delete existing roles (i.e. replacing old roles with new roles)
                        MyCommon.QueryStr = "delete from AdminUserRoles with (RowLock) where AdminUserID=@UserId"
                        MyCommon.DBParameters.Add("@UserId", SqlDbType.Int).Value = iUserId
                        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                    End If
                End If
            End If
        Else
            MyCommon.QueryStr = "dbo.pt_AdminUsers_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@UserName ", SqlDbType.NVarChar, 50).Value = sUsername
            MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            iUserId = MyCommon.LRTsp.Parameters("@AdminUserID").Value
            'Generate a Salt
            Dim USalt As String = HashLib.GenerateNewSalt()
            sPassword = "EXT_" & sUsername
            sPassword = HashLib.SQL_LoginHash(sPassword, USalt)

            MyCommon.QueryStr = "update AdminUsers with (RowLock) set FirstName=@UserName,LastName='EXTERNAL',LanguageId= @DefaultLanguageID ,LastLoginExternal=1,ExternalSourceId= @ExternalSourceId ,Password=@sPassword,LastLogin=getdate(), USalt=@USalt where AdminUserID=@UserId"
            MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar).Value = sUsername
            MyCommon.DBParameters.Add("@DefaultLanguageID", SqlDbType.Int).Value = DefaultLanguageID
            MyCommon.DBParameters.Add("@ExternalSourceId", SqlDbType.Int).Value = iExternalSourceId
            MyCommon.DBParameters.Add("@sPassword", SqlDbType.NVarChar).Value = sPassword
            MyCommon.DBParameters.Add("@UserId", SqlDbType.Int).Value = iUserId
            MyCommon.DBParameters.Add("@USalt", System.Data.SqlDbType.NChar).Value = USalt
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

            MyCommon.Activity_Log(23, iUserId, iUserId, Copient.PhraseLib.Lookup("history.user-createdby", DefaultLanguageID))

            sExtUserRoles = GetExternalRoles()
            If sExtUserRoles = "" Then
                ' get default External Role
                sDefaultExtUserRoles = MyCommon.Fetch_SystemOption(93)
                If sDefaultExtUserRoles <> "" Then
                    sExtUserRoles = "'" & sDefaultExtUserRoles & "'"
                End If
            End If

            If sExtUserRoles <> "" Then
                MyCommon.QueryStr = "select RoleID from AdminRoles where ExtRoleName in (@ExtUserRoles)"
                MyCommon.DBParameters.Add("@ExtUserRoles", SqlDbType.NVarChar).Value = sExtUserRoles
                dstRoles = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If dstRoles.Rows.Count = 0 AndAlso sDefaultExtUserRoles = "" Then
                    ' No rows and have already tried default role, so get default External Role
                    sDefaultExtUserRoles = MyCommon.Fetch_SystemOption(93)
                    If sDefaultExtUserRoles <> "" Then
                        sExtUserRoles = "'" & sDefaultExtUserRoles & "'"
                        MyCommon.QueryStr = "select RoleID from AdminRoles where ExtRoleName in (@ExtUserRoles)"
                        MyCommon.DBParameters.Add("@ExtUserRoles", SqlDbType.NVarChar).Value = sExtUserRoles
                        dstRoles = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    End If
                End If
            End If
        End If

        If iUserId > 0 Then
            If Not dstRoles Is Nothing Then
                If dstRoles.Rows.Count > 0 Then
                    ' insert new roles
                    Dim dr As DataRow
                    Dim iRoleId As Integer
                    For Each dr In dstRoles.Rows
                        iRoleId = MyCommon.NZ(dr.Item("RoleID"), 0)
                        If iRoleId > 0 Then
                            MyCommon.QueryStr = "insert into AdminUserRoles with (RowLock) (RoleID,AdminUserID) values(@RoleID,@UserID)"
                            MyCommon.DBParameters.Add("@RoleID", SqlDbType.Int).Value = iRoleId
                            MyCommon.DBParameters.Add("@UserID", SqlDbType.Int).Value = iUserId
                            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                        End If
                    Next
                End If
            End If
            sAuthToken = Logix.Auth_Token_Create(iUserId, System.Web.HttpContext.Current.Session.SessionID)
        End If

        dstUser = Nothing
        dstRoles = Nothing
        Crypt = Nothing

        Return iUserId
    End Function

    Function GetAdminIdForExternalToken(ByRef sAuthToken As String, ByVal iExternalSourceId As Integer, ByRef sUserName As String, ByRef bLastLoginExternal As Boolean) As Long
        Dim iAdminUserID As Integer = 0
        Dim iUserExtSourceId As Integer
        Dim dst As DataTable

        MyCommon.QueryStr = "select AdminUserID, UserName, ExternalSourceId, LastLoginExternal from AdminUsers where Authtoken=@AuthToken"
        MyCommon.DBParameters.Add("@AuthToken", SqlDbType.VarChar).Value = sAuthToken
        dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        If dst.Rows.Count > 0 Then
            iAdminUserID = MyCommon.NZ(dst.Rows(0).Item("AdminUserID"), 0)
            sUserName = MyCommon.NZ(dst.Rows(0).Item("UserName"), "")
            If MyCommon.NZ(dst.Rows(0).Item("LastLoginExternal"), 0) = 0 Then
                bLastLoginExternal = False
            Else
                bLastLoginExternal = True
            End If
            iUserExtSourceId = MyCommon.NZ(dst.Rows(0).Item("ExternalSourceId"), 0)
            If (iUserExtSourceId <> iExternalSourceId) Then
                iAdminUserID = -1
            End If
        Else
            iAdminUserID = 0
        End If
        If iAdminUserID > 0 Then
            sAuthToken = Logix.Auth_Token_Create(iAdminUserID, System.Web.HttpContext.Current.Session.SessionID)
        End If
        dst = Nothing

        Return iAdminUserID
    End Function

    Sub VerifyExternalLogin()
        Dim sUserName As String
        Dim iAdminUserID As Integer
        Dim sAuthToken As String = ""
        Dim sLogOut As String = ""
        Dim sBounceBack As String
        Dim sTargetURL As String
        Dim dst As System.Data.DataTable

        sUserName = GetExternalUserName()
        If sUserName = "" Then
            LogAttempt(sUserName, Now(), False, "No external user name provided!")
            Response.Cookies("AuthToken").Value = ""
            ReturnToExternalLogin("", Copient.PhraseLib.Lookup("extlogin.noname", DefaultLanguageID))
        Else
            iAdminUserID = GetAdminIdForExternalUser(sUserName, icExternalsourceId, sAuthToken)
            If iAdminUserID = -1 Then
                LogAttempt(sUserName, Now(), False, "Admin user attempting external login.")
                ReturnToExternalLogin("", Copient.PhraseLib.Lookup("extlogin.notextuser", DefaultLanguageID))
            Else
                'Reset the last login for this user
                MyCommon.QueryStr = "update AdminUsers with (RowLock) set LastLoginExternal=1, LastLogin=getdate() where AdminUserID= @AdminUserID"
                MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = iAdminUserID
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                LogAttempt(sUserName, Now(), True, "External")

                sBounceBack = System.Web.HttpUtility.UrlDecode(Request.Form("logixbounceback"))
                If sBounceBack = "" Then
                    sBounceBack = System.Web.HttpUtility.UrlDecode(Request.QueryString("logixbounceback"))
                End If

                If sBounceBack = "" Then
                    'User passed verification - send them to where ever they need to go
                    MyCommon.Activity_Log(1, iAdminUserID, iAdminUserID, Copient.PhraseLib.Lookup("term.loggedin", DefaultLanguageID))
                    MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName from AdminUsers as AU with (NoLock) " & _
                                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                                      "where AU.AdminUserID=@AdminUserID"
                    MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = iAdminUserID
                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    sTargetURL = dst.Rows(0).Item("PageName")
                Else
                    sTargetURL = GetBounceBack(sBounceBack)
                End If

                Response.Cookies("AuthToken").Value = sAuthToken
                Response.Redirect(sTargetURL)
            End If
        End If

    End Sub

    Sub HandleExternalAuthToken()
        Dim sUserName As String = ""
        Dim iAdminUserID As Integer = 0
        Dim sAuthToken As String = ""
        Dim sBounceBack As String
        Dim bExternalLoginTimeOutEnabled As Boolean
        Dim bLastLoginExternal As Boolean = False

        sBounceBack = System.Web.HttpUtility.UrlDecode(Request.Form("bounceback"))
        If sBounceBack = "" Then
            sBounceBack = System.Web.HttpUtility.UrlDecode(Request.QueryString("bounceback"))
        End If

        If Not (Request.Cookies("AuthToken") Is Nothing) Then
            sAuthToken = Request.Cookies("AuthToken").Value
        End If

        iAdminUserID = GetAdminIdForExternalToken(sAuthToken, icExternalsourceId, sUserName, bLastLoginExternal)

        If iAdminUserID > 0 Then
            If bLastLoginExternal Then
                bExternalLoginTimeOutEnabled = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(92))
                If bExternalLoginTimeOutEnabled Then
                    ReturnToExternalLogin(sBounceBack, Copient.PhraseLib.Lookup("login.invalidtimeout", DefaultLanguageID))
                Else
                    If iAdminUserID > 0 Then
                        'Reset the last login for this user
                        MyCommon.QueryStr = "update AdminUsers with (RowLock) set LastLogin=getdate() where AdminUserID=@AdminUserID"
                        MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = iAdminUserID
                        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)

                        Response.Cookies("AuthToken").Value = sAuthToken
                        Response.Redirect(GetBounceBack(sBounceBack))
                    Else
                        LogAttempt(sUserName, Now(), False, "External login failed for token")
                        ReturnToExternalLogin("", Copient.PhraseLib.Lookup("extlogin.badtoken", DefaultLanguageID))
                    End If
                End If
            Else
                ' normal login & time out
                Send_Login_Page()
            End If
        Else
            LogAttempt(sUserName, Now(), False, "External login failed for token")
            ReturnToExternalLogin("", Copient.PhraseLib.Lookup("extlogin.badtoken", DefaultLanguageID))
        End If
    End Sub

    Sub LogOut()
        DeleteAuthTokenCookie()
        ReturnToExternalLogin("", Copient.PhraseLib.Lookup("extlogin.logout", DefaultLanguageID))
    End Sub

    Sub DeleteAuthToken(AuthToken As String)
        If Not (String.IsNullOrWhiteSpace(AuthToken)) Then
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "UPDATE AdminUsers WITH (RowLock) SET AuthToken = null WHERE AuthToken = @AuthToken;"
            MyCommon.DBParameters.Add("@AuthToken", SqlDbType.VarChar, 400).Value = AuthToken
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        End If
    End Sub

    Sub DeleteAuthTokenByusername(username As String)
        ' If Not (String.IsNullOrWhiteSpace(AuthToken)) Then
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "UPDATE AdminUsers WITH (RowLock) SET AuthToken = null WHERE UserName = @username;"
        MyCommon.DBParameters.Add("@username", SqlDbType.VarChar, 50).Value = username
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        'End If
    End Sub


    Sub ReturnToExternalLogin(ByVal sBounceBackUrl As String, ByVal sLogixMessage As String)
        Dim sSeparator = "?"
        Dim sExternalLoginUrl As String = MyCommon.Fetch_SystemOption(91)
        If sExternalLoginUrl.Length > 0 Then
            If Not (String.IsNullOrWhiteSpace(sBounceBackUrl)) Then
                sExternalLoginUrl = sExternalLoginUrl & "?LogixBounceBack=" & GetBounceBack(sBounceBackUrl)
                sSeparator = "&"
            End If
            If Not (String.IsNullOrWhiteSpace(sLogixMessage)) Then
                sExternalLoginUrl = sExternalLoginUrl & sSeparator & "LogixMessage=" & sLogixMessage
                sSeparator = "&"
            End If
            Response.Redirect(sExternalLoginUrl)
        Else
            Send_Login_Page(Copient.PhraseLib.Lookup("extlogin.badurl", DefaultLanguageID))
        End If
    End Sub

    '*************************************  End External Login Section *****************************************

</script>
<%
    Dim Mode As String
    Dim otp As Integer
    Dim IsSSOEnabled As Boolean = False
    CurrentRequest.Resolver.AppName = "login.aspx"
    MyCommon.AppName = "login.aspx"
    Response.Expires = 0
    On Error GoTo ErrorTrap
    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    Mode = Request.Form("mode")

    IsSSOEnabled = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(326))

    DefaultLanguageID = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(1))
    bExternalLoginEnabled = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(90))

    If IsSSOEnabled Then

        Dim StrSSOCookie As String = ""
        If Mode Is Nothing Then
            If Request.QueryString("mode") IsNot Nothing AndAlso Request.QueryString("mode") IsNot "" Then
                Mode = Request.QueryString("mode").ToUpper()
            End If
        End If

        If Request.Cookies("NCR_SSO") IsNot Nothing AndAlso Request.Cookies("NCR_SSO") IsNot "" Then
            StrSSOCookie = Request.Cookies("NCR_SSO").Value
        End If

        If Mode = "LOGOUT" Or Mode = "INVALID" Or Mode = "RETURN" Then
            SSOLogout(MyCommon)
        Else
            Dim AuthToken As String = ""
            Dim BounceBack As String = HttpUtility.UrlDecode(Request.Form("bounceback"))
            Dim IsValidSSOSession As Boolean = False

            'SSO Authentication:
            '   Check if NCR_SSO Cookie exists
            '   If NCR_SSO Cookie Exists Get the value and validate with NEP API
            '       If validation fails redirect the user to NEP Login Page
            '       If validation success
            '           -Verify the User in Logix
            '           -Auto login to logix system
            '           -redirect the user to default logix page
            '   If Not Exists redirect the user to NEP Login Page

            'Check if NCR_SSO Cookie exists
            If StrSSOCookie IsNot "" Then
                'If NCR_SSO Cookie Exists, validate with NEP API
                If ValidateSSOCookie(MyCommon, StrSSOCookie, UserName) Then
                    'Verify the User in Logix
                    If Verify_SSOUser(Logix, MyCommon, UserName, userID, AuthToken, emailId) Then
                        Session("UserID") = userID
                        Session("UserName") = UserName
                        Session("AuthToken") = AuthToken
                        Session("MailAddress") = emailId

                        IsValidSSOSession = True

                        'Auto login to logix system and 'redirect the user to default logix page
                        AllowAccess(userID, BounceBack, AuthToken)
                    End If

                End If
            End If

            If Not IsValidSSOSession Then
                If HttpContext.Current.Request.Url.AbsoluteUri.Contains("login.aspx") Then
                    Response.Redirect(MyCommon.Fetch_SystemOption(327) & "/login?" & "g=" & HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Authority))
                Else
                    Response.Redirect(MyCommon.Fetch_SystemOption(327) & "/login?" & "g=" & HttpContext.Current.Request.Url.AbsoluteUri)
                End If

            End If
        End If


    ElseIf bExternalLoginEnabled Then

        Select Case UCase(Mode)
            Case "RETURN"
                ReturnToExternalLogin("", "")
            Case "INVALID"
                HandleExternalAuthToken()
            Case "VERIFY"
                Verify_User()
            Case "LOCAL"
                Send_Login_Page()
            Case "LOGOUT"
                LogOut()
            Case Else
                VerifyExternalLogin()
        End Select
    Else
        Select Case UCase(Mode)
            Case "VERIFY"
                Verify_User()
                If SecondFactorEnabled AndAlso Not String.IsNullOrEmpty(Session("MailAddress")) Then
                    If BypassAuthorizedForCurrentUser = False Then
                        otp = GenerateOTP()
                        SendOTP(otp, Session("UserName"), Session("MailAddress"))
                    Else
                        AllowAccess(Session("UserID"), HttpUtility.UrlDecode(Request.Form("bounceback")), Session("AuthToken"))
                        Session("UserID") = Nothing
                        Session("Authtoken") = Nothing
                        Session("UserName") = Nothing
                        Session("MailAddress") = Nothing
                    End If
                End If
            Case "OTPEMAIL"
                GetUserCredentials()
                If ValidateEmail() Then
                    Dim m_adminUserDataService As IAdminUserData = CurrentRequest.Resolver.Resolve(Of IAdminUserData)()
                    m_adminUserDataService.SaveEmail(Request.Form("tbMailId"), UserName)
                    Session("MailAddress") = Request.Form("tbMailId")
                    otp = GenerateOTP()
                    If SendOTP(otp, Session("UserName"), Request.Form("tbMailId")) Then
                        Send_Login_Page(String.Empty, True, False, True)
                    End If
                Else
                    isEmailValid = False
                    Send_Login_Page(String.Empty, True, True)
                End If
            Case "RESENDOTP"
                GetUserCredentials()
                otp = GenerateOTP()
                If SendOTP(otp, Session("UserName"), Session("MailAddress")) Then
                    isOTPResend = True
                    Send_Login_Page(String.Empty, True, False, True)
                End If
            Case "OTP"
                If ValidateOTP() Then
                    AllowAccess(Session("UserID"), HttpUtility.UrlDecode(Request.Form("bounceback")), Session("AuthToken"))
                    If SecondFactorBypassEnabled AndAlso Request.Form("bypassAuthentication") = "on" Then
                        bypass2FactorCookieName = String.Concat("SecondFactorAuthentication", Session("UserID"))
                        Dim secondfactorCookie As HttpCookie = New HttpCookie(bypass2FactorCookieName)
                        secondfactorCookie("user") = MyCryptlib.SQL_StringEncrypt(Session("UserName"))
                        secondfactorCookie.Expires = DateTime.Now.AddDays(90)
                        Response.Cookies.Add(secondfactorCookie)
                    End If
                    Session("UserID") = Nothing
                    Session("Authtoken") = Nothing
                    Session("UserName") = Nothing
                    Session("MailAddress") = Nothing
                Else
                    isOTPValid = False
                    GetUserCredentials()
                    Send_Login_Page(String.Empty, True, False, True)
                End If

            Case Else
                Send_Login_Page()
        End Select
    End If

    If Not (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Close_LogixRT()
%>
<%
                    Response.End()
ErrorTrap:
                    Response.Write("<pre>" & MyCommon.Error_Processor() & "</pre>")
                    MyCommon.Close_LogixRT()
                    MyCommon = Nothing
%>
