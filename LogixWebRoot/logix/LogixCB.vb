' version:7.3.1.138972.Official Build (SUSDAY10202)
' *****************************************************************************
' * FILENAME: LogixCB.vb
' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' * Copyright © 2002 - Copyright (c) 2002 - 2019.  All rights reserved by:
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

Imports System
Imports System.Data
Imports System.IO
Imports System.Xml

Public Class LogixCB
    Inherits System.Web.UI.Page
    Protected UPC_NUMERIC_ONLY_SYSTEM_OPTION As Integer = 97
    Public Const MAX_CUST_PASSWORD_CLEARTEXT_LEN As Integer = 120
    Public Const UPC_E_OR_EAN_8_OPTION_NUM As Integer = 52
    Public UPC_E_OR_EAN_8 As Integer = -1
    Public OfferLockedforCollisionDetection As Boolean = False

    'Public Logix As New Copient.LogixInc
    'Public Common As New Copient.CommonInc
    Public AdminName As String
    Public LanguageID As Integer
    Public AdminUserID As Integer
    Public BuyersAssociated As List(Of Int32)

    Private Const AntiXsrfTokenKey As String = "__AntiXsrfToken"
    Private Const AntiXsrfUserNameKey As String = "__AntiXsrfUserName"
    Private _antiXsrfTokenValue As String

    Protected Sub Page_Init(sender As Object, e As EventArgs)
        'First, check for the existence of the Anti-XSS cookie
        Dim requestCookie = Request.Cookies(AntiXsrfTokenKey)
        Dim requestCookieGuidValue As Guid
        

        'If the CSRF cookie is found, parse the token from the cookie.
        'Then, set the global page variable and view state user
        'key. The global variable will be used to validate that it matches in the view state form field in the Page.PreLoad
        'method.
        If requestCookie IsNot Nothing AndAlso Guid.TryParse(requestCookie.Value, requestCookieGuidValue) Then
            'Set the global token variable so the cookie value can be
            'validated against the value in the view state form field in
            'the Page.PreLoad method.
            _antiXsrfTokenValue = requestCookie.Value

            'Set the view state user key, which will be validated by the
            'framework during each request
            Page.ViewStateUserKey = _antiXsrfTokenValue
        Else
            'If the CSRF cookie is not found, then this is a new session.
            'Generate a new Anti-XSRF token
            _antiXsrfTokenValue = Guid.NewGuid().ToString("N")

            'Set the view state user key, which will be validated by the
            'framework during each request
            Page.ViewStateUserKey = _antiXsrfTokenValue

            'Create the non-persistent CSRF cookie
            'Set the HttpOnly property to prevent the cookie from
            'being accessed by client side script

            'Add the Anti-XSRF token to the cookie value
            Dim responseCookie = New HttpCookie(AntiXsrfTokenKey)
            responseCookie.HttpOnly = True
            responseCookie.Value = _antiXsrfTokenValue

            'If we are using SSL, the cookie should be set to secure to
            'prevent it from being sent over HTTP connections
            If FormsAuthentication.RequireSSL AndAlso Request.IsSecureConnection Then
                responseCookie.Secure = True
            End If

            'Add the CSRF cookie to the response
            Response.Cookies.[Set](responseCookie)
        End If

        AddHandler Page.PreLoad, AddressOf master_Page_PreLoad
    End Sub



    Protected Sub master_Page_PreLoad(sender As Object, e As EventArgs)
        'During the initial page load, add the Anti-XSRF token and user
        'name to the ViewState
        If Not IsPostBack Then
            'Set Anti-XSRF token
            ViewState(AntiXsrfTokenKey) = Page.ViewStateUserKey

            'If a user name is assigned, set the user name
            ViewState(AntiXsrfUserNameKey) = If(Context.User.Identity.Name, [String].Empty)
        Else
            'During all subsequent post backs to the page, the token value from
            'the cookie should be validated against the token in the view state
            'form field. Additionally user name should be compared to the
            'authenticated users name
            'Validate the Anti-XSRF token
            If DirectCast(ViewState(AntiXsrfTokenKey), String) <> _antiXsrfTokenValue OrElse DirectCast(ViewState(AntiXsrfUserNameKey), String) <> (If(Context.User.Identity.Name, [String].Empty)) Then
                Throw New InvalidOperationException("Validation of Anti - XSRF token failed.")
            End If
        End If
    End Sub

    Public Shared Function isEmpty(ByVal s As String) As Boolean
        Return s Is Nothing OrElse s.Trim.Length < 1
    End Function

    Public Shared Function ifEmptyString(ByVal s As String, ByVal defaultVal As String) As String
        Dim rs As String = IIf(isEmpty(s), defaultVal, s)
        Return rs.Trim
    End Function

    Public Function Verify_AdminUser(ByRef Common As Object, ByRef MyLogix As Object) As Long
        Dim Authtoken As String = ""
        Dim MyURI As String
        Dim TransferKey As String
        Dim Debug As Boolean = False
        Dim IsSSOEnabled As Boolean = False

        If Not (AdminUserID = 0) Then
            If Debug Then Common.write_log("auth.txt", "AppName=" & Common.AppName & " - Verify_AdminUser was called, but we already know the AdminUserID=" & AdminUserID, True)
            IsSSOEnabled = Common.Extract_Val(Common.Fetch_SystemOption(326))
            If IsSSOEnabled Then
                If Not ValidateSSOCookie(Common, AdminUserID.ToString()) Then
                    Return AdminUserID
                    Exit Function
                End If
            End If
            'we already know who the AdminUser is ... we shouldn't be looking him up more than once
            MyLogix.Load_Roles(Common, AdminUserID)
            BuyersAssociated = MyLogix.LoadBuyersForUser(AdminUserID)
            Return AdminUserID
            Exit Function
        End If

        '1st, check the transferkey and see if the user is being transferred into AMS from another product (PrefMan)
        If GetCgiValue("transferkey") <> "" Then
            If Debug Then Common.write_log("auth.txt", "AppName=" & Common.AppName & " - Checking the TransferKey (" & GetCgiValue("transferkey") & ")  AdminUserID=" & AdminUserID, True)

            TransferKey = GetCgiValue("transferkey")
            AdminUserID = MyLogix.Auth_TransferKey_Verify(Common, TransferKey, AdminName, LanguageID, Authtoken)
            If Debug Then Common.write_log("auth.txt", "AppName=" & Common.AppName & " - After TransferKey_Verify AdminUserID=" & AdminUserID, True)
            If Not (AdminUserID = 0) Then
                Response.Cookies("AuthToken").Value = Authtoken
                MyLogix.Load_Roles(Common, AdminUserID)
                BuyersAssociated = MyLogix.LoadBuyersForUser(AdminUserID)
                Return AdminUserID
                Exit Function
            End If
        End If

        Authtoken = ""
        If Not (Request.Cookies("AuthToken") Is Nothing) Then
            Authtoken = Request.Cookies("AuthToken").Value
        End If
        If Debug Then Common.write_log("auth.txt", "AppName=" & Common.AppName & " - AuthToken='" & Authtoken & "'   Transferkey='" & GetCgiValue("transferkey") & "'", True)
        AdminUserID = 0
        AdminUserID = MyLogix.Auth_Token_Verify(Common, Authtoken, AdminName, LanguageID)
        If Debug Then Common.write_log("auth.txt", "AppName=" & Common.AppName & " - After checking AuthToken, AdminUserID=" & AdminUserID, True)

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
        Else
            MyLogix.Load_Roles(Common, AdminUserID)
            BuyersAssociated = MyLogix.LoadBuyersForUser(AdminUserID)
        End If
        Return AdminUserID

    End Function

    Public Function ValidateSSOCookie(ByRef ObjCommon As Copient.CommonInc, ByVal StrSSOCookie As String, ByRef UserName As String) As Boolean
        ' Validate SSO token with NEP
        Dim RetVal As Boolean = False
        CMS.AMS.CurrentRequest.Resolver.AppName = "LoginCB.vb"
        Dim ObjNEPSSOService As CMS.AMS.Contract.ILoginWithSSO = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Security.LoginWithSSO)()
        Dim ObjAMSResult As CMS.AMS.Models.AMSResult(Of String) = New CMS.AMS.Models.AMSResult(Of String)
        Try
            ObjAMSResult = ObjNEPSSOService.ValidateSSOToken(StrSSOCookie, ObjCommon.Fetch_SystemOption(328))
            If ObjAMSResult.ResultType = CMS.AMS.Models.AMSResultType.Success Then
                RetVal = True
                UserName = ObjAMSResult.MessageString
            End If
        Catch ex As Exception
            RetVal = False
        Finally
            ObjNEPSSOService = Nothing
            ObjAMSResult = Nothing
        End Try

        Return RetVal
    End Function

    ' Check for valid SSO token, validate token with NEP and if invalid session logout from NEP
    Public Function ValidateSSOCookie(ByRef ObjCommon As Copient.CommonInc, ByRef UserName As String) As Boolean
        Dim RetVal As Boolean = False
        Dim StrSSOCookie As String = ""

        If Request.Cookies("NCR_SSO") IsNot Nothing AndAlso Request.Cookies("NCR_SSO") IsNot "" Then
            StrSSOCookie = Request.Cookies("NCR_SSO").Value
            If ValidateSSOCookie(ObjCommon, StrSSOCookie, UserName) Then
                RetVal = True
            End If
        End If

        If Not RetVal Then
            ObjCommon.Write_Log("auth.txt", "AppName=" & ObjCommon.AppName & " - failed to validate SSO session, UserName=" & UserName, True)
            Response.Redirect(ObjCommon.Fetch_SystemOption(327) & "login?" & "g=" & HttpContext.Current.Request.Url.AbsoluteUri)
        End If

        Return RetVal
    End Function


    ' Check Single sign-on is enabled 
    ' Check if user exist, if not "pt_Verify_SSO_User" will create a user with empty password, assign 'Membership Editor' role and generate the AuthToken
    ' If user exists then generate the AuthToken
    Public Function Verify_SSOUser(ByRef Logix As Copient.LogixInc, ByRef Common As Copient.CommonInc, ByVal UserName As String, ByRef UserID As Long, ByRef AuthToken As String, ByRef Email As String) As Boolean
        Dim RetVal As Boolean = True
        Dim dst As New DataTable
        Dim Crypt As New CMS.CryptLib
        Dim IsNewUser As Boolean

        Try
            Common.AppName = "LogixInc Library"
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()

            'Check if Single Sign-On is enabled
            If (Common.Fetch_SystemOption(326) = "1") Then
                'Single sign-on is enabled 
                'Check if user exist, if not "pt_Verify_SSO_User" will create a user with empty password and assign a Membership Editor role and generate the AuthToken
                'If user exists then generate the AuthToken
                Common.QueryStr = "dbo.pt_Verify_SSO_User"
                Common.Open_LRTsp()
                Common.LRTsp.Parameters.Add("@UserName", SqlDbType.NVarChar, 100).Value = UserName
                Common.LRTsp.Parameters.Add("@IsNewUser", SqlDbType.Bit).Direction = ParameterDirection.Output
                dst = Common.LRTsp_select
                Common.Close_LRTsp()

                If dst.Rows.Count > 0 Then
                    UserID = Common.NZ(dst.Rows(0).Item("AdminUserID"), 0)
                    Email = Crypt.SQL_StringDecrypt(dst.Rows(0).Item("Email").ToString())
                    AuthToken = Logix.Auth_Token_Create(UserID, System.Web.HttpContext.Current.Session.SessionID)
                    IsNewUser = Common.LRTsp.Parameters("@IsNewUser").Value

                    If IsNewUser Then
                        Common.Write_Log("auth.txt", "AppName=" & Common.AppName & "Verify_SSOUser: New User " & UserName & " created via NEP/SSO", True)
                    End If
                Else
                    RetVal = False
                    Common.Write_Log("auth.txt", "AppName=" & Common.AppName & "Verify_SSOUser: failed to get user ", True)
                End If

                If UserID = 0 Then
                    RetVal = False
                    Common.Write_Log("auth.txt", "AppName=" & Common.AppName & "Verify_SSOUser: failed to get user ", True)
                End If
            End If
        Catch Ex As Exception
            RetVal = False
            Common.Write_Log("auth.txt", "AppName=" & Common.AppName & "Verify_SSOUser: Exception " & Ex.Message, True)
        Finally
            If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
            dst.Dispose()
            Crypt = Nothing

        End Try

        Return RetVal
    End Function

    ' SSOLogout
    Public Sub SSOLogout(ByRef Common As Copient.CommonInc)

        Response.Redirect(Common.Fetch_SystemOption(327) & "logout")

    End Sub

    '************************************* END SSO Authentication Section *************************************

    Public Sub Send(ByVal WebText As String)
        Response.Write(WebText & vbCrLf)
    End Sub

    Public Sub Sendb(ByVal WebText As String)
        Response.Write(WebText)
    End Sub

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

    Public Function CleanString(ByVal InString As String, Optional ByVal AdditionalValidCharacters As String = "") As String
        Dim tmpString As String = ""
        Dim z As Integer

        If InString IsNot Nothing Then
            For z = 0 To InString.Length - 1
                If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.#$*-_&%@!?/:;+ " & AdditionalValidCharacters, InString(z))) Then
                    tmpString = tmpString & InString(z)
                End If
            Next
        End If

        CleanString = tmpString
    End Function

    Public Function CleanStringRew(ByVal InString As String) As String
        Dim tmpString As String = ""
        Dim z As Integer

        If InString IsNot Nothing Then
            For z = 0 To InString.Length - 1
                If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.#$*-_&%@!?/:;(),'+ ", InString(z))) Then
                    tmpString = tmpString & InString(z)
                End If
            Next
        End If
        CleanStringRew = tmpString
    End Function
    Public Function ValidUPC(ByRef MyCommon As Copient.CommonInc, ByVal UPCString As String, ByRef Message As String) As Boolean
        Dim isValidUPC As Boolean = True
        If String.IsNullOrWhiteSpace(UPCString) Then
            isValidUPC = False
            Message = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
        ElseIf (MyCommon.Fetch_SystemOption(UPC_NUMERIC_ONLY_SYSTEM_OPTION) = 1) Then
            If (MyCommon.Extract_Val(GetCgiValue("ExtProductID")) < 1) Or (Int(MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) <> MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) Then
                isValidUPC = False
                Message = Copient.PhraseLib.Lookup("pgroup-edit.numericonly", LanguageID)
            End If
        ElseIf CleanUPC(UPCString) = False Then
            isValidUPC = False
            Message = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
        End If
        Return isValidUPC
    End Function

    Public Function CleanUPC(ByVal InString As String) As String
        Dim z As Integer
        Dim IsClean As Boolean

        If InString IsNot Nothing Then
            For z = 0 To InString.Length - 1
                If (InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-", InString(z))) Then
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

    Public Function CleanProductCodes(ByVal ProductCodes As String) As String
        Dim cleanedString As String = ""
        Dim tempProducts As String = ""
        tempProducts = ProductCodes.Replace(vbCrLf, ",")
        'Dim charsToTrim() As Char = {","c, " "c}
        'tempProducts = tempProducts.Trim(charsToTrim)
        cleanedString = tempProducts.Replace(",,", ",")
        cleanedString = cleanedString.Replace(", ,", ",")

        Return cleanedString
    End Function

    Public Function GenerateProductCodeWithPadding(ByVal ProductType As Integer, ByVal ProductCode As String) As String
        Dim MyCommon As New Copient.CommonInc
        Dim IDLength As Integer = 0
        Dim bRTConnectionOpened As Boolean = False
        Dim ExtProductID As String = String.Empty
	      Dim rst As DataTable	
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If
            If (Int(ProductType) = 1) Then
        		MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
        		rst = MyCommon.LRT_Select
        		If rst IsNot Nothing Then
        			IDLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
        		End If
            ElseIf (Int(ProductType) = 2) Then
                Integer.TryParse(MyCommon.Fetch_SystemOption(54), IDLength)
            Else
                IDLength = 0
            End If
            If (IDLength > 0) Then
                If (Int(ProductType) = 2) Then
                    ExtProductID = MyCommon.Parse_Quotes(Left(Trim(ProductCode), 120)).PadLeft(IDLength, "0")
                ElseIf (Int(ProductType) = 1) Then
                    ExtProductID = MyCommon.Parse_Quotes(Left(Trim(ProductCode), 19)).PadLeft(IDLength, "0")
                Else
                    ExtProductID = MyCommon.Parse_Quotes(Left(Trim(ProductCode), 26)).PadLeft(IDLength, "0")
                End If
            Else
                If (Int(ProductType) = 2) Then
                    ExtProductID = MyCommon.Parse_Quotes(Left(Trim(ProductCode), 120))
                ElseIf (Int(ProductType) = 1) Then
                    ExtProductID = MyCommon.Parse_Quotes(Left(Trim(ProductCode), 19))
                Else
					ExtProductID = MyCommon.Parse_Quotes(Left(Trim(ProductCode), 26))
				End If
            End If
        Catch ex As Exception

        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try

        Return ExtProductID
    End Function

    Public Function IsValidItemCode(ByVal ProductGroupID As String, ByVal ProductType As Integer, ByRef ProductCode As String, ByVal Operation As Integer, Optional ByRef InfoMsg As String = "") As Boolean

        Dim MyCommon As New Copient.CommonInc
        Dim bGoodItemCode As Boolean = True
        Dim IDLength As Integer = 0
        Dim iProductIdNumericOnly As Integer = 97
        Dim bCreateProducts As Boolean = MyCommon.Fetch_SystemOption(150)
        Dim ExtProductID As String = ""
        Dim rst As DataTable
        Dim bRTConnectionOpened As Boolean = False

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            MyCommon.Open_LogixRT()
            bRTConnectionOpened = True
        End If

        If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(iProductIdNumericOnly)) = 1 AndAlso CleanUPC(ProductCode) = False) Then
            bGoodItemCode = False
        End If

        If ((MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) = True AndAlso MyCommon.Fetch_CM_SystemOption(82) = "1") OrElse (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) = True AndAlso MyCommon.Fetch_UE_SystemOption(144) = "1")) Then
            If (ProductType = 1) Then
                If (CheckItemCode(ProductCode, InfoMsg) = False) Then
                    bGoodItemCode = False
                End If
            End If
        ElseIf (CleanUPC(ProductCode) = False) Then
            bGoodItemCode = False
        End If

        If bGoodItemCode Then
            ExtProductID = GenerateProductCodeWithPadding(ProductType, ProductCode)

            If (MyCommon.Extract_Val(ExtProductID) < 1) Or (Int(MyCommon.Extract_Val(ExtProductID)) <> MyCommon.Extract_Val(ExtProductID)) Then
                bGoodItemCode = False
            End If

            If bGoodItemCode Then
                ProductCode = ExtProductID

                If Operation = 2 Then 'Check whether product exist in the product group before removing
                    Try

                        MyCommon.QueryStr = "select pg.productid from ProdGroupItems as pg with (NoLock) inner join products as p with (NoLock) on pg.productid=p.productid and  pg.ProductGroupID=" & ProductGroupID & " and p.extproductid='" & ProductCode & "' and p.ProductTypeID=" & ProductType & " and pg.Deleted=0;"
                        rst = MyCommon.LRT_Select()
                        If rst.Rows.Count = 0 Then
                            bGoodItemCode = False 'Product doesnot exist in the product group
                        End If
                    Catch ex As Exception

                    Finally
                        If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
                    End Try

                End If
            End If
        End If
        Return bGoodItemCode
    End Function

    Public Function SaveProduct(ByVal ExtProductID As String, ByVal ProductTypeID As Integer, ByVal Operation As Integer) As String

        Dim MyCommon As New Copient.CommonInc
        Dim ProductDesc As String = ""
        Dim querystr As String = ""
        Dim dst As DataTable
        Dim ProductID As Integer
        Dim bRTConnectionOpened As Boolean = False

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If
            MyCommon.QueryStr = "select ProductID from Products with (NoLock) where (ExtProductID = '" & ExtProductID & "' and " &
                                " ProductTypeID = " & ProductTypeID & ")"

            dst = MyCommon.LRT_Select

            If dst.Rows.Count > 0 Then
                querystr &= "insert into #TempProdPK values(" & MyCommon.NZ(dst.Rows(0).Item("ProductID"), 0) & "); "
            Else

                'Create New Product
                If Operation <> 2 Then

                    MyCommon.QueryStr = "dbo.pa_PUA_UpdateProduct"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
                    MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = ProductDesc
                    MyCommon.LRTsp.Parameters.Add("@ProductID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    ProductID = MyCommon.LRTsp.Parameters("@ProductID").Value
                    MyCommon.Close_LRTsp()

                    'Add to table
                    querystr = "insert into #TempProdPK values(" & ProductID & "); "

                End If

            End If
        Catch ex As Exception
            'MyCommon.Write_Log(LogFile, "Method Name : Saveproduct " & vbCr & ex.Message & vbCr, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try
        Return querystr
    End Function

    Public Function SaveProductToProductGroup(ByVal insertStatement As String, ByVal ProductGroupID As Integer, ByVal Operation As String, ByVal ProductType As Integer) As Boolean

        Dim MyCommon As New Copient.CommonInc
        Dim querystr As String
        Dim Status As Integer
        Dim Sucess As Boolean = True
        Dim bRTConnectionOpened As Boolean = False

        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If

            MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempProdPK') IS NOT NULL BEGIN drop table #TempProdPK END "
            MyCommon.LRT_Execute()
            querystr = "create table #TempProdPK([TempPK] int PRIMARY KEY IDENTITY," &
                 "[ProductID] bigint NOT NULL)"

            '0 -Full Replace, 1 - Add to Group, 2- Remove from group

            querystr &= insertStatement

            MyCommon.QueryStr = querystr

            MyCommon.LRT_Execute()

            If Operation = 0 Then
                MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Replace"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.CommandTimeout = 2400
                MyCommon.LRTsp.ExecuteNonQuery()
                Status = MyCommon.LRTsp.Parameters("@Status").Value
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = " drop table #TempProdPK;"
                MyCommon.LRT_Execute()
                If Status = -2 Then
                    Sucess = False
                End If

            ElseIf Operation = 1 Then
                MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.CommandTimeout = 2400
                MyCommon.LRTsp.ExecuteNonQuery()
                Status = MyCommon.LRTsp.Parameters("@Status").Value
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = " drop table #TempProdPK;"
                MyCommon.LRT_Execute()
                If Status = -2 Then
                    Sucess = False
                End If

            ElseIf Operation = 2 Then
                MyCommon.QueryStr = "dbo.pt_LGM_ProductGroup_Remove"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt, 20).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.CommandTimeout = 2400
                MyCommon.LRTsp.ExecuteNonQuery()
                Status = MyCommon.LRTsp.Parameters("@Status").Value
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = " drop table #TempProdPK;"
                MyCommon.LRT_Execute()
                If Status = -2 Then
                    Sucess = False
                End If

            End If

            MyCommon.QueryStr = "update productgroups with (RowLock) set  LastUpdate=getdate() where ProductGroupID=" & ProductGroupID
            MyCommon.LRT_Execute()

        Catch ex As Exception
            Sucess = False
            'MyCommon.Write_Log(LogFile, "Method Name : SaveProductToProductGroup " & vbCr & ex.Message & vbCr, True)
        Finally
            MyCommon.QueryStr = "IF OBJECT_ID('tempdb..#TempProdPK') IS NOT NULL BEGIN drop table #TempProdPK END "
            MyCommon.LRT_Execute()
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try

        Return Sucess
    End Function

    Public Sub Send_HeadBegin(Optional ByVal PageTitle As String = "", Optional ByVal PageSubTitle As String = "", Optional ByVal PageID As String = "", Optional ByVal DefaultLanguageID As Integer = 0)
        Dim TempLanguageID As Integer
        Dim MyCommon As New Copient.CommonInc
        Dim dst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim longDate As New DateTime
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
        Send("<base target=""_self""/>")
        Sendb("<title>" & Copient.PhraseLib.Lookup("term.logix", TempLanguageID))
        If (PageTitle = "") Then
        Else
            Sendb(" > " & Copient.PhraseLib.Lookup(PageTitle, TempLanguageID))
        End If
        If (PageID = "") Then
        Else
            Sendb(" " & Left(PageID, 100))
        End If
        If (PageSubTitle = "") Then
        Else
            Sendb(" > " & Copient.PhraseLib.Lookup(PageSubTitle, TempLanguageID))
        End If
        Send("</title>")
        MyCommon.Close_LogixRT()
        MyCommon = Nothing
        'This header is added to avoid cross-frame scripting for ticket AMS-2318
        Response.AddHeader("X-Frame-Options", "SAMEORIGIN ")
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
        Send_Metas(DefaultLanguageID, "IE=9")
    End Sub

    Public Sub Send_Metas(ByVal DefaultLanguageID As Integer, ByVal CompatibilityMode As String)
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
        Send("<meta http-equiv=""X-UA-Compatible"" content=""" & CompatibilityMode & """ />")
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
        Dim EPMInstalled As Boolean = False
        Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
        Dim StyleValue As Integer = -1

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()

        Send("<link rel=""icon"" href=""/images/favicon.ico"" type=""image/x-icon"" />")
        Send("<link rel=""shortcut icon"" href=""/images/favicon.ico"" type=""image/x-icon"" />")
        Send("<link rel=""apple-touch-icon"" href=""/images/touchicon.png"" />")
        myUrl = Request.CurrentExecutionFilePath.ToLower()
        If (Copient.commonShared.UnauthenticatedPages.Contains(myUrl)) Then
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
        Send("<link rel=""stylesheet"" href=""/css/logix-print.css"" type=""text/css"" media=""braille, embossed, print, tty"" />")
        If Restricted Then
            Send("<link rel=""stylesheet"" href=""/css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
        End If

        EPMInstalled = MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)
        If EPMInstalled Then
            Send("<style type=""text/css"">")
            If Request.Cookies("Style") IsNot Nothing Then
                Integer.TryParse(Request.Cookies("Style").Value, StyleValue)
            End If

            Select Case StyleValue
                Case 1  'override the Logix Blue style
                    Send("#tabs a, #tabs a.on {background: url('/images/tab_narrow1.png') no-repeat scroll 0 0 transparent; left: 7px; width: 82px;}")
                    Send("#tabs a:hover {background: url('/images/tab-hover_narrow1.png') no-repeat;}")
                    Send("#tabs a.on {background: url('/images/tab-on_narrow1.png') no-repeat;}")
                    Send("#tabs a.on:hover {background: url('/images/tab-on_narrow1.png') no-repeat;}")
                Case 2  'override the NCR Blue style
                    Send("#tabs a, #tabs a.on {left: 7px; width: 82px;}")
                    Send("#tabs a:hover {background: url('/images/ncr/tab-hover_narrow1.png') no-repeat;}")
                    Send("#tabs a.on {background: url('/images/ncr/tab-on_narrow1.png') no-repeat; font-weight: bold; height: 25px;}")
                    Send("#tabs a.on:hover {background: url('/images/ncr/tab-on_narrow1.png') no-repeat;}")
                Case 3  'override the GOLD style
                    Send(" ") 'nothing to change here
                Case Else 'override the NCR Green style
                    Send("#tabs a, #tabs a.on {left: 7px; width: 82px;}")
                    Send("#tabs a:hover {background: url('/images/ncrgreen/tab-hover_narrow1.png') no-repeat;}")
                    Send("#tabs a.on {background: url('/images/ncrgreen/tab-on_narrow1.png') no-repeat; font-weight: bold; height: 25px;}")
                    Send("#tabs a.on:hover {background: url('/images/ncrgreen/tab-on_narrow1.png') no-repeat;}")
            End Select
            Send("</style>")
        End If

        If Request.Browser.Browser = "IE" Or Request.Browser.Browser = "Opera" Then
            'Browser-specific multilanguage-input tweak
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

        Send("<script type=""text/javascript"" src=""/javascript/logix.js""></script>")
        '  Send("<script type=""text/javascript"" src=""/javascript/jquery-1.0.2.js""></script>")
        Send("<script type=""text/javascript"" src=""/javascript/jquery.min.js""></script>")
        Send("<script type=""text/javascript"" src=""/javascript/jquery-ui-1.10.3/jquery-ui-min.js""></script>")
        Send("<link rel=""stylesheet"" href=""/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css"" />")

        If (ScriptNames IsNot Nothing) Then
            For i = 0 To ScriptNames.GetUpperBound(0)
                Send("<script src=""/javascript/" & ScriptNames(i) & """ type=""text/javascript""></script>")
            Next
        End If

        Send_JavaScript_Terms()
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
        ElseIf (BodyType = 7) Then
            Send(" class=""report"">")
        ElseIf (BodyType = 11) Then
            Send(" class=""template"">")
        ElseIf (BodyType = 12) Then
            Send(" class=""popup template"" onunload=""ChangeParentDocument()"">")
        ElseIf (BodyType = 13) Then
            Send(" class=""popup template"">")
        ElseIf (BodyType = 14) Then
            Send(" class=""template"" onunload=""updateCookie()"">")
        ElseIf (BodyType = 15) Then
            Send("class=""extrawidepopup""onunload=""ChangeParentDocument()"">")
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
        Send("<!-- Admin UserID =" & MyCommon.GetAdminUser.AdminUserID & " -->")

        Send("  <span id=""user""><a href=""/logix/user-edit.aspx?UserID=" & AdminUserID & """>" & IIf(Len(AdminName) > 20, Left(AdminName, 20) & "...", AdminName) & "</a><span class=""noprint""> | </span></span>")
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

        If (File.Exists(Server.MapPath("/images/logos/licenseecustom.png"))) Then
            Send("  <div id=""licenseecustom"" title=""" & Copient.PhraseLib.Lookup("term.licensee", TempLanguageID) & """></div>")
        Else
            Send("  <div id=""licensee"" title=""" & Copient.PhraseLib.Lookup("term.licensee", TempLanguageID) & """></div>")
        End If

        Send("  <br clear=""all"" />")
        Send("</div>")
        Send("")
    End Sub

    Public Sub Send_Tabs(ByRef MyLogix As Object, ByVal Tabset As Integer)
        Dim MyCommon As New Copient.CommonInc
        Dim dst As System.Data.DataTable
        Dim TabOns() As String = {"", "", "", "", "", "", "", "", "", "", ""}
        Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
        Dim EPMInstalled As Boolean
        Dim TabStyleOverride As String = ""
        Dim EPMPage As String = ""
        Dim EPMHostURI As String = ""

        TabOns(Tabset) = "class=""on"" "

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()

        EPMInstalled = MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)

        Send("<div id=""tabs"">")
        Send("  <a href=""/logix/status.aspx"" accesskey=""1"" " & TabOns(1) & TabStyleOverride & "id=""tab1"" title=""" & Copient.PhraseLib.Lookup("term.systemoverview", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.logix", LanguageID) & "</a>")
        Send("  <a href=""/logix/offer-list.aspx"" accesskey=""2"" " & TabOns(2) & TabStyleOverride & "id=""tab2"" title=""" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & "</a>")
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""3"" " & TabOns(3) & TabStyleOverride & "id=""tab3"" title=""" & Copient.PhraseLib.Lookup("term.customers", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.customers", LanguageID) & "</a>")
        End If
        Send("  <a href=""/logix/pgroup-list.aspx"" accesskey=""4"" " & TabOns(4) & TabStyleOverride & "id=""tab4"" title=""" & Copient.PhraseLib.Lookup("term.products", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.products", LanguageID) & "</a>")
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then
            Send("  <a href=""/logix/point-list.aspx"" accesskey=""5"" " & TabOns(5) & TabStyleOverride & "id=""tab5"" title=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & "</a>")
        End If
        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
            Send("  <a href=""/logix/graphic-list.aspx"" accesskey=""6"" " & TabOns(6) & TabStyleOverride & "id=""tab6"" title=""" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & "</a>")
        End If
        Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""7"" " & TabOns(7) & TabStyleOverride & "id=""tab7"" title=""" & Copient.PhraseLib.Lookup("term.stores", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.locations", LanguageID) & "</a>")

        If (MyLogix.UserRoles.AccessStoreHealth = True) Then
            MyCommon.QueryStr = "select EngineID from PromoEngines where Installed=1 and DefaultEngine=1 and EngineID in (0,2,9);"
            dst = MyCommon.LRT_Select
            If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(91) = "1" AndAlso MyCommon.NZ(dst.Rows(0).Item("EngineID"), -1) = 9 Then
                Send("  <a href=""/logix/UE/UEServerHealthSummary.aspx"" accesskey=""8"" " & TabOns(8) & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
            Else
                MyCommon.QueryStr = "select EngineID from PromoEngines where Installed=1 and DefaultEngine=1 and EngineID in (0,2,9);"
                dst = MyCommon.LRT_Select
                If (dst.Rows.Count > 0) Then
                    Dim tmp As Integer = MyCommon.NZ(dst.Rows(0).Item("EngineID"), -1)
                    Select Case tmp
                        Case 0 ' link to CM store health
                            Send("  <a href=""/logix/store-health-cm.aspx?filterhealth=2"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                        Case 2 ' link to CPE store health
                            Send("  <a href=""/logix/store-health-cpe.aspx?filterhealth=2"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                        Case 9
                            Send("  <a href=""/logix/UE/store-health-ue.aspx?filterhealth=2"" accesskey=""8"" " & TabOns(8) & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                        Case Else
                            Send("  <a href=""/logix/user-list.aspx"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                    End Select
                Else
                    Send("  <a href=""/logix/user-list.aspx"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
                End If
            End If
        Else
            Send("  <a href=""/logix/user-list.aspx"" accesskey=""8"" " & TabOns(8) & TabStyleOverride & "id=""tab8"" title=""" & Copient.PhraseLib.Lookup("term.administration", LanguageID) & """>" & Copient.PhraseLib.Lookup("term.admin", LanguageID) & "</a>")
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
                Send("  <a href=""/logix/authtransfer.aspx?SendToURI=" & EPMHostURI & """ accesskey=""9"" " & TabOns(9) & TabStyleOverride & "id=""tab9"" style=""width:70px"" title=""" & Copient.PhraseLib.Lookup(IntegrationVals.PhraseTerm, LanguageID) & """>" & Copient.PhraseLib.Lookup(IntegrationVals.PhraseTerm, LanguageID) & "</a>")
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
        Dim AdminUserID As Long
        Dim CPEInstalled As Boolean = False
        Dim CMInstalled As Boolean = False
        Dim UEInstalled As Boolean = False
        Dim TempDateTime As New DateTime
        Dim BannersEnabled As Boolean = False
        Dim EnableFuelPartner As Boolean = False
        'Dim CmInstallationType As Integer = 0
        Dim OperateAtEnterprise As Boolean
        Dim OfferEngineID As Int32 = 0
        Dim system_option80 As String = ""
        If DefaultLanguageID > 0 Then
            TempLanguageID = DefaultLanguageID
        Else
            TempLanguageID = LanguageID
        End If

        If {24, 27, 205, 208}.Contains(Subtabset) Then
            OfferEngineID = GetOfferEngineID(MyCommon, ID)
            If OfferEngineID = 9 Then
                system_option80 = MyCommon.Fetch_UE_SystemOption(80)
            Else
                system_option80 = MyCommon.Fetch_CPE_SystemOption(80)
            End If
        End If


        ' determine which engines are installed
        MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select EngineID from PromoEngines with (NoLock) where Installed=1;"
        dst = MyCommon.LRT_Select()
        For Each row In dst.Rows
            If row.Item("EngineID") = 0 Then CMInstalled = True
            If row.Item("EngineID") = 2 Then CPEInstalled = True
            If row.Item("EngineID") = 9 Then UEInstalled = True
        Next

        ' AL-6636 Remove UE@Enterprise system option since UE can now operate in mixed mode
        OperateAtEnterprise = ((CPEInstalled AndAlso (MyCommon.Fetch_CPE_SystemOption(91) = "1")) OrElse (UEInstalled AndAlso (MyCommon.Fetch_UE_SystemOption(91) = "1")))

        'To Get how many Server type locations exists in system 
        MyCommon.QueryStr = "SELECT LocationID from dbo.Locations WHERE LocationTypeID = 2 AND Deleted = 0;"
        Dim TempServer As System.Data.DataTable
        TempServer = MyCommon.LRT_Select

        'Are banners enabled
        MyCommon.QueryStr = "Select OptionValue from SystemOptions with (NoLock) where OptionID=66;"
        dt = MyCommon.LRT_Select()
        If dt.Rows.Count = 1 Then
            If dt.Rows(0).Item("OptionValue") = 0 Then BannersEnabled = False
            If dt.Rows(0).Item("OptionValue") = 1 Then BannersEnabled = True
        End If

        MyCommon.Close_LogixRT()

        EnableFuelPartner = MyCommon.Fetch_CM_SystemOption(56) = "1"

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
            If CMInstalled Or UEInstalled Then
                If (MyCommon.Fetch_SystemOption(167) = "1") Then
                    Send("  <a href=""/logix/Enhanced-extoffer-list.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.external", TempLanguageID) & "</a>")
                Else
                    Send("  <a href=""/logix/extoffer-list.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.external", TempLanguageID) & "</a>")
                End If
            Else
                Send("  <a href=""/logix/extoffer-list.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.external", TempLanguageID) & "</a>")
            End If
            If MyCommon.IsEngineInstalled(6) Then
                Send("  <a href=""/logix/CAM/CAM-offer-list.aspx"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.cam", TempLanguageID) & "</a>")
            End If
            If BannersEnabled Then
                Send("  <a href=""/logix/banneroffer-list.aspx"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"">" & Copient.PhraseLib.Lookup("term.banners", TempLanguageID) & "</a>")
            End If
            If MyCommon.IsEngineInstalled(9) Then
                ' Collision Section
                Send(" <a href=""/logix/CollidingOffers-list.aspx"" accesskey=""^"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.Collisions", TempLanguageID) & "</a>")
            End If
            If MyCommon.IsEngineInstalled(9) AndAlso MyLogix.UserRoles.OfferApproval Then
                ' PendingApproval Section
                Send(" <a href=""/logix/UE/PendingApproval.aspx"" accesskey=""$"" " & SubtabOns(8) & "id=""subtab8"">" & Copient.PhraseLib.Lookup("term.approvals", TempLanguageID) & "</a>")
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
                If (MyCommon.NZ(row.Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 11 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 12) Then
                    DisableSubTabs = True
                End If
                If (Not DisableSubTabs) Then
                    DisableSubTabs = (MyCommon.NZ(row.Item("DeployDeferred"), False) = True)
                End If
            Next
            MyCommon.Close_LogixRT()
            If (DisableSubTabs) Then
                Send("  <a href=""#"" accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""&amp;"" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</span></a>")
                Send("  <a href=""#"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;""><span class=""grey"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</span></a>")
                Send("  <a href=""/logix/offer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/offer-loc.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
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
            Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
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
            ' Offers section -- CPE offers
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
                If (system_option80 = "1") Then
                    TempDateTime = Convert.ToDateTime(MyCommon.NZ(row.Item("EndDate"), New DateTime(1990, 1, 1)))
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
            Else
                Send("  <a href=""/logix/offer-loc.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CPEoffer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CPEoffer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/CPEoffer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/CPEoffer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 25) Then
            ' Offers section -- CPE offer templates
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""("" " & SubtabOns(9) & "id=""subtab9"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/offer-loc.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.locations", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CPEoffer-rew.aspx?" & ID & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.rewards", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CPEoffer-con.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.conditions", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CPEoffer-gen.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/CPEoffer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 26) Then
            ' Offers section -- CPE deleted offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/CPEoffer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")
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
                If (MyCommon.NZ(row.Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 11 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 12) Then
                    DisableSubTabs = True
                End If
                If (Not DisableSubTabs) Then
                    DisableSubTabs = (MyCommon.NZ(row.Item("DeployDeferred"), -1) = True)
                End If
                If (system_option80 = "1") Then
                    TempDateTime = Convert.ToDateTime(MyCommon.NZ(row.Item("EndDate"), New DateTime(1990, 1, 1)))
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
            Send("  <a href=""/logix/offer-channels.aspx?" & ID & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.channels", TempLanguageID) & "</a>")
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
                If (MyCommon.NZ(row.Item("StatusFlag"), -1) = 2 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 11 Or MyCommon.NZ(row.Item("StatusFlag"), -1) = 12) Then
                    DisableSubTabs = True
                End If
                If (Not DisableSubTabs) Then
                    DisableSubTabs = (MyCommon.NZ(row.Item("DeployDeferred"), -1) = True)
                End If
                If (system_option80 = "1") Then
                    TempDateTime = Convert.ToDateTime(MyCommon.NZ(row.Item("EndDate"), New DateTime(1990, 1, 1)))
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
        ElseIf (Subtabset = 208) Then
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
                If (system_option80 = "1") Then
                    TempDateTime = Convert.ToDateTime(MyCommon.NZ(row.Item("EndDate"), New DateTime(1990, 1, 1)))
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
                'BZ2079: UE-feature-removal - hiding the notifications tab for UE offers
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

        ElseIf (Subtabset = 209) Then
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

        ElseIf (Subtabset = 210) Then
            ' Offers section -- UE deleted offers
            ID = "OfferID=" & ID
            Send("  <a href=""/logix/offer-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/temp-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.templates", TempLanguageID) & "</a>")
            If (MyLogix.UserRoles.ViewHistory = True) Then
                Send("  <a href=""/logix/offer-hist.aspx?" & ID & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/UE/UEoffer-sum.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.summary", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 30) Then
            ' Customers section
            ID = "CustPK=" & ID
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx?" & ID & """ accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyLogix.UserRoles.ViewHistory = True) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1") Then
                Send("  <a href=""/logix/CM-cashier-report.aspx" & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.CashierHistory", TempLanguageID) & "</a>")
            End If
            If (MyCommon.Fetch_SystemOption(142) = "1") Then
                Send("  <a href=""/logix/coupon-inquiry.aspx"" accesskey=""!"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.barcode", TempLanguageID) & "</a>")
            End If
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso
             (MyLogix.UserRoles.AccessTransactionInquiry = True)) Then
                Send("  <a href=""/logix/customer-transaction-inquiry.aspx?"" accesskey=""*"" " & SubtabOns(5) & "id=""subtab5""> " & Copient.PhraseLib.Lookup("term.transaction", TempLanguageID) & " " & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 31) Then
            ' Customers section -- groups
            If ID = 0 Then IsNew = True
            ID = "CustomerGroupID=" & ID

            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.Fetch_SystemOption(142) = "1") Then
                Send("  <a href=""/logix/coupon-inquiry.aspx"" accesskey=""!"" " & SubtabOns(5) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.barcode", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/cgroup-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/cgroup-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso
               (MyLogix.UserRoles.AccessTransactionInquiry = True)) Then
                Send("  <a href=""/logix/customer-transaction-inquiry.aspx?"" accesskey=""*"" " & SubtabOns(5) & "id=""subtab5""> " & Copient.PhraseLib.Lookup("term.transaction", TempLanguageID) & " " & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            End If
        ElseIf (Subtabset = 32) Then
            ' Customer section -- inquiry
            CustPK = "CustPK=" & ID
            If SecondaryID <> "" Then
                CardPK = "&amp;CardPK=" & MyCommon.Extract_Val(SecondaryID)
            End If
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.Fetch_SystemOption(142) = "1") Then
                Send("  <a href=""/logix/coupon-inquiry.aspx"" accesskey=""!"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.barcode", TempLanguageID) & "</a>")
            End If
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyLogix.UserRoles.ViewHistory = True) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1" AndAlso ID = 0) Then
                Send("  <a href=""/logix/CM-cashier-report.aspx" & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.CashierHistory", TempLanguageID) & "</a>")
            End If
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso
               (MyLogix.UserRoles.AccessTransactionInquiry = True)) Then
                Send("  <a href=""/logix/customer-transaction-inquiry.aspx?"" accesskey=""*"" " & SubtabOns(9) & "id=""subtab9""> " & Copient.PhraseLib.Lookup("term.transaction", TempLanguageID) & " " & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso ID > 0) Then
                Send("  <a href=""/logix/customer-hist.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""*"" " & SubtabOns(8) & "id=""subtab8"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            If (ID > 0) Then

                If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
                    Send("  <a href=""/logix/customer-prefs.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""&amp;"" " & SubtabOns(7) & "id=""subtab7"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.preferences", TempLanguageID) & "</a>")
                End If

                Send("  <a href=""/logix/customer-transactions.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.transactions", TempLanguageID) & "</a>")
                If (MyLogix.UserRoles.AccessAdjustmentsPage = True) Then
                    Send("  <a href=""/logix/customer-adjustments.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.adjustments", TempLanguageID) & "</a>")
                End If
                Send("  <a href=""/logix/customer-offers.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.offers", TempLanguageID) & "</a>")
                If (MyCommon.Fetch_SystemOption(107) = 1) Then
                    Send("  <a href=""/logix/customer-general.aspx?edit=Edit&editterms=" & ID & "&" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
                Else
                    Send("  <a href=""/logix/customer-general.aspx?" & CustPK & IIf(CardPK <> "", CardPK, "") & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.general", TempLanguageID) & "</a>")
                End If
            End If
        ElseIf (Subtabset = 33) Then
            ' CAM Customer section -- inquiry
            CustPK = "CustPK=" & ID
            If SecondaryID <> "" Then
                CardPK = "&amp;CardPK=" & MyCommon.Extract_Val(SecondaryID)
            End If
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.Fetch_SystemOption(142) = "1") Then
                Send("  <a href=""/logix/coupon-inquiry.aspx"" accesskey=""!"" " & SubtabOns(8) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.barcode", TempLanguageID) & "</a>")
            End If
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
        ElseIf (Subtabset = 34) Then
            'Coupon Inqury
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@""" & SubtabOns(2) & "  id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.Fetch_SystemOption(142) = "1") Then
                Send("  <a href=""/logix/coupon-inquiry.aspx"" accesskey=""!""  class=""on"" id=""subtab3"">" & Copient.PhraseLib.Lookup("term.barcode", TempLanguageID) & "</a>")
            End If
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso
               (MyLogix.UserRoles.AccessTransactionInquiry = True)) Then
                Send("  <a href=""/logix/customer-transaction-inquiry.aspx?"" accesskey=""*"" " & SubtabOns(9) & "id=""subtab9""> " & Copient.PhraseLib.Lookup("term.transaction", TempLanguageID) & " " & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/coupon-inquiry.aspx?" & """ accesskey=""*"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("coupon-inquiry.couponinquiry", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/coupon-batch-report.aspx?" & """ accesskey=""*"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & "Batch Report" & "</a>")
            'Send("  <a href=""/logix/customer-hist.aspx" & """ accesskey=""*"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & "Batch Report" & "</a>")
        ElseIf (Subtabset = 35) Then
            ' Customer section -- transaction inquiry
            Send("  <a href=""/logix/cgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            Send("  <a href=""/logix/customer-inquiry.aspx"" accesskey=""@"" " & SubtabOns(2) & " id=""subtab2"">" & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (MyCommon.Fetch_SystemOption(142) = "1") Then
                Send("  <a href=""/logix/coupon-inquiry.aspx"" accesskey=""!"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.barcode", TempLanguageID) & "</a>")
            End If
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyLogix.UserRoles.ViewHistory = True) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1" AndAlso ID = 0) Then
                Send("  <a href=""/logix/CM-cashier-report.aspx" & """ accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.CashierHistory", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/customer-transaction-inquiry.aspx"" accesskey=""@"" class=""on"" id=""subtab5""> " & Copient.PhraseLib.Lookup("term.transaction", TempLanguageID) & " " & Copient.PhraseLib.Lookup("term.inquiry", TempLanguageID) & "</a>")
            If (ID <> "") Then
                Send("  <a href=""/logix/customer-transactions.aspx?"" accesskey=""^"" class=""on"" id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.transactions", TempLanguageID) & "</a>")
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
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If MyCommon.Fetch_UE_SystemOption(148) = "1" And UEInstalled Then
                Send("  <a href=""/logix/tcp-list.aspx"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.trackablecoupon", TempLanguageID) & "</a>")
            End If

            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
                If (EnableFuelPartner) Then
                    Send("  <a href=""/logix/CM-fuel-partner-report.aspx"" accesskey=""#"" " & SubtabOns(7) & "id=""subtab4"">Fuel Partner Report</a>")
                End If
            End If
        ElseIf (Subtabset = 51) Then
            ' Programs section -- points
            If ID = 0 Then IsNew = True
            ID = "ProgramGroupID=" & ID
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If MyCommon.Fetch_UE_SystemOption(148) = "1" And UEInstalled Then
                Send("  <a href=""/logix/tcp-list.aspx"" accesskey=""$"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.trackablecoupon", TempLanguageID) & "</a>")
            End If

            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
                If (EnableFuelPartner) Then
                    Send("  <a href=""/logix/CM-fuel-partner-report.aspx"" accesskey=""#"" " & SubtabOns(7) & "id=""subtab4"">Fuel Partner Report</a>")
                End If
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/point-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/point-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(5) & "id=""subtab5""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 52) Then
            ' Programs section -- stored value
            If ID = 0 Then IsNew = True
            ID = "ProgramGroupID=" & ID
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If MyCommon.Fetch_UE_SystemOption(148) = "1" And UEInstalled Then
                Send("  <a href=""/logix/tcp-list.aspx"" accesskey=""$"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.trackablecoupon", TempLanguageID) & "</a>")
            End If

            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" " & SubtabOns(4) & "id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
                If (EnableFuelPartner) Then
                    Send("  <a href=""/logix/CM-fuel-partner-report.aspx"" accesskey=""#"" " & SubtabOns(7) & "id=""subtab4"">Fuel Partner Report</a>")
                End If
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/sv-hist.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(6) & "id=""subtab6"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/sv-edit.aspx?" & ID & """ accesskey=""#"" " & SubtabOns(5) & "id=""subtab5""  style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 53) Then
            ' Programs section -- promovars (In DP-only environments, this is under the PromoVars tab)
            ' Promovars deprecated
        ElseIf (Subtabset = 54) Then
            ' Programs section -- Advanced Limits
            If ID = 0 Then IsNew = True
            ID = "LimitID=" & ID
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/point-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.points", TempLanguageID) & "</a>")
            End If
            If (CPEInstalled OrElse CMInstalled OrElse UEInstalled) Then
                Send("  <a href=""/logix/sv-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.storedvalue", TempLanguageID) & "</a>")
            End If
            If MyCommon.Fetch_UE_SystemOption(148) = "1" And UEInstalled Then
                Send("  <a href=""/logix/tcp-list.aspx"" accesskey=""$"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.trackablecoupon", TempLanguageID) & "</a>")
            End If

            If (CMInstalled) Then
                Send("  <a href=""/logix/CM-advlimit-list.aspx"" accesskey=""#"" class=""on"" id=""subtab4"">" & Copient.PhraseLib.Lookup("term.advlimits", TempLanguageID) & "</a>")
                If (EnableFuelPartner) Then
                    Send("  <a href=""/logix/CM-fuel-partner-report.aspx"" accesskey=""#"" " & SubtabOns(7) & "id=""subtab4"">Fuel Partner Report</a>")
                End If
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

            If OperateAtEnterprise OrElse TempServer.Rows.Count > 0 Then
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/store-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            End If
            If UEInstalled Then
                Send("  <a href=""/logix/terminal-sets-list.aspx"" accesskey=""@"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.terminalsets", TempLanguageID) & "</a>")
                If MyCommon.Fetch_UE_SystemOption(135) = "1" Then
                    Send("  <a href=""/logix/uom-sets-list.aspx"" accesskey=""$"" " & SubtabOns(4) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.uomsets", TempLanguageID) & "</a>")
                End If
            End If

        ElseIf (Subtabset = 71) Then
            ' Locations section -- groups
            If ID = "0" Then IsNew = True
            ID = "LocationGroupID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            If OperateAtEnterprise OrElse TempServer.Rows.Count > 0 Then
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/store-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            End If
            If UEInstalled Then
                Send("  <a href=""/logix/terminal-sets-list.aspx"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.terminalsets", TempLanguageID) & "</a>")
                If MyCommon.Fetch_UE_SystemOption(135) = "1" Then
                    Send("  <a href=""/logix/uom-sets-list.aspx"" accesskey=""&"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.uomsets", TempLanguageID) & "</a>")
                End If
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/lgroup-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/lgroup-edit.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 72) Then
            ' Locations section -- stores
            If ID = "0" Then IsNew = True
            ID = "LocationID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            If OperateAtEnterprise OrElse TempServer.Rows.Count > 0 Then
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" " & SubtabOns(3) & "id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/store-list.aspx"" accesskey=""@"" class=""on"" id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            End If
            If UEInstalled Then
                Send("  <a href=""/logix/terminal-sets-list.aspx"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.terminalsets", TempLanguageID) & "</a>")
                If MyCommon.Fetch_UE_SystemOption(135) = "1" Then
                    Send("  <a href=""/logix/uom-sets-list.aspx"" accesskey=""&"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.uomsets", TempLanguageID) & "</a>")
                End If
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/store-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/store-edit.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")
        ElseIf (Subtabset = 73) Then
            ' Locations section -- servers (i.e., stores of LocationTypeID 2)
            If ID = "0" Then IsNew = True
            ID = "LocationID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            If OperateAtEnterprise OrElse TempServer.Rows.Count > 0 Then
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" class=""on"" id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/store-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            End If
            If UEInstalled Then
                Send("  <a href=""/logix/terminal-sets-list.aspx"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.terminalsets", TempLanguageID) & "</a>")
                If MyCommon.Fetch_UE_SystemOption(135) = "1" Then
                    Send("  <a href=""/logix/uom-sets-list.aspx"" accesskey=""&"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.uomsets", TempLanguageID) & "</a>")
                End If
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/store-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/store-edit.aspx?" & ID & IIf(IsNew, "&LocationTypeID=2", "") & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 74) Then
            ' Locations section -- Terminal sets
            If ID = "0" Then IsNew = True
            ID = "TerminalSetID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            If OperateAtEnterprise OrElse TempServer.Rows.Count > 0 Then
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/store-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            End If
            If UEInstalled Then
                Send("  <a href=""/logix/terminal-sets-list.aspx"" accesskey=""^"" class=""on"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.terminalsets", TempLanguageID) & "</a>")
                If MyCommon.Fetch_UE_SystemOption(135) = "1" Then
                    Send("  <a href=""/logix/uom-sets-list.aspx"" accesskey=""&"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.uomsets", TempLanguageID) & "</a>")
                End If
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/terminal-sets-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/terminal-sets-edit.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 75) Then
            ' Locations section -- UOM sets
            If ID = "0" Then IsNew = True
            ID = "UOMSetID=" & ID
            Send("  <a href=""/logix/lgroup-list.aspx"" accesskey=""!"" " & SubtabOns(1) & "id=""subtab1"">" & Copient.PhraseLib.Lookup("term.groups", TempLanguageID) & "</a>")
            If OperateAtEnterprise OrElse TempServer.Rows.Count > 0 Then
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=1"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
                Send("  <a href=""/logix/store-list.aspx?LocationTypeID=2"" accesskey=""#"" id=""subtab3"">" & Copient.PhraseLib.Lookup("term.servers", TempLanguageID) & "</a>")
            Else
                Send("  <a href=""/logix/store-list.aspx"" accesskey=""@"" " & SubtabOns(2) & "id=""subtab2"">" & Copient.PhraseLib.Lookup("term.stores", TempLanguageID) & "</a>")
            End If
            If UEInstalled Then
                Send("  <a href=""/logix/terminal-sets-list.aspx"" accesskey=""^"" " & SubtabOns(6) & "id=""subtab6"">" & Copient.PhraseLib.Lookup("term.terminalsets", TempLanguageID) & "</a>")
                If MyCommon.Fetch_UE_SystemOption(135) = "1" Then
                    Send("  <a href=""/logix/uom-sets-list.aspx"" accesskey=""&"" class=""on"" " & SubtabOns(7) & "id=""subtab7"">" & Copient.PhraseLib.Lookup("term.uomsets", TempLanguageID) & "</a>")
                End If
            End If
            If (MyLogix.UserRoles.ViewHistory = True AndAlso IsNew = False) Then
                Send("  <a href=""/logix/uom-sets-hist.aspx?" & ID & """ accesskey=""%"" " & SubtabOns(5) & "id=""subtab5"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.history", TempLanguageID) & "</a>")
            End If
            Send("  <a href=""/logix/uom-sets-edit.aspx?" & ID & """ accesskey=""$"" " & SubtabOns(4) & "id=""subtab4"" style=""float: right; left: auto; right: 11px;"">" & Copient.PhraseLib.Lookup("term.edit", TempLanguageID) & "</a>")

        ElseIf (Subtabset = 8) Then
            ' Administration section
            Dim PageUrl As String = "/logix/store-health-cpe.aspx"
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
                If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(91) = "1" Then
                    PageUrl = "/logix/UE/UEServerHealthSummary.aspx"
                Else
                    If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                        MyCommon.QueryStr = "select EngineID from PromoEngines where Installed=1 and DefaultEngine=1 and EngineID in (0,2,9);"
                        dst = MyCommon.LRT_Select
                        If (dst.Rows.Count > 0) Then
                            Dim tmp As Integer = MyCommon.NZ(dst.Rows(0).Item("EngineID"), -1)
                            Select Case tmp
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
            If (MyLogix.UserRoles.AccessReports AndAlso MyCommon.Fetch_SystemOption(267) = "1") Then
                Send("  <a href=""/logix/external-list.aspx"" accesskey=""*"" " & SubtabOns(10) & "id=""subtab10"">" & Copient.PhraseLib.Lookup("term.external", TempLanguageID) & "</a>")
            End If
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
            Send("  <a href=""/logix/customer-inquiry.aspx?token=nothing" & subtaburl & """ accesskey=""!"" class=""on"" id=""subtab1"">" & Copient.PhraseLib.Lookup("term.customers", TempLanguageID) & "</a>")
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

    Public Sub Send_Save(Optional ByVal Attributes As String = "", Optional ByVal ButtonType As String = "submit", Optional ByVal onclk As String = "")
        ' Find out if this is an offer were trying to send
        Send("<input type=""" & ButtonType & """ accesskey=""s"" class=""regular"" id=""save"" name=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """" & onclk & Attributes & " />")
    End Sub

    Public Sub Send_Delete(Optional ByVal Attributes As String = "")
        Dim OnClickAttrib As String = ""

        If (Attributes <> "") Then
            OnClickAttrib = ParseAttribute(Attributes, "onclick", False, True)
        End If

        Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""" & OnClickAttrib & "if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Edit(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" accesskey=""e"" class=""regular"" id=""edit"" name=""edit"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_New(Optional ByVal Attributes As String = "")
        Sendb("<input type=""submit"" accesskey=""n"" class=""regular"" id=""new"" name=""new"" value=""" & Copient.PhraseLib.Lookup("term.new", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Import(Optional ByVal Attributes As String = "")
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim AdminUserID As Long
        Dim myUrl As String = ""
        Dim dt As System.Data.DataTable = Nothing
        Dim row As System.Data.DataRow = Nothing
        Dim BannersEnabled As Boolean = False
        Dim AllowMultipleBanners As Boolean = False
        Dim AllBannersCheckBox As String = ""
        Dim AllBannersCount As Integer = 1

        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        myUrl = Request.CurrentExecutionFilePath
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
        Send("    <form action=""" & myUrl & """ method=""post"" enctype=""multipart/form-data"">")
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
        ' Send("<div id=""divfile"" style=""height:0px;overflow:hidden"">")
        'Send("<input type=""file"" id=""fileInput"" name=""fileInput"" onchange=""fileonclick()"" accept=""*"" />")
        ' Send("</div>")
        ' Send("<button type=""button"" onclick=""chooseFile();"">"&Copient.PhraseLib.Lookup("term.browse", LanguageID) &"</button>")
        '     Send("<label id=""lblfileupload"" name=""lblfileupload"">"&Copient.PhraseLib.Lookup("term.nofilesselected", LanguageID)&"</label>")
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
        Send("<input type=""button"" class=""regular"" id=""addoffer"" name=""addoffer"" value=""" & Copient.PhraseLib.Lookup("term.addoffer", LanguageID) & "..."" onclick=""openPopup('/logix/customer-addoffer.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');""" & Attributes & " />")
    End Sub

    Public Sub Send_AssignFolders(Optional ByVal OfferID As Long = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""assignfolders"" name=""assignfolders"" value=""" & Copient.PhraseLib.Lookup("term.assignfolders", LanguageID) & "..."" onclick=""openPopup('/logix/folder-browse.aspx?OfferID=" & OfferID & "');""" & Attributes & " />")
    End Sub

    Public Sub Send_CAMAddOffer(Optional ByVal CustomerPK As Integer = 0, Optional ByVal CardPK As Integer = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""addoffer"" name=""addoffer"" value=""" & Copient.PhraseLib.Lookup("term.addoffer", LanguageID) & "..."" onclick=""openPopup('/logix/CAM/CAM-customer-addoffer.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');""" & Attributes & " />")
    End Sub

    Public Sub Send_CancelDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""canceldeploy"" name=""canceldeploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.canceldeploy", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.canceldeploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Close(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""close"" name=""close"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_CopyOffer(Optional ByVal IsTemplate As Boolean = False, Optional ByVal Attributes As String = "")
        Sendb("<input type=""submit"" class=""regular"" id=""copyoffer"" name=""copyoffer"" ")
        If IsTemplate Then
            Send("value=""" & Copient.PhraseLib.Lookup("term.copytemplate", LanguageID) & """" & Attributes & " />")
        Else
            Send("value=""" & Copient.PhraseLib.Lookup("term.copyoffer", LanguageID) & """" & Attributes & " />")
        End If
    End Sub

    Public Sub Send_CopyGroup(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""copygroup"" name=""copygroup"" value=""" & Copient.PhraseLib.Lookup("term.copygroup", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_CustomerNotes(Optional ByVal CustomerPK As Integer = 0, Optional ByVal CardPK As Integer = 0, Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""addnote"" name=""addnote"" value=""" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & "..."" onclick=""openPopup('/logix/customer-notes.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');"" />")
    End Sub

    Public Sub Send_DeferDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""deferdeploy"" name=""deferdeploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.deferdeploy", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_DeferDeployConditional(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""deferdeploy"" name=""deferdeploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.deferdeploy-condition", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_DeferDeployCollision(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""deferdeploycollision"" name=""deferdeploycollision"" onclick=""confirmDeployCollision(true);"" value=""" & Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Deploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""deploy"" name=""deploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.deploy", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_DeployConditional(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""deploy"" name=""deploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.deploy-condition", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_DeployCollision(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""deploycollision"" name=""deploycollision"" onclick=""confirmDeployCollision(false);"" value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Download(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""download"" name=""download"" value=""" & Copient.PhraseLib.Lookup("term.download", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_ExportToEDW(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""exportedw"" name=""exportedw"" value=""" & Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Export(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""export"" name=""export"" value=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_ExportCME(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""exportCME"" name=""exportCME"" value=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ")""" & Attributes & " />")
    End Sub

    Public Sub Send_ExportCRM(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""exportCRM"" name=""exportCRM"" value=""" & Copient.PhraseLib.Lookup("term.export", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.crm", LanguageID) & ")""" & Attributes & " />")
    End Sub

    Public Sub Send_PreValidate(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""preValidate"" name=""preValidate"" value=""" & Copient.PhraseLib.Lookup("term.prevalidate", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_PostValidate(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""postValidate"" name=""postValidate"" value=""" & Copient.PhraseLib.Lookup("term.postvalidate", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_ReadyToDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""readyToDeploy"" name=""readyToDeploy"" value=""" & Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_GenerateIPL(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""generateIPL"" name=""generateIPL"" value=""" & Copient.PhraseLib.Lookup("term.generateipl", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_OfferFromTemp(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""OfferFromTemp"" name=""OfferFromTemp"" onclick=""javascript: return assignNoofDuplicateOffers(true);"" value=""" & Copient.PhraseLib.Lookup("term.newfromtemp", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_ReDeploy(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""redeploy"" name=""redeploy"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.redeploy", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.redeploy", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_RevalidateAll(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""revalidateall"" name=""revalidateall"" value=""" & Copient.PhraseLib.Lookup("term.revalidate-all", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Saveastemp(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""saveastemp"" name=""saveastemp"" value=""" & Copient.PhraseLib.Lookup("term.saveastemp", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_SendOutbound(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""sendoutbound"" name=""sendoutbound"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.sendoutbound", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID) & """" & Attributes & " />")
    End Sub

    Public Sub Send_Upload(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""upload"" name=""upload"" value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """" & Attributes & " onclick=""javascript:{document.getElementById('uploader').style.display='block';}"" />")
    End Sub

    Public Sub Send_ModifyGroup(Optional ByVal Attributes As String = "")
        Send("<input type=""button"" class=""regular"" id=""modifygroup"" name=""modifygroup"" value=""" & Copient.PhraseLib.Lookup("term.modifygroup", LanguageID) & """" & Attributes & " onclick=""javascript:PopUpModifyGroup();"" />")
    End Sub

    Public Sub Send_Listbar(Optional ByVal linesPerPage As Integer = 0, Optional ByVal sizeOfData As Integer = 0, Optional ByVal PageNum As Integer = 0, Optional ByVal searchString As String = "", Optional ByVal SortText As String = "", Optional ByVal ShowExpired As String = "", Optional ByVal QueryString As String = "", Optional ByVal HideSearcher As Boolean = False, Optional ByVal AdminUserID As Integer = 0)
        Dim myUrl As String
        Dim startVal As Integer
        Dim endVal As Integer
        Dim expiredString As String
        Dim filterString As String = ""
        Dim filterUserString As String = ""
        Dim filterhealth As String = ""
        Dim filterOffer As String = ""
        Dim filterUser As String = ""
        Dim CustomerInquiry As String = ""
        Dim MyCommon As New Copient.CommonInc
        Dim rst As System.Data.DataTable
        Dim dt As System.Data.DataRow

        Dim selectetdEngine As String = ""
        MyCommon.Open_LogixRT()

        Dim Logix As New Copient.LogixInc
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        myUrl = Request.CurrentExecutionFilePath
        expiredString = "&amp;ShowExpired=" & ShowExpired
        filterString = "&amp;filterOffer=" & Request.QueryString("filterOffer")
        filterUserString = "&amp;filterUser=" & Request.QueryString("filterUser")
        If (bEnableRestrictedAccessToUEOfferBuilder) Then
            If ((myUrl = "/logix/offer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers) OrElse ((myUrl = "/logix/Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers) OrElse (myUrl = "/logix/Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.AccessTranslatedUEOffers))) Then

                If (Not String.IsNullOrEmpty(Request.QueryString("filterengine"))) Then selectetdEngine = Request.QueryString("filterengine")

                If (Not String.IsNullOrEmpty(Request.QueryString("engine"))) Then
                    If (Request.QueryString("engine").Trim.ToLower.Equals("cm")) Then
                        selectetdEngine = 0
                    ElseIf (Request.QueryString("engine").Trim.ToLower.Equals("ue")) Then
                        selectetdEngine = 9
                    End If
                ElseIf (Request.Form("engine") <> Nothing) Then
                    If (Request.Form("engine").Trim.ToLower.Equals("cm")) Then
                        selectetdEngine = 0
                    ElseIf (Request.Form("engine").Trim.ToLower.Equals("ue")) Then
                        selectetdEngine = 9
                    End If
                End If
                If (String.IsNullOrEmpty(selectetdEngine)) Then selectetdEngine = "0"
                filterString &= "&amp;filterengine=" & selectetdEngine
            End If
        End If
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


        If (myUrl = "/logix/store-health-cpe.aspx" OrElse myUrl = "/logix/store-health-cm.aspx" OrElse myUrl = "/logix/store-health-UE.aspx" OrElse myUrl = "/logix/offer-health.aspx") Then
            filterString = "&amp;filterhealth=" & Request.QueryString("filterhealth")
        End If

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
            If (myUrl = "/logix/offer-list.aspx") Or (myUrl = "/logix/Enhanced-extoffer-list.aspx") Then
                Send("   <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
                Send("   <input type=""button"" id=""advsearch"" name=""advsearch"" value=""..."" alt=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ onclick=""launchAdvSearch();"" /><br />")
            ElseIf (myUrl = "/logix/CM-cashier-report.aspx") Then
                Send("   <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
                Send("   <input type=""button"" id=""advsearch"" name=""advsearch"" value=""..."" alt=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ onclick=""launchAdvSearch();"" /><br />")
            ElseIf (myUrl = "/logix/pgroup-list.aspx") Then
                Send("   <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
                Send("   <input type=""button"" id=""advsearch"" name=""advsearch"" value=""..."" alt=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.advancedsearch", LanguageID) & """ onclick=""launchAdvSearch();"" /><br />")
            Else
                Send("   <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ /><br />")
            End If
        End If
        Send("  </div>")
        If (bEnableRestrictedAccessToUEOfferBuilder) Then
            If ((myUrl = "/logix/offer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers) OrElse ((myUrl = "/logix/Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers) OrElse (myUrl = "/logix/Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.AccessTranslatedUEOffers))) Then
                Send("  <div id=""paginator1"">")
            Else
                Send("  <div id=""paginator"">")
            End If
        Else
            Send("  <div id=""paginator"">")
        End If
        If (searchString = "") Then
        Else
            Sendb("   <span class=""printonly"">")
            Sendb(Copient.PhraseLib.Lookup("term.searchterms", LanguageID) & ": """ & searchString & """<br />")
            Send("</span>")
        End If
        If (Not String.IsNullOrEmpty(searchString)) Then searchString = Server.UrlEncode(searchString)
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
        ElseIf (myUrl = "/logix/pgroup-list.aspx") Then
            Send("   &nbsp;[ <b>" & startVal & "</b> - <b>" & endVal & "</b> " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " <b>" & sizeOfData & " matches</b> ]&nbsp;")
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
        If (bEnableRestrictedAccessToUEOfferBuilder) Then
            If ((myUrl = "/logix/offer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers) OrElse ((myUrl = "/logix/Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers) OrElse (myUrl = "/logix/Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.AccessTranslatedUEOffers))) Then
                Send("   <div id=""enginefilter"" title=""enginefilter""> ")
                Send("   <select id=""filterengine"" name=""filterengine"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
                Send("    <option value=""0""" & IIf(selectetdEngine = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.cm", LanguageID) & "</option>")
                Send("    <option value=""9""" & IIf(selectetdEngine = "9", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.ue", LanguageID) & "</option>")
                Send("   </select> ")
                Send(" </div>")
            End If
        End If
        Send("  <div id=""filter"" title=""" & Copient.PhraseLib.Lookup("term.filter", LanguageID) & """>")
        If (myUrl = "/logix/offer-list.aspx") Then
            If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Logix.UserRoles.CreateUEOffers) Then
                Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
                filterOffer = Request.QueryString("filterOffer")
                If filterOffer = "" Then filterOffer = "1"
                If (selectetdEngine = 9) Then
                    If (filterOffer = "5" OrElse filterOffer = "6" OrElse filterOffer = "7" OrElse filterOffer = "8") Then filterOffer = 0
                End If
                Send("    <option value=""0""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
                Send("    <option value=""1""" & IIf(filterOffer = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
                Send("    <option value=""2""" & IIf(filterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
                If (CustomerInquiry = "") Then
                    Send("    <option value=""3""" & IIf(filterOffer = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showdiscrepancies", LanguageID) & "</option>")
                End If
                'Allow users to search for offers created by all users under the same banner
                If (MyCommon.Fetch_SystemOption(130) = "1" AndAlso MyCommon.Fetch_SystemOption(66) = "1") Then
                    Send("    <option value=""4""" & IIf(filterOffer = "4", " selected=""selected""", "") & ">" & "View Offers By User" & "</option>")
                End If
                If (selectetdEngine = "0") Then
                    Send("    <option value=""8""" & IIf(filterOffer = "8", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlydraft", LanguageID) & "</option>")
                End If
                If (CustomerInquiry = "") Then
                    If (selectetdEngine = "0" AndAlso MyCommon.Fetch_CM_SystemOption(74) = "1") Then
                        Dim sSystemType As String = MyCommon.Fetch_CM_SystemOption(77)
                        ' Production only
                        If (sSystemType <> "1" AndAlso sSystemType <> "2") Then
                            Send("    <option value=""5""" & IIf(filterOffer = "5", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-prevalidate", LanguageID) & "</option>")
                            Send("    <option value=""6""" & IIf(filterOffer = "6", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-postvalidate", LanguageID) & "</option>")
                            Send("    <option value=""7""" & IIf(filterOffer = "7", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-readytodeploy", LanguageID) & "</option>")
                        End If
                    End If
                End If
                Send("  </select>")
                If (filterOffer = "4" AndAlso myUrl = "/logix/offer-list.aspx") Then
                    Dim datatable As System.Data.DataTable
                    Dim datarow As System.Data.DataRow
                    Dim k As Integer = 0
                    MyCommon.QueryStr = "select UserName, AdminUserID from AdminUsers " &
                                "where AdminUserID in (select distinct AdminUserID " &
                                "from AdminUserBanners Where BannerID IN " &
                                "(select BannerID from AdminUserBanners where AdminUserID = " & AdminUserID & " ));"
                    datatable = MyCommon.LRT_Select
                    'Send("  <div id=""userFilter"" title=""" & Copient.PhraseLib.Lookup("term.user", LanguageID) & """>")
                    Send("   <select id=""filterUser"" name=""filterUser"" onchange=""handleUserFilterRegEx(this.options[this.selectedIndex].value);"">")
                    filterUser = Request.QueryString("filterUser")
                    If filterUser = "" Then filterUser = "0"
                    For Each datarow In datatable.Rows
                        Send("    <option value=""" & MyCommon.NZ(datarow.Item("AdminUserID"), 0) & """" & IIf(filterUser = MyCommon.NZ(datarow.Item("AdminUserID"), 0).ToString, " selected=""selected""", "") & ">" & MyCommon.NZ(datarow.Item("UserName"), "") & "</option>")
                        k = k + 1
                    Next
                    Send("  </select>")
                End If
            Else
                Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
                filterOffer = Request.QueryString("filterOffer")
                If filterOffer = "" Then filterOffer = "1"
                Send("    <option value=""0""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
                Send("    <option value=""1""" & IIf(filterOffer = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
                Send("    <option value=""2""" & IIf(filterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
                If (CustomerInquiry = "") Then
                    Send("    <option value=""3""" & IIf(filterOffer = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showdiscrepancies", LanguageID) & "</option>")
                End If
                'Allow users to search for offers created by all users under the same banner
                If (MyCommon.Fetch_SystemOption(130) = "1" AndAlso MyCommon.Fetch_SystemOption(66) = "1") Then
                    Send("    <option value=""4""" & IIf(filterOffer = "4", " selected=""selected""", "") & ">" & "View Offers By User" & "</option>")
                End If
                If (MyCommon.IsEngineInstalled(0)) Then
                    Send("    <option value=""8""" & IIf(filterOffer = "8", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlydraft", LanguageID) & "</option>")
                End If
                If (CustomerInquiry = "") Then
                    If (MyCommon.IsEngineInstalled(0) AndAlso MyCommon.Fetch_CM_SystemOption(74) = "1") Then
                        Dim sSystemType As String = MyCommon.Fetch_CM_SystemOption(77)
                        ' Production only
                        If (sSystemType <> "1" AndAlso sSystemType <> "2") Then
                            Send("    <option value=""5""" & IIf(filterOffer = "5", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-prevalidate", LanguageID) & "</option>")
                            Send("    <option value=""6""" & IIf(filterOffer = "6", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-postvalidate", LanguageID) & "</option>")
                            Send("    <option value=""7""" & IIf(filterOffer = "7", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-readytodeploy", LanguageID) & "</option>")
                        End If
                    End If
                End If
                Send("  </select>")
                If (filterOffer = "4" AndAlso myUrl = "/logix/offer-list.aspx") Then
                    Dim datatable As System.Data.DataTable
                    Dim datarow As System.Data.DataRow
                    Dim k As Integer = 0
                    MyCommon.QueryStr = "select UserName, AdminUserID from AdminUsers " &
                                "where AdminUserID in (select distinct AdminUserID " &
                                "from AdminUserBanners Where BannerID IN " &
                                "(select BannerID from AdminUserBanners where AdminUserID = " & AdminUserID & " ));"
                    datatable = MyCommon.LRT_Select
                    'Send("  <div id=""userFilter"" title=""" & Copient.PhraseLib.Lookup("term.user", LanguageID) & """>")
                    Send("   <select id=""filterUser"" name=""filterUser"" onchange=""handleUserFilterRegEx(this.options[this.selectedIndex].value);"">")
                    filterUser = Request.QueryString("filterUser")
                    If filterUser = "" Then filterUser = "0"
                    For Each datarow In datatable.Rows
                        Send("    <option value=""" & MyCommon.NZ(datarow.Item("AdminUserID"), 0) & """" & IIf(filterUser = MyCommon.NZ(datarow.Item("AdminUserID"), 0).ToString, " selected=""selected""", "") & ">" & MyCommon.NZ(datarow.Item("UserName"), "") & "</option>")
                        k = k + 1
                    Next
                    Send("  </select>")
                End If
            End If
        ElseIf (myUrl = "/logix/extoffer-list.aspx") Then
            Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
            filterOffer = Request.QueryString("filterOffer")
            If filterOffer = "" Then filterOffer = "0"
            MyCommon.QueryStr = "select ExtInterfaceID, Name, PhraseID from ExtCRMInterfaces with (NoLock) where ExtInterfaceID>0;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                Send("  <option value=""-1""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.alloffers", LanguageID) & "</option>")
                For Each dt In rst.Rows
                    Send("  <option value=""" & dt.Item("ExtInterfaceID") & """" & IIf(filterOffer = dt.Item("ExtInterfaceID"), " selected=""selected""", "") & ">" & IIf(IsDBNull(dt.Item("PhraseID")), Convert.ToString(dt.Item("Name")), Copient.PhraseLib.Lookup(MyCommon.NZ(dt.Item("PhraseID"), 0), LanguageID)) & "</option>")
                Next
                Send(" </select>")
            End If
        ElseIf (myUrl = "/logix/Enhanced-extoffer-list.aspx") Then
            If ((bEnableRestrictedAccessToUEOfferBuilder AndAlso Logix.UserRoles.CreateUEOffers) OrElse (bEnableRestrictedAccessToUEOfferBuilder AndAlso Logix.UserRoles.AccessTranslatedUEOffers)) Then
                Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
                filterOffer = Request.QueryString("filterOffer")
                If filterOffer = "" Then filterOffer = "1"
                If (selectetdEngine = 9) Then
                    If (filterOffer = "5" OrElse filterOffer = "6" OrElse filterOffer = "7" OrElse filterOffer = "8") Then filterOffer = 0
                End If
                Send("    <option value=""0""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
                Send("    <option value=""1""" & IIf(filterOffer = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
                Send("    <option value=""2""" & IIf(filterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
                If (CustomerInquiry = "") Then
                    Send("    <option value=""3""" & IIf(filterOffer = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showdiscrepancies", LanguageID) & "</option>")
                End If
                'Allow users to search for offers created by all users under the same banner
                If (MyCommon.Fetch_SystemOption(130) = "1" AndAlso MyCommon.Fetch_SystemOption(66) = "1") Then
                    Send("    <option value=""4""" & IIf(filterOffer = "4", " selected=""selected""", "") & ">" & "View Offers By User" & "</option>")
                End If
                If (selectetdEngine = "0") Then
                    Send("    <option value=""8""" & IIf(filterOffer = "8", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlydraft", LanguageID) & "</option>")
                End If
                If (CustomerInquiry = "") Then
                    If (selectetdEngine = "0" AndAlso MyCommon.Fetch_CM_SystemOption(74) = "1") Then
                        Dim sSystemType As String = MyCommon.Fetch_CM_SystemOption(77)
                        ' Production only
                        If (sSystemType <> "1" AndAlso sSystemType <> "2") Then
                            Send("    <option value=""5""" & IIf(filterOffer = "5", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-prevalidate", LanguageID) & "</option>")
                            Send("    <option value=""6""" & IIf(filterOffer = "6", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-postvalidate", LanguageID) & "</option>")
                            Send("    <option value=""7""" & IIf(filterOffer = "7", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-readytodeploy", LanguageID) & "</option>")
                        End If
                    End If
                End If
                Send("  </select>")
                If (filterOffer = "4" AndAlso myUrl = "/logix/offer-list.aspx") Then
                    Dim datatable As System.Data.DataTable
                    Dim datarow As System.Data.DataRow
                    Dim k As Integer = 0
                    MyCommon.QueryStr = "select UserName, AdminUserID from AdminUsers " &
                                "where AdminUserID in (select distinct AdminUserID " &
                                "from AdminUserBanners Where BannerID IN " &
                                "(select BannerID from AdminUserBanners where AdminUserID = " & AdminUserID & " ));"
                    datatable = MyCommon.LRT_Select
                    'Send("  <div id=""userFilter"" title=""" & Copient.PhraseLib.Lookup("term.user", LanguageID) & """>")
                    Send("   <select id=""filterUser"" name=""filterUser"" onchange=""handleUserFilterRegEx(this.options[this.selectedIndex].value);"">")
                    filterUser = Request.QueryString("filterUser")
                    If filterUser = "" Then filterUser = "0"
                    For Each datarow In datatable.Rows
                        Send("    <option value=""" & MyCommon.NZ(datarow.Item("AdminUserID"), 0) & """" & IIf(filterUser = MyCommon.NZ(datarow.Item("AdminUserID"), 0).ToString, " selected=""selected""", "") & ">" & MyCommon.NZ(datarow.Item("UserName"), "") & "</option>")
                        k = k + 1
                    Next
                    Send("  </select>")
                End If
            Else
                Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
                filterOffer = Request.QueryString("filterOffer")
                If filterOffer = "" Then filterOffer = "1"
                Send("    <option value=""0""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
                Send("    <option value=""1""" & IIf(filterOffer = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
                Send("    <option value=""2""" & IIf(filterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
                If (CustomerInquiry = "") Then
                    Send("    <option value=""3""" & IIf(filterOffer = "3", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showdiscrepancies", LanguageID) & "</option>")
                End If
                If (MyCommon.IsEngineInstalled(0)) Then
                    Send("    <option value=""8""" & IIf(filterOffer = "8", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlydraft", LanguageID) & "</option>")
                End If
                'Allow users to search for offers created by all users under the same banner
                If (MyCommon.Fetch_SystemOption(130) = "1" AndAlso MyCommon.Fetch_SystemOption(66) = "1") Then
                    Send("    <option value=""4""" & IIf(filterOffer = "4", " selected=""selected""", "") & ">" & "View Offers By User" & "</option>")
                End If
                If (CustomerInquiry = "") Then
                    If (MyCommon.IsEngineInstalled(0) AndAlso MyCommon.Fetch_CM_SystemOption(74) = "1") Then
                        Dim sSystemType As String = MyCommon.Fetch_CM_SystemOption(77)
                        ' Production only
                        If (sSystemType <> "1" AndAlso sSystemType <> "2") Then
                            Send("    <option value=""5""" & IIf(filterOffer = "5", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-prevalidate", LanguageID) & "</option>")
                            Send("    <option value=""6""" & IIf(filterOffer = "6", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-postvalidate", LanguageID) & "</option>")
                            Send("    <option value=""7""" & IIf(filterOffer = "7", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonly-readytodeploy", LanguageID) & "</option>")
                        End If
                    End If
                End If

                If (filterOffer = "4") Then
                    Dim datatable As System.Data.DataTable
                    Dim datarow As System.Data.DataRow
                    Dim k As Integer = 0
                    MyCommon.QueryStr = "select UserName, AdminUserID from AdminUsers " &
                                "where AdminUserID in (select distinct AdminUserID " &
                                "from AdminUserBanners Where BannerID IN " &
                                "(select BannerID from AdminUserBanners where AdminUserID = " & AdminUserID & " ));"
                    datatable = MyCommon.LRT_Select
                    Send("   <select id=""filterUser"" name=""filterUser"" onchange=""handleUserFilterRegEx(this.options[this.selectedIndex].value);"">")
                    filterUser = Request.QueryString("filterUser")
                    If filterUser = "" Then filterUser = "0"
                    For Each datarow In datatable.Rows
                        Send("    <option value=""" & MyCommon.NZ(datarow.Item("AdminUserID"), 0) & """" & IIf(filterUser = MyCommon.NZ(datarow.Item("AdminUserID"), 0).ToString, " selected=""selected""", "") & ">" & MyCommon.NZ(datarow.Item("UserName"), "") & "</option>")
                        k = k + 1
                    Next
                    Send("  </select>")
                End If
                Send("  </select>")
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
            Send("  <option value=""-1""" & IIf(filterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.showall", LanguageID) & "</option>")
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
        ElseIf (myUrl = "/logix/store-health-cpe.aspx" Or myUrl = "/logix/store-health-cm.aspx" OrElse myUrl = "/logix/UE/store-health-UE.aspx") Then
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
            Send("    <option value=""6""" & IIf(filterhealth = "6", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.allfailovers", LanguageID) & "</option>")
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

    Public Sub Send_ListbarBarcodes(Optional ByVal linesPerPage As Integer = 0, Optional ByVal sizeOfData As Integer = 0, Optional ByVal PageNum As Integer = 0, Optional ByVal filter As Integer = 0, Optional ByVal QueryString As String = "", Optional ByVal searchString As String = "", Optional ByVal SortText As String = "", Optional ByVal startDate As String = "", Optional ByVal endDate As String = "", Optional ByVal searchField As Boolean = True)
        Dim myUrl As String
        Dim startVal As Integer
        Dim endVal As Integer
        Dim filterString As String = ""

        startVal = linesPerPage * PageNum
        endVal = linesPerPage * PageNum + linesPerPage
        If startVal = 0 Then
            startVal = 1
        Else
            startVal += 1
        End If

        If endVal > sizeOfData Then
            endVal = sizeOfData
        End If

        myUrl = Request.CurrentExecutionFilePath
        If searchField Then
            Send("<div id=""listbar"">")

            Send("  <div id=""searcher"" title=""" & Copient.PhraseLib.Lookup("term.searchterms", LanguageID) & """>")
            Send("    Barcode:<input type=""text"" id=""searchterms"" name=""searchterms"" maxlength=""30"" value=""" & searchString & """ />")
            Send("    &nbsp;&nbsp;")
            Sendb("    <label for=""startdate"">" & Copient.PhraseLib.Lookup("term.dates", LanguageID) & ":</label>")
            Sendb("<input type=""text"" id=""startdate"" name=""startdate"" class=""short"" maxlength=""10"" value=""" & startDate & """ /><img src=""/images/calendar.png"" class=""calendar"" id=""start-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('startdate', event);"" style=""vertical-align:middle;"" />–")
            Sendb("<input type=""text"" id=""enddate"" name=""enddate"" class=""short"" maxlength=""10"" value=""" & endDate & """ /><img src=""/images/calendar.png"" class=""calendar"" id=""end-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('enddate', event);"" style=""vertical-align:middle;"" />")
            Send("&nbsp;&nbsp;")
            Send("    <input type=""submit"" id=""search"" name=""search"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ /><br />")
            Send("  </div>")

            Send("  <div id=""filter"" title=""" & Copient.PhraseLib.Lookup("term.filter", LanguageID) & """>")
            Send("    <select id=""filtercoupon"" name=""filtercoupon"" onchange=""mainform.submit();"">")
            filterString = "&amp;filtercoupon=" & filter
            Send("      <option value=""0""" & IIf(filter = 0, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.AllCoupons", LanguageID) & "</option>")
            Send("      <option value=""1""" & IIf(filter = 1, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.unredeemed", LanguageID) & " " & Copient.PhraseLib.Lookup("term.nonexpired", LanguageID) & "</option>")
            Send("      <option value=""2""" & IIf(filter = 2, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.unredeemed", LanguageID) & " " & Copient.PhraseLib.Lookup("term.expired", LanguageID) & "</option>")
            Send("      <option value=""3""" & IIf(filter = 3, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.voided", LanguageID) & "</option>")
            Send("    </select>")
            Send("  </div>")
            Send("</div>")
        End If

        Send("<div id=""listbar2"">")
        Send("  <div id=""paginator"">")
        If (PageNum > 0) Then
            Send("   <span id=""first""><a id=""firstPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=0&amp;searchterms=" & searchString & SortText & filterString & """><b>|</b>◄" & Copient.PhraseLib.Lookup("term.first", LanguageID) & "</a></span>&nbsp;")
            Send("   <span id=""previous""><a id=""previousPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=" & PageNum - 1 & "&amp;searchterms=" & searchString & SortText & filterString & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a></span>")
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
            Send("   <span id=""next""><a id=""nextPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=" & PageNum + 1 & "&amp;searchterms=" & searchString & SortText & filterString & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a></span>&nbsp;")
            Send("   <span id=""last""><a id=""lastPageLink"" href=""" & myUrl & "?" & QueryString & "&amp;pagenum=" & (Math.Ceiling(sizeOfData / linesPerPage) - 1) & "&amp;searchterms=" & searchString & SortText & filterString & """>" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></a></span><br />")
        Else
            Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</span>&nbsp;")
            Send("   <span id=""last"">" & Copient.PhraseLib.Lookup("term.last", LanguageID) & "►<b>|</b></span><br />")
        End If
        Send("  </div>")

        Send(" <hr class=""hidden"" />")
        Send("</div>")
    End Sub

    ''' <summary>
    ''' Determines if an offer has expired based on its end date
    ''' </summary>
    ''' <param name="OfferID">OfferID of the offer to be checked</param>
    ''' <param name="EngineID">EngineID of the offer to be checked</param>
    ''' <returns>True if offer is expired, false otherwise</returns>
    ''' <remarks></remarks>
    Public Function IsOfferExpired(ByVal OfferID As Integer, Optional ByVal EngineID As Integer = 1) As Boolean

        Dim Common As New Copient.CommonInc
        Dim dt As DataTable

        Common.Open_LogixRT()

        Select Case EngineID
            Case 2, 3, 5, 6, 9
                Common.QueryStr = "SELECT isnull(EndDate,'1/1/1980') AS prodEndDate FROM CPE_Incentives WHERE IncentiveID=" & OfferID & ";"
            Case Else
                Common.QueryStr = "SELECT isnull(prodEndDate,'1/1/1980') AS prodEndDate FROM Offers WHERE OfferID=" & OfferID & ";"
        End Select
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).Item("prodEndDate") < Today Then Return True
        End If

        Return False

    End Function

    ''' <summary>
    ''' Determines if an offer has been deployed
    ''' </summary>
    ''' <param name="OfferID">OfferID of the offer to be checked</param>
    ''' <param name="EngineID">EngineID of the offer to be checked</param>
    ''' <returns>True if offer has been deployed, false otherwise</returns>
    ''' <remarks>Not able to determine if the offer has been deployed with no locations (undeployed)</remarks>
    Public Function HasOfferBeenDeployed(ByVal OfferID As Integer, Optional ByVal EngineID As Integer = 1) As Boolean

        Dim Common As New Copient.CommonInc
        Dim dt As DataTable

        Common.Open_LogixRT()

        Select Case EngineID
            Case 2, 3, 5, 6, 9
                Common.QueryStr = "SELECT isnull(UpdateLevel, 0) AS UpdateLevel FROM CPE_Incentives CPEI " &
                "INNER JOIN PromoEngineUpdateLevels PEU ON CPEI.UpdateLevel = PEU.LastUpdateLevel " &
                "WHERE CPEI.IncentiveID = " & OfferID & " AND PEU.LinkID = " & OfferID & ";"
            Case Else
                Common.QueryStr = "SELECT isnull(UpdateLevel, 0) AS UpdateLevel FROM Offers O " &
                "INNER JOIN PromoEngineUpdateLevels PEU ON O.UpdateLevel = PEU.LastUpdateLevel " &
                "WHERE O.OfferID = " & OfferID & " AND PEU.LinkID = " & OfferID & ";"
        End Select
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).Item("UpdateLevel") > 0 Then Return True
        End If

        Return False

    End Function

    Public Sub Send_Status(ByVal OfferID As Integer, Optional ByVal EngineID As Integer = 1)
        Dim MyCommon As New Copient.CommonInc
        Dim Logix As New Copient.LogixInc
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim status As Integer
        Dim IsTemplate As Boolean = False
        Dim StatusFlag As Integer = 0
        Dim StatusText As String = ""
        Dim statusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()

        StatusText = Logix.GetOfferStatus(OfferID, LanguageID, statusCode)
        Select Case EngineID
            Case 2, 3, 5, 6, 9
                MyCommon.QueryStr = "select StatusFlag, IsTemplate, isnull(EndDate,'1/1/1980') as prodEndDate, isnull(DeployDeferred, 0) as DeployDeferred from CPE_Incentives where IncentiveID=" & OfferID & ";"
            Case Else
                MyCommon.QueryStr = "select StatusFlag, isnull(IsTemplate, 0) as IsTemplate, isnull(prodEndDate,'1/1/1980') as prodEndDate, isnull(DeployDeferred, 0) as DeployDeferred from Offers where OfferID=" & OfferID & ";"
        End Select
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            StatusFlag = rst.Rows(0).Item("StatusFlag")
            IsTemplate = rst.Rows(0).Item("IsTemplate")
            If rst.Rows(0).Item("prodEndDate") < Today Then status = 3
            If rst.Rows(0).Item("DeployDeferred") = True Then status = 4
            If Not IsTemplate Then
                If (statusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE Or statusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_TESTING) AndAlso (StatusFlag > 0) AndAlso (StatusFlag <> 2) AndAlso OfferLockedforCollisionDetection = False Then
                    If HasOfferBeenDeployed(OfferID, EngineID) Then
                        Send("<div id=""statusbar"" class=""red-background"">" & Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID) & "</div>")
                        MyCommon.Close_LogixRT()
                        MyCommon = Nothing
                        Logix = Nothing
                        Exit Sub
                    End If
                End If
            End If
        End If

        ' Query for and set the offer's status value
        Select Case EngineID
            Case 2, 3, 5, 6, 9
                MyCommon.QueryStr = "Select StatusFlag,isnull(EndDate,0) as prodEndDate, DeployDeferred from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
            Case Else
                MyCommon.QueryStr = "select StatusFlag,isnull(prodEndDate,0) as prodEndDate, DeployDeferred from Offers with (NoLock) where OfferID=" & OfferID
        End Select
        rst = MyCommon.LRT_Select()
        For Each row In rst.Rows
            status = row.Item("StatusFlag")
        Next
        For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("prodEndDate"), Today) < Today) Then
                status = 3
            End If
            If (MyCommon.NZ(row.Item("DeployDeferred"), False) = True) Then
                status = 4
            End If
        Next
        If (OfferID <= 0) Then
        Else
            Sendb("<div id=""statusbar""")
            If (OfferLockedforCollisionDetection = True) Then
                Sendb(" class=""green-background"">" & Copient.PhraseLib.Lookup("offer.status2msg", LanguageID))
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
                                         Optional ByVal IsTemplate As Boolean = False, Optional ByVal bDisallowEditPg As Boolean = False, Optional ByVal bStoreUser As Boolean = False, Optional ByVal sValidLocIds As String = "", Optional ByVal sValidSU As String = "", Optional ByRef SelectedProductGroupId As Integer = 0, Optional ByRef ExcludedProdGroupID As Integer = 0, Optional ByRef ByExistingPGSelector As Boolean = False, Optional ByVal SearchProductGrouptext As String = Nothing, Optional ByVal RadioButtonEnableStart As Boolean = False, Optional ByVal RadioButtonEnableContain As Boolean = False)
        Dim MyCommon As New Copient.CommonInc
        Dim row As System.Data.DataRow
        Dim rst As System.Data.DataTable
        Dim RID As Integer = CType(RewardID, Integer)
        Dim wherestr As String = ""
        Dim sJoin As String = ""
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        Dim RecordLimit As Integer = GroupRecordLimit
        Dim SystemOption235 As String = MyCommon.Fetch_SystemOption(235)

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        Send("<div class=""box"" id=""groups"" " & IIf(ByExistingPGSelector = False, "style=""display: none; overflow:auto;""", "style=""overflow:auto;""") & ">")
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
        If Not (SearchProductGrouptext <> Nothing) Then
            Send("<input type=""radio""  id=""functionradio1"" name=""functionradioSearch"" " & IIf(MyCommon.Fetch_SystemOption(175) = "1", "checked=""checked""", "") & " /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
            Send("<input type=""radio""  id=""functionradio2"" name=""functionradioSearch"" " & IIf(MyCommon.Fetch_SystemOption(175) = "2", "checked=""checked""", "") & "  /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
        Else
            Send("<input type=""radio""  id=""functionradio1"" name=""functionradioSearch"" " & IIf(RadioButtonEnableStart, "checked=""checked""", "") & " /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
            Send("<input type=""radio""  id=""functionradio2"" name=""functionradioSearch"" " & IIf(RadioButtonEnableContain, "checked=""checked""", "") & "  /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")

        End If


        Send("<input type=""text"" class=""medium"" id=""functioninputSearch"" name=""functioninputSearch"" maxlength=""100"" onkeyup=""javascript:xmlPostTimer('OfferFeeds.aspx','ProductGroupsCM');"" value='" & IIf(SearchProductGrouptext <> Nothing, SearchProductGrouptext, String.Empty) & "' /><br />")
        Send("<div id=""searchLoadDiv"" style=""display: block;"">&nbsp;</div>")
        Send("            <div style=""float:left;position:relative;"">")
        Send("<br />")
        Send("                <label for=""pgroup-select"">" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</label><br clear=""all"" />")

        Send("                <select class=""longer"" id=""pgroup-select"" name=""pgroup-select"" size=""2"">")
        MyCommon.QueryStr = "select OFR.ProductGroupID,ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C on OFR.ProductGroupID=C.ProductGroupID and RewardID=" & RID & ";"

        rst = MyCommon.LRT_Select
        If (rst.Rows.Count = 0 And SelectedProductGroupId = 0) Then
            'Send("<option>" & Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "</option>")
            TransactionLevelSelected = True
        End If
        If (SelectedProductGroupId <> 0 And SelectedProductGroupId <> -1) Then
            MyCommon.QueryStr = " select  ProductGroupID,Name from ProductGroups where ProductGroupID =" & SelectedProductGroupId & ";"
            Dim pgi As DataTable = MyCommon.LRT_Select
            For Each row In pgi.Rows
                Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
                selecteditem = row.Item("ProductGroupID")
            Next
        End If
        For Each row In rst.Rows
            If Not (SelectedProductGroupId.ToString().Contains(row.Item("ProductGroupID"))) And SelectedProductGroupId <> -1 And SelectedProductGroupId = 0 Then
                Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
                selecteditem = row.Item("ProductGroupID")
            End If
        Next
        Send("                </select><br />")
        Send("                <br style=""line-height: 11px;"" />")

        Send("                <label for=""pgroup-exclude"">" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ":</label><br clear=""all"" />")
        Send("                <select class=""longer"" id=""pgroup-exclude"" name=""pgroup-exclude"" size=""2"">")

        MyCommon.QueryStr = "select OFR.ProductGroupID,OFR.ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C on OFR.ExcludedProdGroupID=C.ProductGroupID  where not(ExcludedProdGroupID=0) and RewardID=" & RID & ";"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            If Not (ExcludedProdGroupID.ToString().Contains(row.Item("ExcludedProdGroupID"))) And ExcludedProdGroupID <> -1 Then
                Send("<option value=""" & row.Item("ExcludedProdGroupID") & """>" & row.Item("Name") & "</option>")
                ExcludedItem = row.Item("ExcludedProdGroupID")
            End If
        Next
        If (ExcludedProdGroupID <> 0) And (ExcludedProdGroupID <> -1) Then
            MyCommon.QueryStr = " select  ProductGroupID,Name from ProductGroups where ProductGroupID =" & ExcludedProdGroupID & ";"
            Dim pgi As DataTable = MyCommon.LRT_Select
            For Each row In pgi.Rows
                Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
                ExcludedItem = row.Item("ProductGroupID")
            Next
        End If

        Send("                </select>")

        Send("            </div>")
        Send("")
        Send("            <div style=""float:left;position:relative;padding: 14px 1px 0px 1px;"">")
        Send("<br />")
        Sendb("                <input ")
        If Not IsTemplate Then
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
        If Not IsTemplate Then
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
        If Not IsTemplate Then
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
        If Not IsTemplate Then
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
        If (RecordLimit > 0) Then
            If (SystemOption235 = "1") Then
                Send(Copient.PhraseLib.Lookup("groups.displayname", LanguageID) & ": " & RecordLimit.ToString() & "<br />")
            Else
                Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RecordLimit.ToString() & "<br />")
            End If
        End If
        Send("            <div style=""float:left;position:relative;"">")
        Send("                <label for=""pgroup-avail"">" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</label><br clear=""all"" />")

        Send("<div id=""pgavaildiv"" class=""column3x1"">")
        Send("                <select class=""longer"" id=""pgroup-avail"" name=""pgroup-avail"" size=""6"">")
        Dim Limiter As String = ""
        If (ExcludedItem) Then Limiter = "and pg.ProductGroupID <> " & ExcludedItem
        If (selecteditem) Then Limiter = Limiter & " and pg.ProductGroupID <> " & selecteditem

        If bStoreUser Then
            sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID=pglu.ProductGroupID "
            wherestr = " and (LocationID in (" & sValidLocIds & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) and AnyProduct=0"
        End If


        If SearchProductGrouptext <> "" Then
            If (RadioButtonEnableStart) Then
                wherestr = wherestr + " and Name  like '" & SearchProductGrouptext & "%'"
            End If
            If (RadioButtonEnableContain) Then
                wherestr = wherestr + " and Name  like '%" & SearchProductGrouptext & "%'"
            End If

        End If

        MyCommon.QueryStr = "select " & IIf(RecordLimit > 0, "top " & RecordLimit, "") & " pg.ProductGroupID,CreatedDate,Name,LastUpdate,AnyProduct from ProductGroups pg with (NoLock) " & sJoin & "where Deleted=0" & Limiter & wherestr
        If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(pg.TranslatedFromOfferID,0) = 0 "
        If (SystemOption235 = "1") Then
            MyCommon.QueryStr &= " order by AnyProduct desc, Name"
        Else
            MyCommon.QueryStr &= " order by AnyProduct desc, ProductGroupID desc, Name asc"
        End If

        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
        Next
        Send("                </select>")
        Send("            </div>")
        Send("            </div>")
        Send("            <br clear=""left"" /><br class=""zero"" />")
        Send("            <hr class=""hidden"" />")
        Send("          </div>")


    End Sub

    Public Sub Send_DirectProductAddSelector(ByRef Logix As Object, ByRef ByExistingPGSelector As Boolean, ByRef ShowAllItems As Boolean, ByRef GroupSize As Integer, ByRef rstItems As DataTable, ByRef ProductsWithoutDesc As Integer, ByRef descriptionItem As String, ByRef ByAddSingleProduct As Boolean, ByRef IDLength As Integer, ByRef GName As String, Optional ByVal IsPostRequest As Boolean = False)

        Dim MyCommon As New Copient.CommonInc
        Dim rst2 As DataTable
        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()

        Send("<div class=""box"" id=""directprodaddselector"" " & IIf(ByExistingPGSelector = True, "style=""display: none; overflow:auto;""", "style=""overflow:auto;""") & ">")
        Send("  <h2>")
        Send("    <span>" & Copient.PhraseLib.Lookup("term.productcondition", LanguageID) & "</span>")
        Send("  </h2>")
        Send("    <span>")
        Send("<label for=""modprodgrpname"">")
        Send(Copient.PhraseLib.Lookup("pg.productgroupname", LanguageID))
        Send(":</label><br />")
        Send("<input type=""text"" id=""modprodgroupname"" style=""width: 347px;"" name=""modprodgroupname"" maxlength=""200"" value=""" & GName & """ />")
        Send("<br />")
        Send("    </span>")
        Send("<br class=""half"" />")
        Send("<div style=""float: left; width: 310px;"">")
        Send("<span style=""position: relative"">")

        If (ShowAllItems OrElse GroupSize <= 100) AndAlso rstItems IsNot Nothing Then
            Sendb(Copient.PhraseLib.Lookup("pgroup-edit.all-items-note", LanguageID) & " (" & rstItems.Rows.Count & " ")
            If (rstItems.Rows.Count = 1) Then
                Sendb(Copient.PhraseLib.Lookup("term.product", LanguageID).ToString.ToLower & ")<br />")
            Else
                Sendb(Copient.PhraseLib.Lookup("term.products", LanguageID).ToString.ToLower & ")<br />")
            End If
        Else
            Sendb(Copient.PhraseLib.Lookup("pgroup-edit.listnote", LanguageID) & "<br />")
        End If

        Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
        If (ProductsWithoutDesc = 1) Then
            Response.Write(ProductsWithoutDesc & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.product", LanguageID) & " ", VbStrConv.Lowercase))
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.withoutdesc", LanguageID), VbStrConv.Lowercase))
        Else
            Response.Write(ProductsWithoutDesc & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.products", LanguageID) & " ", VbStrConv.Lowercase))
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.withoutdesc", LanguageID), VbStrConv.Lowercase))
        End If
        Send("</span>")
        Send("<select name=""PKID"" id=""PKID"" size=""15"" multiple=""multiple"" onscroll=""handlePageClick(this);"" class=""longer"" style=""width: 290px;"">")

        descriptionItem = String.Empty
        If (GroupSize > 0) Then
            For Each row4 As DataRow In rstItems.Rows
                descriptionItem = MyCommon.NZ(row4.Item("ExtProductID"), " ") & " " & MyCommon.NZ(row4.Item("Description"), " ") & "-"
                If MyCommon.NZ(row4.Item("PhraseID"), 0) > 0 Then
                    descriptionItem &= Copient.PhraseLib.Lookup(MyCommon.NZ(row4.Item("PhraseID"), 0), LanguageID)
                Else
                    If MyCommon.NZ(row4.Item("ProductType"), "") <> "" Then
                        descriptionItem &= row4.Item("ProductType")
                    Else
                        descriptionItem &= Copient.PhraseLib.Lookup("term.unknown", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.type", LanguageID), VbStrConv.Lowercase) & " " & MyCommon.NZ(row4.Item("ProductTypeID"), 0)
                    End If
                End If
                Send("<option value=""" & row4.Item("PKID") & """>" & descriptionItem & "</option>")
            Next
        End If
        Send("</select")
        Send("<br />")
        'If (Not IsSpecialGroup OrElse (IsSpecialGroup And CanEditSpecialGroup)) Then
        If (Logix.UserRoles.EditProductGroups) Then
            Send("    <br class=""half"" /><input type=""submit"" class=""large"" id=""remove"" name=""remove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width:150px;"" value=""" & Copient.PhraseLib.Lookup("term.removefromlist", LanguageID) & """ />")
        End If
        'End If
        'If (Not ShowAllItems AndAlso GroupSize > 100) Then
        '    Send("<input class=""regular"" id=""btnShowAll"" name=""btnShowAll"" type=""button"" value=""" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & """ onclick=""submitShowAll();"" />")
        'End If
        Send("<br />")
        Send("</div>")
        Send("<div style=""margin-left: 310px;"">")
        Send("<table cellpadding=""1"" cellspacing=""1"">")
        Send("<tr>")
        Send("<td>")
        Send(Copient.PhraseLib.Lookup("term.type", LanguageID))
        Send("<br />")
        Send("<select id=""producttype"" name=""producttype"">")
        'BZ2079: UE-feature-removal #: Remove unsupported product types for UE (Mix/Match Code, Manufacturer Family code, Pool Code)
        '        To restore previous functionality: remove the all code in the If statement checking engines except the query without a where clause.

        MyCommon.QueryStr = "select ProductTypeID,PhraseID from ProductTypes with (NoLock)"

        rst2 = MyCommon.LRT_Select
        For Each row3 As DataRow In rst2.Rows
            'ProductTypeID = row3.Item("ProductTypeID")
            Send("     <option value=""" & row3.Item("ProductTypeID") & """>" & Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID) & "</option>")
        Next
        Send("</select>")
        Send("</td>")
        Send("</tr>")
        Send("</table>")
        Send("<br />")

        If ByAddSingleProduct Then
            Send("<input type=""radio"" id=""prodaddselector1"" name=""prodaddselector"" onclick=""javascript:ProductAddSelection();"" checked=""checked"" value=""prodadd"" />")
            Send("<label for=""prodaddselector1"">")
            Send(Copient.PhraseLib.Lookup("gen.addsingleproduct", LanguageID))
            Send("</label>")
            Send("<input type=""radio"" id=""prodaddselector2"" name=""prodaddselector"" onclick=""javascript:ProductAddSelection();"" value=""prodlistadd"" />")
            Send("<label for=""prodaddselector2"">")
            Send(Copient.PhraseLib.Lookup("gen.addproductlist", LanguageID))
            Send("</label>")
        Else
            Send("<input type=""radio"" id=""prodaddselector1"" name=""prodaddselector"" onclick=""javascript:ProductAddSelection();"" value=""prodadd"" />")
            Send("<label for=""prodaddselector1"">")
            Send(Copient.PhraseLib.Lookup("gen.addsingleproduct", LanguageID))
            Send("</label>")
            Send("<input type=""radio"" id=""prodaddselector2"" name=""prodaddselector"" onclick=""javascript:ProductAddSelection();"" checked=""checked"" value=""prodlistadd"" />")
            Send("<label for=""prodaddselector2"">")
            Send(Copient.PhraseLib.Lookup("gen.addproductlist", LanguageID))
            Send("</label>")
        End If
        Send("<br />")
        Send("<br />")
        Send("<div id=""addsingleproduct"" " & IIf(ByAddSingleProduct, """", " style=""display:none;""") & ">")
        Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)
        Send(Copient.PhraseLib.Lookup("term.id", LanguageID))
        Send(":<br />")
        If IDLength > 0 Then
            Send("<input type=""text"" id=""productid"" name=""ExtProductID"" maxlength=""" & IDLength & """ style=""width: 137px;"" value="""" />")
        Else
            Send("<input type=""text"" id=""productid"" name=""ExtProductID"" style=""width: 137px;"" value="""" />")
        End If
        Send("<br />")
        Send(" <br />")
        Send("<label for=""productdesc"">")
        Send(Copient.PhraseLib.Lookup("term.description", LanguageID))
        Send(":</label><br />")
        Send("<input type=""text"" id=""productdesc"" style=""width: 347px;"" name=""productdesc"" maxlength=""200"" value="""" />")
        Send("<br />")
        Send("<br class=""half"" />")
        If (Logix.UserRoles.EditProductGroups) Then
            Send("<div style=""float: left;"">")
            Send("<input type=""submit"" class=""large"" id=""add"" name=""add"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """  onclick=""return isValidID();"" />")
            Send("</div>")
            Send("<div style=""float: right; margin-right: 20px"">")

            Send("<input type=""submit"" class=""large"" id=""mremove"" name=""mremove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width: 150px;"" value=""" & Copient.PhraseLib.Lookup("term.removemanually", LanguageID) & """/>")
            Send("</div>")
            Send("<br />")
        End If

        Send("</div>")
        Send("<div id=""addproductlist"" " & IIf(Not ByAddSingleProduct, """", " style=""display:none;"" ") & ">")

        Send("<textarea name=""pasteproducts"" id=""pasteproducts"" style=""width: 290px; height: 150px"">")
        Send("</textarea>")
        Send("<br />")
        Send("<br />")
        Send("<input type=""radio"" name=""modifyoperation"" value=""0"" checked=""checked"" />")
        Send("<label for=""operation4"">" & Copient.PhraseLib.Lookup("term.FullReplace", LanguageID) & "</label>&nbsp;&nbsp;")
        Send("<input type=""radio"" name=""modifyoperation""  value=""1""  />")
        Send("<label for=""operation5"">" & Copient.PhraseLib.Lookup("term.AddToGroup", LanguageID) & "</label>&nbsp;&nbsp;")
        Send("<input type=""radio"" name=""modifyoperation"" value=""2""  />")
        Send("<label for=""operation6"">" & Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID) & "</label>")
        Send("<br />")

        Send("<br />")
        If (Logix.UserRoles.EditProductGroups) AndAlso Not IsPostRequest Then
            Send("     <input type=""button"" class=""regular"" id=""modifyprodgroup"" name=""modifyprodgroup"" value=""Apply Changes"" onclick=""javascript:ModifyGroup();"" />")
            Send("     <br />")
        ElseIf (Logix.UserRoles.EditProductGroups) AndAlso IsPostRequest Then
            Send("     <input type=""submit"" class=""regular"" id=""modifyprodgroup"" name=""modifyprodgroup"" value=""Apply Changes"" onclick=""return IsValidRegularExpression();"" />")
            Send("     <br />")
        End If
        Send("</div>")
        Send("</div>")
        Send("<hr class=""hidden"" />")
        Send("</div>")

    End Sub

    Public Sub Send_ProductConditionSelector(ByRef Logix As Object, ByRef TransactionLevelSelected As Object, ByRef FromTemplate As Object,
                                             ByRef Disallow_Edit As Object, ByRef selecteditem As Object, ByRef ExcludedItem As Object,
                                             ByRef RewardID As Object, ByRef EngineID As Integer,
                                             Optional ByVal IsTemplate As Boolean = False, Optional ByVal bDisallowEditPg As Boolean = False, Optional ByVal bStoreUser As Boolean = False, Optional ByVal sValidLocIds As String = "", Optional ByVal sValidSU As String = "")
        Dim MyCommon As New Copient.CommonInc
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim disabledattribute As String = ""
        Dim wherestr As String = ""
        Dim sJoin As String = ""
        Dim topString As String = ""
        Dim myUrl As String
        myUrl = Request.CurrentExecutionFilePath
        If GroupRecordLimit > 0 Then topString = "top " & GroupRecordLimit
        Dim orderBy As String = ""

        If (MyCommon.Fetch_SystemOption(235) = "1") Then
            orderBy = " order by AnyProduct desc, Name"
        Else
            orderBy = " order by AnyProduct desc, ProductGroupID desc, Name asc"
        End If

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

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
        If Not (Logix.userroles.editoffer And Not (FromTemplate And Disallow_Edit)) Then
            disabledattribute = " disabled=""disabled"""
        End If

        Send("<input type=""radio"" id=""functionradio1"" name=""functionradio"" " & IIf(MyCommon.Fetch_SystemOption(175) = "1", "checked=""checked""", "") & " " & disabledattribute & " /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
        Send("<input type=""radio"" id=""functionradio2"" name=""functionradio"" " & IIf(MyCommon.Fetch_SystemOption(175) = "2", "checked=""checked""", "") & " " & disabledattribute & " /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
        Send("<input type=""text"" class=""medium"" id=""functioninput"" name=""functioninput"" maxlength=""100"" onkeyup=""javascript:xmlPostTimer('OfferFeeds.aspx','ProductGroupsCM');"" value=""""" & disabledattribute & "/>")
        Send("<br />")
        Send("<div id=""searchLoadDiv"" style=""display: block;"">&nbsp;</div>")
        Send("<div id=""pgList"">")
        Send("<select class=""longer"" id=""functionselect"" name=""functionselect"" size=""12""" & disabledattribute & ">")
        Dim Limiter As String = ""
        If (ExcludedItem) Then Limiter = "and pg.ProductGroupID <> " & ExcludedItem
        If (selecteditem) Then Limiter = Limiter & " and pg.ProductGroupID <> " & selecteditem

        If bStoreUser Then
            sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID=pglu.ProductGroupID "
            wherestr = " and (LocationID in (" & sValidLocIds & ") or CreatedByAdminID in (" & sValidSU & ")) and AnyProduct=0"
        End If

        MyCommon.QueryStr = "select " & topString & " pg.ProductGroupID,CreatedDate,Name,LastUpdate,AnyProduct from ProductGroups pg with (NoLock) " & sJoin & " where Deleted=0" & Limiter & wherestr
        If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(pg.TranslatedFromOfferID,0) = 0 "
        MyCommon.QueryStr &= orderBy
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                Send("<option style=""font-weight:bold;color:brown;"" value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
            Else
                Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
            End If
        Next
        Send("</select>")
        Send("</div>")
        Send("<br />")
        If (GroupRecordLimit > 0) Then
            If (MyCommon.Fetch_SystemOption(235) = "1") Then
                Send(Copient.PhraseLib.Lookup("groups.displayname", LanguageID) & ": " & GroupRecordLimit.ToString() & "<br />")
            Else
                Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & GroupRecordLimit.ToString() & "<br />")
            End If
        End If
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
                Send("  <option style=""font-weight:bold;color:brown;"" value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
            Else
                Send("  <option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
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

    Public Sub Send_ProductConditionSelectorAdv(ByRef Logix As Object, ByRef TransactionLevelSelected As Object, ByRef FromTemplate As Object,
                                             ByRef Disallow_Edit As Object, ByRef selecteditem As Object, ByRef ExcludedItem As Object,
                                             ByRef RewardID As Object, ByRef EngineID As Integer,
                                             Optional ByVal IsTemplate As Boolean = False, Optional ByVal bDisallowEditPg As Boolean = False, Optional ByVal ByExistingPGSelector As Boolean = False, Optional ByVal bStoreUser As Boolean = False, Optional ByVal sValidLocIds As String = "", Optional ByVal sValidSU As String = "")
        Dim MyCommon As New Copient.CommonInc
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim disabledattribute As String = ""
        Dim wherestr As String = ""
        Dim sJoin As String = ""
        Dim orderBy As String = ""
        Dim topString As String = ""
        Dim myUrl As String = Request.CurrentExecutionFilePath
        If GroupRecordLimit > 0 Then topString = "top " & GroupRecordLimit
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        Dim bCreateGroupOrProgramFromOffer As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(134) = "1", True, False)

        If (MyCommon.Fetch_SystemOption(235) = "1") Then
            orderBy = " order by AnyProduct desc, Name"
        Else
            orderBy = " order by AnyProduct desc, ProductGroupID desc, Name asc"
        End If

        MyCommon.AppName = "LogixCB"
        MyCommon.Open_LogixRT()
        Send("<div class=""box"" id=""selector"" " & IIf(ByExistingPGSelector = False, "style=""display: none; overflow:auto;""", "style=""overflow:auto;""") & ">")
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
        If Not (Logix.userroles.editoffer And Not (FromTemplate And Disallow_Edit)) Then
            disabledattribute = " disabled=""disabled"""
        End If

        Send("<input type=""radio"" id=""functionradio1"" name=""functionradio"" " & IIf(MyCommon.Fetch_SystemOption(175) = "1", "checked=""checked""", "") & " " & disabledattribute & " /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
        Send("<input type=""radio"" id=""functionradio2"" name=""functionradio"" " & IIf(MyCommon.Fetch_SystemOption(175) = "2", "checked=""checked""", "") & " " & disabledattribute & " /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
        Send("<input type=""text"" class=""medium"" id=""functioninput"" name=""functioninput"" maxlength=""100"" onkeyup=""javascript:xmlPostTimer('OfferFeeds.aspx','ProductGroupsCM');"" value=""""" & disabledattribute & "/>")
        If (bCreateGroupOrProgramFromOffer AndAlso Logix.userroles.createproductgroups AndAlso myUrl <> "/logix/offer-rew-point.aspx" AndAlso myUrl <> "/logix/offer-rew-sv.aspx") Then
            Send("<input class=""regular"" name=""btncreate"" id=""btncreate"" type=""button"" value=""" & Copient.PhraseLib.Lookup("term.create", LanguageID) & """ onclick=""handleCreateClick('btncreate');""" & disabledattribute & "/>")
        End If
        Send("<br />")
        Send("<div id=""searchLoadDiv"" style=""display: block;"">&nbsp;</div>")
        Send("<div id=""pgList"" class=""column3x1"">")
        Send("<select class=""long"" id=""functionselect"" name=""functionselect"" size=""20""" & disabledattribute & ">")
        Dim Limiter As String = ""
        If (ExcludedItem) Then Limiter = "and pg.ProductGroupID <> " & ExcludedItem
        If (selecteditem) Then Limiter = Limiter & " and pg.ProductGroupID <> " & selecteditem

        If bStoreUser Then
            sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID=pglu.ProductGroupID "
            wherestr = " and (LocationID in (" & sValidLocIds & ") or CreatedByAdminID in (" & sValidSU & ")) and AnyProduct=0"
        End If

        MyCommon.QueryStr = "select " & topString & " pg.ProductGroupID,CreatedDate,Name,LastUpdate,AnyProduct from ProductGroups pg with (NoLock) " & sJoin & " where Deleted=0" & Limiter & wherestr
        If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &= " and isnull(pg.TranslatedFromOfferID,0) = 0 "
        MyCommon.QueryStr &= orderBy
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                Send("<option style=""font-weight:bold;color:brown;"" value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
            Else
                Send("<option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """ >" & row.Item("Name") & "</option>")
            End If
        Next
        Send("</select>")
        Send("</div>")
        Send("<br />")
        If (GroupRecordLimit > 0) Then
            If (MyCommon.Fetch_SystemOption(235) = "1") Then
                Send(Copient.PhraseLib.Lookup("groups.displayname", LanguageID) & ": " & GroupRecordLimit.ToString() & "<br />")
            Else
                Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & GroupRecordLimit.ToString() & "<br />")
            End If
        End If
        Send("<div class=""column3x2"">")
        Send("<center>")
        Send("<br />")
        Send("<br />")
        'Send("<b><label for=""selected"">" & Copient.PhraseLib.Lookup("term.selectedproducts", LanguageID) & ":</label></b><br />")
        Send("<input type=""button"" class=""regular select"" id=""select1"" name=""select1"" value=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & " &#9658;"" onclick=""select1_onclick();""" & disabledattribute & " />")
        Send("<br />")
        Send("<br />")
        Send("<input type=""button"" class=""regular deselect"" id=""deselect1"" name=""deselect1"" value="" &#9668; " & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & """ onclick=""deselect1_onclick();"" disabled=""disabled"" /><br />")
        Send("<br />")
        Send("<br />")
        Send("<br />")
        Send("<br />")
        Send("<br />")
        Send("<br />")
        Send("<br />")
        Send("<input type=""button"" class=""regular select"" id=""select2"" name=""select2"" value=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & " &#9658;"" onclick=""select2_onclick();"" disabled=""disabled"" " & disabledattribute & " />")
        Send("<br />")
        Send("<br />")
        Send("<input type=""button"" class=""regular deselect"" id=""deselect2"" name=""deselect2"" value="" &#9668; " & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & """ onclick=""deselect2_onclick();"" disabled=""disabled"" " & disabledattribute & " /><br />")
        Send("</center>")
        Send("</div>")
        Send("<br />")
        Send("<div class=""column3x3"">")
        Send("<div class=""graybox"">")
        Send("<h3>")
        Send("<span>" & Copient.PhraseLib.Lookup("term.selectedproducts", LanguageID) & "</span>")
        Send("</h3>")
        Send("<select class=""long"" id=""selected"" name=""selected"" size=""7""" & disabledattribute & ">")

        MyCommon.QueryStr = "select OFR.ProductGroupID,ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C with (NoLock) on OFR.ProductGroupID=C.ProductGroupID and RewardID=" & RewardID & ";"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count = 0) Then
            'Send("  <option>" & Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID) & "</option>")
            TransactionLevelSelected = True
        End If
        For Each row In rst.Rows
            If MyCommon.NZ(row.Item("ProductGroupID"), 0) = 1 Then
                Send("  <option style=""font-weight:bold;color:brown;"" value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """ >" & row.Item("Name") & "</option>")
            Else
                Send("  <option value=""" & row.Item("ProductGroupID") & """ title=""" & row.Item("Name") & """>" & row.Item("Name") & "</option>")
            End If
            selecteditem = row.Item("ProductGroupID")
        Next
        Send("</select>")
        Send("</div")
        Send("<br />")


        Send("<br class=""half"" />")
        Send("<div class=""graybox"">")
        Send("<h3>")
        Send("<span>" & Copient.PhraseLib.Lookup("term.excludedproducts", LanguageID) & "</span>")
        Send("</h3>")
        Send("<select class=""long"" id=""excluded"" name=""excluded"" size=""7""" & disabledattribute & ">")

        MyCommon.QueryStr = "select OFR.ProductGroupID,OFR.ExcludedProdGroupID,C.Name from OfferRewards as OFR with (NoLock) join ProductGroups as C with (NoLock) on OFR.ExcludedProdGroupID=C.ProductGroupID  where not(ExcludedProdGroupID=0) and RewardID=" & RewardID & ";"
        rst = MyCommon.LRT_Select
        For Each row In rst.Rows
            Send("<option value=""" & row.Item("ExcludedProdGroupID") & """>" & row.Item("Name") & "</option>")
            ExcludedItem = row.Item("ExcludedProdGroupID")
        Next
        Send("</select>")
        Send("</div>")
        Send("<hr class=""hidden"" />")
        Send("</div>")
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
            If Not String.IsNullOrWhiteSpace(Request.Form("notetext")) Then
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
        If (Logix.UserRoles.CreateNotes) Then
            Send("        <div id=""notesinput"" style=""display:none;"">")
            Send("          <textarea id=""notetext"" name=""notetext""></textarea>")
            Send("          <br />")
            Send("          <input type=""checkbox"" id=""private"" name=""private"" value=""1"" /><label for=""private"">" & Copient.PhraseLib.Lookup("term.private", LanguageID) & "</label>")
            Send("          <input type=""checkbox"" id=""important"" name=""important"" value=""1"" style=""display:none;"" /><label for=""important"" style=""display:none;"">" & Copient.PhraseLib.Lookup("term.important", LanguageID) & "</label>")
            Send("          <br />")
            Send("          <input type=""submit"" class=""regular"" id=""notesave"" name=""notesave"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ style=""margin-top:6px;"" />")
            Send("          <input type=""button"" class=""regular"" id=""notecancel"" name=""notecancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ style=""margin-top:6px;"" onclick=""toggleNotesInput();"" /><br />")
            Send("        </div>")
        End If
        Send("      </div>")
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

    Public Sub Send_SV_Extend(Optional ByVal Attributes As String = "")
        Send("<input type=""submit"" class=""regular"" id=""SavePropDeployExtend"" name=""SavePropDeployExtend"" value=""" & Copient.PhraseLib.Lookup("term.saveextend-sv", LanguageID) & """" & Attributes & " />")
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

    Public Function IsBrowserIE() As Boolean
        Return (Request.Browser.Browser = "IE")
    End Function

    Public Function CPEOffer_Has_AnyCustomer(ByVal Common As Copient.CommonInc, ByVal OfferID As Long) As Boolean
        Dim RetVal As Boolean = False

        Common.QueryStr = "dbo.pa_Check_AnyCustomer_In_Offer"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@AnyCustomerInOffer", SqlDbType.Int).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        If Common.LRTsp.Parameters("@AnyCustomerInOffer").Value = 1 Then RetVal = True
        Common.Close_LRTsp()
        Return RetVal

    End Function

    Public Function GetCgiValue(ByVal VarName As String) As String
        Dim TempVal As String
        TempVal = ""
        If Request.QueryString(VarName) Is Nothing Then
            TempVal = ""
        Else
            TempVal = Request.QueryString(VarName)
        End If

        If TempVal = "" Then
            If Request.Form(VarName) Is Nothing Then
                TempVal = ""
            Else
                TempVal = Request.Form(VarName)
            End If
        End If

        Return Server.HtmlEncode(TempVal)

    End Function

    Public Function Get_Raw_RequestData(ByVal RawStream As System.IO.Stream) As String
        Dim Index As Long
        Dim RawLen, strRead As Long
        Dim RawRequest As String
        Dim GetRequest As String

        'this gets the GET request (if any)
        GetRequest = Request.QueryString.ToString()

        'this get ths FORM post (if any)
        RawRequest = ""
        ' Find number of bytes in stream.
        RawLen = CInt(RawStream.Length)
        ' Create a byte array.
        Dim RawArray(RawLen) As Byte
        ' Read stream into byte array.
        strRead = RawStream.Read(RawArray, 0, RawLen)
        ' Convert byte array to a text string.
        For Index = 0 To RawLen - 1
            RawRequest = RawRequest & Chr(RawArray(Index))
        Next Index

        'now that we have the GET and FORM data, concatenate them
        If (RawRequest <> "") Or (GetRequest <> "") Then
            If Not (RawRequest = "") And Not (GetRequest = "") Then
                RawRequest = GetRequest & "&" & RawRequest
            Else
                RawRequest = GetRequest & RawRequest
            End If
        End If

        Get_Raw_RequestData = RawRequest

    End Function

    Public Sub Send_Calendar_Overrides(ByRef Common As Copient.CommonInc)
        Dim i As Integer
        Dim dt As DataTable
        Dim UserCulture As System.Globalization.CultureInfo = Nothing
        Dim FirstDayOfWeek As Integer = 0
        Dim days As String() = {"sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"}
        Dim months As String() = {"january", "february", "march", "april", "may", "june", "july", "august",
                                  "september", "october", "november", "december"}

        Send("  // localize by overriding phrase text and start of week found in datepicker.js for used by the calendar control")

        ' find the last day of the week in the user's language (region)
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        Common.QueryStr = "select MSNetCode from Languages with (NoLock) where LanguageID=" & LanguageID
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            UserCulture = New System.Globalization.CultureInfo(Common.NZ(dt.Rows(0).Item("MSNetCode"), "en-US").ToString)
            FirstDayOfWeek = CInt(UserCulture.DateTimeFormat.FirstDayOfWeek)
            Send("  firstDayOfWeek=" & FirstDayOfWeek & ";")
            Send("  lastDayOfWeek = " & ((FirstDayOfWeek + 6) Mod 7) & ";")

            Sendb("  dateArray = new Array(")
            For i = 0 To 31
                If i > 0 Then Sendb(",")
                Sendb("'" & Copient.commonShared.TranslateDigits(i, UserCulture) & "'")
            Next
            Send(")")
        End If

        ' send the days of the week in the user's language. start at the first day and wrap around the array to get any days that might remain
        Sendb("  dayArray = new Array(")
        For i = 0 To days.GetUpperBound(0)
            If i > 0 Then Sendb(",")
            Sendb("'" & Copient.PhraseLib.Lookup("calendar." & days(i) & "abbreviation", LanguageID) & "'")
        Next
        Send(")")

        ' send the name of the months in the user's language
        Sendb("  monthArray = new Array(")
        For i = 0 To months.GetUpperBound(0)
            If i > 0 Then Sendb(",")
            Sendb("'" & Copient.PhraseLib.Lookup("term." & months(i), LanguageID) & "'")
        Next
        Send(")")

        If UserCulture IsNot Nothing Then
            Send("  defaultDateSeparator = '" & UserCulture.DateTimeFormat.DateSeparator & "';")
            Send("  defaultDateFormat = '" & GetLocalizedDateFormat(UserCulture) & "';")
        End If

        Send("  calendarPhrase = '" & Copient.PhraseLib.Lookup("term.calendar", LanguageID) & "';")
        Send("  todayPhrase = '" & Copient.PhraseLib.Lookup("term.today", LanguageID) & "';")

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
            Send("  <div style=""position: relative;float: left;""><font size=""3""><b>" & BoxTitle & "</b></font>" & ExtraTitleText & "</div>")
            Send("  <div class=""resizer"" style=""position: relative;"">")
            Send("    <a href="""" onclick=""resizeBox('" & BoxObjectName & "body','img" & BoxObjectName & "body','" & BoxTitle & "', '" & BoxID & "', '" & AdminUserID & "'); return false;"">")
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
                Send("  <div id=""" & BoxObjectName & "body""  style=""display: none;"">")
            End If

        End If  'ValidBoxID

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

    Public Function CheckItemCode(ByRef itemCode As String, Optional ByRef StatusString As String = "") As Boolean
        Dim itemVal As New Copient.ItemCodeValidation
        Dim bRetVal As Boolean = False
        bRetVal = itemVal.ValidateItemCode(itemCode, StatusString)
        itemVal = Nothing
        CheckItemCode = bRetVal
    End Function

    Public Function IsValidOffer(ByVal OfferID As Integer, ByVal EngineID As Integer, ByVal Common As Copient.CommonInc) As Boolean
        Dim rst As DataTable
        Dim bValidOffer As Boolean = False
        Select Case EngineID
            Case 2, 3, 5, 6, 9
                Common.QueryStr = "select * from CPE_Incentives where IncentiveID = " & OfferID & " and EngineID = " & EngineID & ";"
            Case Else
                Common.QueryStr = "select * from Offers where IncentiveID = " & OfferID & " and EngineID = " & EngineID & ";"
        End Select

        rst = Common.LRT_Select
        If rst.Rows.Count > 0 Then
            bValidOffer = True
        End If

        IsValidOffer = bValidOffer
    End Function

    Public Sub RemoveTransformedOffers(ByRef dt As DataTable, ByVal Common As Copient.CommonInc)
        'Removing transformed offers
        Dim Logix As New Copient.LogixInc
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(Common.Fetch_SystemOption(249) = "1", True, False)
        AdminUserID = Verify_AdminUser(Common, Logix)
        Dim conditionalQuery = String.Empty
        If (bEnableRestrictedAccessToUEOfferBuilder) Then
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                Dim OfferIDs As String = String.Empty
                If (dt.Columns.Contains("LinkID")) Then
                    Dim value As String = dt.Columns("LinkID").DataType.Name
                    Select Case value
                        Case "Int32"
                            OfferIDs = String.Join(",", dt.AsEnumerable().Select(Function(x) x.Field(Of Int32)("LinkID").ToString()).ToArray())
                        Case "Int64"
                            OfferIDs = String.Join(",", dt.AsEnumerable().Select(Function(x) x.Field(Of Int64)("LinkID").ToString()).ToArray())
                        Case "Int16"
                            OfferIDs = String.Join(",", dt.AsEnumerable().Select(Function(x) x.Field(Of Int16)("LinkID").ToString()).ToArray())
                    End Select
                ElseIf dt.Columns.Contains("OfferID") Then
                    Dim value As String = dt.Columns("OfferID").DataType.Name
                    Select Case value
                        Case "Int32"
                            OfferIDs = String.Join(",", dt.AsEnumerable().Select(Function(x) x.Field(Of Int32)("OfferID").ToString()).ToArray())
                        Case "Int64"
                            OfferIDs = String.Join(",", dt.AsEnumerable().Select(Function(x) x.Field(Of Int64)("OfferID").ToString()).ToArray())
                        Case "Int16"
                            OfferIDs = String.Join(",", dt.AsEnumerable().Select(Function(x) x.Field(Of Int16)("OfferID").ToString()).ToArray())
                    End Select
                End If

                If (Not String.IsNullOrEmpty(OfferIDs)) Then
                    conditionalQuery = GetRestrictedAccessToUEBuilderQuery(Common, Logix, "")
                End If

                Common.QueryStr = " Select OfferID from Offers where OfferID IN(" & OfferIDs & ") " &
                              " Union " &
                              " Select IncentiveID as OfferID  from CPE_Incentives where IncentiveID IN(" & OfferIDs & ") "
                If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then Common.QueryStr &= conditionalQuery & " "

                Dim rst As DataTable = Common.LRT_Select()
                If rst.Rows.Count > 0 Then
                    Dim table As New DataTable()
                    OfferIDs = String.Join(",", rst.AsEnumerable().Select(Function(x) x.Field(Of Int64)("OfferID").ToString()).ToArray())

                    ' Filter and Sort expressions
                    Dim expression As String = String.Empty
                    Dim sortOrder As String = String.Empty
                    If (dt.Columns.Contains("LinkID")) Then
                        expression = "[LinkID] IN ( " & OfferIDs & ")"
                        sortOrder = "[LinkID] ASC"
                    Else
                        expression = "[OfferID] IN ( " & OfferIDs & ")"
                        sortOrder = "[OfferID] ASC"
                    End If

                    ' Create a DataView using the table as its source and the filter and sort expressions
                    Using dv As New DataView(dt, expression, sortOrder, DataViewRowState.CurrentRows)
                        Using tempDataTable As DataTable = dv.ToTable()
                            dt.Clear()
                            For Each dr As DataRow In tempDataTable.Rows
                                dt.ImportRow(dr)
                            Next
                        End Using
                    End Using
                Else
                    dt.Clear()
                End If
            End If
        End If
    End Sub

    Public Function GetRoleBasedUEOffers(ByVal lstOffers As List(Of CMS.AMS.Models.Offer), ByVal MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc) As List(Of CMS.AMS.Models.Offer)
        Dim conditionalQuery As String = String.Empty
        Dim rst2 As DataTable
        Dim tempOfferLst As New List(Of CMS.AMS.Models.Offer)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        If (bEnableRestrictedAccessToUEOfferBuilder) Then
            conditionalQuery = GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "")
        End If
        If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then
            If lstOffers.Count > 0 Then
                For Each offer As CMS.AMS.Models.Offer In lstOffers
                    MyCommon.QueryStr = "select distinct O.OfferID  from Offers as O where O.OfferID=" & offer.OfferID &
                                           " UNION " &
                                           " SELECT DISTINCT  I.IncentiveID as OfferID from CPE_Incentives I with (NoLock) " &
                                           "  WHERE I.IncentiveID =" & offer.OfferID & " " & conditionalQuery & " "
                    rst2 = MyCommon.LRT_Select
                    If rst2.Rows.Count > 0 Then
                        tempOfferLst.Add(offer)
                    End If
                Next
            End If
        Else
            tempOfferLst = lstOffers
        End If
        Return tempOfferLst
    End Function

    Public Function GetRestrictedAccessToUEBuilderQuery(ByVal MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc, ByVal TableAlias As String) As String
        Dim conditionalQuery As String = String.Empty
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        If (bEnableRestrictedAccessToUEOfferBuilder) Then
            If (Logix.UserRoles.CreateUEOffers AndAlso Not Logix.UserRoles.AccessTranslatedUEOffers) Then
                If (Not String.IsNullOrEmpty(TableAlias)) Then
                    conditionalQuery = "  and isnull( " & TableAlias & ".InboundCRMEngineID,0) <> 10 "
                Else
                    conditionalQuery = "  and isnull(InboundCRMEngineID,0) <> 10 "
                End If
            ElseIf (Not Logix.UserRoles.CreateUEOffers AndAlso Logix.UserRoles.AccessTranslatedUEOffers) Then
                If (Not String.IsNullOrEmpty(TableAlias)) Then
                    conditionalQuery = "  and (isnull( " & TableAlias & ".InboundCRMEngineID,0) = 10  or " & TableAlias & ".EngineID =0 )"
                Else
                    conditionalQuery = "  and (isnull(InboundCRMEngineID,0) = 10 or EngineID =0)"
                End If
            ElseIf (Not Logix.UserRoles.CreateUEOffers AndAlso Not Logix.UserRoles.AccessTranslatedUEOffers) Then
                If (Not String.IsNullOrEmpty(TableAlias)) Then
                    conditionalQuery = "  and " & TableAlias & ".EngineID =0 "
                Else
                    conditionalQuery = "  and EngineID =0  "
                End If
            End If
        End If

        Return conditionalQuery
    End Function
    Public Function ValidateTiers(ByVal offerID As Integer, ByVal NewTierTypeID As Integer, ByVal NumOfTiers As Integer) As Boolean
        Dim rst As New DataTable
        Dim rst1 As New DataTable
        Dim Common As New Copient.CommonInc
        Dim oldtiertypeid As Integer
        Dim oldNumTiers As Integer
        Dim conditionid As Integer
        Dim Rewardid As Integer
        Dim LinkId As Integer
        If offerID <> 0 Then
            If NewTierTypeID = 0 Then
                If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
                Common.QueryStr = "SELECT TierTypeID,NumTiers from offers where offerid=" & offerID
                rst = Common.LRT_Select
                oldNumTiers = rst.Rows(0).Item("NumTiers")
                oldtiertypeid = rst.Rows(0).Item("TierTypeID")
                Common.QueryStr = "select * from offerconditions as OC inner join conditiontiers as C on OC.conditionid=C.conditionid where OC.offerid=" & offerID & " and OC.tiered=1 and C.tierlevel<>0"
                rst1 = Common.LRT_Select
                If rst1 IsNot Nothing AndAlso rst1.Rows.Count > 0 Then
                    Return True
                End If
                Common.QueryStr = "SELECT RewardID,LinkID from offerrewards where offerid=" & offerID & " and tiered<>0"
                rst = Common.LRT_Select
                If rst IsNot Nothing AndAlso rst.Rows.Count > 0 Then
                    For Each row In rst.Rows
                        Rewardid = row.item("rewardid")
                        LinkId = row.item("LinkID")
                        Common.QueryStr = "SELECT * from RewardTiers where TierLevel<>0 and RewardId=" & Rewardid
                        rst = Common.LRT_Select
                        If rst IsNot Nothing AndAlso rst.Rows.Count > 0 Then
                            Return True
                        End If
                        Common.QueryStr = "SELECT * from RewardCustomerGroupTiers where TierLevel<>0 and RewardId=" & Rewardid
                        rst = Common.LRT_Select
                        If rst IsNot Nothing AndAlso rst.Rows.Count > 0 Then
                            Return True
                        End If
                        Common.QueryStr = "SELECT * from CashierMessageTiers where TierLevel<>0 and MessageID=" & LinkId
                        rst = Common.LRT_Select
                        If rst IsNot Nothing AndAlso rst.Rows.Count > 0 Then
                            Return True
                        End If
                        Common.QueryStr = "SELECT * from PrintedMessageTiers where TierLevel<>0 and MessageID=" & LinkId
                        rst = Common.LRT_Select
                        If rst IsNot Nothing AndAlso rst.Rows.Count > 0 Then
                            Return True
                        End If
                    Next
                End If
                Return False
            Else
                Return False
            End If
        Else
            Return False
        End If
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
    Public Function GetOfferEngineID(ByRef Common As Copient.CommonInc, ByVal OfferId As Long) As Int32
        Dim statusFlag As Int32 = 0
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        Common.QueryStr = "select OI.EngineID from OfferIDs as OI (NOLOCK) WHERE OI.OfferID = " & OfferId
        Dim dt As DataTable = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            statusFlag = Convert.ToInt32(dt.Rows(0)(0))
        End If
        Return statusFlag
    End Function
    Public Function IsOfferWaitingForApproval(ByVal OfferId As Long) As Boolean
        CMS.AMS.CurrentRequest.Resolver.AppName = "LogixCB.vb"
        Dim m_OAWService As CMS.AMS.Contract.IOfferApprovalWorkflowService = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOfferApprovalWorkflowService)()
        Dim m_IsOfferAwaitingApproval = m_OAWService.CheckIfOfferIsAwaitingApproval(OfferId).Result
        Return m_IsOfferAwaitingApproval
    End Function

    Public Sub ResetOfferApprovalStatus(ByVal OfferId As Long)
        CMS.AMS.CurrentRequest.Resolver.AppName = "LogixCB.vb"
        Dim m_OAWService As CMS.AMS.Contract.IOfferApprovalWorkflowService = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOfferApprovalWorkflowService)()
        m_OAWService.ResetOfferApprovalStatus(OfferId)
    End Sub
    Public Sub ResetOfferApprovalStatus_MultipleOffers(ByRef Common As Copient.CommonInc, ByVal offersDt As DataTable)
        Common.QueryStr = "dbo.pt_ResetOfferApprovalStatus"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@OffersDT", SqlDbType.Structured).Value = offersDt
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    End Sub

    Public Sub SendNotificationsOfItemChange(ByVal ItemId As Int64, ByVal ItemType As Int32)
        CMS.AMS.CurrentRequest.Resolver.AppName = "LogixCB.vb"
        Dim offersList(-1) As Long
        Dim oawService As CMS.AMS.Contract.IOfferApprovalWorkflowService = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOfferApprovalWorkflowService)()
        Dim MyCommon As Copient.CommonInc = CMS.AMS.CurrentRequest.Resolver.Resolve(Of Copient.CommonInc)()
        Dim linkType As Int32 = -1
        If ItemType = 1 Then
            linkType = 1
        ElseIf ItemType = 2 Then
            linkType = 2
        ElseIf ItemType = 3 Then
            linkType = 8
        End If

        MyCommon.QueryStr = "dbo.pa_AssociatedOffers"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@LinkType", SqlDbType.Int).Value = linkType
        MyCommon.LRTsp.Parameters.Add("@LinkID", SqlDbType.Int).Value = ItemId
        Dim rst As DataTable = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        If (rst.Rows.Count > 0) Then
            ReDim offersList(rst.Rows.Count - 1)
            For i = 0 To rst.Rows.Count - 1
                offersList(i) = MyCommon.NZ(rst.Rows(i).Item("IncentiveID"), "")
            Next
        End If

        If offersList.Count > 0 Then
            Threading.ThreadPool.QueueUserWorkItem(New Threading.WaitCallback(AddressOf BackgroundNotificationWorker), New Object() {ItemId, ItemType, oawService, offersList})
		End If
    End Sub
    Public Sub BackgroundNotificationWorker(State As Object)
        Dim obj As Object() = State
        Dim ItemID As Int64 = obj(0)
        Dim ItemType As Int32 = obj(1)
        Dim oawService As CMS.AMS.Contract.IOfferApprovalWorkflowService = obj(2)
        Dim offersList() As Long = CType(obj(3), Long())

        oawService.SendNotificationEmail(ItemType + 4, ItemID, offerIds:=offersList.ToList())
    End Sub

    ''' <summary>
    ''' This method will retrieve the attributes from XSD (reading xsd from AgentFiles folder) in the form of dictionary
    ''' </summary>
    ''' <param name="xmlSchema">XSD </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertXMLtoSchemaDic(ByVal xmlSchema As String) As Dictionary(Of [String], List(Of [String]))
        Dim dic As Dictionary(Of [String], List(Of [String])) = New Dictionary(Of String, List(Of String))()

        Try
            Dim xsd = XDocument.Load(XmlReader.Create(New StringReader(xmlSchema)))
            Dim ns = xsd.Root.GetDefaultNamespace()
            Dim prefix = xsd.Root.GetNamespaceOfPrefix("xs")
            If prefix Is Nothing Then
                prefix = xsd.Root.GetNamespaceOfPrefix("xsd")
            End If
            Dim rootElement = xsd.Root.Element(prefix + "element")
            Dim sections = rootElement.Element(prefix + "complexType").Element(prefix + "sequence").Elements(prefix + "element").ToList()
            For Each section In sections
                Dim namesList1 As New List(Of String)()
                If Not section.HasElements AndAlso section.HasAttributes Then
                    namesList1.Add(section.Attribute("name").Value)
                Else
                    ' for each section element
                    Dim items = section.Element(prefix + "complexType").Element(prefix + "sequence").Elements(prefix + "element")
                    For Each item In items
                        If section.Attribute("minOccurs").Value = "1" OrElse item.Attribute("minOccurs").Value = "1" Then
                            namesList1.Add(section.Attribute("name").Value + "/" + item.Attribute("name").Value + "/*")
                        Else
                            namesList1.Add(section.Attribute("name").Value + "/" + item.Attribute("name").Value)
                        End If

                    Next
                End If
                dic.Add(section.Attribute("name").Value, namesList1)
            Next
        Catch exception As Exception
        End Try
        Return dic
    End Function
    ''' <summary>
    ''' This method converts the string containing filter attributes to a dictionary
    ''' </summary>
    ''' <param name="s"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
Public Function ConvertFilterStringToDic(s As String) As Dictionary(Of [String], List(Of [String]))
	Dim dic As Dictionary(Of [String], List(Of [String])) = New Dictionary(Of String, List(Of String))()

	If s.Length > 0 Then
		Dim values As String() = s.Split("|"C)
		For Each s1 As String In values
			Dim filteredXMlLst As New List(Of [String])()
			Dim values1 As String() = s1.Split("-"C)
			Dim TagIds As List(Of String) = values1(1).Split(","C).ToList()
			Dim keystr As String = values1(0) + "/"
			TagIds = TagIds.[Select](Function(r) String.Concat(keystr, r)).ToList()
			filteredXMlLst.AddRange(TagIds)
			dic.Add(values1(0), filteredXMlLst)

		Next
	End If
	Return dic
End Function
    ''' <summary>
    ''' This method returns the attributes which are available for user can select in the form of dictionary 
    ''' </summary>
    ''' <param name="fullst"></param>
    ''' <param name="filteredXMlLst"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReturnNonFilteredDic(fullst As Dictionary(Of [String], List(Of [String])), ByRef filteredXMlLst As Dictionary(Of [String], List(Of [String]))) As Dictionary(Of [String], List(Of [String]))
        Dim NonfilteredXMlLst As New Dictionary(Of [String], List(Of [String]))()

        For Each pair As KeyValuePair(Of String, List(Of [String])) In fullst
            If filteredXMlLst.ContainsKey(pair.Key) Then
                Dim templst As List(Of [String]) = filteredXMlLst(pair.Key)
                Dim templst1 As New List(Of [String])
                For Each s As String In pair.Value
                    If s.Contains("/*") AndAlso templst.Contains(s.Substring(0, s.Length - 2)) Then
                        Dim ind As Integer = templst.IndexOf(s.Substring(0, s.Length - 2))
                        templst.RemoveAt(ind)
                        templst.Insert(ind, s)
                    ElseIf s.Contains("/*") AndAlso Not templst.Contains(s.Substring(0, s.Length - 2)) Then
                        templst.Add(s)
                    ElseIf Not templst.Contains(s) Then
                        templst1.Add(s)
                    End If
                Next
                filteredXMlLst(pair.Key) = templst
                If (templst1.Count > 0) Then
                    NonfilteredXMlLst.Add(pair.Key, templst1)
                End If
            Else
                NonfilteredXMlLst.Add(pair.Key, pair.Value)
            End If
        Next
        Return NonfilteredXMlLst
    End Function
    Public Function TruncateWordAppendEllipsis(ByVal input As String, ByVal length As Integer) As String
        If input Is Nothing OrElse input.Length < length Then
            Return input
        End If

        Dim iNextSpace = input.LastIndexOf(" ", length)
        Return String.Format("{0}...", input.Substring(0, (IIf((iNextSpace > 0), iNextSpace, length))).Trim())

    End Function







    Public Sub LoadFilterData(ByVal referenceID As Integer, ByVal MyCommon As Copient.CommonInc, ByVal hasPermission As Boolean, Optional ByVal connectorID As Integer = 0)
        Dim dt As DataTable
        Send("<div style=""float: left; position: relative;"">")
        Send("<label for=""pa""><b>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</b></label><br />")
        Send("<select class=""longer"" id=""SelFullXML"" name=""SelFullXML"" size=""12""  multiple=""multiple"">")
        Dim XMLString = String.Empty
        MyCommon.QueryStr = "SELECT XSDFileName,FilteredAttributes FROM FilterOutputColumns with (NoLock) WHERE ReferenceId = " & referenceID
        MyCommon.QueryStr = MyCommon.QueryStr & " and ConnectorID =" & connectorID
        If (connectorID > 0) Then
        End If
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
            For Each row2 In dt.Rows
                Dim XMLfileName As String = MyCommon.NZ(row2.Item("XSDFileName"), "")
                Dim filterstr As String = MyCommon.NZ(row2.Item("FilteredAttributes"), "")
                Dim xsdPath As String = MyCommon.Get_Install_Path & "AgentFiles\" & XMLfileName
                Dim xsdstr As String = System.IO.File.ReadAllText(xsdPath)

                Dim filteredXMLDic As Dictionary(Of [String], List(Of String)) = ConvertFilterStringToDic(filterstr)
                Dim FullXMLDic As Dictionary(Of [String], List(Of String)) = ConvertXMLtoSchemaDic(xsdstr)
                Dim NonFilteredXMLDic As Dictionary(Of [String], List(Of String)) = ReturnNonFilteredDic(FullXMLDic, filteredXMLDic)

                For Each pair As KeyValuePair(Of String, List(Of [String])) In NonFilteredXMLDic
                    Send("<optgroup style=""font-weight:bold""   label=""" & pair.Key & """>")
                    For Each s As String In pair.Value
                        Dim opnValue As String = ""
                        Dim opnText As String = ""

                        If (s.Contains("/*")) Then
                            opnValue = s.Substring(0, s.Length - 2)
                            opnText = s.Substring(s.IndexOf("/") + 1, s.LastIndexOf("/") - s.IndexOf("/") - 1)
                            Send("<option value=""" & opnValue & """  alt=""" & opnValue & """ title=""" & opnValue & """ disabled=""true"">" & opnText & "</option>")
                        Else
                            opnValue = s
                            opnText = s.Substring(s.IndexOf("/") + 1)
                            Send("<option value=""" & opnValue & """  alt=""" & opnValue & """ title=""" & opnValue & """>" & opnText & "</option>")
                        End If

                    Next
                Next
                Send("</select>")
                Send("</div>")
                Send("<div style=""float: left; padding: 65px 10px 10px 10px; position: relative;"">")
                Send("<input type=""button"" class=""regular select"" id=""select"" " & IIf(NonFilteredXMLDic.Count = 0 OrElse Not hasPermission, "disabled=""disabled""", "") & " name="">>"" value="">>""  ;""/>&nbsp")
                Send("<br clear=""all"">")
                Send("<br class=""half"">")
                Send("<input type=""button"" class=""regular Deselect"" id=""Deselect"" " & IIf(Not hasPermission, "disabled=""disabled""", "") & " name=""<<"" value=""<<""  ;""/>&nbsp")
                Send("</div>")

                Send("<div style=""float: left; position: relative;"">")
                Send("<label for=""ps""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label><br />")
                Send("<select class=""longer"" id=""selFilterXML"" name=""selFilterXML"" size=""12""  multiple=""multiple"" >")
                Dim defaultFilters As String = String.Empty
                For Each pair As KeyValuePair(Of String, List(Of [String])) In filteredXMLDic
                    Send("<optgroup style=""font-weight:bold"" label=""" & pair.Key & """>")
                    For Each s As String In pair.Value
                        Dim opnValue As String = ""
                        Dim opnText As String = ""

                        If (s.Contains("/*")) Then
                            opnValue = s.Substring(0, s.Length - 2)
                            opnText = s.Substring(s.IndexOf("/") + 1, s.LastIndexOf("/") - s.IndexOf("/") - 1)
                            defaultFilters = defaultFilters & opnValue & ","
                            Send("<option  value=""" & opnValue & """ alt=""" & opnValue & """ title=""" & opnValue & """ disabled=""true"">" & opnText & "</option>")
                        Else
                            opnValue = s
                            opnText = s.Substring(s.IndexOf("/") + 1)
                            Send("<option value=""" & opnValue & """  alt=""" & opnValue & """ title=""" & opnValue & """>" & opnText & "</option>")
                        End If
                    Next
                Next
                Send("</select>")
                Send("<input type=""hidden"" id=""defaultFilters"" name=""defaultFilters"" value=""" & defaultFilters & """/>")
                Send("<br>")
                Send("</div>")
                Send("<br clear=""left"">")
            Next
        End If
    End Sub

End Class
