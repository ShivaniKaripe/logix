﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %><%@ Import Namespace="System.Net.Http" %>

<%@ Import Namespace="System.Collections.Generic" %>
<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<script type="text/javascript">

    var nVer = navigator.appVersion;
    var nAgt = navigator.userAgent;
    var browserName = navigator.appName;
    var nameOffset, verOffset, ix;

    var browser = navigator.appName;

    // In Opera, the true version is after "Opera" or after "Version"
    if ((verOffset = nAgt.indexOf("Opera")) != -1) {
        browserName = "Opera";
    }
    // In MSIE, the true version is after "MSIE" in userAgent
    else if ((verOffset = nAgt.indexOf("MSIE")) != -1) {
        browserName = "IE";
    }
    // In Chrome, the true version is after "Chrome" 
    else if ((verOffset = nAgt.indexOf("Chrome")) != -1) {
        browserName = "Chrome";
    }
    // In Safari, the true version is after "Safari" or after "Version" 
    else if ((verOffset = nAgt.indexOf("Safari")) != -1) {
        browserName = "Safari";
    }
    // In Firefox, the true version is after "Firefox" 
    else if ((verOffset = nAgt.indexOf("Firefox")) != -1) {
        browserName = "Firefox";
    }
    // In most other browsers, "name/version" is at the end of userAgent 
    else if ((nameOffset = nAgt.lastIndexOf(' ') + 1) <
          (verOffset = nAgt.lastIndexOf('/'))) {
        browserName = nAgt.substring(nameOffset, verOffset);
        fullVersion = nAgt.substring(verOffset + 1);
        if (browserName.toLowerCase() == browserName.toUpperCase()) {
            browserName = navigator.appName;
        }
    }


    if (browserName == "IE") {
        document.attachEvent("onclick", PageClick);
    }
    else {
        document.onclick = function (evt) {
            var target = document.all ? event.srcElement : evt.target;
            if (target.href) {
                if (IsFormChanged(document.mainform)) {
                    var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                    return bConfirm;
                }
            }
        };
    }
    function PageClick(evt) {
        var target = document.all ? event.srcElement : evt.target;

        if (target.href) {
            if (IsFormChanged(document.mainform)) {
                var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                return bConfirm;
            }
        }
    }

   
</script>
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

    Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""

    Dim AdminUserID As Long
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim MyCryptLib As New Copient.CryptLib
    Dim MyAltID As New Copient.AlternateID
    Dim AltIDResponse As New Copient.AlternateID.CreateUpdateResponse
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim rst4 As DataTable
    Dim dt As DataTable
    Dim dt2 As DataTable
    'Card identifiers
    Dim CardPK As Long = 0
    Dim ExtCardID As String = ""
    Dim CardTypeID As Integer = 0

    'Customer identifiers
    Dim CustomerPK As Long = 0
    Dim CustomerLockPK As Long
    Dim CustomerTypeID As Integer = 0
    Dim IsHouseholdID As Boolean = False
    Dim Prefix As String = ""
    Dim FirstName As String = ""
    Dim MiddleName As String = ""
    Dim LastName As String = ""
    Dim Suffix As String = ""
    Dim FullName As String = ""

    '**
    Dim Comments As String = ""
    Dim DriverLicenseID As String = ""
    Dim TaxExemptID As String = ""
    Dim DateOpened As String = ""
    Dim DateOpened_month As String = ""
    Dim DateOpened_day As String = ""
    Dim DateOpened_year As String = ""
    Dim CreditLimit As Double = 0.0
    Dim APR As Double = 0.0
    Dim FullAddress As String = ""
    Dim Phone1 As String = ""
    Dim Phone2 As String = ""
    Dim Phone3 As String = ""
    Dim FullPhone As String = ""
    Dim Email As String = ""
    Dim Household As String = ""
    Dim DOB_month As String = ""
    Dim DOB_day As String = ""
    Dim DOB_year As String = ""
    Dim DOB As String = ""
    Dim AltIDValue As String = ""
    Dim AltIDVerifier As String = ""
    Dim EmployeeID As String = ""
    Dim Employee As Integer = 0
    Dim TestCard As Integer = 0
    Dim sHouseholdText As String
    Dim HHPK As Integer = 0
    Dim HouseholdID As String = ""
    Dim CustomerStatusID As Integer = 0
    Dim AirmileMemberID As String = ""
    Dim HasAirmileMemberID As Boolean = False

    Dim i As Integer = 0
    Dim j As Integer = 0
    Dim Shaded As String = " class=""shaded"""
    Dim Edit As Boolean
    Dim UnknownPhrase As String = ""
    Dim HasSearchResults As Boolean = False
    Dim ClientUserID1 As String = ""
    Dim searchterms As String = ""
    Dim restrictLinks As Boolean = False
    Dim SavedCardStatus As Integer = 0
    Dim AltIDField As Object = Nothing
    Dim AltIDTable As String = ""
    Dim AltIDCol As String = ""
    Dim AltIDVerCol As String = ""
    Dim NewAltID As String = ""
    Dim SaveFailed As Boolean = False
    Dim BannerID As Integer = 0
    Dim DateValid As Boolean = False
    Dim NullifyAltID As Boolean = False

    'Default URLs for links from this page
    Dim URLOfferSum As String = "offer-sum.aspx"
    Dim URLCPEOfferSum As String = "CPEoffer-sum.aspx"
    Dim URLcgroupedit As String = "cgroup-edit.aspx"
    Dim URLpointedit As String = "point-edit.aspx"
    Dim URLtrackBack As String = ""
    Dim inCardNumber As String = ""
    Dim extraLink As String = "" '<- The customer care remote links, if needed

    Response.Expires = 0
    CurrentRequest.Resolver.AppName ="Customer General"
    MyCommon.AppName = "customer-general.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()

    Dim MyLookup As New Copient.CustomerLookup(MyCommon)
    Dim Customers(-1) As Copient.Customer
    Dim HHCustomers(-1) As Copient.Customer
    Dim Cust As New Copient.Customer
    '**
    Dim CustNotes(-1) As Copient.CustomerNote
    Dim CustNote As New Copient.CustomerNote
    Dim Note As String = ""
    Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
    Dim TempDate As Date
    
    Dim DemotionPolicy As Integer = 0
    Dim DemotionDisplayed As Boolean = False
    Dim HHCount As Integer = 0
    Dim ValidTypeChange As Boolean = False

    Dim RulesEngine As New Copient.HouseholdRules(MyCommon)
    Dim HHOptions(-1) As Copient.HouseholdRules.InterfaceOption
    Dim HHQueuePKID As Long = 0
    Dim PendingRemoval As Boolean
    Dim QueueStatus As String = ""

    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim EnrollmentDate As String = ""
    Dim EnrollmentDate_month As String = ""
    Dim EnrollmentDate_day As String = ""
    Dim EnrollmentDate_year As String = ""
    Dim headers As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))
    Dim DatePartOrder As New List(Of Copient.LogixInc.DATE_PART)
    Dim CustomerServiceURL As String = ""
    dim InputBitSize as Integer
    Dim LockID As Integer = 0
    Dim DigitalReceipt as Integer
    Dim PaperReceipt as Boolean

    Dim bLinkedCards As Boolean = False
    Dim LinkedCardID As String = ""
    Dim LinkedCardTypeID As Integer = 0

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    MyLookup.SetAdminUserID(AdminUserID)
    MyLookup.SetLanguageID(LanguageID)

    UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)

    AltIDCol = ParseTableCol(MyCommon.Fetch_SystemOption(60))
    AltIDTable = ParseTable(MyCommon.Fetch_SystemOption(60))
    AltIDVerCol = ParseTableCol(MyCommon.Fetch_SystemOption(61))

    If MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(143), -1) = "1" Then
        bLinkedCards = True
    End If

    'Check the logged in user and see if they are to be restricted to this page
    restrictLinks = MyLookup.IsRestrictedUser(AdminUserID)

    CustomerPK = 0
    Edit = False

    'Special handling for customer inquiry direct link-in 
    If (restrictLinks) Then
        URLOfferSum = ""
        URLCPEOfferSum = ""
        URLcgroupedit = ""
        URLpointedit = ""
    End If

    DatePartOrder = Logix.GetShortDatePartOrder(MyCommon)

    'Set session to nothing, just to be sure
    Session.Add("extraLink", "")

    If (GetCgiValue("mode") = "summary") Then
        URLtrackBack = GetCgiValue("exiturl")
        inCardNumber = GetCgiValue("cardnumber")
        extraLink = "&mode=summary&exiturl=" & URLtrackBack & "&cardnumber=" & inCardNumber
        Session.Add("extraLink", extraLink)
    End If

    'Hack for pop-ups; check session for extra link
    If (Session("extraLink").ToString = "") Then
        extraLink = Session("extraLink")
    End If

    If (MyCommon.Extract_Val(GetCgiValue("CustPK")) = 0 AndAlso (MyCommon.Extract_Val(GetCgiValue("CustomerPK")) = 0)) Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "customer-inquiry.aspx")
    End If

    If (MyCommon.Extract_Val(GetCgiValue("CardPK")) > 0) Then
        CardPK = MyCommon.Extract_Val(GetCgiValue("CardPK"))
        ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
        CardTypeID = MyLookup.FindCardTypeFromCardPK(CardPK)
    End If

    headers.Add(New KeyValuePair(Of String, String)("Accept-Language", "en-us"))
    headers.Add(New KeyValuePair(Of String, String)("Application-Key", "ECOM_AURORA"))
    headers.Add(New KeyValuePair(Of String, String)("Organization-Key", "NCR"))

    CustomerServiceURL = MyCommon.Fetch_SystemOption(264)
    CustomerServiceURL = IIf((CustomerServiceURL Is Nothing) OrElse (CustomerServiceURL = "") OrElse (CustomerServiceURL.EndsWith("/")), CustomerServiceURL, CustomerServiceURL & "/")

    'Included (GetCgiValue("mode") <> "removeCardPoints") condition for Delete Customer & Points Button
    '<2.3.26>- AMS_FIS_hbc_J
    If ((GetCgiValue("mode") <> "removeCard") AndAlso _
        (GetCgiValue("mode") <> "addCard") AndAlso _
        (GetCgiValue("mode") <> "saveCard") AndAlso _
        (GetCgiValue("mode") <> "saveLinkedCard") AndAlso _
        (GetCgiValue("mode") <> "removeFromHH") AndAlso _
        (GetCgiValue("mode") <> "demoteFromHH") AndAlso _
        (GetCgiValue("mode") <> "unqueueFromHH") AndAlso _
        (GetCgiValue("save") = "")) AndAlso _
        (GetCgiValue("mode") <> "removeCardPoints") AndAlso _
        (MyCommon.Extract_Val(GetCgiValue("CardPK")) > 0 Or (GetCgiValue("searchterms") <> "" And _
        (GetCgiValue("Search") <> "" Or GetCgiValue("searchPressed") <> "")) Or _
        inCardNumber <> "" _
        ) Then
        If (GetCgiValue("mode") = "add") Then
            Cust = New Copient.Customer()
            If MyLookup.AddCustomer(Cust, ReturnCode) Then
                CustomerPK = Cust.GetCustomerPK
            End If
        End If

        ' Someone wants to search for a customer.  First, let's get their primary key from our database
        If (MyCommon.Extract_Val(GetCgiValue("CustPK")) > 0 Or (MyCommon.Extract_Val(GetCgiValue("CustomerPK")) > 0)) Then
            If (MyCommon.Extract_Val(GetCgiValue("CustPK")) > 0) Then
                CustomerPK = MyCommon.Extract_Val(GetCgiValue("CustPK"))
            ElseIf (MyCommon.Extract_Val(GetCgiValue("CustomerPK")) > 0) Then
                CustomerPK = MyCommon.Extract_Val(GetCgiValue("CustomerPK"))
            End If
            ReDim Customers(0)
            Customers(0) = MyLookup.FindCustomerInfo(CustomerPK, ReturnCode)
            If ReturnCode <> Copient.CustomerAbstract.RETURN_CODE.OK Then
                ReDim Customers(-1) 'If there is no customer found then there should be no record in the array
            End If
        Else
            ' If the page was called from an outside application, set ClientUserID1 to the outside passed in value
            If (inCardNumber <> "" And GetCgiValue("mode") = "summary") Then
                ClientUserID1 = MyCommon.Pad_ExtCardID(inCardNumber, CardTypeID)
                searchterms = GetCgiValue("searchterms")
                Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.ALL_CUSTOMER_TYPES, -1, ClientUserID1, ReturnCode)
            End If
            If (GetCgiValue("searchterms") <> "" And ClientUserID1 = "") Then
                ClientUserID1 = MyCommon.Pad_ExtCardID(GetCgiValue("searchterms"), CardTypeID)
                Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.ALL_CUSTOMER_TYPES, -1, ClientUserID1, ReturnCode)
                If Customers Is Nothing OrElse Customers.Length = 0 Then
                    Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.PHONE, -1, GetCgiValue("searchterms"), ReturnCode)
                End If
                If Customers Is Nothing OrElse Customers.Length = 0 Then
                    Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.LAST_NAME, -1, GetCgiValue("searchterms"), ReturnCode)
                End If
            End If
        End If

        If (Customers.Length = 1) Then
            CustomerPK = Customers(0).GetCustomerPK
            CustomerTypeID = Customers(0).GetCustomerTypeID
            If Customers(0).GetCustomerTypeID = 1 Then
                IsHouseholdID = True
            End If
            If (IsHouseholdID) Then
                HHCustomers = MyLookup.GetCustomersInHousehold(CustomerPK, ReturnCode)
            Else
                'Find household ID if one exists
                HHPK = Customers(0).GetHHPK
                If HHPK > 0 Then
                    HouseholdID = Customers(0).GetHouseHoldID
                End If
            End If
            FirstName = Customers(0).GetFirstName
            MiddleName = Customers(0).GetMiddleName
            LastName = Customers(0).GetLastName
            EmployeeID = Customers(0).GetEmployeeID
            HasSearchResults = False
            If (GetCgiValue("unlock") <> "") Then
                'MyLookup.UnlockCustomer(CustomerPK, HHPK, ReturnCode)
                Try
                    Integer.TryParse(GetCgiValue("LockID"), LockID)
                    Dim RESTServiceHelper As CMS.AMS.Contract.IRestServiceHelper = CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IRestServiceHelper)()
                    RESTServiceHelper.CallService(Of CMS.AMS.Models.CustomerLockInfo)(RESTServiceList.CustomerService, IIf((CustomerServiceURL Is Nothing) OrElse (CustomerServiceURL = ""), "", CustomerServiceURL & "unlock"), LanguageID, HttpMethod.Post, "{""lockId"": " & LockID.ToString() & "}", False, headers)
                Catch ex As Exception
                    infoMessage = ex.Message
                End Try
                
                MyCommon.Activity_Log(25, MyCommon.Extract_Val(GetCgiValue("CustPK")), AdminUserID, Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " " & ExtCardID & " " & Copient.PhraseLib.Lookup("term.unlocked", LanguageID))
            End If
        ElseIf Customers.Length > 1 Then
            infoMessage = "" & Copient.PhraseLib.Lookup("customer.multiplefound", LanguageID) & ""
        Else
            infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
            infoMessage = infoMessage & " <a href=""customer-general.aspx?mode=add&Search=Search" & extraLink & "&searchterms=" & GetCgiValue("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
        End If

        If (GetCgiValue("Edit") <> "") Then
            Edit = True
        End If

    ElseIf (GetCgiValue("Edit") <> "") Then
        'Displaying inputs for editing customer information/attributes
        Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.ALL_CUSTOMER_TYPES, -1, GetCgiValue("editterms"), ReturnCode)
        If (Customers.Length > 0) Then
            CustomerPK = Customers(0).GetCustomerPK
            If Customers(0).GetCustomerTypeID = 1 Then
                IsHouseholdID = True
            End If
            If (IsHouseholdID) Then
                HHCustomers = MyLookup.GetCustomersInHousehold(CustomerPK, ReturnCode)
            Else
                'Find household if one exists
                HHPK = Customers(0).GetHHPK
                If (HHPK > 0) Then
                    HouseholdID = Customers(0).GetHouseHoldID
                End If
            End If
        Else
            infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
        End If
        If (GetCgiValue("Edit") <> "") Then
            Edit = True
        End If


    ElseIf (GetCgiValue("save") <> "") Then


        'Saving customer information; first setup the page so it draws correctly
        If (GetCgiValue("CustomerPK") <> "") Then
            MyCommon.QueryStr = "select CustomerPK, CustomerTypeID, HHPK from Customers with (NoLock) where CustomerPK=" & GetCgiValue("CustomerPK") & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
                ReDim Customers(0)
                Customers(0) = MyLookup.FindCustomerInfo(MyCommon.Extract_Val(GetCgiValue("CustomerPK")), ReturnCode)
            End If
        Else
            Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.ALL_CUSTOMER_TYPES, -1, GetCgiValue("editterms"), ReturnCode)
        End If
        If (Customers.Length > 0) Then
            CustomerPK = Customers(0).GetCustomerPK
            SavedCardStatus = Customers(0).GetCardStatusID
            If Customers(0).GetCustomerTypeID = 1 Then
                IsHouseholdID = True
            End If
            If (IsHouseholdID) Then
                HHCustomers = MyLookup.GetCustomersInHousehold(CustomerPK, ReturnCode)
            Else
                'Find household if one exists
                HHPK = MyCommon.NZ(Customers(0).GetHHPK, 0)
                If (HHPK > 0) Then
                    HouseholdID = Customers(0).GetHouseHoldID
                End If
            End If
        End If

        'If the customer doesn't already have a CustomerExt record, create a blank one to work with
        MyCommon.QueryStr = "select CustomerPK from CustomerExt with (NoLock) where CustomerPK=" & CustomerPK & ";"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count = 0 Then
            MyCommon.QueryStr = "insert into CustomerExt with (RowLock) (CustomerPK) values (" & CustomerPK & ");"
            MyCommon.LXS_Execute()
        End If

        Household = MyCommon.Extract_Val(GetCgiValue("Household"))
        If (IsHouseholdID = True AndAlso Household = 0) OrElse (IsHouseholdID = False AndAlso Household = 1) Then
            'If we get in here, we know the user's changing a customer to a household (or vice versa),
            'which means we'll also need to change the CardTypeID of any associated cards.
            'This routine first determines if that's even possible, and if so, makes the change.
            '(It wouldn't be possible if there's already a customer card *and* a household card with the same number.)
            MyCommon.QueryStr = "select C.CardPK, C.ExtCardID, C.CardTypeId from CardIDs as C with (NoLock) " & _
                                "inner join (select ExtCardID, CardTypeId from CardIDs with (NoLock) where CustomerPK= " & CustomerPK & ") as ExistingCards " & _
                                "  on ExistingCards.ExtCardID = C.ExtCardID " & _
                                "where ExistingCards.CardTypeID = " & IIf(IsHouseholdID, 1, 0) & " and C.CardTypeID = " & IIf(IsHouseholdID, 0, 1)
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
                If IsHouseholdID Then
                    infoMessage = Copient.PhraseLib.Lookup("customer-general.UniquenessConstraintHousehold", LanguageID)
                Else
                    infoMessage = Copient.PhraseLib.Lookup("customer-general.UniquenessConstraintCustomer", LanguageID)
                End If
            Else
                ValidTypeChange = True
                MyCommon.QueryStr = "update CardIDs set CardTypeID=" & Household & " where CustomerPK=" & CustomerPK & ";"
                MyCommon.LXS_Execute()
            End If
        Else
            ValidTypeChange = True
        End If

        If ValidTypeChange Then
            If Logix.UserRoles.EditCustomerIdData Then
                'Set and save the customer's information

                Customers(0).SetPrefix(Logix.TrimAll(GetCgiValue("Prefix")))
                Customers(0).SetFirstName(Logix.TrimAll(GetCgiValue("FirstName")))
                Customers(0).SetMiddleName(Logix.TrimAll(GetCgiValue("MiddleName")))
                Customers(0).SetLastName(Logix.TrimAll(GetCgiValue("LastName")))
                Customers(0).SetSuffix(Logix.TrimAll(GetCgiValue("Suffix")))
                Customers(0).SetCustomerStatusID(MyCommon.Extract_Val(GetCgiValue("CustomerStatusID")))
                If Logix.UserRoles.AccessCustomerIdData_Address Then Customers(0).GetGeneralInfo.SetAddress(LTrim(RTrim(GetCgiValue("Address"))))
                If Logix.UserRoles.AccessCustomerIdData_City Then Customers(0).GetGeneralInfo.SetCity(LTrim(RTrim(GetCgiValue("City"))))
                If Logix.UserRoles.AccessCustomerIdData_State Then Customers(0).GetGeneralInfo.SetState(LTrim(RTrim(GetCgiValue("State"))))
                If Logix.UserRoles.AccessCustomerIdData_ZIP Then Customers(0).GetGeneralInfo.SetZip(LTrim(RTrim(GetCgiValue("Zip"))))
                If Logix.UserRoles.AccessCustomerIdData_Country Then Customers(0).GetGeneralInfo.SetCountry(LTrim(RTrim(GetCgiValue("Country"))))
                If Logix.UserRoles.AccessCustomerIdData_Phone Then
                    Phone1 = (Logix.TrimAll(GetCgiValue("Phone1")))
                    Phone2 = (Logix.TrimAll(GetCgiValue("Phone2")))
                    Phone3 = (Logix.TrimAll(GetCgiValue("Phone3")))
                    If (Phone1.Equals("")) Then
                        Customers(0).GetGeneralInfo.SetPhone(Phone1 & Phone2 & Phone3)
                    Else
                        Dim tmp As Boolean
                        If (Not UInt64.TryParse(Phone1, tmp)) Then
                            infoMessage = "Please enter a valid phone number"
                        Else
                            Customers(0).GetGeneralInfo.SetPhone(Phone1 & Phone2 & Phone3)
                        End If
                    End If
                End If
                If Logix.UserRoles.AccessCustomerIdData_MobilePhone Then
                    Phone1 = MyCommon.Parse_Quotes(Logix.TrimAll(GetCgiValue("MobilePhone1")))
                    Phone2 = MyCommon.Parse_Quotes(Logix.TrimAll(GetCgiValue("MobilePhone2")))
                    Phone3 = MyCommon.Parse_Quotes(Logix.TrimAll(GetCgiValue("MobilePhone3")))
                    Customers(0).GetGeneralInfo.SetMobilePhone(Phone1 & Phone2 & Phone3)
                End If

                If Not (String.IsNullOrEmpty(GetCgiValue("Email"))) Then
                    If (Logix.UserRoles.AccessCustomerIdData_Email AndAlso MyCommon.EmailAddressCheck((Logix.TrimAll(GetCgiValue("Email"))))) Then
                        Customers(0).GetGeneralInfo.SetEmail(GetCgiValue("Email"))
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("emailValidation", LanguageID)
                        SaveFailed = True
                    End If
                Else
                    Customers(0).GetGeneralInfo.SetEmail(GetCgiValue("Email"))
                End If

                Household = MyCommon.Parse_Quotes(Logix.TrimAll(GetCgiValue("Household")))
                Customers(0).SetCustomerTypeID(IIf(Household = "1", 1, 0))
                CustomerTypeID = Customers(0).GetCustomerTypeID
                Customers(0).SetPassword(Logix.TrimAll(GetCgiValue("Password")))
                AltIDValue = GetCgiValue("AltIDValue")
                Customers(0).SetAltID(AltIDValue)
                AltIDVerifier = GetCgiValue("AltIDVerifier")
                Customers(0).SetAltIdVerifier(AltIDVerifier)
                Customers(0).GetGeneralInfo.SetAirmileMemberID(GetCgiValue("AirmileMemberID"))
                AirmileMemberID = MyCommon.Parse_Quotes(GetCgiValue("AirmileMemberID"))
                Customers(0).SetBannerID(MyCommon.Extract_Val(GetCgiValue("BannerID")))
                EmployeeID = (Logix.TrimAll(GetCgiValue("EmployeeID")))
                If (GetCgiValue("Employee") <> "") Then
                    Customers(0).SetEmployeeID(EmployeeID)
                Else
                    Customers(0).SetEmployeeID("")
                    EmployeeID = ""
                End If
                If (GetCgiValue("Employee") = "on") Then
                    Employee = 1
                    Customers(0).SetEmployee(True)
                Else
                    Employee = 0
                    Customers(0).SetEmployee(False)
                End If
                If (GetCgiValue("TestCard") = "on") Then
                    TestCard = 1
                    Customers(0).SetTestCard(True)
                Else
                    TestCard = 0
                    Customers(0).SetTestCard(False)
                End If
                Customers(0).SetRestrictedRedemption(GetCgiValue("RestrictedRdmpt") = "on")

                If Logix.UserRoles.AccessCustomerIdData_DOB Then
                    'Format the date of birth
                    DOB_month = GetCgiValue("dob1")
                    DOB_day = GetCgiValue("dob2")
                    DOB_year = GetCgiValue("dob3")
                    DOB = ""
                    If (DOB_month = "" And DOB_day = "" And DOB_year = "") Then  'Allows a null to be set for the DOB when there is nothing saved in the DOB fields
                        DateValid = True
                        Customers(0).GetGeneralInfo.SetDateOfBirth(Nothing)
                    ElseIf (ValidateMonth(DOB_month) = False Or ValidateDay(DOB_day) = False Or ValidateYear(DOB_year) = False) Then   'If any part of the DOB is invalid then give the proper infomessage
                        Dim TempMessage As String = ""
                        If (ValidateMonth(DOB_month) = False) Then
                            TempMessage = "" & Copient.PhraseLib.Lookup("customer-general.invalidmonth", LanguageID) & "<br />"
                        End If
                        If (ValidateDay(DOB_day) = False) Then
                            TempMessage = TempMessage & "" & Copient.PhraseLib.Lookup("customer-general.invalidday", LanguageID) & "<br />"
                        End If
                        If (ValidateYear(DOB_year) = False) Then
                            TempMessage = TempMessage & "" & Copient.PhraseLib.Lookup("customer-general.invalidyear", LanguageID) & "<br />"
                        End If
                        If TempMessage <> "" Then DateValid = False
                        SaveFailed = True
                        infoMessage = TempMessage
                    Else
                        DOB = DOB_month.Trim.PadLeft(2, "0") & "/" & DOB_day.Trim.PadLeft(2, "0") & "/" & DOB_year.Trim.PadLeft(4, "0")
                        DateValid = Date.TryParse(DOB, New System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None, TempDate)
                        If DateValid Then
                            Customers(0).GetGeneralInfo.SetDateOfBirth(TempDate)
                        Else
                            infoMessage = Copient.PhraseLib.Lookup("customer-general.InvalidDOB", LanguageID)
                            SaveFailed = True
                        End If
                    End If
                Else
                    DateValid = True
                End If

                If Logix.UserRoles.AccessCustomerIdData_ARCustomer Then
                    '**Set ARCustomer flag
                    Customers(0).GetGeneralInfo.SetARCustomer(GetCgiValue("ARCustomer") = "on")
                End If

                If Logix.UserRoles.AccessCustomerIdData_Comments Then
                    Customers(0).GetGeneralInfo.SetComments(GetCgiValue("Comments"))
                    Comments = MyCommon.Parse_Quotes(GetCgiValue("Comments"))
                End If
                If Logix.UserRoles.AccessCustomerIdData_DriverLicenseID Then
                    Customers(0).GetGeneralInfo.SetDriverLicenseID(GetCgiValue("DriverLicenseID"))
                    DriverLicenseID = MyCommon.Parse_Quotes(GetCgiValue("DriverLicenseID"))
                End If
                If Logix.UserRoles.AccessCustomerIdData_TaxExemptID Then
                    Customers(0).GetGeneralInfo.SetTaxExemptID(GetCgiValue("TaxExemptID"))
                    TaxExemptID = MyCommon.Parse_Quotes(GetCgiValue("TaxExemptID"))
                End If
                If Logix.UserRoles.AccessCustomerIdData_CreditLimit Then
                    If IsNumeric(MyCommon.Parse_Quotes(GetCgiValue("CreditLimit"))) Then
                        Customers(0).GetAccountReceivableInfo.SetCreditLimit(GetCgiValue("CreditLimit"))
                        CreditLimit = MyCommon.Parse_Quotes(GetCgiValue("CreditLimit"))
                    End If
                End If
                If Logix.UserRoles.AccessCustomerIdData_APR Then
                    If IsNumeric(MyCommon.Parse_Quotes(GetCgiValue("APR"))) Then
                        Customers(0).GetAccountReceivableInfo.SetAPR(GetCgiValue("APR"))
                        APR = MyCommon.Parse_Quotes(GetCgiValue("APR"))
                    End If
                End If

                'Set CompoundCharge flag
                Customers(0).GetAccountReceivableInfo.SetCompoundCharge(GetCgiValue("CompoundCharge") = "on")
                'Set FinanceCharge flag
                Customers(0).GetAccountReceivableInfo.SetFinanceCharge(GetCgiValue("FinanceCharge") = "on")

                'Format date opened
                If Logix.UserRoles.AccessCustomerIdData_DateOpened Then
                    DateOpened_month = GetCgiValue("dateopened1")
                    DateOpened_day = GetCgiValue("dateopened2")
                    DateOpened_year = GetCgiValue("dateopened3")
                    DateOpened = ""
                    If (DateOpened_month = "" And DateOpened_day = "" And DateOpened_year = "") Then  'Allows a null to be set for the DateOpened when there is nothing saved in the DateOpened fields
                        DateValid = True
                        Customers(0).GetGeneralInfo.SetDateOpened(Nothing)
                    ElseIf (ValidateMonth(DateOpened_month) = False Or ValidateDay(DateOpened_day) = False Or ValidateYear(DateOpened_year) = False) Then   'If any part of the DateOpened is invalid then give the proper infomessage
                        Dim TempMessage As String = ""
                        If (ValidateMonth(DateOpened_month) = False) Then
                            TempMessage = "" & Copient.PhraseLib.Lookup("customer-general.invalidmonth", LanguageID) & "<br />"
                        End If
                        If (ValidateDay(DateOpened_day) = False) Then
                            TempMessage = TempMessage & "" & Copient.PhraseLib.Lookup("customer-general.invalidday", LanguageID) & "<br />"
                        End If
                        If (ValidateYear(DateOpened_year) = False) Then
                            TempMessage = TempMessage & "" & Copient.PhraseLib.Lookup("customer-general.invalidyear", LanguageID) & "<br />"
                        End If
                        If TempMessage <> "" Then DateValid = False
                        SaveFailed = True
                        infoMessage = TempMessage
                    Else
                        DateOpened = DateOpened_month.Trim.PadLeft(2, "0") & "/" & DateOpened_day.Trim.PadLeft(2, "0") & "/" & DateOpened_year.Trim.PadLeft(4, "0")
                        DateValid = Date.TryParse(DateOpened, New System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None, TempDate)
                        If DateValid Then
                            Customers(0).GetGeneralInfo.SetDateOpened(TempDate)
                        Else
                            infoMessage = Copient.PhraseLib.Lookup("customer-general.InvalidDateOpenedFormat", LanguageID)
                            SaveFailed = True
                        End If
                    End If
                Else
                    DateValid = True
                End If

                If Logix.UserRoles.AccessCustomerIdData_EnrollmentDate Then
                    'RT 4125 Format the ENROLLMENT date
                    EnrollmentDate_month = GetCgiValue("ed1")
                    EnrollmentDate_day = GetCgiValue("ed2")
                    EnrollmentDate_year = GetCgiValue("ed3")
                    EnrollmentDate = ""
                    If (EnrollmentDate_month = "" And EnrollmentDate_day = "" And EnrollmentDate_year = "") Then  'Allows a null to be set for the EnrollmentDate when there is nothing saved in the field
                        DateValid = True
                        Customers(0).SetEnrollmentDate(Nothing)
                    ElseIf (ValidateMonth(EnrollmentDate_month) = False Or ValidateDay(EnrollmentDate_day) = False Or ValidateYear(EnrollmentDate_year) = False) Then   'If any part of the date is invalid then give the proper infomessage
                        Dim TempMessage As String = ""
                        If (ValidateMonth(EnrollmentDate_month) = False) Then
                            TempMessage = "" & Copient.PhraseLib.Lookup("customer-general.invalidmonth", LanguageID) & "<br />"
                        End If
                        If (ValidateDay(EnrollmentDate_day) = False) Then
                            TempMessage = TempMessage & "" & Copient.PhraseLib.Lookup("customer-general.invalidday", LanguageID) & "<br />"
                        End If
                        If (ValidateYear(EnrollmentDate_year) = False) Then
                            TempMessage = TempMessage & "" & Copient.PhraseLib.Lookup("customer-general.invalidyear", LanguageID) & "<br />"
                        End If
                        If TempMessage <> "" Then DateValid = False
                        SaveFailed = True
                        infoMessage = TempMessage
                    Else
                        EnrollmentDate = EnrollmentDate_month.Trim.PadLeft(2, "0") & "/" & EnrollmentDate_day.Trim.PadLeft(2, "0") & "/" & EnrollmentDate_year.Trim.PadLeft(4, "0")
                        DateValid = Date.TryParse(EnrollmentDate, New System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None, TempDate)
                        If DateValid Then
                            Customers(0).SetEnrollmentDate(TempDate)
                        Else
                            infoMessage = Copient.PhraseLib.Lookup("customer-general.InvalidDateEnrolledFormat", LanguageID)
                            SaveFailed = True
                        End If
                    End If
                Else
                    DateValid = True
                End If

                'Handle updates to Alternate Identifier 
                If (infoMessage.Trim = "" AndAlso AltIDCol.Trim <> "") Then
                    NewAltID = GetNewAltID(AltIDCol, AltIDField)
                    If (AltIDCol.Trim <> "" And NewAltID.Trim <> "") Then
                        AltIDResponse = MyAltID.UpdateCustomerAltID(CustomerPK, NewAltID, MyCommon.NZ(Customers(0).GetBannerID, 0))
                        Select Case AltIDResponse
                            Case Copient.AlternateID.CreateUpdateResponse.ALTIDINUSE
                                infoMessage = Copient.PhraseLib.Detokenize("customer-general.AltIDInUse", LanguageID, AltIDField, NewAltID) 'Changes were not saved.  The unique alternate identifier is already in use by another customer<br />(AltIDField=NewAltID)
                                SaveFailed = True
                            Case Copient.AlternateID.CreateUpdateResponse.MEMBERNOTFOUND
                                infoMessage = Copient.PhraseLib.Detokenize("customer-general.CustomerNotFound", LanguageID, "")
                                SaveFailed = True
                            Case Copient.AlternateID.CreateUpdateResponse.ERROR_APPLICATION
                                infoMessage = Copient.PhraseLib.Detokenize("customer-general.AltIDError", LanguageID, MyAltID.ErrorMessage)
                                SaveFailed = True
                        End Select
                    ElseIf (AltIDTable.Trim <> "" AndAlso AltIDCol.Trim <> "" AndAlso NewAltID.Trim = "") Then
                        ' check to determine if the AltID something other than NULL, if so then we need to nullify it.
                        MyCommon.QueryStr = "select AltID from Customers with (NoLock) where AltID is not NULL and CustomerPK = " & CustomerPK
                        dt = MyCommon.LXS_Select
                        If dt.Rows.Count > 0 Then
                            NullifyAltID = True
                        End If
                    End If
                End If

                'Supplemental fields
                If MyCommon.Fetch_SystemOption(110) Then
                    Dim ActivityText As String = ""
                    Dim dtSup As DataTable
                    MyCommon.QueryStr = "select CSF.FieldID, CSF.Name, CSF.FieldTypeID, CSFT.Name As FieldTypeName, Length, Value, ISNULL(CS.Deleted, 1) as NoCustomerRecord " & _
                                        "from CustomerSupplementalFields as CSF with (NoLock) " & _
                                        "left join CustomerSupplementalFieldTypes as CSFT on CSFT.FieldTypeID=CSF.FieldTypeID " & _
                                        "left join CustomerSupplemental as CS on CS.FieldID=CSF.FieldID and CustomerPK=" & CustomerPK & " and CS.Deleted=0 " & _
                                        "where CSF.Deleted=0 " & IIf(Logix.UserRoles.AccessProtectedSupplementalFields, "", "and Editable=1 ") & _
                                        "order by CSF.FieldID;"
                    dtSup = MyCommon.LXS_Select
                    If dtSup.Rows.Count > 0 Then
                        Dim FieldID As Integer = 0
                        Dim FieldName As String = ""
                        Dim FieldTypeName As String = ""
                        Dim OldValue As String = ""
                        Dim SubmittedValue As String = ""
                        Dim SubmittedValueDate As Date
                        For Each row In dtSup.Rows
                            FieldID = row.Item("FieldID")
                            FieldName = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                            FieldTypeName = MyCommon.NZ(row.Item("FieldTypeName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                            If FieldTypeName = "Date" Then
                                If (MyCommon.NZ(GetCgiValue("CS" & FieldID & "-1"), "") <> "") AndAlso (MyCommon.NZ(GetCgiValue("CS" & FieldID & "-2"), "") <> "") AndAlso (MyCommon.NZ(GetCgiValue("CS" & FieldID & "-3"), "") <> "") Then
                                    SubmittedValue = GetCgiValue("CS" & FieldID & "-1") & "/" & GetCgiValue("CS" & FieldID & "-2") & "/" & GetCgiValue("CS" & FieldID & "-3")
                                Else
                                    SubmittedValue = ""
                                End If
                            Else
                                SubmittedValue = MyCommon.Parse_Quotes(MyCommon.NZ(GetCgiValue("CS" & FieldID), ""))
                            End If
                            If (SubmittedValue <> "") Then
                                'Validate value against field type
                                If (FieldTypeName = "Bit" And (SubmittedValue <> "0") AndAlso (SubmittedValue <> "1")) OrElse _
                                   (FieldTypeName = "Integer" And (Not IsNumeric(SubmittedValue) OrElse (Int(SubmittedValue) <> SubmittedValue))) OrElse _
                                   (FieldTypeName = "Decimal" And (Not IsNumeric(SubmittedValue))) OrElse _
                                   (FieldTypeName = "Date" And (Not IsDate(SubmittedValue))) Then
                                    SaveFailed = True
                                    infoMessage &= Copient.PhraseLib.Detokenize("customer-general.InvalidValue", LanguageID, StrConv(FieldTypeName, VbStrConv.Lowercase), row.Item("Name"))  'Invalid value submitted for {0} field "{1}".
                                End If
                            End If
                        Next
                        Dim IsPrivacyOn As Boolean = False
                        Dim NerID As Integer = 0
                        If infoMessage = "" Then
                            For Each row In dtSup.Rows
                                FieldID = MyCommon.NZ(row.Item("FieldID"), 0)
                                FieldName = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                FieldTypeName = MyCommon.NZ(row.Item("FieldTypeName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                OldValue = MyCommon.NZ(row.Item("Value"), "")
                                If FieldName = "Do not earn/redeem" Then
                                    NerID = FieldID
                                End If
                                If FieldTypeName = "Date" Then
                                    If (GetCgiValue("CS" & FieldID & "-1") <> "") AndAlso (GetCgiValue("CS" & FieldID & "-2") <> "") AndAlso (GetCgiValue("CS" & FieldID & "-3") <> "") Then
                                        SubmittedValue = GetCgiValue("CS" & FieldID & "-1") & "/" & GetCgiValue("CS" & FieldID & "-2") & "/" & GetCgiValue("CS" & FieldID & "-3")
                                    End If
                                    If Date.TryParse(SubmittedValue, SubmittedValueDate) Then
                                        SubmittedValue = SubmittedValueDate.ToString("dd MMM yyyy")
                                    End If
                                Else
                                    SubmittedValue = MyCommon.Parse_Quotes(MyCommon.NZ(GetCgiValue("CS" & FieldID), ""))
                                End If
                                If OldValue <> SubmittedValue Then
                                    ActivityText &= ", " & FieldName & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " "
                                    If FieldTypeName = "Bit" Then
                                        ActivityText &= IIf(OldValue = "1", Copient.PhraseLib.Lookup("term.true", LanguageID), Copient.PhraseLib.Lookup("term.false", LanguageID))
                                    Else
                                        ActivityText &= """" & MyCommon.Parse_Quotes(OldValue) & """"
                                    End If
                                    ActivityText &= " " & StrConv(Copient.PhraseLib.Lookup("term.to", LanguageID), VbStrConv.Lowercase) & " "
                                    If FieldTypeName = "Bit" Then
                                        ActivityText &= IIf(SubmittedValue = "1", Copient.PhraseLib.Lookup("term.true", LanguageID), Copient.PhraseLib.Lookup("term.false", LanguageID))
                                    Else
                                        ActivityText &= """" & MyCommon.Parse_Quotes(Left(Logix.TrimAll(SubmittedValue), 1000)) & """"
                                    End If
                                End If
                                If MyCommon.Fetch_SystemOption(128) Then
                                    If (SubmittedValue <> "") Then
                                        If (FieldID = NerID) Then
                                            IsPrivacyOn = True
                                        End If
                                    Else
                                        SubmittedValue = "0"
                                    End If

                                    Dim MyMassUpdate As Copient.MassUpdate
                                    MyMassUpdate = New Copient.MassUpdate(MyCommon, Copient.MassUpdate.CALLER.LOGIX_UI)
                                    Dim MyCustSuppAttribute As New Copient.CustomerSupplemental
                                    MyCustSuppAttribute.SetCardTypeID(CustomerTypeID)
                                    MyCustSuppAttribute.SetDeleted(False)
                                    MyCustSuppAttribute.SetExtCardID(ExtCardID)
                                    MyCustSuppAttribute.SetFieldID(FieldID)
                                    MyCustSuppAttribute.SetFieldValue(SubmittedValue)
                                    Dim bUpdated As Boolean = MyMassUpdate.EditCustomerSupplemental(MyCustSuppAttribute)

                                Else
                                    If (row.Item("NoCustomerRecord") = "1") Then
                                        'No preexisting value/record for this field
                                        If (SubmittedValue <> "") Then
                                            'The user's submitting one, so create it
                                            MyCommon.QueryStr = "insert into CustomerSupplemental (CustomerPK, FieldID, Value, LastUpdate) " & _
                                                             "values (" & CustomerPK & ", " & FieldID & ", '" & SubmittedValue & "', getdate());"
                                            MyCommon.LXS_Execute()
                                        End If
                                    Else
                                        'There's a preexisting value/record for this field
                                        If (SubmittedValue <> "") Then
                                            'Update the record
                                            MyCommon.QueryStr = "update CustomerSupplemental set Value='" & SubmittedValue & "' " & _
                                                             "where CustomerPK=" & CustomerPK & " and FieldID=" & FieldID & ";"
                                            MyCommon.LXS_Execute()
                                        Else
                                            'To keep record counts low, any field being set to empty gets its record physically deleted
                                            MyCommon.QueryStr = "delete from CustomerSupplemental " & _
                                                             "where CustomerPK=" & CustomerPK & " and FieldID=" & FieldID & ";"
                                            MyCommon.LXS_Execute()
                                        End If
                                    End If
                                End If  'closing of If MyCommon.Fetch_SystemOption(128) Then
                            Next
                            If IsPrivacyOn Then
                                infoMessage = "All members of household have been marked private."
                            End If

                            If ActivityText <> "" Then
                                ActivityText = Copient.PhraseLib.Lookup("history.customer-edited-supp", LanguageID) & ActivityText
                                If ActivityText.Length > 1000 Then
                                    ActivityText = Left(ActivityText, 997) & "..."
                                End If
                                MyCommon.Activity_Log2(25, 11, CustomerPK, AdminUserID, ActivityText)
                            End If
                        End If
                    End If
                End If
            End If

            If (MyCommon.Fetch_SystemOption(212) <> 0) Then

                DigitalReceipt = (GetCgiValue("hiddenDigitalReceipt"))>>1

                MyCommon.QueryStr = "update CustomerExt set DigitalReceipt = "& DigitalReceipt &" where CustomerPK= "& CustomerPK


                MyCommon.LXS_Execute()
                If MyCommon.RowsAffected > 0 Then
                    Dim ActivityText as String = "Digital Receipt Updated"
                    MyCommon.Activity_Log2(25, 11, GetCgiValue("CustPK"), AdminUserID, ActivityText)
                End If
            End If

                'Print Receipt DB submit				

            If (MyCommon.Fetch_SystemOption(310) <> 0) Then
                PaperReceipt= GetCgiValue("hiddenPaperReceipt")
                MyCommon.QueryStr = "update CustomerExt set PaperReceipt = "& IIF(PaperReceipt,1,0) &" where CustomerPK= "& CustomerPK
                MyCommon.LXS_Execute()
                If MyCommon.RowsAffected > 0 Then
                    Dim ActivityText as String = "Paper Receipt Updated"
                    MyCommon.Activity_Log2(25, 11, GetCgiValue("CustPK"), AdminUserID, ActivityText)
                End If
            End If

            'Attributes
            If MyCommon.Fetch_SystemOption(111) AndAlso Logix.UserRoles.AssignAttributes Then
                Dim ActivityText As String = ""
                Dim AttributeTypeID As Integer = 0
                Dim AttributeTypeName As String = ""
                Dim OldValueID As Integer = 0

                'Get master list of attributes
                MyCommon.QueryStr = "select AttributeTypeID, Description from AttributeTypes with (NoLock) where Deleted=0;"
                Dim dtAT As DataTable = MyCommon.LRT_Select

                If dtAT.Rows.Count > 0 Then

                    Dim anAttributeWasChanged As Boolean = False
                    For Each row In dtAT.Rows

                        AttributeTypeID = MyCommon.NZ(row.Item("AttributeTypeID"), 0)
                        AttributeTypeName = MyCommon.NZ(row.Item("Description"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))

                        'For each attribute, get the customer's currently value (if any)
                        MyCommon.QueryStr = "select AttributeValueID from CustomerAttributes with (NoLock) where CustomerPK=" & CustomerPK & " and AttributeTypeID=" & AttributeTypeID & ";"
                        Dim dtAV As DataTable = MyCommon.LXS_Select
                        If dtAV.Rows.Count > 0 Then
                            OldValueID = MyCommon.NZ(dtAV.Rows(0).Item("AttributeValueID"), 0)
                        End If

                        Dim SubmittedValueID As Integer = MyCommon.Parse_Quotes(MyCommon.NZ(GetCgiValue("at-" & AttributeTypeID), OldValueID))

                        If OldValueID <> SubmittedValueID Then

                            anAttributeWasChanged = True

                            If SubmittedValueID = 0 Then
                                'The user is clearing the value for this attribute, so delete the record
                                MyCommon.QueryStr = "update CustomerAttributes set Deleted = 1, AttributeValueID = 0, CPEStoreSendFlag = 1 where CustomerPK=" & CustomerPK & " and AttributeTypeID=" & AttributeTypeID & ";"
                                MyCommon.LXS_Execute()
                            Else
                                'The user's setting a value for this attribute...
                                If dtAV.Rows.Count > 0 Then
                                    '...and a record for it already exists, so update it
                                    MyCommon.QueryStr = _
                                    String.Format("update CustomerAttributes set Deleted = 0, AttributeValueID = {0}, CPEStoreSendFlag = 1 where CustomerPK = {1} and AttributeTypeID = {2};", _
                                              SubmittedValueID, CustomerPK, AttributeTypeID)

                                    MyCommon.LXS_Execute()
                                Else
                                    '...and no record for it exists, so insert it
                                    MyCommon.QueryStr = _
                                      String.Format("insert into CustomerAttributes ( CustomerPK, AttributeTypeID, AttributeValueID, CPEStoreSendFlag, Deleted, LastUpdate) values ( {0}, {1}, {2}, 1, 0, getdate() );", _
                                          CustomerPK, AttributeTypeID, SubmittedValueID)
                                    MyCommon.LXS_Execute()
                                End If

                            End If

                            'Assemble history
                            ActivityText &= ", " & AttributeTypeName & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " "
                            ActivityText &= """" & GetAttributeValueName(OldValueID) & """"
                            ActivityText &= " " & StrConv(Copient.PhraseLib.Lookup("term.to", LanguageID), VbStrConv.Lowercase) & " "
                            ActivityText &= """" & GetAttributeValueName(SubmittedValueID) & """"

                        End If

                    Next

                    If anAttributeWasChanged Then ' ActivityText <> "" Then

                        MyLookup.UpdateStoreSendFlag(CustomerPK)

                        ActivityText = Copient.PhraseLib.Lookup("history.customer-edited-attributes", LanguageID) & ActivityText
                        If ActivityText.Length > 1000 Then
                            ActivityText = Left(ActivityText, 997) & "..."
                        End If
                        MyCommon.Activity_Log2(25, 11, CustomerPK, AdminUserID, ActivityText)
                    End If
                End If
            End If

            If (infoMessage.Trim = "" AndAlso DateValid) Then

                If MyLookup.SaveCustomerInfo(Customers(0), ReturnCode) Then

                    If MyCommon.Fetch_SystemOption(146) Then
                        If ReturnCode = Copient.CustomerAbstract.RETURN_CODE.CUSTOMER_ALREADY_EXISTS Then 'if card already exists for another customer. 
                            infoMessage = Copient.PhraseLib.Lookup("customer-general.duplicatecardsdetected", LanguageID)
                        End If
                        If ReturnCode = Copient.CustomerAbstract.RETURN_CODE.INVALID_CUSTOMERTYPEID Then 'general information is deleted and causes cards to be deleted. 
                            infoMessage = Copient.PhraseLib.Lookup("cusomter-general.autoaddcardsdelete", LanguageID)
                        End If
                    End If

                    'If an empty string is sent for the AltID then NULL it out.
                    If NullifyAltID Then
                        MyCommon.QueryStr = String.Format("update {0} with (RowLock) set {1} = NULL where CustomerPK = {2}", AltIDTable, AltIDCol, CustomerPK)
                        MyCommon.LXS_Execute()

                        MyCommon.QueryStr = String.Format("update Customers with (RowLock) set CPEStoreSendFlag = 1 where CustomerPK = {0}", CustomerPK)
                        MyCommon.LXS_Execute()

                    End If

                    ' store the customer edit for later reporting use
                    MyCommon.QueryStr = "dbo.pt_CustomerEdits_Insert"
                    MyCommon.Open_LXSsp()
                    MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                    MyCommon.LXSsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LXSsp.Parameters.Add("@EditPK", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LXSsp.ExecuteNonQuery()
                    MyCommon.Close_LXSsp()

                Else
                    infoMessage = Copient.PhraseLib.Detokenize("customer-general.CouldNotSaveInfo", LanguageID, ReturnCode)
                End If

            End If
            Edit = True

            If (Not SaveFailed) AndAlso (infoMessage = "") Then
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "customer-general.aspx?edit=Edit&editterms=" & CustomerPK & "&CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, ""))
            End If

        End If

    ElseIf (GetCgiValue("mode") = "removeCard") And Not (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER)) Then
        MyCommon.QueryStr = "select CardPK from CardIDs with (NoLock) where CustomerPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & " and CardPK=" & MyCommon.Extract_Val(GetCgiValue("RemoveCardPK")) & ";"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
            If MyLookup.RemoveCardFromCustomer(MyCommon.Extract_Val(GetCgiValue("CustPK")), MyCommon.Extract_Val(GetCgiValue("RemoveCardPK")), ReturnCode) Then
                Response.Status = "301 Moved Permanently"
                If (MyCommon.Extract_Val(GetCgiValue("CustPK")) = 0) OrElse (MyCommon.Extract_Val(GetCgiValue("CustPK")) = MyCommon.Extract_Val(GetCgiValue("RemoveCardPK"))) Then
                    Response.AddHeader("Location", "customer-general.aspx?CustPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & extraLink)
                Else
                    Response.AddHeader("Location", "customer-general.aspx?CustPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & "&CardPK=" & MyCommon.Extract_Val(GetCgiValue("CardPK")) & extraLink)
                End If
                GoTo done
            End If

        End If

    ElseIf (GetCgiValue("mode") = "addCard") And Not (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER)) Then

        Dim NewExtCardID As String = ""

        Dim ExtCardTypeID As Integer = Convert.ToInt32(MyCommon.Extract_Val(GetCgiValue("AddCardTypeID")))

        NewExtCardID = MyCommon.Pad_ExtCardID(GetCgiValue("AddExtCardID"), ExtCardTypeID)
        If MyCommon.AllowToProcessCustomerCard(NewExtCardID, ExtCardTypeID, Nothing) Then
            MyCommon.QueryStr = "select CID.ExtCardIDOriginal as ExtCardID, CT.PhraseID, CT.Description from CardIDs as CID with (NoLock)" & _
                                " inner join CardTypes as CT on CT.CardTypeID=CID.CardTypeID " & _
                                " where CustomerPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & _
                                " and ExtCardID='" & MyCryptLib.SQL_StringEncrypt(NewExtCardID, True) & "'" & _
                                " and CT.CardTypeID=" & MyCommon.Extract_Val(GetCgiValue("AddCardTypeID")) & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
                infoMessage = Copient.PhraseLib.Detokenize("customer-general.CustomerAlreadyHas", LanguageID, IIf(IsDBNull(dt.Rows(0).Item("PhraseID")), MyCommon.NZ(dt.Rows(0).Item("Description"), ""), Copient.PhraseLib.Lookup(MyCommon.NZ(dt.Rows(0).Item("PhraseID"), 0), LanguageID)), NewExtCardID) 'This customer already has a {0} with number {1}.
            Else
                MyCommon.QueryStr = "select CID.ExtCardIDOriginal as ExtCardID, CT.PhraseID, CT.Description from CardIDs as CID with (NoLock)" & _
                                    " inner join CardTypes as CT on CT.CardTypeID=CID.CardTypeID " & _
                                    " where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(NewExtCardID, True) & "'" & _
                                    " and CT.CardTypeID=" & MyCommon.Extract_Val(GetCgiValue("AddCardTypeID")) & ";"
                dt2 = MyCommon.LXS_Select
                If dt2.Rows.Count > 0 Then
                    infoMessage = Copient.PhraseLib.Detokenize("customer-general.CardIDInUse", LanguageID, IIf(IsDBNull(dt2.Rows(0).Item("PhraseID")), MyCommon.NZ(dt2.Rows(0).Item("Description"), ""), Copient.PhraseLib.Lookup(MyCommon.NZ(dt2.Rows(0).Item("PhraseID"), 0), LanguageID)), NewExtCardID)  'Another customer already has a {0} with number {1}.
                Else
                    If MyLookup.AddCardToCustomer(MyCommon.Extract_Val(GetCgiValue("CustPK")), NewExtCardID, MyCommon.Extract_Val(GetCgiValue("AddCardTypeID")), MyCommon.Extract_Val(GetCgiValue("AddCardStatusID")), ReturnCode) Then
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "customer-general.aspx?CustPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & IIf(GetCgiValue("CardPK") <> "", "&CardPK=" & MyCommon.Extract_Val(GetCgiValue("CardPK")), "") & "&searchterms=" & GetCgiValue("HHPK") & "&search=Search" & extraLink)
                        GoTo done
                    Else
                        If ReturnCode = Copient.CustomerAbstract.RETURN_CODE.INVALID_CUSTOMERTYPEID Then
                            infoMessage = Copient.PhraseLib.Lookup("customer-general-addcardnotemployee", LanguageID)
                        ElseIf ReturnCode = Copient.CustomerAbstract.RETURN_CODE.INVALID_REWARDCARD Then
                            infoMessage = Copient.PhraseLib.Lookup("customer.invalidcard", LanguageID) & " (" & NewExtCardID & ")"
                        End If
                    End If
                End If
            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("term.invalidnumericcard", LanguageID)
        End If

    ElseIf (GetCgiValue("mode") = "removeCardPoints") Then
        '<2.3.26>- AMS_FIS_hbc_J
        ' Delete Customer Card & Points information
        Dim DeleteCustPoints As Integer = 0
        Try
            DeleteCustPoints = CInt(MyCommon.Fetch_SystemOption(121))
        Catch ex As Exception
            DeleteCustPoints = -1
        End Try
        If DeleteCustPoints = 1 Then
            If (MyCommon.Extract_Val(GetCgiValue("CardPK")) > 0) Then
                CardPK = MyCommon.Extract_Val(GetCgiValue("CardPK"))
                ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
            End If
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MyCommon.QueryStr = "delete from CardIds with (RowLock) where CustomerPK=" & GetCgiValue("CustPK")
            MyCommon.LXS_Execute()
            MyCommon.QueryStr = "delete from Points with (RowLock) where CustomerPK=" & GetCgiValue("CustPK")
            MyCommon.LXS_Execute()
            MyCommon.QueryStr = "delete from Customers with (RowLock) where CustomerPK=" & GetCgiValue("CustPK")
            MyCommon.LXS_Execute()
            Dim ActivityText As String = ""
            ActivityText = "Deleted customer record for card : " & ExtCardID
            MyCommon.Activity_Log2(25, 19, GetCgiValue("CustPK"), AdminUserID, ActivityText)
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "/logix/cgroup-list.aspx")
            GoTo done
        End If
    ElseIf (GetCgiValue("mode") = "saveCard") Then

        ' Save change to the card status
        CustomerPK = MyCommon.Extract_Val(GetCgiValue("CustPK"))
        Dim targetCardPK As String = MyCommon.Extract_Val(GetCgiValue("ChangedCardPK"))
        Dim newCardStatus As String = MyCommon.Extract_Val(GetCgiValue("CardStatusID"))
        Dim SaveAllowed As Boolean
        If Not MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
            MyCommon.QueryStr = "select CID.ExtCardIDOriginal as ExtCardID, CID.CardStatusID, CID.CardTypeID, CT.Description as CType,CT.PhraseID as PhraseID, CS.Description as CStatus from CardIDs as CID with (NoLock)" & _
                            "inner join CardTypes as CT on CT.CardTypeID = CID.CardTypeID " & _
                            "inner join CardStatus as CS on CS.CardStatusID = CID.CardStatusID " & _
                            "where CID.CardPK = " & targetCardPK & ";"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
            SaveAllowed= IsSavePermitted(CustomerPK,dt.Rows(0).Item("CardTypeID"),targetCardPK,MyCommon)
            If SaveAllowed Then
                MyLookup.ChangeCustomerCardStatus(CustomerPK, targetCardPK, newCardStatus)
            Else
                SaveFailed=True
                infoMessage=Copient.PhraseLib.Lookup("term.OneActive", LanguageID)&" " & dt.Rows(0).Item("CType")
            End If
        End If
        End If
        

        If GetCgiValue("EditMode") <> "" Then
            Edit = True
        End If

    ElseIf (GetCgiValue("mode") = "removeFromHH") Then

        Cust = MyLookup.FindCustomerInfo(MyCommon.Extract_Val(GetCgiValue("RemoveCustPK")), ReturnCode)
        If Cust IsNot Nothing Then
            HHOptions = RulesEngine.GetHouseholdingOptions()
            RulesEngine.SetLogFile("Householding.txt")
            If RulesEngine.SendToQueue(GetQueueEntryForRemove(MyCommon.Extract_Val(GetCgiValue("RemoveCustPK")), MyCommon.Extract_Val(GetCgiValue("CustPK")), AdminUserID, HHOptions), HHQueuePKID) Then
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "customer-general.aspx?CustPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & IIf(GetCgiValue("CardPK") <> "", "&CardPK=" & MyCommon.Extract_Val(GetCgiValue("CardPK")), "") & "&searchterms=" & GetCgiValue("HHPK") & "&search=Search" & extraLink)
                GoTo done
            End If
        End If
    ElseIf (GetCgiValue("mode") = "demoteFromHH") Then
        Cust = MyLookup.FindCustomerInfo(MyCommon.Extract_Val(GetCgiValue("DemoteCustPK")), ReturnCode)
        HHCount = MyLookup.GetCustomersInHousehold(MyCommon.Extract_Val(GetCgiValue("CustPK")), ReturnCode).Length
        If Cust IsNot Nothing Then
            HHOptions = RulesEngine.GetHouseholdingOptions()
            RulesEngine.SetLogFile("DemoteHouseholding.txt")
            If RulesEngine.SendToQueue(GetQueueEntryForRemove(MyCommon.Extract_Val(GetCgiValue("DemoteCustPK")), MyCommon.Extract_Val(GetCgiValue("CustPK")), AdminUserID, HHOptions), HHQueuePKID) Then
                MyCommon.Activity_Log(25, MyCommon.Extract_Val(GetCgiValue("CustPK")), AdminUserID, Copient.PhraseLib.Lookup("history.customer-remove-household", LanguageID))
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "customer-general.aspx?CustPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & IIf(GetCgiValue("CardPK") <> "", "&CardPK=" & MyCommon.Extract_Val(GetCgiValue("CardPK")), "") & "&searchterms=" & GetCgiValue("HHPK") & "&search=Search" & extraLink)
                GoTo done
            End If
        End If
    ElseIf (GetCgiValue("mode") = "unqueueFromHH") Then
        HHQueuePKID = MyCommon.Extract_Val(GetCgiValue("HHQueuePKID"))
        If HHQueuePKID > 0 Then
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MyCommon.QueryStr = "delete from HouseholdQueue where PKID=" & HHQueuePKID & ";"
            MyCommon.LXS_Execute()
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "customer-general.aspx?CustPK=" & MyCommon.Extract_Val(GetCgiValue("CustPK")) & IIf(GetCgiValue("CardPK") <> "", "&CardPK=" & MyCommon.Extract_Val(GetCgiValue("CardPK")), "") & "&searchterms=" & GetCgiValue("HHPK") & "&search=Search" & extraLink)
            GoTo done
        End If
    ElseIf (GetCgiValue("mode") = "saveLinkedCard") Then
        CustomerPK = MyCommon.Extract_Val(GetCgiValue("CustPK"))
        LinkedCardID = GetCgiValue("LinkedCardID")
        LinkedCardTypeID = MyCommon.Extract_Val(GetCgiValue("LinkedCardTypeID"))

        ' Make sure the card being linked to exists or that we're deleting an existing card
        MyCommon.QueryStr = "select CardPK from CardIDs where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(LinkedCardID, True) & "' and CardTypeID=" & LinkedCardTypeID & ";"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Or LinkedCardID = "" Then
            If IsNothing(LinkedCardID) Then
                LinkedCardID = "null"
            End If

            ' See if the customer has a row in the CustomerExt table
            MyCommon.QueryStr = "select CustomerPK from CustomerExt where CustomerPK=" & CustomerPK & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
                MyCommon.QueryStr = "update CustomerExt set LinkedCard=" & LinkedCardID & " where CustomerPK=" & CustomerPK & ";"
            Else
                MyCommon.QueryStr = "insert into CustomerExt (CustomerPK, LinkedCard) values (" & CustomerPK & "," & LinkedCardID & ");"
            End If
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MyCommon.LXS_Execute()
        Else
            infoMessage = Copient.PhraseLib.Lookup("term.linkedcardinvalid", LanguageID)
        End If
    End If

    'Load the customer
    If CustomerPK = 0 Then
        CustomerPK = MyCommon.Extract_Val(GetCgiValue("CustPK"))
    End If
    Cust = MyLookup.FindCustomerInfo(CustomerPK, ReturnCode)
    If Cust Is Nothing OrElse ReturnCode <> Copient.CustomerAbstract.RETURN_CODE.OK OrElse Cust.GetCustomerPK = 0 Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "customer-inquiry.aspx")
        GoTo done
    End If

    'Determine if the customer has an AirMile member ID
    MyCommon.QueryStr = "select AirmileMemberID from CustomerExt with (NoLock) where CustomerPK=" & CustomerPK & ";"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
        If (MyCommon.NZ(dt.Rows(0).Item("AirmileMemberID"), "") <> "") Then
            HasAirmileMemberID = True
        End If
    End If

    If CardPK > 0 Then
        Send_HeadBegin("term.customer", "term.general", MyCommon.TruncateString(ExtCardID, 34))
    Else
        Send_HeadBegin("term.customer", "term.general")
    End If
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    If restrictLinks Then
        Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
    End If
    Send_Scripts()
    Send_HeadEnd()

    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    If (Not restrictLinks) Then
        Send_Tabs(Logix, 3)
        If CardPK > 0 Then
            Send_Subtabs(Logix, 32, 3, LanguageID, CustomerPK, , CardPK)
        Else
            Send_Subtabs(Logix, 32, 3, LanguageID, CustomerPK)
        End If
    Else
        If CardPK > 0 Then
            Send_Subtabs(Logix, 91, 4, LanguageID, CustomerPK, extraLink, CardPK)
        Else
            Send_Subtabs(Logix, 91, 4, LanguageID, CustomerPK, extraLink)
        End If
    End If

    If (Logix.UserRoles.AccessCustomerInquiry = False) Then
        Send_Denied(1, "perm.customer-access")
        GoTo done
    End If

    If infoMessage = "" Then
        If GetCgiValue("infomessage") <> "" Then
            infoMessage = GetCgiValue("infomessage")
        End If
    End If
%>
<script type="text/javascript">
    function ChangeEmployeeID() {
        var emp = document.getElementById("EmployeeID");
        var foremp = document.getElementById("forEmployeeID");
        var empCheck = document.getElementById("Employee");

        if (empCheck.checked == false) {
            foremp.style.display = "none";
            emp.style.display = "none";
        } else {
            foremp.style.display = "block";
            emp.style.display = "block";
        }
    }

function popupPro() {
        var ProPopup = window.open("RapidSearch.aspx?DataID=&Callback=handleProClose","ProPopup","scrollbars=no,resizable=yes,width=590,height=560");
    }

function handleProClose(asAddress) {
	  
  var address = document.getElementById("Address");
  var city = document.getElementById("City");
  var state = document.getElementById("State");
  var zip = document.getElementById("Zip");
	
   if (asAddress[1] == ""){        
       address.value = asAddress[0]
    }
    else{
        address.value = asAddress[0]  + ', ' + asAddress[1]; 
    }    
        
    city.value = asAddress[2];
	state.value = asAddress[3];
	zip.value = asAddress[4];
    address.focus();
}

function ChangeARInfo() {
}

    function removeFromHH(CustPK, CardPK, ExtCardID, HHPK, HHCardPK, HHExtCardID, HHQueuePKID, action) {
        var tokenValues = [];
        var msg = '';

        tokenValues[0] = ExtCardID;

        if (action == 'remove') {
            msg = '<%Sendb(Copient.PhraseLib.Lookup("customer-general.ConfirmRemove", LanguageID))%>';
            msg = detokenizeString(msg, tokenValues);
            var response = confirm(msg);
            if (response) {
                if (CardPK > 0) {
                    document.location = "customer-general.aspx?mode=removeFromHH&CustPK=" + HHPK + "&CardPK=" + HHCardPK + "&RemoveCustPK=" + CustPK;
                } else {
                    document.location = "customer-general.aspx?mode=removeFromHH&CustPK=" + HHPK + "&RemoveCustPK=" + CustPK;
                }
            }
        } else if (action == 'demote') {
            msg = '<%Sendb(Copient.PhraseLib.Lookup("customer-general.ConfirmDemote", LanguageID))%>';
            msg = detokenizeString(msg, tokenValues);
            var response = confirm(msg);
            if (response) {
                if (CardPK > 0) {
                    document.location = "customer-general.aspx?mode=demoteFromHH&CustPK=" + HHPK + "&CardPK=" + HHCardPK + "&DemoteCustPK=" + CustPK;
                } else {
                    document.location = "customer-general.aspx?mode=demoteFromHH&CustPK=" + HHPK + "&DemoteCustPK=" + CustPK;
                }
            }
        } else if (action == 'unqueue') {
            msg = '<%Sendb(Copient.PhraseLib.Lookup("customer-general.ConfirmUnqueue", LanguageID))%>';
            msg = detokenizeString(msg, tokenValues);
            var response = confirm(msg); //customer-general.ConfirmUnqueue
            if (response) {
                document.location = "customer-general.aspx?mode=unqueueFromHH&CustPK=" + HHPK + "&CardPK=" + HHCardPK + "&HHQueuePKID=" + HHQueuePKID;
            }
        }
    }

    function addToHousehold(CustomerPK, CardPK) {
        if (CardPK > 0) {
            openPopup('customer-addhousehold.aspx?HHPK=' + CustomerPK + '&CardPK=' + CardPK);
        } else {
            openPopup('customer-addhousehold.aspx?HHPK=' + CustomerPK);
        }
    }

    function removeCard(CustPK, CardPK, Household, RemoveCardPK, RemoveExtCardID) {
        var tokenValues = [];
        var text = "";

        tokenValues[0] = RemoveExtCardID;


        text = '<%Sendb(Copient.PhraseLib.Lookup("customer-general.ConfirmDeleteCustomer", LanguageID))%>';

        text = detokenizeString(text, tokenValues);

        var response = confirm(text);
        if (response) {
            if (CardPK > 0) {
                document.location = "customer-general.aspx?mode=removeCard&CustPK=" + CustPK + "&CardPK=" + CardPK + "&RemoveCardPK=" + RemoveCardPK;
            } else {
                document.location = "customer-general.aspx?mode=removeCard&CustPK=" + CustPK + "&RemoveCardPK=" + RemoveCardPK;
            }
        }
    }

    function addCard(CustPK, CardPK) {
        var AddExtCardID = document.getElementById("NewExtCardID").value;
        var AddCardTypeID = document.getElementById("NewCardTypeID").value;
        var AddCardStatusID = document.getElementById("NewCardStatusID").value;
        AddExtCardID = encodeURI(AddExtCardID)

        if (AddExtCardID == "") {
            alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.SpecifyCard", LanguageID))%>');
        } else {
            if (CardPK > 0) {
                document.location = "customer-general.aspx?mode=addCard&CustPK=" + CustPK + "&CardPK=" + CardPK + "&AddExtCardID=" + AddExtCardID + "&AddCardTypeID=" + AddCardTypeID + "&AddCardStatusID=" + AddCardStatusID;
            } else {
                document.location = "customer-general.aspx?mode=addCard&CustPK=" + CustPK + "&AddExtCardID=" + AddExtCardID + "&AddCardTypeID=" + AddCardTypeID + "&AddCardStatusID=" + AddCardStatusID;
            }
        }
    }

    function saveCard(CustPK, CardPK, ChangedCardPK, EditMode) {
        var elem = document.getElementById('cardStatus' + ChangedCardPK);
        var qryStr = "";

        if (elem != null && CardPK > 0) {
            qryStr = "customer-general.aspx?mode=saveCard&CustPK=" + CustPK + "&CardPK=" + CardPK + "&CardStatusID=" + elem.value + "&ChangedCardPK=" + ChangedCardPK;
            if (EditMode) {
                qryStr = qryStr + "&EditMode=Edit"
            }
            document.location = qryStr
        } else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.UpdateError", LanguageID))%>');
        }
    }
     var input;
    function createHdnInput(ctrlId, value) {

        input = document.createElement("input");
        input.setAttribute("type", "hidden");
        input.setAttribute("name", ctrlId);
        input.setAttribute("id", ctrlId);
        input.setAttribute("value", value);

        return input;
    }

    function saveCard_Logix(ChangedCardPK, CustomerPK, CustPK, CardPK) {
        var elem = document.getElementById('cardstatusid' + ChangedCardPK);
        var cardid = document.getElementById('extcardid' + ChangedCardPK);
        var pin = document.getElementById('extpin' + ChangedCardPK);
        var pinverify = document.getElementById('extpinverify' + ChangedCardPK);

        var form = document.createElement("form");
        form.method = "post";
        form.action = "customer-general.aspx?mode=saveCard&CustomerPK=" + CustomerPK + "&CustPK=" + CustPK + "&CardPK=" + CardPK;

        form.appendChild(createHdnInput("CardStatusID", elem.value));
        form.appendChild(createHdnInput("ChangedCardPK", ChangedCardPK));
        form.appendChild(createHdnInput("extcardid", escape(cardid.value)));
        form.appendChild(createHdnInput("sTemplate", elem.value));
        form.appendChild(createHdnInput("pin", escape(pin.value)));
        form.appendChild(createHdnInput("pinverify", escape(pinverify.value)));
        form.appendChild(createHdnInput("exiturl", ""));

        document.body.appendChild(form);
        form.submit();
    }

    function removeCard_Logix(RemoveCardPK, RemoveExtCardID, CustomerPK, CustPK, CardPK) {
        var tokenValues = [];
        var text = "";

        tokenValues[0] = RemoveExtCardID;
        text = '<%Sendb(Copient.PhraseLib.Lookup("customer-general.ConfirmDeleteCustomer", LanguageID))%>';
        text = detokenizeString(text, tokenValues);

        var response = confirm(text);
        if (response) {
            var form = document.createElement("form");
            form.method = "post";
            form.action = "customer-general.aspx?mode=removeCard&CustomerPK=" + CustomerPK + "&CustPK=" + CustPK + "&CardPK=" + CardPK;
            
            form.appendChild(createHdnInput("RemoveCardPK", RemoveCardPK));
            form.appendChild(createHdnInput("exiturl", ""));

            document.body.appendChild(form);
            form.submit();
        }
    }

    function savePwd_Logix(CustomerPK, CustPK, CardPK) {
        var elem = document.getElementById('sharedpassword');
        var elemverify = document.getElementById('sharedpasswordverify');

        var form = document.createElement("form");
        form.method = "post";
        form.action = "customer-general.aspx?mode=savePwd&CustomerPK=" + CustomerPK + "&CustPK=" + CustPK + "&CardPK=" + CardPK;

        form.appendChild(createHdnInput("pin", escape(elem.value)));
        form.appendChild(createHdnInput("pinverify", escape(elemverify.value)));
        form.appendChild(createHdnInput("exiturl", ""));

        document.body.appendChild(form);
        form.submit();
    }

    function addCard_Logix(CustomerPK, CustPK, CardPK) {

        var AddExtCardID = document.getElementById("NewExtCardID").value;
        var AddCardTypeID = document.getElementById("NewCardTypeID").value;
        var AddCardStatusID = document.getElementById("NewCardStatusID").value;

        if (AddExtCardID == "") {
            alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.SpecifyCard", LanguageID))%>'); //To add a card to this customer,\nplease specify a card number.
        } else if (AddExtCardID.indexOf('<?') > -1) {
            alert("Invalid identifier format."); //The characters <? cause a server security error and should not be allowed as identifier
        } else {
            var form = document.createElement("form");
            form.method = "post";
            form.action = "customer-general.aspx?mode=addCard&CustomerPK=" + CustomerPK + "&CustPK=" + CustPK + "&CardPK=" + CardPK;

            form.appendChild(createHdnInput("AddExtCardID", escape(AddExtCardID)));
            form.appendChild(createHdnInput("AddCardTypeID", escape(AddCardTypeID)));
            form.appendChild(createHdnInput("AddCardStatusID", escape(AddCardStatusID)));
            form.appendChild(createHdnInput("exiturl", ""));

            document.body.appendChild(form);
            form.submit();
        }
    }


    function isValidEntry() {
        var retVal = true;

        // post localization, we are storing phone number as entered, so validation is not necessary
        // validate phone number
        //for (var i=1; i <= 3 && retVal; i++) {
        //  retVal = retVal && isValidPhonePart("Phone", i);
        //}
        //retVal = retVal && isValidPhoneCombo("Phone");

        // validate mobile phone number
        //for (var i=1; i <= 3 && retVal; i++) {
        //  retVal =  retVal && isValidPhonePart("MobilePhone", i);
        //}
        //retVal = retVal && isValidPhoneCombo("MobilePhone");

        // validate date of birth
        for (var i = 1; i <= 3 && retVal; i++) {
            retVal = retVal && isValidDobPart(i);
        }
        //**
        // validate date opened
        for (var i = 1; i <= 3 && retVal; i++) {
            retVal = retVal && isValidDateOpenedPart(i);
        }
        return retVal;
    }

    function isValidPhonePart(prefix, partNum) {
        var retVal = true;
        var elemPart = document.getElementById(prefix + partNum);

        if (elemPart != null) {
            if (partNum == 1 && elemPart.value != "" && elemPart.value.length != 3) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.PhoneAreaCode", LanguageID))%>');
                retVal = false;
            }
            if (partNum == 1 && (isNaN(elemPart.value) || parseInt(elemPart.value) < 0)) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.PhoneAreaCode", LanguageID))%>');
                retVal = false;
            }
            if (partNum == 2 && elemPart.value != "" && (elemPart.value.length != 3 || isNaN(elemPart.value) || parseInt(elemPart.value) < 0)) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.PhoneExchange", LanguageID))%>');
                retVal = false;
            }
            if (partNum == 3 && elemPart.value != "" && (elemPart.value.length != 4 || isNaN(elemPart.value) || parseInt(elemPart.value) < 0)) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.PhoneLocal", LanguageID))%>');
                retVal = false;
            }
            if (!retVal) {
                elemPart.focus();
                elemPart.select();
            }
        }

        return retVal;
    }

    function isValidPhoneCombo(prefix) {
        var retVal = false;
        var elemP1 = document.getElementById(prefix + "1");
        var elemP2 = document.getElementById(prefix + "2");
        var elemP3 = document.getElementById(prefix + "3");

        if (elemP1 != null && elemP2 != null && elemP3 != null) {
            if (elemP1.value != "" || elemP2.value != "" || elemP3.value != "") {
                // validate to acceptable formats (xxx) xxx-xxxx and xxx-xxxx
                retVal = (elemP1.value != "" && elemP1.value.length == 3 && elemP2.value != "" && elemP2.value.length == 3 && elemP3.value != "" && elemP3.value.length == 4);
                retVal = retVal || (elemP1.value == "" && elemP2.value != "" && elemP2.value.length == 3 && elemP3.value != "" && elemP3.value.length == 4);
            } else {
                // all phone parts are blank, so no phone number was provided to validate
                retVal = true;
            }
        }
        if (!retVal) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.PhoneNumber", LanguageID))%>');
        }

        return retVal;
    }

    function isValidDobPart(partNum) {
        var retVal = true;
        var elemPart = document.getElementById("dob" + partNum);

        if (elemPart != null) {
            if (elemPart.value != "" && isNaN(elemPart.value)) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.DateOfBirthMustBeNumber", LanguageID))%>');
                retVal = false;
                elemPart.focus();
                elemPart.select();
            }
        }

        return retVal;
    }
    function isValidDateOpenedPart(partNum) {
        //**
        var retVal = true;
        var elemPart = document.getElementById("dateopened" + partNum);

        if (elemPart != null) {
            if (elemPart.value != "" && isNaN(elemPart.value)) {
                alert('<%Sendb(Copient.PhraseLib.Lookup("customer-general.InvalidDateOpenedFormat", LanguageID))%>');
                retVal = false;
                elemPart.focus();
                elemPart.select();
            }
        }

        return retVal;
    }
    //<2.3.26>-AMS_FIS_hbc_J
    //Delete card & Points information for particular customer
    function removeCardIdPoints(CustPK, CardPK) {
        var response = confirm('<%Sendb(Copient.PhraseLib.Lookup("customer-general.ConfirmDeleteAll", LanguageID))%>');
        if (response) {
            if (CardPK > 0) {
                document.location = "customer-general.aspx?mode=removeCardPoints&CustPK=" + CustPK + "&CardPK=" + CardPK;
            }
        }
    }
    //-->

    function saveLinkedCard(CustPK,CardPK) {
       var saveLinkedExtCardID = document.getElementById("linkedCardID").value;
       var saveLinkedCardTypeID = document.getElementById("linkedCardTypeID").value;
       saveLinkedExtCardID = escape(saveLinkedExtCardID)

       if (CustPK > 0) {
          if (saveLinkedExtCardID == "") {
              document.location = "customer-general.aspx?mode=saveLinkedCard&CustPK=" + CustPK + "&CardPK=" + CardPK;
          }
          else {
              document.location = "customer-general.aspx?mode=saveLinkedCard&LinkedCardID=" + saveLinkedExtCardID + "&LinkedCardTypeID=" + saveLinkedCardTypeID + "&CustPK=" + CustPK + "&CardPK=" + CardPK;
          }
       }
    }

</script>
<script type="text/javascript" src="../javascript/jquery.min.js"></script>
<script type="text/javascript" src="../javascript/thickbox.js"></script>
<form id="mainform" name="mainform" method="post" action="customer-general.aspx" onsubmit="return isValidEntry();">
<input type="hidden" name="altid" value="<%Sendb(AltIDCol) %>" />
<input type="hidden" name="verifier" value="<%Sendb(AltIDVerCol) %>" />
	<%
	If (MyCommon.Fetch_SystemOption(212) <> 0) Then
			MyCommon.QueryStr = "select DigitalReceipt from CustomerExt where CustomerPK= "& CustomerPK &""
		  rst = MyCommon.LXS_Select	
		  If(rst IsNot Nothing AndAlso rst.Rows.Count > 0) 
		    DigitalReceipt=MyCommon.NZ(rst.Rows(0).Item("DigitalReceipt"), 0)
		  End If 
		  If (DigitalReceipt<>0)
			  DigitalReceipt=(DigitalReceipt<<1 OR 1)
		  End If	
	  End If  
	    %>
		<%Send(" <input type=""hidden"" name=""hiddenDigitalReceipt"" value="""& DigitalReceipt &""" id=""hiddenDigitalReceipt"">")%>
	
	
	
    <%
    If (MyCommon.Fetch_SystemOption(310) <> 0) Then						
        MyCommon.QueryStr = "select PaperReceipt from CustomerExt where CustomerPK= "& CustomerPK &""
        rst = MyCommon.LXS_Select
        If(rst IsNot Nothing AndAlso rst.Rows.Count > 0)
            PaperReceipt=(MyCommon.NZ(rst.Rows(0).Item("PaperReceipt"), False)) 
        End If
    End If
    %>
		<%Send(" <input type=""hidden"" id=""hiddenPaperReceipt"" name=""hiddenPaperReceipt"" value="""& PaperReceipt &""" >")%>
	 
<div id="intro">
    <h1 id="title">
        <%
            If CardPK = 0 Then
                If (IsHouseholdID OrElse Cust.GetCustomerTypeID = 1) Then
                    Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
                    IsHouseholdID = True
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
                End If
            Else
                'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                Dim TmpExtCardID As String = ExtCardID
                If MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1 Then
                    If MyLookup.FindCardTypeFromCardPK(CardPK) = 3 AndAlso TmpExtCardID.Length >= 14 Then
                        TmpExtCardID = TmpExtCardID.Substring(0, TmpExtCardID.Length - 4) & "****"
                    End If
                End If
                If (IsHouseholdID OrElse Cust.GetCustomerTypeID = 1) Then
                    Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & MyCommon.TruncateString(TmpExtCardID, 34))
                    IsHouseholdID = True
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & MyCommon.TruncateString(TmpExtCardID, 29))
                End If
            End If
            If SaveFailed Then
                MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & Cust.GetCustomerPK & ";"
                rst2 = MyCommon.LXS_Select
                If rst2.Rows.Count > 0 Then
                    FullName = IIf(MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") & " ", "")
                    FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") & " ", "")
                    FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
                    FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, MyCommon.NZ(rst2.Rows(0).Item("LastName"), ""), "")
                    FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & MyCommon.NZ(rst2.Rows(0).Item("Suffix"), ""), "")
                Else
                    FullName = IIf(Cust.GetPrefix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, Cust.GetPrefix & " ", "")
                    FullName &= IIf(Cust.GetFirstName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, Cust.GetFirstName & " ", "")
                    FullName &= IIf(Cust.GetMiddleName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(Cust.GetMiddleName, 1) & ". ", "")
                    FullName &= IIf(Cust.GetLastName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, Cust.GetLastName, "")
                    FullName &= IIf(Cust.GetSuffix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & Cust.GetSuffix, "")
                End If
            Else
                FullName = IIf(Cust.GetPrefix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, Cust.GetPrefix & " ", "")
                FullName &= IIf(Cust.GetFirstName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, Cust.GetFirstName & " ", "")
                FullName &= IIf(Cust.GetMiddleName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(Cust.GetMiddleName, 1) & ". ", "")
                FullName &= IIf(Cust.GetLastName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, Cust.GetLastName, "")
                FullName &= IIf(Cust.GetSuffix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & Cust.GetSuffix, "")
            End If
            If FullName <> "" Then
                Sendb(": " & MyCommon.TruncateString(FullName, 30))
            End If
            If (restrictLinks AndAlso URLtrackBack <> "") Then
                Send(" - <a href=""" & URLtrackBack & """>" & Copient.PhraseLib.Lookup("customer-inquiry.return", LanguageID) & "</a>")
            End If
        %>
    </h1>
    <div id="controls" <% Sendb(IIf(restrictLinks AndAlso URLtrackBack <> "", " style=""width:115px;""", "")) %>>
        <%
            If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
                Send_CustomerNotes(CustomerPK, CardPK)
            End If
        %>
    </div>
</div>
<div id="main">
    <%
        If (infoMessage <> "") Then
            Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br />")
        End If
    %>
    <input type="hidden" id="CustomerPK" name="CustomerPK" value="<%Sendb(Cust.GetCustomerPK)%>" />
    <%
        If (CardPK > 0) Then
            Send("    <input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
        End If
        If (GetCgiValue("mode") = "summary") Then
            Send("    <input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
            Send("    <input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
        End If
    %>
    <div id="column">
        <%
            Send("<div id=""statusbar"" class=""red-background"" style=""display: none;""></div>")
            Send("<input type=""hidden"" id=""editterms"" name=""editterms"" value=""" & CustomerPK & """ />")
            If (CustomerPK > 0) Then
                Send("<div class=""box"" id=""identity""" & IIf(Cust.GetCustomerPK = 0, " style=""display: none;""", "") & ">")
                If (Cust IsNot Nothing) Then
                    Sendb("<h2><span>")
                    If (IsHouseholdID) Then
                        Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.identity", LanguageID), VbStrConv.Lowercase))
                    Else
                        Sendb(Copient.PhraseLib.Lookup("term.cardholder", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.identity", LanguageID), VbStrConv.Lowercase))
                    End If
                    Send("</span></h2>")
                    If (Not Edit) Then
                        'Build up the full address
                        FullAddress = ""
                        If Cust.GetCustomerPK > 0 Then
                            If Cust.GetGeneralInfo.GetAddress <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Address Then
                                FullAddress &= MyCommon.TruncateString(MyCommon.NZ(Cust.GetGeneralInfo.GetAddress, ""), 25) & "<br />"
                            End If
                            If Cust.GetGeneralInfo.GetCity <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_City Then
                                FullAddress &= MyCommon.TruncateString(MyCommon.NZ(Cust.GetGeneralInfo.GetCity, ""), 25) & ", "
                            End If
                            If Cust.GetGeneralInfo.GetState <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_State Then
                                FullAddress &= MyCommon.TruncateString(MyCommon.NZ(Cust.GetGeneralInfo.GetState, ""), 25) & "&nbsp;"
                            End If
                            If Cust.GetGeneralInfo.GetZip <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_ZIP Then
                                FullAddress &= MyCommon.NZ(Cust.GetGeneralInfo.GetZip, "") & "&nbsp;"
                            End If
                        End If
                        'Build up the full phone
                        FullPhone = ""
                        If Cust.GetCustomerPK > 0 Then
                            If Cust.GetGeneralInfo.GetPhone <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Phone Then
                                FullPhone &= MyLookup.FormatPhoneNumber(MyCommon.NZ(Cust.GetGeneralInfo.GetPhone, "&nbsp;")) & "<br />"
                            End If
                            If Cust.GetGeneralInfo.GetMobilePhone <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MobilePhone Then
                                FullPhone &= MyLookup.FormatPhoneNumber(MyCommon.NZ(Cust.GetGeneralInfo.GetMobilePhone, "&nbsp;"))
                            End If
                        End If
                        'Get email address
                        If Cust.GetCustomerPK > 0 Then
                            If Cust.GetGeneralInfo.GetEmail <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Email Then
                                Email = MyCommon.NZ(Cust.GetGeneralInfo.GetEmail, "&nbsp;") & "<br />"
                            End If
                        End If
                        'Get date of birth
                        If Cust.GetCustomerPK > 0 Then
                            If MyCommon.NZ(Cust.GetGeneralInfo.GetDateOfBirth, "1/1/1900") > "1/1/1900" AndAlso Logix.UserRoles.AccessCustomerIdData_DOB Then
                                DOB = Cust.GetGeneralInfo.GetDateOfBirth & "<br />"
                            End If
                        End If
                
                        If FullName <> "" OrElse FullAddress <> "" OrElse FullPhone <> "" OrElse Email <> "" Then
                            Send("<table summary=""" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & """>")
                            Send("<thead>")
                            Send("  <tr>")
                            If FullName <> "" Then
                                Send("    <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
                            End If
                            If FullAddress <> "" Then
                                Send("    <th class=""th-address"" scope=""col"">" & Copient.PhraseLib.Lookup("term.address", LanguageID) & "</th>")
                            End If
                            If FullPhone <> "" Then
                                Send("    <th class=""th-phone"" scope=""col"">" & Copient.PhraseLib.Lookup("term.phone", LanguageID) & "</th>")
                            End If
                            If Email <> "" Then
                                Send("    <th class=""th-email"" scope=""col"">" & Copient.PhraseLib.Lookup("term.email", LanguageID) & "</th>")
                            End If
                            If DOB <> "" Then
                                Send("    <th class=""th-email"" scope=""col"">" & Copient.PhraseLib.Lookup("term.dob", LanguageID) & "</th>")
                            End If
                            Send("  </tr>")
                            Send("</thead>")
                            Send("<tbody>")
                            Send("  <tr class=""shaded"">")
                            If FullName <> "" Then
                                Send("    <td>" & FullName & "</td>")
                            End If
                            If FullAddress <> "" Then
                                Send("    <td>" & FullAddress & "</td>")
                            End If
                            If FullPhone <> "" Then
                                Send("    <td>" & FullPhone & "</td>")
                            End If
                            If Email <> "" Then
                                Send("    <td>" & Email & "</td>")
                            End If
                            If DOB <> "" Then
                                Send("    <td>" & DOB & "</td>")
                            End If
                            Send("  </tr>")
                            Send("</tbody>")
                            Send("</table>")
                            Send("<br class=""half"" />")
                        End If
              
                        'Supplemental data
                        If MyCommon.Fetch_SystemOption(110) Then
                            MyCommon.QueryStr = "select CSF.FieldID, CSF.Name, CSF.FieldTypeID, CSFT.Name As FieldTypeName, Visible, Value " & _
                                                "from CustomerSupplementalFields as CSF with (NoLock) " & _
                                                "left join CustomerSupplementalFieldTypes as CSFT on CSFT.FieldTypeID=CSF.FieldTypeID " & _
                                                "left join CustomerSupplemental as CS on CS.FieldID=CSF.FieldID and CustomerPK=" & CustomerPK & " and CS.Deleted=0 " & _
                                                "where CSF.Deleted=0 And Visible=1 " & _
                                                "order by CSF.Name;"
                            rst2 = MyCommon.LXS_Select
                            If rst2.Rows.Count > 0 Then
                                Send("<table summary=""" & Copient.PhraseLib.Lookup("term.customersupplementalfields", LanguageID) & """>")
                                Send("</table>")
                                Send("<br class=""half"" />")
                            End If
                        End If
              
                        ' Test card enabled?  (NB: this field is currently on the Customers table, so really it's a test *customer*, not a test card --HUW)
                        If MyCommon.Fetch_SystemOption(88) Then
                            If MyCommon.NZ(Cust.GetTestCard, False) Then
                                TestCard = 1
                                If Logix.UserRoles.AccessCustomerIdData_Test Then
                                    Send("<span class=""red"">" & Copient.PhraseLib.Lookup("customer.testcard", LanguageID) & "</span><br />")
                                End If
                            Else
                                TestCard = 0
                            End If
                        Else
                            TestCard = 0
                        End If
              
                        'Active/inactive?
                        If Logix.UserRoles.AccessCustomerIdData_Status Then
                            If MyCommon.NZ(Cust.GetCustomerStatusID, 0) = 1 Then
                                Send(Copient.PhraseLib.Lookup("customer.active", LanguageID) & "<br />")
                            ElseIf MyCommon.NZ(Cust.GetCustomerStatusID, 0) = 2 Then
                                Send(Copient.PhraseLib.Lookup("customer.inactive", LanguageID) & "<br />")
                            End If
                        End If
              
                        'Banner
                        If MyCommon.NZ(Cust.GetBannerID, 0) > 0 Then
                            MyCommon.QueryStr = "select BannerID, Name, Description from Banners with (NoLock) where BannerID=" & MyCommon.NZ(Cust.GetBannerID, 0) & ";"
                            rst2 = MyCommon.LRT_Select
                            If rst2.Rows.Count > 0 Then
                                Sendb("Cardholder is in the ")
                                If Logix.UserRoles.AccessBanners Then
                                    Sendb("<a href=""banner-edit.aspx?BannerID=" & MyCommon.NZ(rst2.Rows(0).Item("BannerID"), 0) & """>" & MyCommon.NZ(rst2.Rows(0).Item("Name"), "(unnamed)") & "</a>")
                                Else
                                    Sendb(MyCommon.NZ(rst2.Rows(0).Item("Name"), "(unnamed)"))
                                End If
                                Send(" banner.<br />")
                            End If
                        End If
              
                        ' Report employee and household status
                        If Logix.UserRoles.AccessCustomerIdData_Employee Then
                            If (MyCommon.NZ(Cust.GetEmployee, 0) = 0) Then
                                Employee = 0
                                Send(Copient.PhraseLib.Lookup("customer.nonemployee", LanguageID) & "<br />")
                            Else
                                Employee = 1
                                Sendb(Copient.PhraseLib.Lookup("customer.employee", LanguageID))
                                If Logix.UserRoles.AccessCustomerIdData_EmployeeID Then
                                    If Cust.GetEmployeeID <> "" Then Sendb(" (" & Cust.GetEmployeeID & ")")
                                End If
                                Send("<br />")
                            End If
                        End If
              
                        If Logix.UserRoles.ViewHHCardholders Then
                            If (Cust.GetHHPK > 0) Then
                                Sendb(Copient.PhraseLib.Lookup("customer.household", LanguageID))
                                MyCommon.QueryStr = "select CardPK,ExtCardIDOriginal as ExtCardID, CardTypeID from CardIDs with (NoLock) where CustomerPK=" & Cust.GetHHPK & ";"
                                rst2 = MyCommon.LXS_Select
                                For k As Integer = 1 To rst2.Rows.Count
                                    rst2.Rows(k - 1).Item("ExtCardID") = MyCryptLib.SQL_StringDecrypt(rst2.Rows(k - 1).Item("ExtCardID"))
                                Next
                                'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                                If rst2.Rows.Count > 0 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1 Then
                                    For k As Integer = 1 To rst2.Rows.Count
                                        If MyCommon.NZ(rst2.Rows(k - 1).Item("CardTypeID"), 0) = 3 Then
                                            rst2.Rows(k - 1).Item("ExtCardID") = rst2.Rows(k - 1).Item("ExtCardID").Substring(0, rst2.Rows(k - 1).Item("ExtCardID").Length - 4) & "****"
                                        End If
                                    Next
                                End If
                                If rst2.Rows.Count > 0 Then
                                    i = 1
                                    For Each row In rst2.Rows
                                        Send("<a href=""customer-general.aspx?CustPK=" & Cust.GetHHPK & "&amp;CardPK=" & MyCommon.NZ(row.Item("CardPK"), 0) & extraLink & """>" & MyCommon.NZ(row.Item("ExtCardID"), "") & "</a><br />")
                                        If i < rst2.Rows.Count Then
                                            Send("/")
                                        End If
                                        i += 1
                                    Next
                                End If
                            ElseIf Cust.GetHouseHoldID <> "" Then
                                Send(Copient.PhraseLib.Lookup("customer.household", LanguageID) & "<a href=""customer-general.aspx?searchterms=" & HouseholdID & "&amp;search=Search" & extraLink & """>" & HouseholdID & "</a><br />")
                            ElseIf (Not IsHouseholdID) Then
                                Send(Copient.PhraseLib.Lookup("customer.nonhousehold", LanguageID) & "<br />")
                            End If
                        End If
        								If Cust.GetHHPK > 0 Then
        										CustomerLockPK = Cust.GetHHPK
        										sHouseholdText = " (" & Copient.PhraseLib.Lookup("term.household", LanguageID) & ") "
        								Else
        										CustomerLockPK = Cust.GetCustomerPK
        										sHouseholdText = ""
        								End If
        								Dim IsCustomerLocked As Boolean = False
        								
        								If (Logix.UserRoles.ViewCustomerLockInfo OrElse Logix.UserRoles.ForceUnlockCustomer) AndAlso MyCommon.IsEngineInstalled(9) Then
                            'Customer Lock information 
                            Try
                                Dim RESTServiceHelper As CMS.AMS.Contract.IRestServiceHelper = CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IRestServiceHelper)()
                                Dim result As KeyValuePair(Of CMS.AMS.Models.CustomerLockInfo, String) = RESTServiceHelper.CallService(Of CMS.AMS.Models.CustomerLockInfo)(RESTServiceList.CustomerService, IIf((CustomerServiceURL Is Nothing) OrElse (CustomerServiceURL = ""), "", CustomerServiceURL & "lock/status?id=" & ExtCardID & "&type=" & CardTypeID.ToString()), LanguageID, HttpMethod.Get, "", False, headers)
                
                                infoMessage = result.Value
                                Dim lockDetails As CMS.AMS.Models.CustomerLockInfo = result.Key
                                If (lockDetails IsNot Nothing) Then
                                    Send(Copient.PhraseLib.Lookup("term.lockstatus", LanguageID) & ": " & IIf(lockDetails.lockinfo.locked, "<span style=""color:red;"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</span>", Copient.PhraseLib.Lookup("term.unlocked", LanguageID)) & "<br />")
                                    If lockDetails.lockinfo.locked Then
                                        IsCustomerLocked = True
                                        Send("<table class=""noPaddingMargin"" style=""padding: 0px 0px 0px 4px;margin:0 0 0 0"">")
                                        Send("<tr><td>" & Copient.PhraseLib.Lookup("term.lockedby", LanguageID) & ":</td><td>" & lockDetails.lockinfo.lockedBy.id & " (" & GetDisplayTextForCardType(lockDetails.lockinfo.lockedBy.type) & ")</td></tr>")
                                        Send("<tr><td>" & Copient.PhraseLib.Lookup("term.lockedon", LanguageID) & ":</td><td>" & lockDetails.lockinfo.lockedDate.ConvertToLocalDateTime() & "</td></tr>")
                                        Send("<tr><td>" & Copient.PhraseLib.Lookup("term.store", LanguageID) & ":</td><td>" & lockDetails.location.storeCode & "(" & lockDetails.location.storeName & ")</td></tr>")
                                        Send("<tr><td>" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & ":</td><td>" & lockDetails.location.terminalID & "</td></tr>")
                                        Send("<tr><td valign=""top"">" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & ":</td><td><div style=""display: inline-block;width: 500px;word-wrap:break-word;"">" & lockDetails.location.transactionID & "</div></td></tr></table><br />")
                                        LockID = lockDetails.lockinfo.lockId
                                    End If
                                Else
                                    Send("<span style=""color:red;"">" & Copient.PhraseLib.Lookup("term.customerserviceerror", LanguageID) & " " & infoMessage & "</span><br />")
                                End If
                            Catch ex As Exception
                                infoMessage = ex.Message
                            End Try

        								End If
        								
        								If HasAirmileMemberID AndAlso Logix.UserRoles.AccessCustomerIdData_AirmileMemberID Then
        										Send(Copient.PhraseLib.Lookup("term.airmilememberid", LanguageID) & ": " & MyCommon.TruncateString(MyCommon.NZ(Cust.GetGeneralInfo.GetAirmileMemberID, "&nbsp;"), 25) & "<br />")
        								End If
        								
        								
        								Send("<br class=""half"" />")
        								' Buttons
        								Dim EditableSupplementals As Integer = 0
        								If MyCommon.Fetch_SystemOption(110) Then
        										MyCommon.QueryStr = "select FieldID from CustomerSupplementalFields with (NoLock) where Deleted=0 and Visible=1" & IIf(Logix.UserRoles.AccessProtectedSupplementalFields, "", " and Editable=1") & ";"
        										rst = MyCommon.LXS_Select
        										EditableSupplementals = rst.Rows.Count
        								End If
        								Dim IsAllowedToEdit As Boolean = (Logix.UserRoles.EditCustomerIdData) OrElse (Logix.UserRoles.AssignAttributes AndAlso MyCommon.Fetch_SystemOption(111)) OrElse (EditableSupplementals > 0)
        								
        								If IsAllowedToEdit Then
        										Send("<input type=""button"" class=""regular"" id=""edit"" name=""edit"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""window.location.href='customer-general.aspx?edit=Edit" & extraLink & "&amp;editterms=" & Cust.GetCustomerPK & "&amp;CustPK=" & Cust.GetCustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "';"" />")
        								End If
        								If IsCustomerLocked AndAlso Logix.UserRoles.ForceUnlockCustomer Then
        										Send("<input type=""button"" class=""regular"" id=""unlock"" name=""unlock"" value=""" & Copient.PhraseLib.Lookup("term.unlock", LanguageID) & """ onclick=""window.location.href='customer-general.aspx?unlock=Unlock&amp;LockID=" & LockID.ToString() & "&amp;CustPK=" & Cust.GetCustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "';"" />")
        								End If
        								If (IsAllowedToEdit AndAlso MyCommon.Fetch_SystemOption(121)) Then
        										Send("<input type=""button"" class=""regular"" id=""DeleteCustomerPoints"" name=""DeleteCustomerPoints"" value=""" & Copient.PhraseLib.Lookup("customer-general.DeleteCustomerAndPoints", LanguageID) & """ style=""width:200px"" onclick=""removeCardIdPoints(" & Cust.GetCustomerPK & "," & CardPK & ");"" />")	'removeCardIdPoints
        								End If
        						Else
        								'EDITING
        								Dim TempValue As String = ""
        								Dim PhoneAsEntered As String = ""
        								Dim MobilePhoneAsEntered As String = ""
        								Dim DOBParts() As String = {"", "", ""}
        								'**
        								Dim DateOpenedParts() As String = {"", "", ""}
        								Dim EnrollmentDateParts() As String = {"", "", ""}
        								ExtCardID = MyCommon.Extract_Val(GetCgiValue("ExtCardID"))
        								'CARDHOLDER IDENTITY INPUTS, COLUMN 1
        								Send("<table style=""width:355px;float:left;position:relative;"" summary=""" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & " 1"">")
        								'Name prefix
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Prefix, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Prefix"">" & Copient.PhraseLib.Lookup("term.prefix", LanguageID) & GetDesignatorText("Prefix", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Prefix = IIf(SaveFailed, GetCgiValue("Prefix"), MyCommon.NZ(Cust.GetPrefix, UnknownPhrase).Replace("""", "&quot;"))
        								Send("  <td><input type=""text"" class=""medium"" id=""Prefix"" name=""Prefix"" maxlength=""20"" value=""" & Prefix & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'First name
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_FirstName, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""FirstName"">" & Copient.PhraseLib.Lookup("term.firstname", LanguageID) & GetDesignatorText("FirstName", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								FirstName = IIf(SaveFailed, GetCgiValue("FirstName"), MyCommon.NZ(Cust.GetFirstName, UnknownPhrase).Replace("""", "&quot;"))
        								Send("  <td><input type=""text"" class=""medium"" id=""FirstName"" name=""FirstName"" maxlength=""50"" value=""" & FirstName & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Middle name
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_MiddleName, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""MiddleName"">" & Copient.PhraseLib.Lookup("term.middlename", LanguageID) & GetDesignatorText("MiddleName", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								MiddleName = IIf(SaveFailed, GetCgiValue("MiddleName"), MyCommon.NZ(Cust.GetMiddleName, UnknownPhrase).Replace("""", "&quot;"))
        								Send("  <td><input type=""text"" class=""medium"" id=""MiddleName"" name=""MiddleName"" maxlength=""50"" value=""" & MiddleName & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Last name
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_LastName, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""LastName"">" & Copient.PhraseLib.Lookup("term.lastname", LanguageID) & GetDesignatorText("LastName", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								LastName = IIf(SaveFailed, GetCgiValue("LastName"), MyCommon.NZ(Cust.GetLastName, UnknownPhrase).Replace("""", "&quot;"))
        								Send("  <td><input type=""text"" class=""medium"" id=""LastName"" name=""LastName"" maxlength=""50"" value=""" & LastName & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Name suffix
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Suffix, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Suffix"">" & Copient.PhraseLib.Lookup("term.suffix", LanguageID) & GetDesignatorText("Suffix", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Suffix = IIf(SaveFailed, GetCgiValue("Suffix"), MyCommon.NZ(Cust.GetSuffix, UnknownPhrase).Replace("""", "&quot;"))
        								Send("  <td><input type=""text"" class=""medium"" id=""Suffix"" name=""Suffix"" maxlength=""20"" value=""" & Suffix & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'AltID
        								If (AltIDCol.ToUpper = "ALTID") Then
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_AltID, "", " style=""display:none;""") & ">")
        										Dim sAltIdPhraseName As String = "term.alternateid"
        										If MyCommon.IsEngineInstalled(0) Then
        												' if CM is installed may need to use different phrase
        												Dim sCmAltIdPhraseName As String = MyCommon.Fetch_CM_SystemOption(67)
        												If sCmAltIdPhraseName <> "" Then
        														sAltIdPhraseName = sCmAltIdPhraseName
        												End If
        										End If
        										Send("  <td class=""medium""><label for=""AltIDValue"">" & Copient.PhraseLib.Lookup(sAltIdPhraseName, LanguageID) & ":</label> </td>")
        										AltIDValue = IIf(SaveFailed, AltIDValue, MyCommon.NZ(Cust.GetAltID, "").Replace("""", "&quot;"))
        										Send("  <td><input type=""text"" class=""medium"" id=""AltIDValue"" name=""AltIDValue"" maxlength=""20"" value=""" & AltIDValue & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        										Send("</tr>")
        								End If
        								'AltID verifier
        								If (AltIDVerCol.ToUpper = "VERIFIER") Then
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_AltID, "", " style=""display:none;""") & ">")
        										Dim sVerifierPhraseName As String = "term.alternate-id-verifier"
        										If MyCommon.IsEngineInstalled(0) Then
        												' if CM is installed may need to use different phrase
        												Dim sCmVerifierPhraseName As String = MyCommon.Fetch_CM_SystemOption(68)
        												If sCmVerifierPhraseName <> "" Then
        														sVerifierPhraseName = sCmVerifierPhraseName
        												End If
        										End If
        										Send("  <td class=""medium""><label for=""AltIDVerifier"">" & Copient.PhraseLib.Lookup(sVerifierPhraseName, LanguageID) & ":</label> </td>")
        										AltIDVerifier = IIf(SaveFailed, AltIDVerifier, MyCommon.NZ(Cust.GetAltIdVerifier, "").Replace("""", "&quot;"))
        										Send("  <td><input type=""text"" class=""medium"" id=""AltIDVerifier"" name=""AltIDVerifier"" maxlength=""20"" value=""" & AltIDVerifier & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        										Send("</tr>")
        								End If
        								'Employee-related
        								If Not SaveFailed Then
        										Employee = IIf(MyCommon.NZ(Cust.GetEmployee, False), 1, 0)
        								End If
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Employee, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Employee"">" & Copient.PhraseLib.Lookup("term.employee", LanguageID) & ":</label> </td>")
        								Send("  <td><input type=""checkbox"" id=""Employee"" name=""Employee"" onclick=""javascript:ChangeEmployeeID();""" & IIf(Employee = 1, " checked=""checked""", "") & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								EmployeeID = IIf(SaveFailed, GetCgiValue("EmployeeID"), MyCommon.NZ(Cust.GetEmployeeID, ""))
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_EmployeeID, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""EmployeeID"" id=""forEmployeeID"" style=""display:" & IIf(Employee = 1, "block", "none") & ";"">" & Copient.PhraseLib.Lookup("term.employeeid", LanguageID) & ":</label></td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""EmployeeID"" name=""EmployeeID"" maxlength=""26"" style=""display:" & IIf(Employee = 1, "block", "none") & """ value=""" & EmployeeID & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", "disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Test card
        								If MyCommon.Fetch_SystemOption(88) Then
        										If Not SaveFailed Then
        												TestCard = IIf(MyCommon.NZ(Cust.GetTestCard, False), 1, 0)
        										End If
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Test, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""TestCard"">" & Copient.PhraseLib.Lookup("term.testcustomer", LanguageID) & ":</label> </td>")
        										Send("  <td><input type=""checkbox"" id=""TestCard"" name=""TestCard""" & IIf(TestCard = 1, " checked=""checked""", "") & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        										Send("</tr>")
        								End If
        								'Banner
        								If MyCommon.Fetch_SystemOption(66) Then
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Banner, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banner", LanguageID) & ":</label> </td>")
        										Send("  <td><select class=""medium"" id=""BannerID"" name=""BannerID""" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & ">")
        										Send("    <option value=""0"">** " & Copient.PhraseLib.Lookup("term.noneselected", LanguageID) & " **</option>")
        										MyCommon.QueryStr = "select BannerID, Name, Description from Banners with (NoLock) where Deleted=0;"
        										rst2 = MyCommon.LRT_Select
        										For Each row2 In rst2.Rows
        												BannerID = IIf(SaveFailed, GetCgiValue("BannerID"), MyCommon.NZ(Cust.GetBannerID, 0))
        												If (BannerID = MyCommon.NZ(row2.Item("BannerID"), 0)) Then
        														Send("    <option value=""" & MyCommon.NZ(row2.Item("BannerID"), 0) & """ selected=""selected"">" & MyCommon.TruncateString(MyCommon.NZ(row2.Item("Name"), "&nbsp;"), 25) & "</option>")
        												Else
        														Send("    <option value=""" & MyCommon.NZ(row2.Item("BannerID"), 0) & """>" & MyCommon.TruncateString(MyCommon.NZ(row2.Item("Name"), "&nbsp;"), 25) & "</option>")
        												End If
        										Next
        										Send("  </select></td>")
        										Send("</tr>")
        								End If
        								'Card status
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Status, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""CustomerStatusID"">" & Copient.PhraseLib.Lookup("customer.customerstatus", LanguageID) & ":</label> </td>")
        								Send("  <td><select class=""medium"" id=""CustomerStatusID"" name=""CustomerStatusID""" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & ">")
        								MyCommon.QueryStr = "select CustomerStatusID, PhraseID from CustomerStatus with (NoLock);"
        								rst2 = MyCommon.LXS_Select
        								For Each row2 In rst2.Rows
        										CustomerStatusID = IIf(SaveFailed, GetCgiValue("CustomerStatusID"), MyCommon.NZ(Cust.GetCustomerStatusID, 0))
        										If (CustomerStatusID = MyCommon.NZ(row2.Item("CustomerStatusID"), 0)) Then
        												Send("<option value=""" & MyCommon.NZ(row2.Item("CustomerStatusID"), 0) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</option>")
        										Else
        												Send("<option value=""" & MyCommon.NZ(row2.Item("CustomerStatusID"), 0) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
        										End If
        								Next
        								Send("  </select></td>")
        								Send("</tr>")
        								'Household
        								Send("<tr style=""display:none;"">")
        								Send("  <td class=""medium""><label for=""Household"">" & Copient.PhraseLib.Lookup("term.household", LanguageID) & ":</label> </td>")
        								Send("  <td><select class=""medium"" id=""Household"" name=""Household""" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & ">")
        								If SaveFailed Then
        										Send("      <option value=""0"" " & IIf(Household = "0", "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.no", LanguageID) & "</option>")
        										Send("      <option value=""1"" " & IIf(Household = "1", "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.yes", LanguageID) & "</option>")
        								Else
        										Send("      <option value=""0"" " & IIf(IsHouseholdID, "", "selected=""selected""") & ">" & Copient.PhraseLib.Lookup("term.no", LanguageID) & "</option>")
        										Send("      <option value=""1"" " & IIf(IsHouseholdID, "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.yes", LanguageID) & "</option>")
        								End If
        								Send("      </select></td>")
        								Send("</tr>")
        								'Restricted Redemption
        								If MyCommon.Fetch_CPE_SystemOption(129) Then
        										Dim IsRestrictedRedemption As Boolean = False
        										If Not SaveFailed Then
        												IsRestrictedRedemption = MyCommon.NZ(Cust.GetRestrictedRedemption, False)
        										End If
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_RestrictedRdmpt, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""RestrictedRdmpt"">" & Copient.PhraseLib.Lookup("term.restrictedredemption", LanguageID) & ":</label> </td>")
        										Send("  <td><input type=""checkbox"" id=""RestrictedRdmpt"" name=""RestrictedRdmpt""" & IIf(IsRestrictedRedemption, " checked=""checked""", "") & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        										Send("</tr>")
        								End If

        								'**
        								'Comments
        								TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Comments AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetComments, ""), GetCgiValue("Comments"))
        								'TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Comments AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetComments, ""), IIf(Logix.UserRoles.AccessCustomerIdData_Comments, GetCgiValue("Comments"), MyCommon.NZ(Cust.GetGeneralInfo.GetComments, "")))
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Comments, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Comments"">" & Copient.PhraseLib.Lookup("term.comments", LanguageID) & GetDesignatorText("Comments", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""Comments"" name=""Comments"" maxlength=""50"" value=""" & TempValue & """ /></td>")
        								Send("</tr>")

        								'DriverLicenseID
                        TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_DriverLicenseID AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetDriverLicenseID, ""), GetCgiValue("DriverLicenseID"))
        								'TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_DriverLicenseID AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetDriverLicenseID, ""), IIf(Logix.UserRoles.AccessCustomerIdData_DriverLicenseID,GetCgiValue("DriverLicenseID"),MyCommon.NZ(Cust.GetGeneralInfo.GetDriverLicenseID, "")))
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_DriverLicenseID, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""DriverLicenseID"">" & Copient.PhraseLib.Lookup("term.driverslicense", LanguageID) & GetDesignatorText("DriverLicenseID", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""DriverLicenseID"" name=""DriverLicenseID"" maxlength=""50"" value=""" & TempValue & """ /></td>")
        								Send("</tr>")

        								'TaxExemptID
                        TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_TaxExemptID AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetTaxExemptID, ""), GetCgiValue("TaxExemptID"))
        								'TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_TaxExemptID AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetTaxExemptID, ""), IIf(Logix.UserRoles.AccessCustomerIdData_TaxExemptID, GetCgiValue("TaxExemptID"),MyCommon.NZ(Cust.GetGeneralInfo.GetTaxExemptID, "")))
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_TaxExemptID, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""TaxExemptID"">" & Copient.PhraseLib.Lookup("term.taxexemptid", LanguageID) & GetDesignatorText("TaxExemptID", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""TaxExemptID"" name=""TaxExemptID"" maxlength=""100"" value=""" & TempValue & """ /></td>")
        								Send("</tr>")
						
        								'Digital Receipt	
        								If (MyCommon.Fetch_SystemOption(212) <> 0) Then
        										If (Logix.UserRoles.AccessCustomerIdData_ReceiptPref) Then
													Dim disabledflag as string = ""
													If (MyCommon.Fetch_SystemOption(212) = 2) Then
														disabledflag = " disabled=""disabled"" "
													End If

        												MyCommon.QueryStr = "select Description from ThirdPartyReceiptPref"
        												rst = MyCommon.LXS_Select
        												InputBitSize = rst.Rows.Count + 1
						
        												Dim checked(InputBitSize) As Integer
        												MyCommon.QueryStr = "select DigitalReceipt from CustomerExt where CustomerPK= " & CustomerPK & ""
													rst = MyCommon.LXS_Select
													If (rst IsNot Nothing AndAlso rst.Rows.Count > 0) Then	 													
													    DigitalReceipt = MyCommon.NZ(rst.Rows(0).Item("DigitalReceipt"), 0)
													End If	
        												If (DigitalReceipt <> 0) Then
        														DigitalReceipt = (DigitalReceipt << 1 Or 1)
        												End If
							
        												For index As Integer = 0 To InputBitSize - 1
								
        														checked(index) = (DigitalReceipt >> index) And 1
        												Next
        												Send("<tr>")
        												Send("  <td class=""medium""><label for=""Digital Receipt"">" & Copient.PhraseLib.Lookup("term.dreceiptpref", LanguageID) & GetDesignatorText("Digital Receipt", AltIDCol, AltIDVerCol) & "</label> </td>")
        												MyCommon.QueryStr = "select Description from ThirdPartyReceiptPref"
        												rst = MyCommon.LXS_Select
							
        												For num As Integer = 0 To InputBitSize - 1
        														If (num > 0) Then
        																' one of the 3rd party checkboxes
        																Dim text As String = MyCommon.NZ(rst.Rows(num - 1).Item("Description"), 0)
        																If (checked(num) <> 1) Then
        																		' NOT checked
        																		Send("<DD><input type=""checkbox"" id=""c" & num & """ name=""c" & num & disabledflag & """ onclick=""bitToggle( " & num & ");""" & disabledflag & ">" & text & "</DD>")
        																Else
        																		' CHECKED
        																		Send("<DD><input type=""checkbox"" id=""c" & num & """ name=""c" & num & disabledflag & """ onclick=""bitToggle(" & num & ");"" checked=""checked""" & disabledflag & ">" & text & "</DD>")
        																End If
								
        														ElseIf (checked(num) <> 1) Then
        																' main checkbox (not one of the 3rd parties), and it's NOT checked
        																Send("<td><DT><input type=""checkbox"" id=""c" & num & """  name=""c" & num & """ onclick=""showMe('div1');bitToggle(0);""" & disabledflag & " ></DT>")
        																Send("<div id=""div1"" style=""display:none""><p>")
        														Else
        																' main checkbox (not one of the 3rd parties), and it IS checked
        																Send("<td><DT><input type=""checkbox"" id=""c" & num & """ name=""c" & num & """ onclick=""showMe('div1');bitToggle(0);"" checked=""checked""" & disabledflag & "></DT>")
        																Send("<div id=""div1"" ><p>")
        														End If
        												Next
        												Send("</td>")
        												Send("</div>")
        												Send("</tr>")
						
							
        										End If
							
        								End If
							
        								'Paper Receipt	
        								If (MyCommon.Fetch_SystemOption(310) <> 0) Then
        										If (Logix.UserRoles.AccessCustomerIdData_ReceiptPref) Then
													Dim disabledflag as string = ""
													If (MyCommon.Fetch_SystemOption(310) = 2) Then
														disabledflag = " disabled=""disabled"""
													End If
						
        												MyCommon.QueryStr = "select PaperReceipt from CustomerExt where CustomerPK= " & CustomerPK & ""
        												rst = MyCommon.LXS_Select
													If (rst IsNot Nothing AndAlso rst.Rows.Count > 0) Then
														PaperReceipt = (MyCommon.NZ(rst.Rows(0).Item("PaperReceipt"), False))
													End If
							
        												Send("<tr>")
        												Send("  <td class=""medium""><label for=""PaperReceipt"">" & Copient.PhraseLib.Lookup("term.preceiptpref", LanguageID) & GetDesignatorText("Paper Receipt", AltIDCol, AltIDVerCol) & "</label> </td>")

        												Send("<td><DT><input type=""checkbox"" name=""preceiptpref"" onclick=""paperCheck();"" id=""preceiptpref""" & IIf(PaperReceipt, " checked=""checked""", "") & disabledflag & "></DT>")

        												Send("</tr>")
        										End If
							
        								End If
        								'only show these fields if accounts receivable system option is enabled
        								If (MyCommon.Fetch_SystemOption(113)) Then
        										'**ARCustomer
        										Dim IsARCustomer As Boolean = False
        										If Not SaveFailed Then
        												IsARCustomer = MyCommon.NZ(Cust.GetGeneralInfo.GetARCustomer, False)
        										End If
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_ARCustomer, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""ARCustomer"">" & Copient.PhraseLib.Lookup("term.araccount", LanguageID) & ":</label> </td>")
        										Send("  <td><input type=""checkbox"" id=""ARCustomer"" name=""ARCustomer""" & IIf(IsARCustomer, " checked=""checked""", "") & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        										Send("</tr>")

        										'style=""display:" & IIf(Employee = 1, "block", "none") & """
        										'id=""forEmployeeID"" style=""display:" & IIf(Employee = 1, "block", "none") & ";"">" & Copient.Phrase
        										'FinanceCharge
        										Dim IsFinanceCharge As Boolean = False
        										If Not SaveFailed Then
        												IsFinanceCharge = MyCommon.NZ(Cust.GetAccountReceivableInfo.GetFinanceCharge, False)
        										End If
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_FinanceCharge, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""FinanceCharge"">" & Copient.PhraseLib.Lookup("term.financecharge", LanguageID) & ":</label> </td>")
        										Send("  <td><input type=""checkbox"" id=""FinanceCharge"" name=""FinanceCharge""" & IIf(IsFinanceCharge, " checked=""checked""", "") & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        										Send("</tr>")

        										'CompoundCharge
        										Dim IsCompoundCharge As Boolean = False
        										If Not SaveFailed Then
        												IsCompoundCharge = MyCommon.NZ(Cust.GetAccountReceivableInfo.GetCompoundCharge, False)
        										End If
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_CompoundCharge, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""CompoundCharge"">" & Copient.PhraseLib.Lookup("term.compoundcharge", LanguageID) & ":</label> </td>")
        										Send("  <td><input type=""checkbox"" id=""CompoundCharge"" name=""CompoundCharge""" & IIf(IsCompoundCharge, " checked=""checked""", "") & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        										Send("</tr>")
        								End If
        								Send("</table>")
              
        								'CARDHOLDER IDENTITY INPUTS, COLUMN 2
        								Send("<table style=""width:355px;float:left;position:relative;"" summary=""" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & " 2"">")
        								'Address
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Address, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Address"">" & Copient.PhraseLib.Lookup("customer.address", LanguageID) & GetDesignatorText("Address", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Address AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetAddress, ""), GetCgiValue("Address"))
        								If TempValue <> "" Then
        										TempValue = TempValue.Replace("""", "&quot;")
        								End If
        								Send("  <td><input type=""text"" class=""medium"" id=""Address"" name=""Address"" maxlength=""200"" value=""" & TempValue & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								If (MyCommon.Fetch_SystemOption(162)) Then
        										Send("<td><input type=""button"" class=""regular"" id=""verifyAddress"" name=""verifyAddress"" value=""" & Copient.PhraseLib.Lookup("customer-general.VerfyAddress", LanguageID) & """ style=""width:120px"" onclick=""javascript:popupPro();"" /></td>")
        								End If
        								Send("</tr>")
        								'City
        								TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_City AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetCity, ""), GetCgiValue("City"))
        								If TempValue <> "" Then
        										TempValue = TempValue.Replace("""", "&quot;")
        								End If
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_City, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""City"">" & Copient.PhraseLib.Lookup("customer.city", LanguageID) & GetDesignatorText("City", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""City"" name=""City"" maxlength=""100"" value=""" & TempValue & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'State
        								TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_State AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetState, ""), GetCgiValue("State"))
        								If TempValue <> "" Then
        										TempValue = TempValue.Replace("""", "&quot;")
        								End If
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_State, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""State"">" & Copient.PhraseLib.Lookup("customer.state", LanguageID) & GetDesignatorText("State", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""State"" name=""State"" maxlength=""50"" value=""" & TempValue & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Postal code
        								TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_ZIP AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetZip, ""), GetCgiValue("Zip"))
        								If TempValue <> "" Then
        										TempValue = TempValue.Replace("""", "&quot;")
        								End If
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_ZIP, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Zip"">" & Copient.PhraseLib.Lookup("customer.zip", LanguageID) & GetDesignatorText("Zip", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""Zip"" name=""Zip"" maxlength=""20"" value=""" & TempValue & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Country
        								'TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Country AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetCountry, ""), GetCgiValue("Country"))
        								TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Country AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetCountry, ""), IIf(Logix.UserRoles.AccessCustomerIdData_Country, GetCgiValue("Country"), MyCommon.NZ(Cust.GetGeneralInfo.GetCountry, "")))
        								If TempValue <> "" Then
        										TempValue = TempValue.Replace("""", "&quot;")
        								End If
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Country, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Country"">" & Copient.PhraseLib.Lookup("term.country", LanguageID) & GetDesignatorText("Country", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td><input type=""text"" class=""medium"" id=""Country"" name=""Country"" maxlength=""50"" value=""" & TempValue & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Phone number
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Phone, "", " style=""display:none;""") & ">")
                        PhoneAsEntered = IIf(Logix.UserRoles.AccessCustomerIdData_Phone AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetPhone, ""), GetCgiValue("Phone1"))
        								Send("  <td class=""medium""><label for=""Phone1"">" & Copient.PhraseLib.Lookup("customer.phone", LanguageID) & GetDesignatorText("Phone", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td style=""font-size:18px;"" >")
        								Send("    <input type=""text"" class=""medium"" id=""Phone1"" name=""Phone1"" maxlength=""50"" value=""" & PhoneAsEntered & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        								Send("  </td>")
        								Send("</tr>")
        								'Mobile phone number
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_MobilePhone, "", " style=""display:none;""") & ">")
                        MobilePhoneAsEntered = IIf(Logix.UserRoles.AccessCustomerIdData_MobilePhone AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetMobilePhone, ""), GetCgiValue("MobilePhone1"))
        								Send("  <td class=""medium""><label for=""MobilePhone1"">" & Copient.PhraseLib.Lookup("customer.mobilephone", LanguageID) & GetDesignatorText("MobilePhone", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								Send("  <td style=""font-size:18px;"" >")
        								Send("    <input type=""text"" class=""medium"" id=""MobilePhone1"" name=""MobilePhone1"" maxlength=""50"" value=""" & MobilePhoneAsEntered & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        								Send("  </td>")
        								Send("</tr>")
        								'Email
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Email, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Email"">" & Copient.PhraseLib.Lookup("customer.email", LanguageID) & GetDesignatorText("email", AltIDCol, AltIDVerCol) & ":</label> </td>")
                        TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Email AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetEmail, ""), GetCgiValue("Email"))
        								'TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Email AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetEmail, ""), IIf(Logix.UserRoles.AccessCustomerIdData_Email, GetCgiValue("Email"), MyCommon.NZ(Cust.GetGeneralInfo.GetEmail, "")))
        								Send("  <td><input type=""text"" class=""medium"" id=""Email"" name=""Email"" maxlength=""200"" value=""" & HttpUtility.HtmlEncode(TempValue) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'Date of birth
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_DOB, "", " style=""display:none;""") & ">")
        								TempValue = ""
        								If Logix.UserRoles.AccessCustomerIdData_DOB AndAlso Not SaveFailed Then
        										If Cust.GetGeneralInfo.GetDateOfBirth <> Nothing Then
        												TempValue = MyCommon.NZ(Cust.GetGeneralInfo.GetDateOfBirth.ToString("MMddyyyy"), "")
        										End If
        										'ElseIf Logix.UserRoles.AccessCustomerIdData_DOB = False Then
        										'  If Cust.GetGeneralInfo.GetDateOfBirth <> Nothing Then
        										'    TempValue = MyCommon.NZ(Cust.GetGeneralInfo.GetDateOfBirth.ToString("MMddyyyy"), "")
        										'  End If			    
        								Else
        										TempValue = GetCgiValue("dob1") & GetCgiValue("dob2") & GetCgiValue("dob3")
        								End If
        								DOBParts = ParseDateOfBirth(TempValue)
        								Send("  <td><label for=""dob1"">" & Copient.PhraseLib.Lookup("term.dateofbirth", LanguageID) & GetDesignatorText("DOB", AltIDCol, AltIDVerCol) & "<br />(" & GetShortDatePattern(MyCommon) & "):</label> </td>")
        								Send("  <td style=""font-size:18px;"" >")
        								' ensure that the order of the text boxes are in the localized short date pattern for the user's language
        								For i = 0 To DatePartOrder.Count - 1
        										Select Case DatePartOrder(i)
        												Case Copient.LogixInc.DATE_PART.MONTH
        														Sendb("    <input type=""text"" style=""width:39px;"" id=""dob1"" name=""dob1"" maxlength=""2"" value=""" & DOBParts(0) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        												Case Copient.LogixInc.DATE_PART.DAY
        														Sendb("<input type=""text"" style=""width:40px;"" id=""dob2"" name=""dob2"" maxlength=""2"" value=""" & DOBParts(1) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        												Case Copient.LogixInc.DATE_PART.YEAR
        														Sendb("<input type=""text"" class=""shorter"" id=""dob3"" name=""dob3"" maxlength=""4"" value=""" & DOBParts(2) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        										End Select
        										If i < 2 Then Sendb("/")
        								Next
        								Send("")
        								Send("  </td>")
        								Send("</tr>")
        								'Password
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Password, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""Password"">" & Copient.PhraseLib.Lookup("term.password", LanguageID) & ":</label> </td>")
        								Dim tmpPass As String
        								tmpPass = IIf(SaveFailed, GetCgiValue("Password"), MyCommon.NZ(Cust.GetPassword, ""))
        								Send("  <td><input type=""text"" class=""medium"" id=""Password"" name=""Password"" maxlength=""" & MAX_CUST_PASSWORD_CLEARTEXT_LEN & """ value=""" & tmpPass & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Send("</tr>")
        								'AirMile member ID
        								TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_AirmileMemberID AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetAirmileMemberID, ""), GetCgiValue("AirmileMemberID"))
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_AirmileMemberID AndAlso HasAirmileMemberID, "", " style=""display:none;""") & ">")
        								Send("  <td class=""medium""><label for=""AirmileMemberID"">" & Copient.PhraseLib.Lookup("term.airmilememberid", LanguageID) & GetDesignatorText("AirmileMemberID", AltIDCol, AltIDVerCol) & ":</label> </td>")
        								If (MyCommon.NZ(MyCommon.Fetch_CPE_SystemOption(153), 1)) Then
        										Send("  <td><input type=""text"" class=""medium"" id=""AirmileMemberID"" name=""AirmileMemberID"" maxlength=""50"" value=""" & TempValue & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " /></td>")
        								Else
        										Send("  <td><input type=""text"" class=""medium"" id=""AirmileMemberID"" name=""AirmileMemberID"" maxlength=""50"" value=""" & TempValue & """" & "readonly" & " /></td>")
        								End If
        								Send("</tr>")
        								'Enrollment date
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_EnrollmentDate, "", " style=""display:none;""") & ">")
        								TempValue = ""
        								If Logix.UserRoles.AccessCustomerIdData_EnrollmentDate AndAlso Not SaveFailed Then
        										If Cust.GetEnrollmentDate <> Nothing Then
        												TempValue = MyCommon.NZ(Cust.GetEnrollmentDate.ToString("MMddyyyy"), "")
        										End If
        								Else
        										TempValue = GetCgiValue("ed1") & GetCgiValue("ed2") & GetCgiValue("ed3")
        								End If
        								EnrollmentDateParts = ParseEnrollmentDate(TempValue)
        								Send("  <td><label for=""ed1"">" & Copient.PhraseLib.Lookup("customer.EnrollmentDate", LanguageID) & GetDesignatorText("EnrollmentDate", AltIDCol, AltIDVerCol) & "<br />(" & GetShortDatePattern(MyCommon) & "):</label> </td>")
        								Send("  <td style=""font-size:18px;"">")
        								' ensure that the order of the text boxes are in the localized short date pattern for the user's language
        								For i = 0 To DatePartOrder.Count - 1
        										Select Case DatePartOrder(i)
        												Case Copient.LogixInc.DATE_PART.MONTH
        														Sendb("    <input type=""text"" style=""width:39px;"" id=""ed1"" name=""ed1"" maxlength=""2"" value=""" & EnrollmentDateParts(0) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        												Case Copient.LogixInc.DATE_PART.DAY
        														Sendb("<input type=""text"" style=""width:40px;"" id=""ed2"" name=""ed2"" maxlength=""2"" value=""" & EnrollmentDateParts(1) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        												Case Copient.LogixInc.DATE_PART.YEAR
        														Send("<input type=""text"" class=""shorter"" id=""ed3"" name=""ed3"" maxlength=""4"" value=""" & EnrollmentDateParts(2) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & " />")
        										End Select
        										If i < 2 Then Sendb("/")
        								Next
        								Send("")
        								Send("  </td>")
        								Send("</tr>")
        								'**
        								'Date opened
        								Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_DateOpened, "", " style=""display:none;""") & ">")
        								TempValue = ""
        								If Logix.UserRoles.AccessCustomerIdData_DateOpened AndAlso Not SaveFailed Then
        										If Cust.GetGeneralInfo.GetDateOpened <> Nothing Then
        												TempValue = MyCommon.NZ(Cust.GetGeneralInfo.GetDateOpened.ToString("MMddyyyy"), "")
        										End If
        										'ElseIf Logix.UserRoles.AccessCustomerIdData_DateOpened = False Then
        										'  If Cust.GetGeneralInfo.GetDateOpened <> Nothing Then
        										'    TempValue = MyCommon.NZ(Cust.GetGeneralInfo.GetDateOpened.ToString("MMddyyyy"), "")
        										'  End If			  
        								Else
        										TempValue = GetCgiValue("dateopened1") & GetCgiValue("dateopened2") & GetCgiValue("dateopened3")
        								End If
        								DateOpenedParts = ParseDateOpened(TempValue)
        								Send("  <td><label for=""dateopened1"">" & Copient.PhraseLib.Lookup("term.dateopened", LanguageID) & GetDesignatorText("DateOpened", AltIDCol, AltIDVerCol) & "<br />(" & GetShortDatePattern(MyCommon) & "):</label> </td>")
        								Send("  <td style=""font-size:18px;"">")
        								' ensure that the order of the text boxes are in the localized short date pattern for the user's language
        								For i = 0 To DatePartOrder.Count - 1
        										Select Case DatePartOrder(i)
        												Case Copient.LogixInc.DATE_PART.MONTH
        														Sendb("    <input type=""text"" style=""width:39px;"" id=""dateopened1"" name=""dateopened1"" maxlength=""2"" value=""" & DateOpenedParts(0) & """ />")
        												Case Copient.LogixInc.DATE_PART.DAY
        														Sendb("<input type=""text"" style=""width:40px;"" id=""dateopened2"" name=""dateopened2"" maxlength=""2"" value=""" & DateOpenedParts(1) & """ />")
        												Case Copient.LogixInc.DATE_PART.YEAR
        														Sendb("<input type=""text"" class=""shorter"" id=""dateopened3"" name=""dateopened3"" maxlength=""4"" value=""" & DateOpenedParts(2) & """ />")
        										End Select
        										If i < 2 Then Sendb("/")
        								Next
        								Send("")
        								Send("  </td>")
        								Send("</tr>")
        								'only show these fields if accounts receivable system option is enabled
        								If (MyCommon.Fetch_SystemOption(113)) Then
        										'CreditLimit
        										TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_CreditLimit AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetAccountReceivableInfo.GetCreditLimit, ""), GetCgiValue("CreditLimit"))
        										If TempValue <> "" Then
        												TempValue = TempValue.Replace("""", "&quot;")
        										End If
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_CreditLimit, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""CreditLimit"" id=""forCreditLimit"">" & Copient.PhraseLib.Lookup("term.creditlimit", LanguageID) & GetDesignatorText("CreditLimit", AltIDCol, AltIDVerCol) & ":</label> </td>")
        										Send("  <td><input type=""text"" class=""medium"" id=""CreditLimit"" name=""CreditLimit"" maxlength=""50"" value=""" & TempValue & """ /></td>")
        										Send("</tr>")

        										'APR
        										TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_APR AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetAccountReceivableInfo.GetAPR, ""), GetCgiValue("APR"))
        										If TempValue <> "" Then
        												TempValue = TempValue.Replace("""", "&quot;")
        										End If
        										Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_APR, "", " style=""display:none;""") & ">")
        										Send("  <td class=""medium""><label for=""APR"" id=""forAPR"">" & Copient.PhraseLib.Lookup("term.apr", LanguageID) & GetDesignatorText("APR", AltIDCol, AltIDVerCol) & ":</label> </td>")
        										Send("  <td><input type=""text"" class=""medium"" id=""APR"" name=""APR"" maxlength=""50"" value=""" & TempValue & """ /></td>")
        										Send("</tr>")
        								End If
        								Send("</table>")
        								Send("<br clear=""left"" />")
        								'Supplemental fields
        								If MyCommon.Fetch_SystemOption(110) Then
        										Dim OrderString As String = "order by CSF.Name;"
        										Dim sortOption As Integer = 0
        										' Default to ordering by the name of the suppl field, which was the original behavior
        										' the new option is 0 to sort by name.  If it's 1 then sort by the External ID (assuming alphanumeric sorting).  If it's 2 then
        										' sort by the external ID cast to a integer value. Thus, if the option is 1, then ordering would be 1->10->11-> ... 2->20->, but if
        										' the option is 2 then it would be 1->2->3->...->10->11, etc.
        										sortOption = MyCommon.Fetch_SystemOption(265)
        										If (sortOption = 1) Then
        										    OrderString = "order by CSF.ExtFieldID;"
        										ElseIf (sortOption = 2) Then
        										    OrderString = "order by CAST(CSF.ExtFieldID as bigint);"
        										End If
        										MyCommon.QueryStr = "select CSF.FieldID, CSF.ExtFieldID, CSF.Name, CSF.FieldTypeID, CSFT.Name As FieldTypeName, Length, Editable, Value " & _
        																				"from CustomerSupplementalFields as CSF with (NoLock) " & _
        																				"left join CustomerSupplementalFieldTypes as CSFT on CSFT.FieldTypeID=CSF.FieldTypeID " & _
        																				"left join CustomerSupplemental as CS on CS.FieldID=CSF.FieldID and CustomerPK=" & CustomerPK & " and CS.Deleted=0 " & _
        																				"where CSF.Deleted=0 " & IIf(Logix.UserRoles.AccessProtectedSupplementalFields, "", "and Visible=1 ") & _
        																				OrderString
        										rst2 = MyCommon.LXS_Select
        										If rst2.Rows.Count > 0 Then
        												Dim Table1Count As Integer = Math.Ceiling(rst2.Rows.Count / 2)
        												Dim FieldID As Integer = 0
        												Dim Name As String = ""
        												Dim FieldTypeID As Integer = 0
        												Dim FieldTypeName As String = ""
        												Dim Value As String = ""
        												Dim Length As Integer = 0
        												Dim Editable As Boolean = False
        												For i = 1 To rst2.Rows.Count
        														FieldID = rst2.Rows(i - 1).Item("FieldID")
        														Name = MyCommon.NZ(rst2.Rows(i - 1).Item("Name"), Copient.PhraseLib.Lookup("term.unnamed", LanguageID))
        														FieldTypeID = MyCommon.NZ(rst2.Rows(i - 1).Item("FieldTypeID"), 0)
        														FieldTypeName = MyCommon.NZ(rst2.Rows(i - 1).Item("FieldTypeName"), "")
        														Value = MyCommon.NZ(rst2.Rows(i - 1).Item("Value"), "")
        														Length = MyCommon.NZ(rst2.Rows(i - 1).Item("Length"), 0)
        														Editable = IIf(rst2.Rows(i - 1).Item("Editable"), True, False)
        														If (i = 1) Or (i = Table1Count + 1) Then
        																Send("<table style=""width:355px;float:left;position:relative;"" summary=""" & Copient.PhraseLib.Lookup("term.customersupplementalfields", LanguageID) & """>")
        														End If
        														Send("  <tr>")
        														Send("    <td class=""medium"" style=""float:left;overflow:hidden;""><label for=""CS" & FieldID & """>" & Name & "</label>:</td>")
        														Sendb("    <td>")
        														If FieldTypeName = "Bit" Then
        																If (Editable OrElse Logix.UserRoles.AccessProtectedSupplementalFields) Then
        																		Send("<input type=""checkbox"" id=""CS" & FieldID & """ name=""CS" & FieldID & """" & IIf(Value = "1", " checked=""checked""", "") & " value=""1"" />")
        																Else
        																		Send("<input type=""hidden"" id=""CS" & FieldID & """ name=""CS" & FieldID & """ value=""" & Value & """ />")
        																		Send("<input type=""checkbox"" disabled=""disabled""" & IIf(Value = "1", " checked=""checked""", "") & " />")
        																End If
        														ElseIf FieldTypeName = "Date" Then
        																Dim ValueDate As Date
        																Dim ValueMonth As String = ""
        																Dim ValueDay As String = ""
        																Dim ValueYear As String = ""
        																If Value <> "" Then
        																		If DateTime.TryParse(Value, ValueDate) Then
        																				ValueMonth = ValueDate.Month.ToString
        																				ValueDay = ValueDate.Day.ToString
        																				ValueYear = ValueDate.Year.ToString
        																		End If
        																End If
        																If (Editable OrElse Logix.UserRoles.AccessProtectedSupplementalFields) Then
        																		'Sendb("<input type=""text"" id=""CS" & FieldID & """ name=""CS" & FieldID & """ maxlength=""23"" value=""" & Value & """ />")
        																		Sendb("<input type=""text"" id=""CS" & FieldID & "-1"" name=""CS" & FieldID & "-1"" maxlength=""2"" style=""width:39px;"" value=""" & ValueMonth & """ />/")
        																		Sendb("<input type=""text"" id=""CS" & FieldID & "-2"" name=""CS" & FieldID & "-2"" maxlength=""2"" style=""width:39px;"" value=""" & ValueDay & """ />/")
        																		Send("<input type=""text"" id=""CS" & FieldID & "-3"" name=""CS" & FieldID & "-3"" maxlength=""4"" style=""width:50px;"" value=""" & ValueYear & """ />")
        																Else
        																		'Send("<input type=""hidden"" id=""CS" & FieldID & """ name=""CS" & FieldID & """ value=""" & Value & """ />")
        																		'Send("<input type=""text"" disabled=""disabled"" value=""" & Value & """ />")
        																		Send("<input type=""hidden"" id=""CS" & FieldID & "-1"" name=""CS" & FieldID & "-1"" value=""" & ValueMonth & """ />")
        																		Send("<input type=""hidden"" id=""CS" & FieldID & "-2"" name=""CS" & FieldID & "-2"" value=""" & ValueDay & """ />")
        																		Send("<input type=""hidden"" id=""CS" & FieldID & "-3"" name=""CS" & FieldID & "-3"" value=""" & ValueYear & """ />")
        																		Sendb("<input type=""text"" disabled=""disabled"" value=""" & ValueMonth & """ style=""width:39px;"" />/")
        																		Sendb("<input type=""text"" disabled=""disabled"" value=""" & ValueDay & """ style=""width:39px;"" />/")
        																		Send("<input type=""text"" disabled=""disabled"" value=""" & ValueYear & """ style=""width:50px;"" />")
        																End If
        														Else
        																If (Editable OrElse Logix.UserRoles.AccessProtectedSupplementalFields) Then
        																		Sendb("<input type=""text"" id=""CS" & FieldID & """ name=""CS" & FieldID & """")
        																		If Length > 0 Then
        																				Sendb(" maxlength=""" & Length & """" & IIf(Length <= 15, " style=""width:" & Length * 11 & "px;""", " class=""medium"""))
        																		End If
        																		Sendb(" value=""" & Value & """ />")
        																Else
        																		Send("<input type=""hidden"" id=""CS" & FieldID & """ name=""CS" & FieldID & """ value=""" & Value & """ />")
        																		Send("<input type=""text"" disabled=""disabled""" & IIf(Length <= 15, " style=""width:" & Length * 11 & "px;""", " class=""medium""") & " value=""" & Value & """ />")
        																End If
        														End If
        														Send("</td>")
        														Send("  </tr>")
        														If (i = Table1Count) Or (i = rst2.Rows.Count) Then
        																Send("</table>")
        														End If
        												Next
        												Send("<br clear=""left"" />")
        										End If
        								End If
              
        								' Attributes
        								Dim CustomerAttributeValueID As Integer = 0
        								If Logix.UserRoles.AccessCustomerIdData_Attributes Then
        										MyCommon.QueryStr = "select AT.AttributeTypeID, AT.Description as AttributeType, AT.ReadOnlyAttribute " & _
        																				"from AttributeTypes as AT with (NoLock) " & _
        																				"where AT.Deleted=0 " & _
        																				"order by AttributeType;"
        										Dim dtAttributeTypes As DataTable = MyCommon.LRT_Select
        										If dtAttributeTypes.Rows.Count > 0 Then
        												Send("<br class=""half"" />")
        												Send("<table summary=""" & Copient.PhraseLib.Lookup("term.attribute", LanguageID) & """>")
        												Send("  <tr>")
        												Send("    <th style=""width:250px;"">" & Copient.PhraseLib.Lookup("term.attribute", LanguageID) & "</th>")
        												Send("    <th>" & Copient.PhraseLib.Lookup("term.value", LanguageID) & "</th>")
        												Send("  </tr>")
        												For Each row In dtAttributeTypes.Rows
        														CustomerAttributeValueID = 0
        														MyCommon.QueryStr = "select AV.AttributeValueID, AV.Description as AttributeValue, AV.DefaultValue " & _
        																								"from AttributeValues as AV with (NoLock) " & _
        																								"where AV.AttributeTypeID=" & MyCommon.NZ(row.Item("AttributeTypeID"), 0) & " and AV.Deleted=0 " & _
        																								"order by AttributeValueID;"
        														Dim dtAttributeValues As DataTable = MyCommon.LRT_Select
        														If dtAttributeValues.Rows.Count > 0 Then
        																MyCommon.QueryStr = "select AttributeValueID " & _
        																										"from CustomerAttributes with (NoLock) " & _
        																										"where AttributeTypeID=" & MyCommon.NZ(row.Item("AttributeTypeID"), 0) & " and CustomerPK=" & CustomerPK & " and Deleted=0;"
        																Dim dtCustomerAttributes As DataTable = MyCommon.LXS_Select
        																If dtCustomerAttributes.Rows.Count > 0 Then
        																		CustomerAttributeValueID = MyCommon.NZ(dtCustomerAttributes.Rows(0).Item("AttributeValueID"), 0)
        																End If
        																Send("  <tr class=""shaded"">")
        																Send("    <td>" & MyCommon.NZ(row.Item("AttributeType"), "") & "</td>")
        																Send("    <td>")
        																If MyCommon.Fetch_SystemOption(111) AndAlso Logix.UserRoles.AssignAttributes Then
        																		Send("      <select id=""at-" & MyCommon.NZ(row.Item("AttributeTypeID"), 0) & """ name=""at-" & MyCommon.NZ(row.Item("AttributeTypeID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ReadOnlyAttribute"), 0) AndAlso (MyCommon.Fetch_SystemOption(119) = 1), " disabled=""disabled""", "") & ">")
        																		Send("        <option value=""0"">--</option>")
        																		For Each row2 In dtAttributeValues.Rows
        																				Send("        <option value=""" & MyCommon.NZ(row2.Item("AttributeValueID"), 0) & """" & IIf((MyCommon.NZ(row2.Item("DefaultValue"), False) AndAlso (dtCustomerAttributes.Rows.Count <= 0)), " selected=""selected""", "") & IIf(MyCommon.NZ(row2.Item("AttributeValueID"), 0) = CustomerAttributeValueID, " selected=""selected""", "") & ">" & MyCommon.NZ(row2.Item("AttributeValue"), "") & "</option>")
        																		Next
        																		Send("      </select>")
        																Else
        																		Send("      " & GetAttributeValueName(CustomerAttributeValueID))
        																End If
        																Send("    </td>")
        																Send("  </tr>")
        														End If
        												Next
        												Send("</table>")
        										End If
        								End If
              
        								'Controls
        								Send("<br class=""half"" />")
        								If (Logix.UserRoles.EditCustomerIdData OrElse (Logix.UserRoles.AssignAttributes AndAlso MyCommon.Fetch_SystemOption(111))) Then
        										Send_Save()
        										Send("<input type=""button"" class=""regular"" id=""cancel"" name=""cancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""window.location.href='customer-general.aspx?CustPK=" & Cust.GetCustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & extraLink & "';"" />")
        								End If
        						End If
        				End If
        				Send("<hr class=""hidden"" />")
        				Send("</div>")
        		End If
        %>
        <%
            If MyCommon.IsEngineInstalled(2) Then
                Send("<div class=""box"" id=""basepoints""" & IIf(Logix.UserRoles.AccessPointsBalances, "", " style=""display:none;""") & ">")
                Send("  <h2>")
                Send("    <span>")
                Send("      " & Copient.PhraseLib.Lookup("term.points", LanguageID))
                Send("    </span>")
                Send("  </h2>")
      
                Dim sProgramID As String = ""
                Dim sProgramName As String
                Dim dPoints As Double
                Dim sSeparator As String = "  :  "
         
                'The following section displays the value of points in the Base Points program for that customer. The Base Points program is defined in CPE_SystemOptions table.
                sProgramID = MyCommon.Fetch_CPE_SystemOption(104).Trim
                If (sProgramID.IndexOf(";") > -1) Then
                    sProgramID = sProgramID.Substring(sProgramID.IndexOf(";") + 1)
                Else
                    sProgramID = 0
                End If
         
                MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & sProgramID & ";"
                Dim dtPointsPrograms As DataTable = MyCommon.LRT_Select
                If (dtPointsPrograms.Rows.Count > 0) Then
                    sProgramName = Trim(dtPointsPrograms.Rows(0).Item(0))
                    Send("      " & sProgramName & " ")
                Else
                    Send("      " & Copient.PhraseLib.Lookup("term.PointsProgram", LanguageID))
                    Send("      " & Copient.PhraseLib.Lookup("term.none", LanguageID))
                End If
         
                MyCommon.QueryStr = "select P.Amount from Points P with (NoLock) where ProgramID=" & sProgramID & " and CustomerPK=" & CustomerPK & ";"
                Dim dtPoints As DataTable = MyCommon.LXS_Select
                If (dtPoints.Rows.Count > 0) Then
                    dPoints = Trim(dtPoints.Rows(0).Item(0))
                    Send("      " & sSeparator)
                    Send("      " & dPoints)
                Else
                    Send("      " & sSeparator)
                    Send("      " & "0")
                End If
         
                Send("  <hr class=""hidden"" />")
                Send("</div>")
            End If
        
            'The following section displays the YTD savings for that customer.
            Dim DisplayYTDSavings As Integer = 0
            Try
                DisplayYTDSavings = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(118))
            Catch ex As Exception
                DisplayYTDSavings = 0
            End Try
            If DisplayYTDSavings = 1 Then
                Dim decCurrSTD As Decimal = 0.0
                Dim decCurrHhSTD As Decimal = 0.0
                Dim bHH As Boolean = False
                Dim bDisplaySTD As Boolean = False
                Dim bDisplayHhSTD As Boolean = False
                MyCommon.QueryStr = "select CurrYearSTD, HHPK  from Customers where CustomerPK=" & CustomerPK
                Dim dtCurrYTD As DataTable = MyCommon.LXS_Select
                If dtCurrYTD IsNot Nothing Then
                    If dtCurrYTD.Rows.Count > 0 Then
                        decCurrSTD = MyCommon.NZ(dtCurrYTD.Rows(0)("CurrYearSTD"), 0.0)
                        If dtCurrYTD.Rows(0)("HHPK") <> 0 Then
                            bHH = True
                            MyCommon.QueryStr = "select CurrYearSTD from Customers where CustomerPK=" & dtCurrYTD.Rows(0)("HHPK")
                            Dim dtHHCurrYTD As DataTable = MyCommon.LXS_Select
                            If dtHHCurrYTD IsNot Nothing Then
                                If dtHHCurrYTD.Rows.Count > 0 Then
                                    decCurrHhSTD = MyCommon.NZ(dtHHCurrYTD.Rows(0)("CurrYearSTD"), 0.0)
                                End If
                            End If
                        End If
                    End If
                End If
                If bHH Then
                    ' original row is customer, 2nd row is household
                    ' need to determine which to display          
                    Dim iHouseholdUpdateForSTD As Integer = 0
                    If Not Integer.TryParse(MyCommon.Fetch_SystemOption(6), iHouseholdUpdateForSTD) Then iHouseholdUpdateForSTD = 0
                    If iHouseholdUpdateForSTD = 3 Then
                        bDisplaySTD = True
                        bDisplayHhSTD = True
                    ElseIf iHouseholdUpdateForSTD = 2 Then
                        bDisplayHhSTD = True
                    Else
                        bDisplaySTD = True
                    End If
                Else
                    ' customer or household row, so only display this value
                    bDisplaySTD = True
                End If
          
                Send("<div class=""box"" id=""Savings"">")
                Send("  <h2>")
                Send("    <span>")
                Send("      " & Copient.PhraseLib.Lookup("term.savings", LanguageID))
                Send("    </span>")
                Send("  </h2>")
                If bDisplaySTD Then
                    Send("      " & Copient.PhraseLib.Lookup("tag.tsd", LanguageID) & " : " & decCurrSTD.ToString("$ 0.00"))
                End If
                If bDisplayHhSTD Then
                    If bDisplaySTD Then
                        Sendb("<br />")
                    End If
                    Send("      " & Copient.PhraseLib.Lookup("tag.tsd", LanguageID) & " : " & decCurrHhSTD.ToString("$ 0.00") & " (" & Copient.PhraseLib.Lookup("term.household", LanguageID) & ")")
                End If
                Send("</div> <!-- id=Savings -->")
            End If

            If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
                Send_Identifiers_Box(CustomerPK, MyCommon, Logix)
            Else
                Send_CardIDs_Box(Logix, MyCommon, MyLookup, CustomerPK, CardPK, IsHouseholdID, CustomerTypeID)
            End If

            If ((IsHouseholdID OrElse Cust.GetCustomerTypeID = 1) AndAlso Logix.UserRoles.ViewHHCardholders AndAlso Cust.GetCustomerPK > 0) Then%>
        <div class="box" id="cardholders" <%if (Cust.GetCustomerPK = 0) then sendb(" style=""display: none;""") %>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.cardholders", LanguageID))%>
                </span>
            </h2>
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.current", LanguageID) & " " & Copient.PhraseLib.Lookup("term.cardholders", LanguageID)) %>">
                <%
                    HHCustomers = MyLookup.GetCustomersInHousehold(CustomerPK, ReturnCode, 2)
                    DemotionPolicy = MyCommon.Fetch_SystemOption(98)
                    If ((DemotionPolicy = 1 AndAlso HHCustomers.Length = 1) OrElse (DemotionPolicy = 2) OrElse (DemotionPolicy = 3)) Then
                        DemotionDisplayed = True
                    End If
                    Send("<thead>")
                    Send("  <tr>")
                    Send("    <th align=""left"" class=""th-del"" scope=""col"">" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & "</th>")
                    If DemotionDisplayed Then
                        Send("    <th align=""left"" class=""th-del"" scope=""col"">" & Copient.PhraseLib.Lookup("term.demote", LanguageID) & "</th>")
                    End If
                    Send("    <th align=""left"" class=""th-cardholder"" scope=""col"">" & Copient.PhraseLib.Lookup("term.cardnumber", LanguageID) & "</th>")
                    Send("    <th align=""left"" class=""th-cardstatus"" scope=""col"">" & Copient.PhraseLib.Lookup("term.cardstatus", LanguageID) & "</th>")
                    If Logix.UserRoles.AccessCustomerIdData_FirstName Then
                        Send("    <th align=""left"" class=""th-firstname"" scope=""col"">" & Copient.PhraseLib.Lookup("term.firstname", LanguageID) & "</th>")
                    End If
                    If Logix.UserRoles.AccessCustomerIdData_MiddleName Then
                        Send("    <th align=""left"" class=""th-middlename"" scope=""col"">" & Copient.PhraseLib.Lookup("term.middlename", LanguageID) & "</th>")
                    End If
                    If Logix.UserRoles.AccessCustomerIdData_LastName Then
                        Send("    <th align=""left"" class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.lastname", LanguageID) & "</th>")
                    End If
                    Send("  </tr>")
                    Send("</thead>")
                    Send("<tbody>")
                    If HHCustomers.Length > 0 Then
                        For i = 0 To HHCustomers.GetUpperBound(0)
                            ' Check if this customer is queued for removal from the household.
                            MyCommon.QueryStr = "select PKID from HouseholdQueue with (NoLock) where CustomerPK = " & HHCustomers(i).GetCustomerPK & _
                                                "  and HHPK=" & CustomerPK
                            dt = MyCommon.LXS_Select
                            PendingRemoval = (dt.Rows.Count > 0)
                
                            Send("<tr>")
                            Send("  <td align=""center"">")
                    
                            Sendb("    <input type=""button"" class=""ex"" name=""ex"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ onclick=""removeFromHH(")
                            Sendb(MyCommon.NZ(HHCustomers(i).GetCustomerPK, "") & ",")
                            If HHCustomers(i).GetCards.Length = 0 Then
                                Sendb("0,'',")
                            Else
                                Sendb(MyCommon.NZ(HHCustomers(i).GetCards(0).GetCardPK, "") & ",")
                                'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                                Sendb("'" & IIf(HHCustomers(i).GetCards(0).GetCardTypeID = 3 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1, HHCustomers(i).GetCards(0).GetExtCardID.Substring(0, HHCustomers(i).GetCards(0).GetExtCardID.Length - 4) & "****", MyCommon.NZ(HHCustomers(i).GetCards(0).GetExtCardID, "")) & "',")
                            End If
                            Sendb(Cust.GetCustomerPK & ",")
                            Sendb(CardPK & ",")
                            Sendb("'" & ExtCardID & "',")
                            Sendb("0,'remove'")
                            Send(");"" " & IIf((Logix.UserRoles.RemoveHHCardholders = False) OrElse (DemotionPolicy = 3 AndAlso HHCustomers.Length = 1), "disabled=""disabled""", "") & " />")
                    
                            Send("  </td>")
                    
                            If DemotionDisplayed Then
                                Send("  <td align=""center"">")
                                If PendingRemoval Then
                                    Sendb(Copient.PhraseLib.Lookup("term.pending", LanguageID))
                                Else
                                    Sendb("    <input type=""button"" class=""ex"" name=""ex"" value=""D"" title=""" & Copient.PhraseLib.Lookup("term.demote", LanguageID) & """")
                                    Sendb("     onclick=""removeFromHH(")
                                    Sendb(MyCommon.NZ(HHCustomers(i).GetCustomerPK, "") & ",")
                                    If HHCustomers(i).GetCards.Length = 0 Then
                                        Sendb("0,'',")
                                    Else
                                        Sendb(MyCommon.NZ(HHCustomers(i).GetCards(0).GetCardPK, "") & ",")
                                        'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                                        Sendb("'" & IIf(HHCustomers(i).GetCards(0).GetCardTypeID = 3 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1, HHCustomers(i).GetCards(0).GetExtCardID.Substring(0, HHCustomers(i).GetCards(0).GetExtCardID.Length - 4) & "****", MyCommon.NZ(HHCustomers(i).GetCards(0).GetExtCardID, "")) & "',")
                                    End If
                                    Sendb(Cust.GetCustomerPK & ",")
                                    Sendb(CardPK & ",")
                                    Sendb("'" & ExtCardID & "',")
                                    Sendb("0,'demote'")
                                    Send(");"" " & IIf(Logix.UserRoles.RemoveHHCardholders, "", "disabled=""disabled""") & " />")
                                End If
                                Send("  </td>")
                            End If
                            Sendb("  <td>")
                            If (HHCustomers(i).GetCards.Length = 0) Then
                                Sendb("<a href=""customer-general.aspx?CustPK=" & HHCustomers(i).GetCustomerPK & "&amp;searchterms=" & MyCommon.NZ(HHCustomers(i).GetCustomerPK, "") & extraLink & """>")
                                Sendb("(" & Copient.PhraseLib.Lookup("term.nocard", LanguageID) & ")")
                                Sendb("</a>")
                            Else
                                For j = 0 To (HHCustomers(i).GetCards.Length - 1)
                                    If j > 0 Then
                                        Sendb("<br />")
                                    End If
                                    Sendb("<a href=""customer-general.aspx?CustPK=" & HHCustomers(i).GetCustomerPK & "&amp;CardPK=" & HHCustomers(i).GetCards(j).GetCardPK & "&amp;searchterms=" & MyCommon.NZ(HHCustomers(i).GetCustomerPK, "") & "&amp;search=Search" & extraLink & """>")
                                    'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                                    Sendb(IIf(HHCustomers(i).GetCards(j).GetCardTypeID = 3 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1, HHCustomers(i).GetCards(j).GetExtCardID.Substring(0, HHCustomers(i).GetCards(j).GetExtCardID.Length - 4) & "****", MyCommon.NZ(HHCustomers(i).GetCards(j).GetExtCardID, "")))
                                    Sendb("</a>")
                                Next
                            End If
                            Send("</td>")
                            Send("  <td>")
                            For j = 0 To (HHCustomers(i).GetCards.Length - 1)
                                If j > 0 Then
                                    Sendb("<br />")
                                End If
                                Sendb(MyCommon.SplitNonSpacedString(GetCardStatus(HHCustomers(i).GetCards(j).GetCardStatusID, MyLookup), 25).ToUpper)
                            Next
                            Send("</td>")
                            If Logix.UserRoles.AccessCustomerIdData_FirstName Then
                                Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(HHCustomers(i).GetFirstName, ""), 25) & "</td>")
                            End If
                            If Logix.UserRoles.AccessCustomerIdData_MiddleName Then
                                Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(HHCustomers(i).GetMiddleName, ""), 25) & "</td>")
                            End If
                            If Logix.UserRoles.AccessCustomerIdData_LastName Then
                                Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(HHCustomers(i).GetLastName, ""), 25) & "</td>")
                            End If
                            Send("</tr>")
                        Next
                    Else
                        Send("<tr>")
                        Send("  <td></td>")
                        Send("</tr>")
                    End If
                    Send("</tbody>")
                %>
            </table>
            <%
                'Add button
                If (Logix.UserRoles.AddHHCardholders) Then
                    Send("<a class=""hidden"" href=""customer-addhousehold.aspx?HHPK=" & Cust.GetCustomerPK & """>â–º</a>")
                    Send("<input type=""button"" class=""regular"" id=""addToHH"" name=""addToHH"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""addToHousehold('" & Cust.GetCustomerPK & "','" & CardPK & "')"" />")
                End If
                'Second cardholders table for pending members
                HHCustomers = MyLookup.GetCustomersInHousehold(CustomerPK, ReturnCode, 3)
                If HHCustomers.Length > 0 Then
                    Send("<table summary=""" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & " " & Copient.PhraseLib.Lookup("term.cardholders", LanguageID) & """>")
                    Send("<thead>")
                    Send("  <tr>")
                    Send("    <th align=""left"" class=""th-cardstatus"" scope=""col"">" & Copient.PhraseLib.Lookup("term.status", LanguageID) & "</th>")
                    Send("    <th align=""left"" class=""th-cardholder"" scope=""col"">" & Copient.PhraseLib.Lookup("term.cardnumber", LanguageID) & "</th>")
                    Send("    <th align=""left"" class=""th-cardstatus"" scope=""col"">" & Copient.PhraseLib.Lookup("term.cardstatus", LanguageID) & "</th>")
                    Send("    <th align=""left"" class=""th-firstname"" scope=""col"">" & Copient.PhraseLib.Lookup("term.firstname", LanguageID) & "</th>")
                    Send("    <th align=""left"" class=""th-middlename"" scope=""col"">" & Copient.PhraseLib.Lookup("term.middlename", LanguageID) & "</th>")
                    Send("    <th align=""left"" class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.lastname", LanguageID) & "</th>")
                    Send("  </tr>")
                    Send("</thead>")
                    Send("<tbody>")
                    For i = 0 To HHCustomers.GetUpperBound(0)
                        Send("<tr>")
                        Sendb("  <td>")
                        MyCommon.QueryStr = "select PKID, ActionTypeID, StatusCode from HouseholdQueue with (NoLock) where CustomerPK=" & HHCustomers(i).GetCustomerPK & ";"
                        rst = MyCommon.LXS_Select
                        If rst.Rows.Count > 0 Then
                            QueueStatus = ""
                            If MyCommon.NZ(rst.Rows(0).Item("StatusCode"), 0) < 0 Then
                                QueueStatus &= "<input type=""button"" class=""ex"" name=""ex"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """"
                                QueueStatus &= " onclick=""removeFromHH("
                                QueueStatus &= MyCommon.NZ(HHCustomers(i).GetCustomerPK, "") & ","
                                If HHCustomers(i).GetCards.Length = 0 Then
                                    QueueStatus &= "0,'',"
                                Else
                                    QueueStatus &= MyCommon.NZ(HHCustomers(i).GetCards(0).GetCardPK, "") & ","
                                    'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                                    QueueStatus &= "'" & IIf(HHCustomers(i).GetCards(0).GetCardTypeID = 3 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1, HHCustomers(i).GetCards(0).GetExtCardID.Substring(0, HHCustomers(i).GetCards(0).GetExtCardID.Length - 4) & "****", MyCommon.NZ(HHCustomers(i).GetCards(0).GetExtCardID, "")) & "',"
                                End If
                                QueueStatus &= Cust.GetCustomerPK & ","
                                QueueStatus &= CardPK & ","
                                QueueStatus &= "'" & ExtCardID & "',"
                                QueueStatus &= MyCommon.NZ(rst.Rows(0).Item("PKID"), 0) & ",'unqueue'"
                                QueueStatus &= ");"" />"
                                QueueStatus &= Copient.PhraseLib.Lookup("term.failed", LanguageID) & "&nbsp;"
                            End If
                            If MyCommon.NZ(rst.Rows(0).Item("ActionTypeID"), 0) = 1 Then
                                QueueStatus &= Copient.PhraseLib.Lookup("term.incoming", LanguageID)
                            ElseIf MyCommon.NZ(rst.Rows(0).Item("ActionTypeID"), 0) = 2 Then
                                QueueStatus &= Copient.PhraseLib.Lookup("term.outgoing", LanguageID)
                            End If
                            Sendb(QueueStatus)
                        End If
                        Send("</td>")
                        Sendb("  <td>")
                        If (HHCustomers(i).GetCards.Length = 0) Then
                            Sendb("<a href=""customer-general.aspx?CustPK=" & HHCustomers(i).GetCustomerPK & "&amp;searchterms=" & MyCommon.NZ(HHCustomers(i).GetCustomerPK, "") & extraLink & """>")
                            Sendb("(" & Copient.PhraseLib.Lookup("term.nocard", LanguageID) & ")")
                            Sendb("</a>")
                        Else
                            For j = 0 To (HHCustomers(i).GetCards.Length - 1)
                                If j > 0 Then
                                    Sendb("<br />")
                                End If
                                Sendb("<a href=""customer-general.aspx?CustPK=" & HHCustomers(i).GetCustomerPK & "&amp;CardPK=" & HHCustomers(i).GetCards(j).GetCardPK & "&amp;searchterms=" & MyCommon.NZ(HHCustomers(i).GetCustomerPK, "") & "&amp;search=Search" & extraLink & """>")
                                'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                                Dim cardIdLength As Integer = If(HHCustomers(i).GetCards(j).GetExtCardID.Length >= 4, HHCustomers(i).GetCards(j).GetExtCardID.Length - 4, HHCustomers(i).GetCards(j).GetExtCardID.Length)
                                Sendb(IIf(HHCustomers(i).GetCards(j).GetCardTypeID = 3 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1, HHCustomers(i).GetCards(j).GetExtCardID.Substring(0, cardIdLength) & "****", MyCommon.NZ(HHCustomers(i).GetCards(j).GetExtCardID, "")))
                                Sendb("</a>")
                            Next
                        End If
                        Send("</td>")
                        Send("  <td>")
                        For j = 0 To (HHCustomers(i).GetCards.Length - 1)
                            If j > 0 Then
                                Sendb("<br />")
                            End If
                            Sendb(MyCommon.SplitNonSpacedString(GetCardStatus(HHCustomers(i).GetCards(j).GetCardStatusID, MyLookup), 25).ToUpper)
                        Next
                        Send("</td>")
                        Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(HHCustomers(i).GetFirstName, ""), 25) & "</td>")
                        Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(HHCustomers(i).GetMiddleName, ""), 25) & "</td>")
                        Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(HHCustomers(i).GetLastName, ""), 25) & "</td>")
                        Send("</tr>")
                    Next
                    Send("</tbody>")
                    Send("</table>")
                End If
            %>
            <hr class="hidden" />
        </div>
        <% End If%>
        <div class="box" id="uniquecoupons" <%Sendb(IIf(MyCommon.Fetch_SystemOption(112) = "0", " style=""display:none;""", ""))%>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.uniquecoupons", LanguageID))%>
                </span>
            </h2>
            <%
                Dim BarcodesDT As DataTable
                MyCommon.QueryStr = "select COUNT(BD1.Barcode) as Total, COUNT(BD2.Barcode) as Voided from BarcodeDetails as BD1 " & _
                                    "left join BarcodeDetails as BD2 on BD1.Barcode=BD2.Barcode and BD2.Voided=1 " & _
                                    "where BD1.CustomerPK=" & CustomerPK & ";"
                BarcodesDT = MyCommon.LXS_Select
                Sendb("Customer has " & BarcodesDT.Rows(0).Item("Total") & " unique coupon" & IIf(BarcodesDT.Rows(0).Item("Total") = 1, "", "s"))
                If BarcodesDT.Rows(0).Item("Total") > 0 Then
                    Sendb(", " & IIf(BarcodesDT.Rows(0).Item("Voided") = 0, "none", BarcodesDT.Rows(0).Item("Voided")) & " of which ")
                    Sendb(IIf(BarcodesDT.Rows(0).Item("Voided") = 1, "is", "are") & " void")
                End If
                Send(".<br />")
                If Logix.UserRoles.AccessCustomerCoupons Then
                    Send("<input type=""button"" id=""uniquecoupons"" name=""uniquecoupons"" value=""" & Copient.PhraseLib.Lookup("term.uniquecoupons", LanguageID) & "..."" onclick=""openExtraWidePopup('/logix/customer-coupons.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "');"" />")
                End If
            %>
        </div>

        <!-- CR119 Linked Cards -->
        <%If (bLinkedCards) Then%>
        <div class="box" id="Div1">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.linkedcards", LanguageID))%>
                </span>
            </h2>

            <%
            MyCommon.QueryStr = "select LinkedCard from CustomerExt with (NoLock) where CustomerPK=" & CustomerPK & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
               LinkedCardID = (MyCommon.NZ(dt.Rows(0).Item("LinkedCard"), ""))
            End If
            Send("<label for=""LinkedCards"">" & Copient.PhraseLib.Lookup("term.linkedto", LanguageID) & ":</label>")
            Send("<input type=""text"" class=""medium"" id=""linkedCardID"" name=""linkedCardID"" style=""width:25%;"" maxlength=""26"" value=""" & LinkedCardID & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & """ />")
        
            Dim DefaultCustTypeID As Integer = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(147), 5)
            MyCommon.QueryStr = "select CardTypeID, Description, PhraseID from CardTypes with (NoLock) "
            MyCommon.QueryStr &= "where CardTypeID = " & DefaultCustTypeID.ToString() & ";"
            rst2 = MyCommon.LXS_Select
            Send("<select id=""linkedCardTypeID"" name=""linkedCardTypeID"" style=""width:20%;"">")
            For Each row2 In rst2.Rows
               Send("<option value=""" & MyCommon.NZ(row2.Item("CardTypeID"), 0) & """" & IIf(DefaultCustTypeID = MyCommon.NZ(row2.Item("CardTypeID"), 0), "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row2.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row2.Item("Description"), "")) & "</option>")
            Next
            Send("</select>")
            If (Logix.UserRoles.EditCustomerGroups) Then
               Sendb("<input type=""button"" class=""regular"" id=""saveLink"" name=""saveLink"" onclick=""saveLinkedCard(" & CustomerPK & "," & CardPK & ");"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
            End If
            %>

        <br />
        <br />
        <%
        Send("<label for=""LinkedCards"">" & Copient.PhraseLib.Lookup("term.linkedfrom", LanguageID) & ":</label>")
        If (GetCgiValue("ShowLinks") = "true") Then
           Dim CardID As String = ""
           Dim CardAndType As String = ""
           Send("<br />")
                MyCommon.QueryStr = "select C.ExtCardIDOriginal as ExtCardId, T.PhraseID, T.Description from CardIDs as C " & _
                               "join CustomerExt as E on C.CustomerPK=E.CustomerPK " & _
                               "join CardTypes as T on C.CardTypeID=T.CardTypeID " & _
                               "where E.LinkedCard in (select ExtCardID from CardIDs where CustomerPK=" & CustomerPK & ")"
           rst2 = MyCommon.LXS_Select
           Send("<select class=""longer"" id=""linkedfromlist"" name=""linkedfromlist"" size=""10"" multiple=""multiple"">")
           For Each row2 In rst2.Rows
              CardID = MyCryptLib.SQL_StringDecrypt(row2.Item("ExtCardID").ToString())
              CardAndType = CardID & " (" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row2.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row2.Item("Description"), "")) & ")"
              Send("<option value=" & CardID & " alt=" & CardAndType & " title=" & CardAndType & ">" & CardAndType & "</option>")
           Next
        Else
           Dim ButtonPhrase = Copient.PhraseLib.Lookup("term.showcards", LanguageID)
           Send("<input type=""button"" class=""regular"" id=""showLinked"" name=""showLinked"" value=" & ButtonPhrase & " title=" & ButtonPhrase & " onclick=""window.location.href='customer-general.aspx?ShowLinks=true" & extraLink & "&amp;editterms=" & Cust.GetCustomerPK & "&amp;CustPK=" & Cust.GetCustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "';"" />")
        End If%>

        </div>
        <%End If%>

        <br clear="all" />
        <% If (Logix.UserRoles.ViewCustomerNotes) Then%>
        <div class="box" id="recentnotes">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.notes", LanguageID))%>
                </span>
            </h2>
            <%
            
                Dim UseNotesAndActivity As Integer = 0
                UseNotesAndActivity = MyCommon.Fetch_SystemOption(140)
            
                CustNotes = MyLookup.GetCustomerNotes(CustomerPK, "CreatedDate", False, ReturnCode)
                If CustNotes.Length > 0 Then
                    If CustNotes.Length > 10 Then
                        Send(Copient.PhraseLib.Lookup("customer-inquiry.notestopten", LanguageID) & "<br />")
                    End If
                    Send("<table summary=""" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & """>")
                    Send("  <thead>")
                    Send("    <tr>")
                    Send("      <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.created", LanguageID) & "</th>")
                    Send("      <th class=""th-author"" scope=""col"">" & Copient.PhraseLib.Lookup("term.author", LanguageID) & "</th>")
                    Send("      <th class=""th-note"" scope=""col"">" & Copient.PhraseLib.Lookup("term.note", LanguageID) & "</th>")
                    Send("    </tr>")
                    Send("  </thead>")
                    Send("  <tbody>")
                    i = 1
                    For Each CustNote In CustNotes
                        If i > 10 Then
                            GoTo closenotes
                        Else
                            Send("    <tr" & Shaded & ">")
                            Send("      <td>" & Logix.ToShortDateTimeString(MyCommon.NZ(CustNote.GetCreatedDate, New Date(1900, 1, 1)), MyCommon) & "</td>")
                            If CustNote.GetFirstName = "" OrElse CustNote.GetLastName = "" Then
                                MyCommon.QueryStr = "select FirstName, LastName from AdminUsers where AdminUserID=" & CustNote.GetAdminUserID & ";"
                                rst4 = MyCommon.LRT_Select()
                                If rst4.Rows.Count > 0 Then
                                    Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst4.Rows(0).Item("FirstName"), ""), 25) & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst4.Rows(0).Item("LastName"), ""), 25) & "</td>")
                                End If
                            Else
                                Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(CustNote.GetFirstName, ""), 25) & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(CustNote.GetLastName, ""), 25) & "</td>")
                            End If
                            Note = MyCommon.SplitNonSpacedString(MyCommon.NZ(CustNote.GetNote, ""), 40)
                            Note = Note.Replace(vbCrLf, "<br />")
                            Send("      <td>" & Note & "</td>")
                            Send("    </tr>")
                            If Shaded = " class=""shaded""" Then
                                Shaded = ""
                            Else
                                Shaded = " class=""shaded"""
                            End If
                        End If
                        i = i + 1
                    Next
closenotes:
                
                    If (UseNotesAndActivity = 1) Then
                        Send("    <tr>")
                        Send("      <td colspan=""2"">")
                  Send("     <br><a href=""javascript:openPopup('customer-activities.aspx?CardPK=" & CardPK & "&amp;CustPK=" & CustomerPK & "&amp;SortDirection=True')"">" & Copient.PhraseLib.Lookup("term.consnotesandhist", LanguageID) & "</a>")
                        Send("    </td>")
                        Send("   </tr>")
                        Send("<br>")
                    End If
                
                    Send("  </tbody>")
                    Send("</table>")
                Else
                    Send(Copient.PhraseLib.Lookup("customer.nonotesposted", LanguageID) & "<br />")
                
                    If (UseNotesAndActivity = 1) Then
                        
                  Send("     <br><a href=""javascript:openPopup('customer-activities.aspx?CardPK=" & CardPK & "&amp;CustPK=" & CustomerPK & "&amp;SortDirection=True')"">" & Copient.PhraseLib.Lookup("term.consnotesandhist", LanguageID) & "</a>")
                        
                    End If
                
                End If
            
            
            %>
        </div>
        <% End If%>
    </div>
    <br clear="all" />
</div>
</form>
<script runat="server">  

    '-------------------------------------------------------------------------------------------------------------  

    Private CardStatusTable As Hashtable = Nothing

    Function ParseDateOfBirth(ByVal DateOfBirth As String) As String()
        Dim DOBParts() As String = {"", "", ""}

        If (DateOfBirth IsNot Nothing) Then
            Select Case DateOfBirth.Length
                Case 4
                    DOBParts(0) = DateOfBirth.Substring(0, 2).PadLeft(2, "0")
                    DOBParts(1) = DateOfBirth.Substring(2, 2).PadLeft(2, "0")
                    DOBParts(2) = ""
                Case 8
                    DOBParts(0) = DateOfBirth.Substring(0, 2).PadLeft(2, "0")
                    DOBParts(1) = DateOfBirth.Substring(2, 2).PadLeft(2, "0")
                    DOBParts(2) = DateOfBirth.Substring(4).PadLeft(4, "0")
            End Select
        End If

        Return DOBParts

    End Function

    Function ParseDateOpened(ByVal DateOpened As String) As String()
        Dim DateOpenedParts() As String = {"", "", ""}

        If (DateOpened IsNot Nothing) Then
            Select Case DateOpened.Length
                Case 4
                    DateOpenedParts(0) = DateOpened.Substring(0, 2).PadLeft(2, "0")
                    DateOpenedParts(1) = DateOpened.Substring(2, 2).PadLeft(2, "0")
                    DateOpenedParts(2) = ""
                Case 8
                    DateOpenedParts(0) = DateOpened.Substring(0, 2).PadLeft(2, "0")
                    DateOpenedParts(1) = DateOpened.Substring(2, 2).PadLeft(2, "0")
                    DateOpenedParts(2) = DateOpened.Substring(4).PadLeft(4, "0")
            End Select
        End If

        Return DateOpenedParts

    End Function
    'RT 4125- Customer Enrollment Date
    Function ParseEnrollmentDate(ByVal EnrollmentDate As String) As String()
        Dim EnrollmentDateParts() As String = {"", "", ""}

        If (EnrollmentDate IsNot Nothing) Then
            Select Case EnrollmentDate.Length
                Case 4
                    EnrollmentDateParts(0) = EnrollmentDate.Substring(0, 2).PadLeft(2, "0")
                    EnrollmentDateParts(1) = EnrollmentDate.Substring(2, 2).PadLeft(2, "0")
                    EnrollmentDateParts(2) = ""
                Case 8
                    EnrollmentDateParts(0) = EnrollmentDate.Substring(0, 2).PadLeft(2, "0")
                    EnrollmentDateParts(1) = EnrollmentDate.Substring(2, 2).PadLeft(2, "0")
                    EnrollmentDateParts(2) = EnrollmentDate.Substring(4).PadLeft(4, "0")
            End Select
        End If

        Return EnrollmentDateParts

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function ParseTableCol(ByVal TableCol As String) As String
        Dim Col As String = ""

        If (TableCol IsNot Nothing) Then
            Col = TableCol.ToString.Trim
            If (Col.IndexOf(".") > -1) Then
                Col = Col.Substring(Col.IndexOf("."))
                If (Left(Col, 1) = ".") Then Col = Col.Substring(1)
            End If
        End If

        Return Col
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function ParseTable(ByVal TableCol As String) As String
        Dim Table As String = ""
        Dim Fields() As String

        If (TableCol IsNot Nothing) Then
            Fields = TableCol.Split(".")
            Table = Fields(0)
        End If

        Return Table
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function GetDesignatorText(ByVal Field As String, ByVal AltID As String, ByVal Verifier As String) As String
        Dim Tag As String = ""

        If (AltID.ToUpper = Field.ToUpper) Then
            Tag = " <span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.alternateid", LanguageID) & ")</span>"
        ElseIf (Verifier.ToUpper = Field.ToUpper) Then
            Tag = " <span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.alternate-id-verifier", LanguageID) & ")</span>"
        End If

        Return Tag
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function GetNewAltID(ByVal AltIDColumn As String, ByRef FieldName As String) As String
        Dim NewAltID As String = ""

        Select Case AltIDColumn.ToUpper
            Case "PHONEDIGITSONLY"
                NewAltID = GetCgiValue("Phone1") & GetCgiValue("Phone2") & GetCgiValue("Phone3")
                FieldName = Copient.PhraseLib.Lookup("customer.phone", LanguageID)
            Case "LASTNAME"
                NewAltID = GetCgiValue("LastName")
                FieldName = Copient.PhraseLib.Lookup("term.lastname", LanguageID)
            Case "FIRSTNAME"
                NewAltID = GetCgiValue("FirstName")
                FieldName = Copient.PhraseLib.Lookup("term.firstname", LanguageID)
            Case "ALTID"
                NewAltID = GetCgiValue("AltIDValue")
                FieldName = Copient.PhraseLib.Lookup("term.alternateid", LanguageID)
            Case "EMAIL"
                NewAltID = GetCgiValue("Email")
                FieldName = Copient.PhraseLib.Lookup("term.email", LanguageID)
            Case "DOB"
                NewAltID = GetCgiValue("dob1") & GetCgiValue("dob2") & GetCgiValue("dob3")
                FieldName = Copient.PhraseLib.Lookup("customer.dateofbirth", LanguageID)
            Case Else
                NewAltID = ""
                FieldName = ""
        End Select

        Return NewAltID
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    'Checks to see if the given month is a number between 1 and 12
    Function ValidateMonth(ByVal DOB_Month As String) As Boolean
        Dim Month As String = DOB_Month.Trim
        Dim Validated As Boolean = False
        Dim MonthNumber As Integer

        If (Month <> "" And IsNumeric(Month)) Then
            MonthNumber = Val(Month)
            If (MonthNumber <= 12 And MonthNumber > 0) Then
                Validated = True
            End If
        End If

        Return Validated
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    'Checks that the give day is a number between 1 and 31
    Function ValidateDay(ByVal DOB_Day As String) As Boolean
        Dim Day As String = DOB_Day.Trim
        Dim Validated As Boolean = False
        Dim DayNumber As Integer

        If (Day <> "" And IsNumeric(Day)) Then
            DayNumber = Val(Day)
            If (DayNumber <= 31 And DayNumber > 0) Then
                Validated = True
            End If
        End If

        Return Validated
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function ValidateYear(ByVal DOB_Year As String) As Boolean
        Dim Year As String = DOB_Year.Trim
        Dim Validated As Boolean = False
        Dim YearNumber As Integer

        If (Year <> "" And IsNumeric(Year)) Then
            YearNumber = Val(Year)
            If (YearNumber <= 2100 And YearNumber > 1900) Then
                Validated = True
            End If
        End If

        Return Validated
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function Demote_Balances(ByVal HHPK As String, ByVal CustPK As String, ByVal HHMemberCt As Integer, ByVal AdminUserID As Long) As Boolean
        Dim MyCommon As New Copient.CommonInc
        Dim Demoted As Boolean = False
        Dim SV As New Copient.StoredValue
        Dim Pts As New Copient.Points
        Dim Lookup As New Copient.CustomerLookup
        Dim RetCode As New Copient.CustomerLookup.RETURN_CODE
        Dim RunAgain As Boolean = True
        Dim PassCt As Integer = 0
        Const INACTIVE_STATUS As Integer = 2

        MyCommon.Open_LogixRT()
        MyCommon.Open_LogixXS()
        'MyCommon.Write_Log("cgtest.txt", "0) HHPK=" & HHPK & ", CustPK=" & CustPK & ", HHMemberCt=" & HHMemberCt & ", AdminUserID=" & AdminUserID & "     ")

        If HHPK > 0 AndAlso CustPK > 0 Then
            ' move the balances from the HHID to the CustID
            SV.TransferBalances(1, HHPK, CustPK, False)
            Pts.TransferBalances(1, HHPK, CustPK, False)

            ' change the history records for HHID to the CustID - only for adjustments made to points and stored value
            PassCt = 1
            While RunAgain AndAlso PassCt <= 2
                MyCommon.QueryStr = "dbo.pa_ActivityLog_TransferAdjustments"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourcePK", SqlDbType.BigInt).Value = HHPK
                MyCommon.LRTsp.Parameters.Add("@DestPK", SqlDbType.BigInt).Value = CustPK
                MyCommon.LRTsp.Parameters.Add("@RunAgain", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                RunAgain = MyCommon.LRTsp.Parameters("@RunAgain").Value
                MyCommon.Close_LRTsp()
                PassCt += 1
            End While

            If HHMemberCt = 0 Then
                ' set the household to inactive when no cardholders remain in the household
                MyCommon.QueryStr = "Update Customers with (RowLock) set CardStatusID=" & 2 & " where CustomerPK=" & HHPK
                MyCommon.LXS_Execute()
                If MyCommon.RowsAffected > 0 Then
                    Demoted = True
                    MyCommon.Activity_Log(25, HHPK, AdminUserID, Copient.PhraseLib.Detokenize("customer-general.HouseholdInactivated", LanguageID, HHPK))
                End If
            End If
        End If

        MyCommon.Close_LogixRT()
        MyCommon.Close_LogixXS()

        Return Demoted
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function GetCardStatus(ByVal CardStatusID As Integer, ByRef MyLookup As Copient.CustomerLookup) As String
        Dim CardStatus As String = ""

        If CardStatusTable Is Nothing Then
            CardStatusTable = MyLookup.GetCardStatuses()
        End If

        If CardStatusTable.ContainsKey(CardStatusID.ToString) Then
            CardStatus = CardStatusTable.Item(CardStatusID.ToString).ToString
        Else
            CardStatus = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
        End If

        Return CardStatus
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function GetQueueEntryForRemove(ByVal CustomerPK As Long, ByVal HHPK As Long, ByVal AdminUserID As Integer, _
                                    ByVal HHOptions As Copient.HouseholdRules.InterfaceOption(), _
                                    Optional ByVal TransferPercent As Decimal = -1) As Copient.HouseholdRules.QUEUE_DATA
        Dim qData As New Copient.HouseholdRules.QUEUE_DATA

        qData.ActionTypeID = Copient.HouseholdRules.ACTION_TYPES.REMOVE
        qData.SourceTypeID = Copient.HouseholdRules.SOURCE_TYPES.LOGIX
        qData.CustomerPK = CustomerPK
        qData.AdminUserID = AdminUserID
        qData.HHPK = HHPK

        If HHOptions IsNot Nothing AndAlso HHOptions.Length >= 5 Then
            qData.Option5Value = HHOptions(0).Value
            qData.Option6Value = HHOptions(1).Value
            qData.Option7Value = HHOptions(2).Value
            qData.Option8Value = HHOptions(3).Value
            qData.Option9Value = HHOptions(4).Value
        End If

        If TransferPercent > -1 AndAlso TransferPercent <= 100 Then
            qData.Option6Percent = TransferPercent.ToString
        End If

        Return qData
    End Function

    Function GetAttributeValueName(ByVal AttributeValueID As Integer) As String
        Dim AttributeValueName As String = ""
        Dim MyCommon As New Copient.CommonInc
        Dim dt As DataTable

        MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "select Description from AttributeValues with (NoLock) where AttributeValueID=" & AttributeValueID & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            AttributeValueName = MyCommon.NZ(dt.Rows(0).Item("Description"), "")
        End If
        MyCommon.Close_LogixRT()

        Return AttributeValueName
    End Function

    Function GetDisplayTextForCardType(ByVal CradTypeID As Integer) As String
        Dim text As String = ""
        Select Case CradTypeID
            Case 0
                text = "term.customercard"
            Case 1
                text = "term.householdcard"
            Case 2
                text = "term.camcard"
            Case 3
                text = "term.alternateid"
            Case 4
                text = "term.username"
            Case 5
                text = "term.associateid"
            Case 6
                text = "term.emailaddress"
            Case 7
                text = "term.secondarymembercard"
        End Select
        Return (Copient.PhraseLib.Lookup(text, LanguageID))
    End Function
    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_CardIDs_Box(ByRef Logix As Copient.LogixInc, ByRef MyCommon As Copient.CommonInc, ByRef MyLookup As Copient.CustomerLookup, _
                         ByVal CustomerPK As Long, ByVal CardPK As Long, ByVal IsHouseholdID As Long, ByVal CustomerTypeID As Integer)

        Dim rst2 As DataTable
        Dim row As DataRow
        Dim AddPermitted As Boolean = False
        Dim StatusEnumerator As IDictionaryEnumerator
        Dim AddCardBuf As New StringBuilder()
        Dim bReadOnlyAlternateID = False
        Dim AlternateIDCardType As Integer = -1
        Dim maxLength As Integer = 256
        Dim MyCryptLib As New Copient.CryptLib

        ' AL-1069 If new Card Number is invalid, CustomerTypeID will always be 0 which does not work for households
        If (IsHouseholdID) Then
            CustomerTypeID = 1
        End If

        Send("<div class=""box"" id=""cards"" " & IIf(Logix.UserRoles.AccessCustomerIdData_Cards, "", " style=""display:none;""") & ">")
        Send("<h2><span>")
        Sendb(Copient.PhraseLib.Lookup("term.cards", LanguageID))
        Send("</span></h2>")

        CardStatusTable = MyLookup.GetCardStatuses()
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.cards", LanguageID) & """>")
        Send("<thead>")
        Send("  <tr>")
        Send("    <th scope=""col"" class=""th-id"">&nbsp;</th>")
        Send("    <th scope=""col"">" & Copient.PhraseLib.Lookup("term.cardnumber", LanguageID) & "</th>")
        Send("    <th scope=""col"">" & Copient.PhraseLib.Lookup("term.cardtype", LanguageID) & "</th>")
        Send("    <th scope=""col"">" & Copient.PhraseLib.Lookup("term.status", LanguageID) & "</th>")
        Send("  </tr>")
        Send("</thead>")
        Send("<tbody>")

        'If the correct System Option is set, do not allow Alternate ID Cards to be changed
        If MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1 Then
            bReadOnlyAlternateID = True
        End If
        MyCommon.QueryStr = "SELECT CardTypeID FROM CardTypes as c " & _
                            "where c.CardTypeID like '%Alternate%ID%' " & _
                            "or c.extcardtypeid like '%Alternate%ID%'" & _
                            "or c.description like '%Alternate%ID%';"
        rst2 = MyCommon.LXS_Select
        If rst2.Rows.Count > 0 Then
            AlternateIDCardType = MyCommon.NZ(rst2.Rows(0).Item("CardTypeID"), -1)
        End If
        MyCommon.QueryStr = "select CID.CardPK, CID.ExtCardIDOriginal as ExtCardID, CID.CardStatusID, CS.Description as StatusDescription, CS.PhraseID as StatusPhraseID, " & _
                            " CID.CardTypeID, CT.Description as TypeDescription, CT.PhraseID as TypePhraseID, CT.ExtCardTypeID " & _
                            "from CardIDs as CID with (NoLock) " & _
                            "inner join CardStatus as CS on CS.CardStatusID=CID.CardStatusID " & _
                            "inner join CardTypes as CT on CT.CardTypeID=CID.CardTypeID " & _
                            "where CustomerPK=" & CustomerPK & " and CT.CardTypeID <> 8 order by ExtCardID;"
        rst2 = MyCommon.LXS_Select
        For k As Integer = 1 To rst2.Rows.Count
            rst2.Rows(k - 1).Item("ExtCardID") = MyCryptLib.SQL_StringDecrypt(rst2.Rows(k - 1).Item("ExtCardID"))
        Next
        'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
        If rst2.Rows.Count > 0 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1 Then
            For k As Integer = 1 To rst2.Rows.Count
                If MyCommon.NZ(rst2.Rows(k - 1).Item("CardTypeID"), 0) = 3 AndAlso MyCommon.NZ(rst2.Rows(k - 1).Item("ExtCardID"), "").Length >= 14 Then
                    rst2.Rows(k - 1).Item("ExtCardID") = rst2.Rows(k - 1).Item("ExtCardID").Substring(0, rst2.Rows(k - 1).Item("ExtCardID").Length - 4) & "****"
                End If
            Next
        End If

        If rst2.Rows.Count > 0 Then
            For Each row In rst2.Rows
                Dim bReadOnly As Boolean = False
                If row.Item("CardTypeID") = AlternateIDCardType AndAlso bReadOnlyAlternateID Then
                    bReadOnly = True
                End If
                Send("  <tr class=""shaded"">")
                Send("    <td align=""center"">")
                If Logix.UserRoles.EditCustomerIdData And Logix.UserRoles.DeleteCustomerCard And Not bReadOnly And  rst2.Rows.Count > 1 Then
                    Sendb("      <input type=""button"" class=""ex"" name=""ex"" value=""X""" & IIf(MyCommon.Fetch_SystemOption(312) = "1", "", " style=""display:none;""") & " title=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ onclick=""removeCard(")
                    Sendb(CustomerPK & "," & CardPK & "," & IIf(IsHouseholdID, 1, 0) & "," & MyCommon.NZ(row.Item("CardPK"), 0) & ",'" & MyCommon.NZ(row.Item("ExtCardID").Replace("'", "\'").Replace("""", "&quot;") & "'", "&nbsp;"))
                    Send(");"" />")
                End If
                If Logix.UserRoles.EditCustomerIdData And Not bReadOnly Then
                    Sendb("      <input type=""button"" class=""adjust"" name=""adjust" & MyCommon.NZ(row.Item("CardPK"), 0) & """ value=""S"" title=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ onclick=""saveCard(")
                    Send(CustomerPK & "," & CardPK & "," & MyCommon.NZ(row.Item("CardPK"), 0) & "," & IIf(GetCgiValue("edit") <> "", "true", "false") & ");"" />")
                End If
                Send("    </td>")
                If bReadOnly Then
                    If CardPK = MyCommon.NZ(row.Item("CardPK"), 0) Then
                        'Send("    <td>" & MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;test").Substring(0,MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;test").Length-4) & "</td>")
                        Send("    <td>" & MyCommon.TruncateString(MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;test"), 34) & "</td>")
                    Else
                        'Send("    <td><a href=""customer-general.aspx?CustPK=" & CustomerPK & "&amp;CardPK=" & MyCommon.NZ(row.Item("CardPK"), 0) & """>" & MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;").Substring(0,MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;test").Length-4) & "</a></td>")
                        Send("    <td><a href=""customer-general.aspx?CustPK=" & CustomerPK & "&amp;CardPK=" & MyCommon.NZ(row.Item("CardPK"), 0) & """>" & MyCommon.TruncateString(MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;"), 34) & "</a></td>")
                    End If
                Else
                    If CardPK = MyCommon.NZ(row.Item("CardPK"), 0) Then
                        Send("    <td>" & MyCommon.TruncateString(MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;"), 34) & "</td>")
                    Else
                        Send("    <td><a href=""customer-general.aspx?CustPK=" & CustomerPK & "&amp;CardPK=" & MyCommon.NZ(row.Item("CardPK"), 0) & """>" & MyCommon.TruncateString(MyCommon.NZ(row.Item("ExtCardID"), "&nbsp;"), 34) & "</a></td>")
                    End If
                End If
                Send("    <td>" & IIf(IsDBNull(row.Item("TypePhraseID")), MyCommon.NZ(row.Item("TypeDescription"), "&nbsp;"), Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("TypePhraseID"), 0), LanguageID)) & "</td>")
                Send("    <td>")
                If bReadOnly Then
                    Send("      <select name=""cardStatus" & MyCommon.NZ(row.Item("CardPK"), 0) & """ id=""cardStatus" & MyCommon.NZ(row.Item("CardPK"), 0) & """" & " disabled=""disabled""" & ">")
                Else
                    Send("      <select name=""cardStatus" & MyCommon.NZ(row.Item("CardPK"), 0) & """ id=""cardStatus" & MyCommon.NZ(row.Item("CardPK"), 0) & """" & IIf(Logix.UserRoles.EditCustomerIdData, "", " disabled=""disabled""") & ">")
                End If
                StatusEnumerator = CardStatusTable.GetEnumerator
                While (StatusEnumerator.MoveNext())

                    If (StatusEnumerator.Key.ToString = 6 AndAlso Not (MyCommon.IsEngineInstalled(2) OrElse MyCommon.IsEngineInstalled(0))) Then
                        Continue While
                    End If

                    Send("        <option value=""" & StatusEnumerator.Key.ToString & """")
                    If MyCommon.NZ(row.Item("CardStatusID"), 0) = Integer.Parse(StatusEnumerator.Key) Then
                        Sendb(" selected=""selected""")
                    End If
                    Send(">" & StatusEnumerator.Value.ToString & "</option>")
                End While
                Send("      </select>")
                Send("    </td>")

                'If (MyCommon.NZ(row.Item("CardStatusID"), 0) = 1) Then
                '  Sendb("    <td><span class=""darkgreen""><b>")
                'ElseIf (MyCommon.NZ(row.Item("CardStatusID"), 0) = 2) Then
                '  Sendb("    <td><span class=""darkgray""><b>")
                'ElseIf (MyCommon.NZ(row.Item("CardStatusID"), 0) = 6) Then
                '  Sendb("    <td><span><b>")
                'Else
                '  Sendb("    <td><span class=""red""><b>")
                'End If
                'Sendb(IIf(IsDBNull(row.Item("StatusPhraseID")), MyCommon.NZ(row.Item("StatusDescription"), "&nbsp;"), Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("StatusPhraseID"), 0), LanguageID)) & "</b></span></td>")
                Send("  </tr>")
            Next
        Else
            Send("  <tr class=""shaded"">")
            Send("    <td></td>")
            Send("    <td colspan=""3"">" & Copient.PhraseLib.Lookup("customer-inquiry.nocards", LanguageID) & "</td>")
            Send("  </tr>")
        End If

        'Add card
        MyCommon.QueryStr = "select CardTypeID, Description, PhraseID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & " and CardTypeID <> 8;" 'Consumer Account Number shouldn't be displayed on UI
        rst2 = MyCommon.LXS_Select
        If rst2.Rows.Count > 0 Then
            For Each row In rst2.Rows
                ' determine if adding a card of this type is permitted for this customer
                MyCommon.QueryStr = "pa_CUA_AddCardPermitted"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = MyCommon.NZ(row.Item("CardTypeID"), -1)
                MyCommon.LXSsp.Parameters.Add("@ReadOnlyAlternate", SqlDbType.BigInt).Value = bReadOnlyAlternateID
                MyCommon.LXSsp.Parameters.Add("@Permitted", SqlDbType.Bit).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                AddPermitted = MyCommon.LXSsp.Parameters("@Permitted").Value
                MyCommon.Close_LXSsp()
                If AddPermitted Then
                    AddCardBuf.Append("        <option value=""" & MyCommon.NZ(row.Item("CardTypeID"), 0) & """>" & IIf(IsDBNull(row.Item("PhraseID")), MyCommon.NZ(row.Item("Description"), "&nbsp;"), Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)) & "</option>")
                End If
            Next
        End If

        ' show the add row only if at least one card, of a card type, is permitted to be added to the customer's account
        If (AddCardBuf.Length > 0) AndAlso (Logix.UserRoles.EditCustomerIdData) Then
            Send("  <tr>")
            Send("    <td align=""center""><input type=""button"" class=""add"" name=""add"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""addCard(" & CustomerPK & "," & CardPK & ");"" /></td>")
            Send("    <td><input type=""text"" id=""NewExtCardID"" name=""NewExtCardID"" value="""" maxlength=""" & maxLength & """ /></td>")
            Send("    <td>")
            Send("      <select id=""NewCardTypeID"" name=""NewCardTypeID"">")
            Send(AddCardBuf.ToString)
            Send("      </select>")
            Send("    </td>")
            Send("    <td>")
            Send("      <select id=""NewCardStatusID"" name=""NewCardStatusID"">")
            MyCommon.QueryStr = "select CardStatusID, Description, PhraseID from CardStatus with (NoLock);"
            rst2 = MyCommon.LXS_Select
            If rst2.Rows.Count > 0 Then
                For Each row In rst2.Rows
                    If (MyCommon.NZ(row.Item("CardStatusID"), 0) = 6 AndAlso Not (MyCommon.IsEngineInstalled(2) OrElse MyCommon.IsEngineInstalled(0))) Then
                        Continue For
                    End If
                    Send("        <option value=""" & MyCommon.NZ(row.Item("CardStatusID"), 0) & """>" & IIf(IsDBNull(row.Item("PhraseID")), MyCommon.NZ(row.Item("Description"), "&nbsp;"), Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)) & "</option>")
                Next
            End If
            Send("      </select>")
            Send("    </td>")
            Send("  </tr>")
        End If

        Send("</tbody>")
        Send("</table>")
        Send("<br class=""half"" />")

        Send("</div> <!-- id=cards -->")

    End Sub


    '-------------------------------------------------------------------------------------------------------------  


    Sub Send_Identifiers_Box(ByVal CustomerPK As Long, ByRef Common As Copient.CommonInc, ByRef Logix As Copient.LogixInc)

        Dim FormData As String
        Dim RawRequest As String
        Dim RawURI As String = ""
        Dim HostURI As String = ""
        Dim TargetAddress As String
        Dim dst As DataTable
        Dim ConnInc As New Copient.ConnectorInc()

        'The code for the contents of this box lives in customer.general.identifiers.aspx.  

        RawRequest = Get_Raw_RequestData(Request.InputStream)
        RawRequest = Replace(RawRequest, "+", "%2b")
        Send("<!-- Raw data:")
        Send(RawRequest)
        Send("-->")

        Common.QueryStr = "select isnull(HTTP_RootURI, '') as HTTP_RootURI from Integrations where IntegrationID=1;"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            HostURI = dst.Rows(0).Item("HTTP_RootURI")
        End If
        dst = Nothing

        HostURI = Trim(HostURI)
        If HostURI = "" Then
            Send("HTTP_RootURI for PrefMan integration is not set in LogixRT.Integrations")
        Else
            If Not (Right(HostURI, 1) = "/") Then HostURI = HostURI & "/"
            Send("<!-- HostURI=" & HostURI & " -->")
            If Not (Right(HostURI, 1) = "/") Then HostURI = HostURI & "/"
            Send("<!-- HostURI=" & HostURI & " -->")
            TargetAddress = HostURI & "UI/customer.general.identifiersbox.aspx"
            Send("<! -- TargetAddress=" & TargetAddress & " -->")

            'open the display box
            Send("<div class=""box"" id=""identifiers"" " & IIf(Logix.UserRoles.AccessCustomerIdData_Cards, "", " style=""display:none;""") & ">")
            Send("<h2><span>")
            Sendb(Copient.PhraseLib.Lookup("term.identifiers", LanguageID))
            Send("</span></h2>")

            'make the call over to PrefMan to get the contents of the box
            FormData = "AuthToken=" & HttpUtility.UrlEncode(Request.Cookies("AuthToken").Value) & "&ParentURI=customer-general.aspx&customerpk=" & CustomerPK & _
                "&IsLogixCall=1&CanUserEditCustomerInLogix=" & If(Logix.UserRoles.EditCustomerIdData, 1, 0) & _
                "&CanUserDeleteCustomerInLogix=" & If(Logix.UserRoles.DeleteCustomerCard, 1, 0)
            ' If EPM is Installed and System Option 116 is Set in Logix, then AlternateID will be set read-only on customer-general.aspx
            If Common.NZ(Integer.Parse(Common.Fetch_SystemOption(116)), -1) = 1 Then
                FormData = FormData & "&bReadOnlyAlternateID=True"
            End If
            FormData = FormData & "&" & RawRequest
            Send(ConnInc.Retrieve_HttpResponse(TargetAddress, FormData))
        End If
        'close the display box
        Send("</div> <!-- id=identifiers -->")

    End Sub

    Function ReplaceSpecialChar(ByVal inputString As String) As String
        Dim strReplace As String
        strReplace = inputString.Replace("&amp;", "&").Replace("&apos;", "'").Replace("&quot;", """").Replace("&lt;", "<").Replace("&gt;", ">")
        strReplace = strReplace.Replace("&", "&amp;").Replace("'", "&apos;").Replace("""", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;")
        Return strReplace
    End Function
    Function IsSavePermitted(ByVal customerPk As Integer,ByVal cardTypeId As Integer,ByVal targetCardPK As Integer,ByRef MyCommon As Copient.CommonInc) As Boolean
        Dim dt As DataTable
        Dim OnePerCustomer As Boolean=False
        If Mycommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
        'check if onepercustomer option is set for cardtypeid
        MyCommon.QueryStr="SELECT OnePerCustomer from cardtypes where cardtypeid="&cardTypeId
        dt=MyCommon.LXS_Select
        MyCommon.Close_LogixXS()
        If dt IsNot Nothing AndAlso dt.Rows.Count>0 Then
            OnePerCustomer=dt.Rows(0).Item("OnePerCustomer")
        End If
        If OnePerCustomer Then
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MyCommon.QueryStr="select COUNT(*) as cnt from CardIDs where CustomerPK=@CustomerPK  and CardTypeID=@CardTypeId and CardStatusID in (1,6) and CardPK <> @CardPk"
            MyCommon.DBParameters.Add("@CustomerPK",SqlDbType.BigInt).Value=customerPk
            MyCommon.DBParameters.Add("@CardTypeId",SqlDbType.Int).Value=cardTypeId
            MyCommon.DBParameters.Add("@CardPk",SqlDbType.BigInt).Value=targetCardPK
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
            If(dt IsNot Nothing AndAlso dt.Rows.Count>0) Then
                If(dt.Rows(0).Item("cnt")=1) Then
                    Return False
                End If
            End If
        End If
        Return True
    End Function

    Function GetShortDatePattern(ByRef MyCommon As Copient.CommonInc) As String
        Dim Pattern As String = ""

        If MyCommon IsNot Nothing AndAlso MyCommon.GetAdminUser.Culture IsNot Nothing Then
            Pattern = MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern
        End If

        Return Pattern
    End Function

</script>
<script>
function showMe (box) {

    var chboxs = document.getElementsByName("c0");
    var vis = "none";

    for(var i=0;i<chboxs.length;i++) { 
        if(chboxs[i].checked){
         vis = "block";
            break;
        }
    }
    document.getElementById(box).style.display = vis;


}
function bitToggle(place){
        var toggle = (1<<place);  
        var  num = document.getElementById("hiddenDigitalReceipt").value;
		num=num^toggle;
		document.getElementById("hiddenDigitalReceipt").value=num;
		
        
        
    }
	
function paperCheck(){
	var check=document.getElementById("hiddenPaperReceipt").value
	
		if((check=="True")){
			check="False";}
			
		else if(check=="False"){
			check="True";}
			
		document.getElementById("hiddenPaperReceipt").value=check;
		
		}	
</script>	

<%
done:
    Send_BodyEnd()
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    MyCommon = Nothing
    Logix = Nothing
%>