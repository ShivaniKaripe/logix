<%@ Page Language="vb" Debug="true" CodeFile="..\..\logix\LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-cashier-inquiry.aspx 
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
  Const iNumFields As Integer = 24
  Dim bDisplay(iNumFields) As Boolean
  Dim bEdit(iNumFields) As Boolean
  Dim sLogDescription(iNumFields) As String
  Dim sOld As String
  Dim sNew As String
  Dim iOld As Integer
  Dim iNew As Integer
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim MyCryptLib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim MyAltID As New Copient.AlternateID
  Dim AltIDResponse As New Copient.AlternateID.CreateUpdateResponse
  Dim rstResults As DataTable = Nothing
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim dt As DataTable
  
  'Card identifiers
  Dim CardPK As Long = 0
  Dim ExtCardID As String = ""
  Dim CardTypeID As Integer = 0
  
  'Customer identifiers
  Dim CustomerPK As Long = 0
  Dim CustomerTypeID As Integer = 0
  Dim Prefix As String = ""
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim Suffix As String = ""
  Dim FullName As String = ""
  Dim Address As String = ""
  Dim City As String = ""
  Dim State As String = ""
  Dim Zip As String = ""
  Dim Country As String = ""
  Dim FullAddress As String = ""
  Dim Phone1 As String = ""
  Dim FullPhone As String = ""
  Dim Email As String = ""
  Dim Household As String = ""
  Dim Password As String = ""
  Dim DOB_month As String = ""
  Dim DOB_day As String = ""
  Dim DOB_year As String = ""
  Dim DOB As String = ""
  Dim AltIDValue As String = ""
  Dim AltIDVerifier As String = ""
  Dim EmployeeID As String = ""
  Dim Employee As Integer = 0
  Dim TestCard As Integer = 0
  Dim CustomerStatusID As Integer = 0
  Dim AirmileMemberID As String = ""
  Dim HasAirmileMemberID As Boolean = False
  Dim Enrollment_month As String = ""
  Dim Enrollment_day As String = ""
  Dim Enrollment_year As String = ""
  Dim EnrollmentDate As String = ""
  
  Dim i As Integer = 0
  Dim j As Integer = 0
  Dim Shaded As String = " class=""shaded"""
  Dim Edit As Boolean = False
  Dim UnknownPhrase As String = ""
  Dim SavedCardStatus As Integer = 0
  Dim AltIDField As Object = Nothing
  Dim IDVerifierField As Object = Nothing
  Dim AltIDTable As String = ""
  Dim AltIDCol As String = ""
  Dim AltIDVerCol As String = ""
  Dim bUpdateAlt As Boolean = True
  Dim NewAltID As String = ""
  Dim SaveFailed As Boolean = False
  Dim BannerID As Integer = 0
  Dim DateValid As Boolean = False
  Dim NullifyAltID As Boolean = False
  
  'Default urls for links from this page
  
  Dim MyLookup As New Copient.CustomerLookup()
  Dim Customers(-1) As Copient.Customer
  Dim Cust As New Copient.Customer
  Dim CustExt As New Copient.CustomerExt
  Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
  Dim TempDate As Date
  Dim Fields As New Copient.CommonInc.ActivityLogFields
  
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim sDisabled As String
  Dim sGUID As String = ""
  Dim sCashierID As String = ""
  Dim sStoreId As String = ""
  
  Dim iPrefix As Integer = 0
  Dim iFirstName As Integer = 1
  Dim iMiddleName As Integer = 2
  Dim iLastName As Integer = 3
  Dim iAltID As Integer = 4
  Dim iVerifier As Integer = 5
  Dim iSuffix As Integer = 6
  Dim iEmployee As Integer = 7
  Dim iEmployeeID As Integer = 8
  Dim iTestCard As Integer = 9
  Dim iBanner As Integer = 10
  Dim iStatus As Integer = 11
  Dim iAddress As Integer = 12
  Dim iCity As Integer = 13
  Dim iState As Integer = 14
  Dim iZip As Integer = 15
  Dim iCountry As Integer = 16
  Dim iPhone As Integer = 17
  Dim iMobilePhone As Integer = 18
  Dim iEmail As Integer = 19
  Dim iDOB As Integer = 20
  Dim iPassword As Integer = 21
  Dim iAirMile As Integer = 22
  Dim iEnrollmentDate As Integer = 23
  
  Dim bContinue As Boolean = True
 
  'AdminUserID = Verify_AdminUser(Logix)
  AdminUserID = 1
  LanguageID = 1
  
  Response.Expires = 0
  MyCommon.AppName = "CM-cashier-inquiry.aspx"
  
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()

  MyLookup.SetAdminUserID(AdminUserID)
  MyLookup.SetLanguageID(LanguageID)
  
  UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  
  AltIDCol = ParseTableCol(MyCommon.Fetch_SystemOption(60))
  AltIDTable = ParseTable(MyCommon.Fetch_SystemOption(60))
  AltIDVerCol = ParseTableCol(MyCommon.Fetch_SystemOption(61))
  
  'Set session to nothing, just to be sure
  Session.Add("extraLink", "")
  
  ' Exit?
  If (Request.QueryString("exit") <> "") Then
    infoMessage = "Please return to POS!"
    bContinue = False
  End If
  
  ' CM and feature (CM option 35) is turned on?
  If bContinue Then
    If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1") Then
      bContinue = True
    Else
      infoMessage = "Cashier customer inquiry is disabled at host, please return to POS!"
      bContinue = False
    End If
  End If
  
  ' Valid GUID?
  If bContinue Then
    sGUID = MyCommon.NZ(Request.QueryString("GUID"), "")
    If Not IsValidGUID(sGUID, MyCommon) Then
      infoMessage = "Invalid GUID, please return to POS!"
      bContinue = False
    End If
  End If
  
  ' Cashier specified?
  If bContinue Then
    sCashierID = MyCommon.NZ(Request.QueryString("CashierID"), "")
    If sCashierID.Length = 0 Then
      infoMessage = "Invalid Cashier ID, please return to POS!"
      bContinue = False
    End If
  End If
  
  ' Store specified?
  If bContinue Then
    sStoreId = MyCommon.NZ(Request.QueryString("StoreID"), "")
    If sStoreId.Length = 0 Then
      infoMessage = "Invalid Store ID, please return to POS!"
      bContinue = False
    End If
  End If

  If bContinue Then
    For i = 0 To iNumFields - 1
      MyCommon.QueryStr = "select Display, AllowEdit from CM_Cashier_Inquiry_Options with (NoLock) where FieldID=" & i & ";"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        bDisplay(i) = MyCommon.NZ(dt.Rows(0).Item(0), False)
        If (bDisplay(i)) Then
          bEdit(i) = MyCommon.NZ(dt.Rows(0).Item(1), False)
        Else
          bEdit(i) = False
        End If
      Else
        bDisplay(i) = False
        bEdit(i) = False
      End If
      sLogDescription(i) = ""
    Next
  End If
  
    If bContinue Then
        Try
        If (Request.QueryString("save") = "Save") Then
            'Saving customer information; first setup the page so it draws correctly
            Dim StrCustomerPk As String = Request.QueryString("CustomerPK").Trim()
            If (StrCustomerPk <> "") Then
                MyCommon.QueryStr = "select CustomerPK, CustomerTypeID, HHPK from Customers with (NoLock) where CustomerPK=@CustomerPK"
                MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = Convert.ToInt64(StrCustomerPk)
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
                If dt.Rows.Count > 0 Then
                    ReDim Customers(0)
                    Customers(0) = MyLookup.FindCustomerInfo(MyCommon.Extract_Val(Request.QueryString("CustomerPK")), ReturnCode)
                    If (Not Customers(0) Is Nothing) Then
                        CustomerPK = Customers(0).GetCustomerPK
                        SavedCardStatus = Customers(0).GetCardStatusID
            
                        'Set and save the customer's information
                        If bEdit(iPrefix) Then
                            sNew = Request.QueryString("Prefix").Trim
                            sOld = Customers(0).GetPrefix.Trim
                            Customers(0).SetPrefix(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iPrefix) = BuildLogDescription("Prefix", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iFirstName) Then
                            sNew = Request.QueryString("FirstName").Trim
                            sOld = Customers(0).GetFirstName.Trim
                            Customers(0).SetFirstName(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iFirstName) = BuildLogDescription("First Name", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iMiddleName) Then
                            sNew = Request.QueryString("MiddleName").Trim
                            sOld = Customers(0).GetMiddleName.Trim
                            Customers(0).SetMiddleName(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iMiddleName) = BuildLogDescription("Middle Name", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iLastName) Then
                            sNew = Request.QueryString("LastName").Trim
                            sOld = Customers(0).GetLastName.Trim
                            Customers(0).SetLastName(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iLastName) = BuildLogDescription("Last Name", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iSuffix) Then
                            sNew = Request.QueryString("Suffix").Trim
                            sOld = Customers(0).GetSuffix.Trim
                            Customers(0).SetSuffix(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iSuffix) = BuildLogDescription("Suffix", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iStatus) Then
                            sNew = Request.QueryString("CustomerStatusID").Trim
                            Integer.TryParse(sNew, iNew)
                            sOld = Customers(0).GetCustomerStatusID.ToString
                            Integer.TryParse(sOld, iOld)
                            Customers(0).SetCustomerStatusID(iNew)
                            If iOld <> iNew Then
                                sLogDescription(iStatus) = BuildLogDescription("Customer Status", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iAddress) Then
                            sNew = Request.QueryString("Address").Trim
                            sOld = Customers(0).GetGeneralInfo.GetAddress.Trim
                            Customers(0).GetGeneralInfo.SetAddress(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iAddress) = BuildLogDescription("Address", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iCity) Then
                            sNew = Request.QueryString("City").Trim
                            sOld = Customers(0).GetGeneralInfo.GetCity.Trim
                            Customers(0).GetGeneralInfo.SetCity(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iCity) = BuildLogDescription("City", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iState) Then
                            sNew = Request.QueryString("State").Trim
                            sOld = Customers(0).GetGeneralInfo.GetState.Trim
                            Customers(0).GetGeneralInfo.SetState(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iState) = BuildLogDescription("State", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iZip) Then
                            sNew = Request.QueryString("Zip").Trim
                            sOld = Customers(0).GetGeneralInfo.GetZip.Trim
                            Customers(0).GetGeneralInfo.SetZip(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iZip) = BuildLogDescription("Zip", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iCountry) Then
                            sNew = Request.QueryString("Country").Trim
                            sOld = Customers(0).GetGeneralInfo.GetCountry.Trim
                            Customers(0).GetGeneralInfo.SetCountry(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iCountry) = BuildLogDescription("Country", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iPhone) Then
                            Phone1 = MyCommon.Parse_Quotes(Request.QueryString("Phone1")).Trim
                            sNew = Phone1
                            sOld = Customers(0).GetGeneralInfo.GetPhone.Trim
                            Customers(0).GetGeneralInfo.SetPhone(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iPhone) = BuildLogDescription("Phone", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iMobilePhone) Then
                            Phone1 = MyCommon.Parse_Quotes(Request.QueryString("MobilePhone1")).Trim
                            sNew = Phone1
                            sOld = Customers(0).GetGeneralInfo.GetMobilePhone.Trim
                            Customers(0).GetGeneralInfo.SetMobilePhone(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iMobilePhone) = BuildLogDescription("MobilePhone", sOld, sNew)
                            End If
                        End If
          
                        If bEdit(iEmail) Then
                            sNew = MyCommon.NZ(Request.QueryString("Email"), "").Trim
                            sOld = Customers(0).GetGeneralInfo.GetEmail.Trim
                            Customers(0).GetGeneralInfo.SetEmail(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iEmail) = BuildLogDescription("Email", sOld, sNew)
                            End If
                        End If
          
                        Customers(0).SetPassword(MyCryptLib.SQL_StringDecrypt(Customers(0).GetPassword))
                        If bEdit(iPassword) Then
                            sNew = MyCommon.NZ(Request.QueryString("Password"), "").Trim
                            sOld = Customers(0).GetPassword
                            Customers(0).SetPassword(sNew)
                            If sOld <> sNew Then
                                sLogDescription(iPassword) = "Modified Password"
                            End If
                        End If
          
                        If bEdit(iAltID) Then
                            AltIDValue = MyCommon.NZ(Request.QueryString("AltIDValue"), "").Trim
                            sOld = Customers(0).GetAltID
                            Customers(0).SetAltID(AltIDValue)
                            If sOld <> AltIDValue Then
                                sLogDescription(iAltID) = BuildLogDescription("AltID", sOld, AltIDValue)
                            End If
                        End If
          
                        If bEdit(iVerifier) Then
                            AltIDVerifier = MyCommon.NZ(Request.QueryString("AltIDVerifier"), "").Trim
                            sOld = Customers(0).GetAltIdVerifier.Trim
                            Customers(0).SetAltIdVerifier(AltIDVerifier)
                            If sOld <> AltIDVerifier Then
                                sLogDescription(iVerifier) = BuildLogDescription("AltIDVerifier", sOld, AltIDVerifier)
                            End If
                        End If
          
                        If bEdit(iAirMile) Then
                            AirmileMemberID = MyCommon.Parse_Quotes(MyCommon.NZ(Request.QueryString("AirmileMemberID"), "")).Trim
                            sOld = Customers(0).GetGeneralInfo.GetAirmileMemberID.Trim
                            Customers(0).GetGeneralInfo.SetAirmileMemberID(AirmileMemberID)
                            If sOld <> AirmileMemberID Then
                                sLogDescription(iAirMile) = BuildLogDescription("AirmileMemberID", sOld, AirmileMemberID)
                            End If
                        End If
          
                        If bEdit(iBanner) Then
                            sNew = MyCommon.NZ(Request.QueryString("BannerID"), "").Trim
                            Integer.TryParse(sNew, iNew)
                            iOld = Customers(0).GetBannerID
                            Customers(0).SetBannerID(iNew)
                            If iOld <> iNew Then
                                sLogDescription(iEmail) = BuildLogDescription("BannerID", iOld.ToString, sNew)
                            End If
                        End If

                        If bEdit(iEmployee) Then
                            sOld = Customers(0).GetEmployee.ToString
                            If (MyCommon.NZ(Request.QueryString("Employee"), "") = "on") Then
                                Employee = 1
                                sNew = "True"
                                Customers(0).SetEmployee(True)
                            Else
                                Employee = 0
                                sNew = "False"
                                Customers(0).SetEmployee(False)
                            End If
                            If sOld <> sNew Then
                                sLogDescription(iEmployee) = BuildLogDescription("Employee", sOld, sNew)
                            End If
                        End If

                        If bEdit(iEmployeeID) Then
                            sOld = Customers(0).GetEmployeeID
                            EmployeeID = MyCommon.NZ(Request.QueryString("EmployeeID"), "")
                            Customers(0).SetEmployeeID(EmployeeID)
                            If sOld <> EmployeeID Then
                                sLogDescription(iEmployeeID) = BuildLogDescription("EmployeeID", sOld, EmployeeID)
                            End If
                        End If

                        If bEdit(iTestCard) Then
                            sOld = Customers(0).GetTestCard.ToString
                            If (MyCommon.NZ(Request.QueryString("TestCard"), "") = "on") Then
                                TestCard = 1
                                sNew = "True"
                                Customers(0).SetTestCard(True)
                            Else
                                TestCard = 0
                                sNew = "False"
                                Customers(0).SetTestCard(False)
                            End If
                            If sOld <> sNew Then
                                sLogDescription(iEmployee) = BuildLogDescription("TestCard", sOld, sNew)
                            End If
                        End If

                        'Format the date of birth
                        If bEdit(iDOB) Then
                            sOld = Customers(0).GetGeneralInfo.GetDateOfBirth.ToString("MM/dd/yyyy")
                            DOB_month = Request.QueryString("dob1")
                            DOB_day = Request.QueryString("dob2")
                            DOB_year = Request.QueryString("dob3")
                            DOB = ""
                            If (DOB_month = "" And DOB_day = "" And DOB_year = "") Then
                                'Allows a null to be set for the DOB when there is nothing saved in the DOB fields
                                DateValid = True
                                Customers(0).GetGeneralInfo.SetDateOfBirth(Nothing)
                            ElseIf (ValidateMonth(DOB_month) = False Or ValidateDay(DOB_day) = False Or ValidateYear(DOB_year) = False) Then
                                'If any part of the DOB is invalid then give the proper infomessage
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
                                DateValid = Date.TryParse(DOB, TempDate)
                                If DateValid Then
                                    Customers(0).GetGeneralInfo.SetDateOfBirth(TempDate)
                                    If sOld <> DOB Then
                                        sLogDescription(iEmployee) = BuildLogDescription("Date of Birth", sOld, DOB)
                                    End If
                                Else
                                    infoMessage = "Invalid date format for date of birth."
                                    SaveFailed = True
                                End If
                            End If
                        Else
                            DateValid = True
                        End If
            
                        'Format the enrollment date
                        If bEdit(iEnrollmentDate) Then
                            sOld = Customers(0).GetEnrollmentDate.ToString("MM/dd/yyyy")
                            Enrollment_month = Request.QueryString("enrollment1")
                            Enrollment_day = Request.QueryString("enrollment2")
                            Enrollment_year = Request.QueryString("enrollment3")
                            EnrollmentDate = ""
                            If (Enrollment_month = "" And Enrollment_day = "" And Enrollment_year = "") Then
                                'Allows a null to be set for the Enrollment Date when there is nothing saved in the Enrollment fields
                                DateValid = True
                                Customers(0).SetEnrollmentDate(Nothing)
                            ElseIf (ValidateMonth(Enrollment_month) = False Or ValidateDay(Enrollment_day) = False Or ValidateYear(Enrollment_year) = False) Then
                                'If any part of the Enrollment Date is invalid then give the proper infomessage
                                Dim TempMessage As String = ""
                                If (ValidateMonth(Enrollment_month) = False) Then
                                    TempMessage = Copient.PhraseLib.Lookup("term.enrollmentdate", LanguageID) & " - " & Copient.PhraseLib.Lookup("customer-inquiry.invalid-date", LanguageID) & "<br />"
                                ElseIf (ValidateDay(Enrollment_day) = False) Then
                                    TempMessage = Copient.PhraseLib.Lookup("term.enrollmentdate", LanguageID) & " - " & Copient.PhraseLib.Lookup("customer-inquiry.invalid-date", LanguageID) & "<br />"
                                ElseIf (ValidateYear(Enrollment_year) = False) Then
                                    TempMessage = Copient.PhraseLib.Lookup("term.enrollmentdate", LanguageID) & " - " & Copient.PhraseLib.Lookup("customer-inquiry.invalid-date", LanguageID) & "<br />"
                                End If
                                If TempMessage <> "" Then DateValid = False
                                SaveFailed = True
                                infoMessage = TempMessage
                            Else
                                EnrollmentDate = Enrollment_month.Trim.PadLeft(2, "0") & "/" & Enrollment_day.Trim.PadLeft(2, "0") & "/" & Enrollment_year.Trim.PadLeft(4, "0")
                                DateValid = Date.TryParse(EnrollmentDate, TempDate)
                                If DateValid Then
                                    Customers(0).SetEnrollmentDate(TempDate)
                                    If sOld <> EnrollmentDate Then
                                        sLogDescription(iEmployee) = BuildLogDescription("Enrollment Date", sOld, EnrollmentDate)
                                    End If
                                Else
                                    infoMessage = "Invalid date format for enrollment date."
                                    SaveFailed = True
                                End If
                            End If
                        Else
                            DateValid = True
                        End If
      
                        'Handle updates to Alternate Identifier
                        If bEdit(iAltID) Then
                            If (infoMessage.Trim = "" AndAlso AltIDCol.Trim <> "") Then
                                NewAltID = GetNewAltID(AltIDCol, AltIDField)
                                If (AltIDCol.Trim <> "" And NewAltID.Trim <> "") Then
                                    AltIDResponse = MyAltID.UpdateCustomerAltID(CustomerPK, NewAltID, BannerID)
                                    Select Case AltIDResponse
                                        Case Copient.AlternateID.CreateUpdateResponse.ALTIDINUSE
                                            infoMessage = "Changes were not saved.  The unique Alternate Identifier is already in use by another customer " & _
                                                          "<br />(" & AltIDField & " = " & NewAltID & ")"
                                            SaveFailed = True
                                        Case Copient.AlternateID.CreateUpdateResponse.MEMBERNOTFOUND
                                            infoMessage = "Customer not found."
                                            SaveFailed = True
                                        Case Copient.AlternateID.CreateUpdateResponse.ERROR_APPLICATION
                                            infoMessage = "Error encountered during Alternate Identifier update: " & MyAltID.ErrorMessage
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
                        End If
      
                        If (infoMessage.Trim = "" AndAlso DateValid) Then
                            If MyLookup.SaveCustomerInfo(Customers(0), ReturnCode) Then
                                'If an empty string is sent for the AltID then NULL it out.
                                If NullifyAltID Then
                                    MyCommon.QueryStr = "update " & AltIDTable & " with (RowLock) set " & AltIDCol & "=NULL where CustomerPK=" & CustomerPK & ";"
                                    MyCommon.LXS_Execute()
                                    MyCommon.QueryStr = "update Customers with (RowLock) set CPEStoreSendFlag=1 " & _
                                                        "where CustomerPK=" & CustomerPK & ";"
                                    MyCommon.LXS_Execute()
                                End If
                            Else
                                infoMessage = "Error encountered: " & ReturnCode
                            End If
                        End If
          
                        Edit = True
                        ExtCardID = MyCommon.Extract_Val(Request.QueryString("ExtCardID"))
            
                        If Not SaveFailed Then
                            Fields.ActivityTypeID = 25
                            Fields.ActivitySubTypeID = 2
                            Fields.LinkID = CustomerPK
                            Fields.AdminUserID = 1
                            Fields.LinkID3 = sCashierID
                            Fields.LinkID4 = sStoreId
                            Fields.LinkID5 = ExtCardID
                            Fields.LinkID6 = CardTypeID
                            For i = 0 To iNumFields - 1
                                If sLogDescription(i) <> "" Then
                                    Fields.Description = sLogDescription(i)
                                    MyCommon.Activity_Log3(Fields)
                                End If
                            Next
                        End If
                    Else
                        infoMessage = "Customer NOT found!"
                    End If
                Else
                    infoMessage = "Customer NOT found!"
                End If
            Else
                infoMessage = "Customer NOT found!"
            End If
        Else
            If MyCommon.NZ(Request.QueryString("CustomerPK"), "") = "" Then
                If (MyCommon.NZ(Request.QueryString("ExtCardID"), "") = "") Then
                    infoMessage = "Invalid card ID!"
                    CustomerPK = 0
                Else
                    ExtCardID = Request.QueryString("ExtCardID")
                    If (MyCommon.NZ(Request.QueryString("CardTypeID"), "") = "") Then
                        ' default to CardTypeId = 0
                        CardTypeID = 0
                    Else
                        CardTypeID = MyCommon.Extract_Val(MyCommon.NZ(Request.QueryString("CardTypeID"), "0"))
                    End If
                    CustomerPK = MyLookup.GetCustomerPK(ExtCardID, CardTypeID, ReturnCode)
                    If CustomerPK = 0 Then
                        infoMessage = "Customer with CardID = " & ExtCardID & " and Card Type = " & CardTypeID & " was not found!"
                    End If
                End If
                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
            Else
                CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
                ExtCardID = Request.QueryString("ExtCardID")
                If CustomerPK = 0 Then
                    infoMessage = "Invalid CustomerPK!"
                End If
                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
            End If
  
            If (CustomerPK > 0) Then
                ReDim Customers(0)
                Customers(0) = MyLookup.FindCustomerInfo(CustomerPK, ReturnCode)
                If (Customers.Length = 1) Then
                    CustomerPK = Customers(0).GetCustomerPK
                    CustomerTypeID = Customers(0).GetCustomerTypeID
                    FirstName = Customers(0).GetFirstName
                    MiddleName = Customers(0).GetMiddleName
                    LastName = Customers(0).GetLastName
                    EmployeeID = Customers(0).GetEmployeeID
                ElseIf Customers.Length > 1 Then
                    infoMessage = "" & Copient.PhraseLib.Lookup("customer.multiplefound", LanguageID) & ""
                Else
                    infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
                End If
                Edit = True
            End If
        End If
         Catch ex As Exception
            infoMessage = "Customer NOT found!"
        End Try
    End If
  
  'Load the customer
  If CustomerPK > 0 Then
    Cust = MyLookup.FindCustomerInfo(CustomerPK, ReturnCode)
  End If
   
  Send_HeadBegin("term.customer", "term.general", MyCommon.Extract_Val(ExtCardID))
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links_Cashier(MyCommon, Handheld)
  Send_Scripts()
  Send_HeadEnd()
  
  Send_BodyBegin(1)
  
  Send("<div id=""bar"">")
  Send("  <span id=""skip""><a href=""#skiptabs"">▼</a></span>")
  Send("  <span id=""time"" title=""" & DateTime.Now.ToString("HH:mm:ss, G\MT zzz") & """>" & DateTime.Now.ToString("HH:mm") & " | </span>")
  Send("  <span id=""date"">" & DateTime.Now.ToString("dddd, d MMMM yyyy") & "</span>")
  Send("</div>")
  Send("")
  
  Send_Logos()
  
  If infoMessage = "" Then
    If Request.QueryString("infomessage") <> "" Then
      infoMessage = Request.QueryString("infomessage")
    End If
  End If
%>

<script type="text/javascript">
function ChangeEmployeeID() { 
  var emp = document.getElementById("EmployeeID");
  var foremp = document.getElementById("forEmployeeID");
  var empCheck = document.getElementById("Employee");
  
  if(empCheck.checked == false) {
    foremp.style.display = "none";
    emp.style.display = "none";
  } else {
    foremp.style.display = "block";
    emp.style.display = "block";
  }
}


function isValidEntry() {
  var retVal = true;
  
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
    
  // validate DOB date
  for (var i=1; i <= 3 && retVal; i++) {
    retVal = retVal && isValidDatePart("dob", i);
  }
  
  // validate Enrollment date
  for (var i=1; i <= 3 && retVal; i++) {
    retVal = retVal && isValidDatePart("enrollment", i);
  }

  return retVal;
}

//function isValidPhonePart(prefix, partNum) {
//  var retVal = true;
//  var elemPart = document.getElementById(prefix + partNum);
//  
//  if (elemPart != null) {
//    if (partNum == 1 && elemPart.value!="" && elemPart.value.length != 3) { 
//      alert('Area code should be either blank or contain 3 digits');
//      retVal = false;
//    }
//    if (partNum == 1 && (isNaN(elemPart.value) || parseInt(elemPart.value) < 0) ) {
//      alert('Area code should be either blank or contain 3 digits');
//      retVal = false;
//    }
//    if (partNum == 2 && elemPart.value != "" && (elemPart.value.length != 3 || isNaN(elemPart.value) || parseInt(elemPart.value) < 0) ) {
//      alert('Prefix for phone number must be 3 digits.');
//      retVal = false;
//    }
//    if (partNum == 3  && elemPart.value != "" && (elemPart.value.length != 4 || isNaN(elemPart.value) || parseInt(elemPart.value) < 0) ) {
//      alert('The final part of the phone number must be 4 digits.');
//      retVal = false;
//    }
//    if (!retVal) {
//      elemPart.focus();
//      elemPart.select();
//    }
//  }
//  
//  return retVal;
//}

//function isValidPhoneCombo(prefix) {
//  var retVal = false;
//  var elemP1 = document.getElementById(prefix + "1");
//  var elemP2 = document.getElementById(prefix + "2");
//  var elemP3 = document.getElementById(prefix + "3");
//  
//  if (elemP1 != null && elemP2 != null && elemP3 != null) {
//    if (elemP1.value != "" || elemP2.value != "" || elemP3.value != "") {
//      // validate to acceptable formats (xxx) xxx-xxxx and xxx-xxxx
//      retVal = (elemP1.value != "" && elemP1.value.length==3 && elemP2.value != "" && elemP2.value.length==3 && elemP3.value !="" && elemP3.value.length==4);
//      retVal = retVal || (elemP1.value == "" && elemP2.value != "" && elemP2.value.length==3 && elemP3.value !="" && elemP3.value.length==4 );
//    } else {
//      // all phone parts are blank, so no phone number was provided to validate
//      retVal = true;
//    }
//  }
//  if (!retVal) {
//    alert("Phone number should be in either 7 or 10 digit phone number format.");
//  }
//  
//  return retVal;
//}

function isValidDatePart(prefix, partNum) {
  var retVal = true;
  var elemPart = document.getElementById(prefix + partNum);
  
  if (elemPart != null) {
    if (elemPart.value != "" && isNaN(elemPart.value)) {
      alert('Date must be a number.');
      retVal = false;
      elemPart.focus();
      elemPart.select();
    }
  }
  
  return retVal;  
}
//-->
</script>

<script type="text/javascript" src="../javascript/jquery.min.js"></script>

<script type="text/javascript" src="../javascript/thickbox.js"></script>

<form id="mainform" name="mainform" action="CM-cashier-inquiry.aspx" onsubmit="return isValidEntry();">
  <input type="hidden" name="altid" value="<%Sendb(AltIDCol) %>" />
  <input type="hidden" name="verifier" value="<%Sendb(AltIDVerCol) %>" />
  <input type="hidden" id="CustomerPK" name="CustomerPK" value="<%Sendb(CustomerPK)%>" />
  <input type="hidden" id="ExtCardID" name="ExtCardID" value="<%Sendb(ExtCardID)%>" />
  <input type="hidden" id="GUID" name="GUID" value="<%Sendb(sGUID)%>" />
  <input type="hidden" id="CashierID" name="CashierID" value="<%Sendb(sCashierID)%>" />
  <input type="hidden" id="StoreID" name="StoreID" value="<%Sendb(sStoreID)%>" />
  <div id="intro">
    <h1 id="title">
      <%
        If Edit Then
          Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
          If SaveFailed Then
            MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & Cust.GetCustomerPK & ";"
            rst2 = MyCommon.LXS_Select
            If rst2.Rows.Count > 0 Then
              FullName = IIf(MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") <> "" AndAlso bDisplay(iPrefix), MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") & " ", "")
              FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") <> "" AndAlso bDisplay(iFirstName), MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") & " ", "")
              FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), "") <> "" AndAlso bDisplay(iMiddleName), Left(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
              FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("LastName"), "") <> "" AndAlso bDisplay(iLastName), MyCommon.NZ(rst2.Rows(0).Item("LastName"), ""), "")
              FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("Suffix"), "") <> "" AndAlso bDisplay(iSuffix), " " & MyCommon.NZ(rst2.Rows(0).Item("Suffix"), ""), "")
            Else
              FullName = IIf(Cust.GetPrefix <> "" AndAlso bDisplay(iPrefix), Cust.GetPrefix & " ", "")
              FullName &= IIf(Cust.GetFirstName <> "" AndAlso bDisplay(iFirstName), Cust.GetFirstName & " ", "")
              FullName &= IIf(Cust.GetMiddleName <> "" AndAlso bDisplay(iMiddleName), Left(Cust.GetMiddleName, 1) & ". ", "")
              FullName &= IIf(Cust.GetLastName <> "" AndAlso bDisplay(iLastName), Cust.GetLastName, "")
              FullName &= IIf(Cust.GetSuffix <> "" AndAlso bDisplay(iSuffix), " " & Cust.GetSuffix, "")
            End If
          Else
            FullName = IIf(Cust.GetPrefix <> "" AndAlso bDisplay(iPrefix), Cust.GetPrefix & " ", "")
            FullName &= IIf(Cust.GetFirstName <> "" AndAlso bDisplay(iFirstName), Cust.GetFirstName & " ", "")
            FullName &= IIf(Cust.GetMiddleName <> "" AndAlso bDisplay(iMiddleName), Left(Cust.GetMiddleName, 1) & ". ", "")
            FullName &= IIf(Cust.GetLastName <> "" AndAlso bDisplay(iLastName), Cust.GetLastName, "")
            FullName &= IIf(Cust.GetSuffix <> "" AndAlso bDisplay(iSuffix), " " & Cust.GetSuffix, "")
          End If
          If FullName <> "" Then
            Sendb(": " & MyCommon.TruncateString(FullName, 30))
          End If
        End If
      %>
    </h1>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br />")
      End If
    %>
    <div id="column">
      <%
        If (CustomerPK > 0) Then
          Send("<div class=""box"" id=""identity""" & IIf(Cust.GetCustomerPK = 0, " style=""display: none;""", "") & ">")
          If (Cust IsNot Nothing) Then
            Sendb("<h2><span>")
            Sendb(Copient.PhraseLib.Lookup("term.cardholder", LanguageID))
            Send("</span></h2>")
            If (Edit) Then
              Dim TempValue As String = ""
              Dim DOBParts() As String = {"", "", ""}
              Dim EnrollmentParts() As String = {"", "", ""}
              
              Send("<table style=""width:355px;float:left;position:relative;"" summary=""" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & " 1"">")
              
              'Name prefix
              Send("<tr" & IIf(bDisplay(iPrefix), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Prefix"">" & Copient.PhraseLib.Lookup("term.prefix", LanguageID) & GetDesignatorText("Prefix", AltIDCol, AltIDVerCol) & ":</label> </td>")
              Prefix = IIf(SaveFailed, Request.QueryString("Prefix"), MyCommon.NZ(Cust.GetPrefix, UnknownPhrase).Replace("""", "&quot;"))
              sDisabled = IIf(bEdit(iPrefix), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""Prefix"" name=""Prefix"" maxlength=""20"" value=""" & Prefix & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'First name
              Send("<tr" & IIf(bDisplay(iFirstName), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""FirstName"">" & Copient.PhraseLib.Lookup("term.firstname", LanguageID) & GetDesignatorText("FirstName", AltIDCol, AltIDVerCol) & ":</label> </td>")
              FirstName = IIf(SaveFailed, Request.QueryString("FirstName"), MyCommon.NZ(Cust.GetFirstName, UnknownPhrase).Replace("""", "&quot;"))
              sDisabled = IIf(bEdit(iFirstName), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""FirstName"" name=""FirstName"" maxlength=""50"" value=""" & FirstName & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Middle name
              Send("<tr" & IIf(bDisplay(iMiddleName), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""MiddleName"">" & Copient.PhraseLib.Lookup("term.middlename", LanguageID) & GetDesignatorText("MiddleName", AltIDCol, AltIDVerCol) & ":</label> </td>")
              MiddleName = IIf(SaveFailed, Request.QueryString("MiddleName"), MyCommon.NZ(Cust.GetMiddleName, UnknownPhrase).Replace("""", "&quot;"))
              sDisabled = IIf(bEdit(iMiddleName), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""MiddleName"" name=""MiddleName"" maxlength=""50"" value=""" & MiddleName & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Last name
              Send("<tr" & IIf(bDisplay(iLastName), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""LastName"">" & Copient.PhraseLib.Lookup("term.lastname", LanguageID) & GetDesignatorText("LastName", AltIDCol, AltIDVerCol) & ":</label> </td>")
              LastName = IIf(SaveFailed, Request.QueryString("LastName"), MyCommon.NZ(Cust.GetLastName, UnknownPhrase).Replace("""", "&quot;"))
              sDisabled = IIf(bEdit(iLastName), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""LastName"" name=""LastName"" maxlength=""50"" value=""" & LastName & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Name suffix
              Send("<tr" & IIf(bDisplay(iSuffix), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Suffix"">" & Copient.PhraseLib.Lookup("term.suffix", LanguageID) & GetDesignatorText("Suffix", AltIDCol, AltIDVerCol) & ":</label> </td>")
              Suffix = IIf(SaveFailed, Request.QueryString("Suffix"), MyCommon.NZ(Cust.GetSuffix, UnknownPhrase).Replace("""", "&quot;"))
              sDisabled = IIf(bEdit(iSuffix), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""Suffix"" name=""Suffix"" maxlength=""20"" value=""" & Suffix & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'AltID
              If (AltIDCol.ToUpper = "ALTID") Then
                Send("<tr" & IIf(bDisplay(iAltID), "", " style=""display:none;""") & ">")
                Send("  <td><label for=""AltIDValue"">" & Copient.PhraseLib.Lookup("term.alternateid", LanguageID) & ":</label> </td>")
                AltIDValue = IIf(SaveFailed, AltIDValue, MyCommon.NZ(Cust.GetAltID, "").Replace("""", "&quot;"))
                sDisabled = IIf(bEdit(iAltID), "", " disabled=""disabled""")
                Send("  <td><input type=""text"" class=""medium"" id=""AltIDValue"" name=""AltIDValue"" maxlength=""20"" value=""" & AltIDValue & """ " & sDisabled & " /></td>")
                Send("</tr>")
              End If
              
              'AltID verifier
              If (AltIDVerCol.ToUpper = "VERIFIER") Then
                Send("<tr" & IIf(bDisplay(iVerifier), "", " style=""display:none;""") & ">")
                Send("  <td><label for=""AltIDVerifier"">" & Copient.PhraseLib.Lookup("term.alternate-id-verifier", LanguageID) & ":</label> </td>")
                AltIDVerifier = IIf(SaveFailed, AltIDVerifier, MyCommon.NZ(Cust.GetAltIdVerifier, "").Replace("""", "&quot;"))
                sDisabled = IIf(bEdit(iVerifier), "", " disabled=""disabled""")
                Send("  <td><input type=""text"" class=""medium"" id=""AltIDVerifier"" name=""AltIDVerifier"" maxlength=""20"" value=""" & AltIDVerifier & """ " & sDisabled & " /></td>")
                Send("</tr>")
              End If
              
              'Employee-related
              If Not SaveFailed Then
                Employee = IIf(MyCommon.NZ(Cust.GetEmployee, False), 1, 0)
              End If
              Send("<tr" & IIf(bDisplay(iEmployee), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Employee"">" & Copient.PhraseLib.Lookup("term.employee", LanguageID) & ":</label> </td>")
              sDisabled = IIf(bEdit(iEmployee), "", " disabled=""disabled""")
              Send("  <td><input type=""checkbox"" id=""Employee"" name=""Employee"" onclick=""javascript:ChangeEmployeeID();""" & IIf(Employee = 1, " checked=""checked""", "") & sDisabled & " /></td>")
              Send("</tr>")
              EmployeeID = IIf(SaveFailed, Request.QueryString("EmployeeID"), MyCommon.NZ(Cust.GetEmployeeID, ""))
              Send("<tr" & IIf(bDisplay(iEmployeeID), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""EmployeeID"" id=""forEmployeeID"" style=""display:" & IIf(Employee = 1, "block", "none") & ";"">" & Copient.PhraseLib.Lookup("term.employeeid", LanguageID) & ":</label></td>")
              sDisabled = IIf(bEdit(iEmployeeID), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""EmployeeID"" name=""EmployeeID"" " & sDisabled & " maxlength=""26"" style=""display:" & IIf(Employee = 1, "block", "none") & """ value=""" & EmployeeID & """ /></td>")
              Send("</tr>")
              
              'Test card
              If MyCommon.Fetch_SystemOption(88) Then
                If Not SaveFailed Then
                  TestCard = IIf(MyCommon.NZ(Cust.GetTestCard, False), 1, 0)
                End If
                Send("<tr" & IIf(bDisplay(iTestCard), "", " style=""display:none;""") & ">")
                Send("  <td><label for=""TestCard"">" & Copient.PhraseLib.Lookup("term.testcustomer", LanguageID) & ":</label> </td>")
                sDisabled = IIf(bEdit(iTestCard), "", " disabled=""disabled""")
                Send("  <td><input type=""checkbox"" id=""TestCard"" name=""TestCard""" & IIf(TestCard = 1, " checked=""checked""", "") & sDisabled & " /></td>")
                Send("</tr>")
              End If
              
              'Banner
              If MyCommon.Fetch_SystemOption(66) Then
                Send("<tr" & IIf(bDisplay(iBanner), "", " style=""display:none;""") & ">")
                Send("  <td><label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banner", LanguageID) & ":</label> </td>")
                sDisabled = IIf(bEdit(iBanner), "", " disabled=""disabled""")
                Send("  <td><select class=""medium"" id=""BannerID"" name=""BannerID"" " & sDisabled & ">")
                Send("    <option value=""0"">** None Selected **</option>")
                MyCommon.QueryStr = "select BannerID, Name, Description from Banners with (NoLock) where Deleted=0;"
                rst2 = MyCommon.LRT_Select
                For Each row2 In rst2.Rows
                  BannerID = IIf(SaveFailed, Request.QueryString("BannerID"), MyCommon.NZ(Cust.GetBannerID, 0))
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
              Send("<tr" & IIf(bDisplay(iStatus), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""CustomerStatusID"">" & Copient.PhraseLib.Lookup("customer.customerstatus", LanguageID) & ":</label> </td>")
              sDisabled = IIf(bEdit(iStatus), "", " disabled=""disabled""")
              Send("  <td><select class=""medium"" id=""CustomerStatusID"" name=""CustomerStatusID"" " & sDisabled & ">")
              MyCommon.QueryStr = "select CustomerStatusID, PhraseID from CustomerStatus with (NoLock);"
              rst2 = MyCommon.LXS_Select
              For Each row2 In rst2.Rows
                CustomerStatusID = IIf(SaveFailed, Request.QueryString("CustomerStatusID"), MyCommon.NZ(Cust.GetCustomerStatusID, 0))
                If (CustomerStatusID = MyCommon.NZ(row2.Item("CustomerStatusID"), 0)) Then
                  Send("<option value=""" & MyCommon.NZ(row2.Item("CustomerStatusID"), 0) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</option>")
                Else
                  Send("<option value=""" & MyCommon.NZ(row2.Item("CustomerStatusID"), 0) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
                End If
              Next
              Send("  </select></td>")
              Send("</tr>")
              Send("</table>")
              
              Send("<table style=""width:355px;float:left;position:relative;"" summary=""" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & " 2"">")
              
              'Address
              Send("<tr" & IIf(bDisplay(iAddress), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Address"">" & Copient.PhraseLib.Lookup("customer.address", LanguageID) & GetDesignatorText("Address", AltIDCol, AltIDVerCol) & ":</label> </td>")
              TempValue = IIf(bDisplay(iAddress) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetAddress, ""), Request.QueryString("Address"))
              If TempValue <> "" Then
                TempValue.Replace("""", "&quot;")
              End If
              sDisabled = IIf(bEdit(iAddress), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""Address"" name=""Address"" maxlength=""200"" value=""" & TempValue & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'City
              TempValue = IIf(bDisplay(iCity) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetCity, ""), Request.QueryString("City"))
              If TempValue <> "" Then
                TempValue.Replace("""", "&quot;")
              End If
              Send("<tr" & IIf(bDisplay(iCity), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""City"">" & Copient.PhraseLib.Lookup("customer.city", LanguageID) & GetDesignatorText("City", AltIDCol, AltIDVerCol) & ":</label> </td>")
              sDisabled = IIf(bEdit(iCity), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""City"" name=""City"" maxlength=""100"" value=""" & TempValue & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'State
              TempValue = IIf(bDisplay(iState) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetState, ""), Request.QueryString("State"))
              If TempValue <> "" Then
                TempValue.Replace("""", "&quot;")
              End If
              Send("<tr" & IIf(bDisplay(iState), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""State"">" & Copient.PhraseLib.Lookup("customer.state", LanguageID) & GetDesignatorText("State", AltIDCol, AltIDVerCol) & ":</label> </td>")
              sDisabled = IIf(bEdit(iState), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""State"" name=""State"" maxlength=""50"" value=""" & TempValue & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Postal code
              TempValue = IIf(bDisplay(iZip) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetZip, ""), Request.QueryString("Zip"))
              If TempValue <> "" Then
                TempValue.Replace("""", "&quot;")
              End If
              Send("<tr" & IIf(bDisplay(iZip), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Zip"">" & Copient.PhraseLib.Lookup("customer.zip", LanguageID) & GetDesignatorText("Zip", AltIDCol, AltIDVerCol) & ":</label> </td>")
              sDisabled = IIf(bEdit(iZip), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""Zip"" name=""Zip"" maxlength=""20"" value=""" & TempValue & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Country
              TempValue = IIf(bDisplay(iCountry) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetCountry, ""), Request.QueryString("Country"))
              If TempValue <> "" Then
                TempValue.Replace("""", "&quot;")
              End If
              Send("<tr" & IIf(bDisplay(iCountry), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Country"">" & Copient.PhraseLib.Lookup("term.country", LanguageID) & GetDesignatorText("Country", AltIDCol, AltIDVerCol) & ":</label> </td>")
              sDisabled = IIf(bEdit(iCountry), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""Country"" name=""Country"" maxlength=""50"" value=""" & TempValue & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Phone number
              Send("<tr" & IIf(bDisplay(iPhone), "", " style=""display:none;""") & ">")
              TempValue = IIf(bDisplay(iPhone) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetPhone, ""), Request.QueryString("Phone1"))
              Send("  <td><label for=""Phone1"">" & Copient.PhraseLib.Lookup("customer.phone", LanguageID) & GetDesignatorText("Phone", AltIDCol, AltIDVerCol) & ":</label> </td>")
              Send("  <td style=""font-size:18px;"" >")
              sDisabled = IIf(bEdit(iPhone), "", " disabled=""disabled""")
              Send("<input type=""text"" class=""medium"" id=""Phone1"" name=""Phone1"" maxlength=""50"" value=""" & TempValue & """ " & sDisabled & " />&nbsp;")
              Send("  </td>")
              Send("</tr>")
              
              'Mobile phone number
              Send("<tr" & IIf(bDisplay(iMobilePhone), "", " style=""display:none;""") & ">")
              TempValue = IIf(bDisplay(iMobilePhone) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetMobilePhone, ""), Request.QueryString("MobilePhone1"))
              Send("  <td><label for=""MobilePhone1"">" & Copient.PhraseLib.Lookup("customer.mobilephone", LanguageID) & GetDesignatorText("MobilePhone", AltIDCol, AltIDVerCol) & ":</label> </td>")
              Send("  <td style=""font-size:18px;"" >")
              sDisabled = IIf(bEdit(iMobilePhone), "", " disabled=""disabled""")
              Send("<input type=""text""class=""medium"" id=""MobilePhone1"" name=""MobilePhone1"" maxlength=""50"" value=""" & TempValue & """ " & sDisabled & " />&nbsp;")
              Send("  </td>")
              Send("</tr>")
              
              'Email
              Send("<tr" & IIf(bDisplay(iEmail), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Email"">" & Copient.PhraseLib.Lookup("customer.email", LanguageID) & GetDesignatorText("email", AltIDCol, AltIDVerCol) & ":</label> </td>")
              TempValue = IIf(bDisplay(iEmail) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetEmail, ""), Request.QueryString("Email"))
              sDisabled = IIf(bEdit(iEmail), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""Email"" name=""Email"" maxlength=""200"" value=""" & TempValue & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Date of birth
              Send("<tr" & IIf(bDisplay(iDOB), "", " style=""display:none;""") & ">")
              TempValue = ""
              If bDisplay(iDOB) AndAlso Not SaveFailed Then
                If Cust.GetGeneralInfo.GetDateOfBirth <> Nothing Then
                  TempValue = MyCommon.NZ(Cust.GetGeneralInfo.GetDateOfBirth.ToString("MMddyyyy"), "")
                End If
              Else
                TempValue = Request.QueryString("dob1") & Request.QueryString("dob2") & Request.QueryString("dob3")
              End If
              DOBParts = ParseDate(TempValue)
              Send("  <td><label for=""dob1"">" & Copient.PhraseLib.Lookup("customer.dateofbirth", LanguageID) & GetDesignatorText("DOB", AltIDCol, AltIDVerCol) & ":</label> </td>")
              Send("  <td style=""font-size:18px;"" >")
              sDisabled = IIf(bEdit(iDOB), "", " disabled=""disabled""")
              Send("    <input type=""text"" style=""width:39px;"" id=""dob1"" name=""dob1"" maxlength=""2"" value=""" & DOBParts(0) & """ " & sDisabled & " />&nbsp;/&nbsp;")
              Send("    <input type=""text"" style=""width:40px;"" id=""dob2"" name=""dob2"" maxlength=""2"" value=""" & DOBParts(1) & """ " & sDisabled & " />&nbsp;/&nbsp;")
              Send("    <input type=""text"" class=""shorter"" id=""dob3"" name=""dob3"" maxlength=""4"" value=""" & DOBParts(2) & """ " & sDisabled & " />")
              Send("  </td>")
              Send("</tr>")
              
              'Enrollment date
              Send("<tr" & IIf(bDisplay(iEnrollmentDate), "", " style=""display:none;""") & ">")
              TempValue = ""
              If bDisplay(iEnrollmentDate) AndAlso Not SaveFailed Then
                If Cust.GetEnrollmentDate <> Nothing Then
                  TempValue = MyCommon.NZ(Cust.GetEnrollmentDate.ToString("MMddyyyy"), "")
                End If
              Else
                TempValue = Request.QueryString("enrollment1") & Request.QueryString("enrollment2") & Request.QueryString("enrollment3")
              End If
              EnrollmentParts = ParseDate(TempValue)
              Send("  <td><label for=""enrollment1"">" & Copient.PhraseLib.Lookup("term.enrollmentdate", LanguageID) & GetDesignatorText("EnrollmentDate", AltIDCol, AltIDVerCol) & ":</label> </td>")
              Send("  <td style=""font-size:18px;"" >")
              sDisabled = IIf(bEdit(iEnrollmentDate), "", " disabled=""disabled""")
              Send("    <input type=""text"" style=""width:39px;"" id=""enrollment1"" name=""enrollment1"" maxlength=""2"" value=""" & EnrollmentParts(0) & """ " & sDisabled & " />&nbsp;/&nbsp;")
              Send("    <input type=""text"" style=""width:40px;"" id=""enrollment2"" name=""enrollment2"" maxlength=""2"" value=""" & EnrollmentParts(1) & """ " & sDisabled & " />&nbsp;/&nbsp;")
              Send("    <input type=""text"" class=""shorter"" id=""enrollment3"" name=""enrollment3"" maxlength=""4"" value=""" & EnrollmentParts(2) & """ " & sDisabled & " />")
              Send("  </td>")
              Send("</tr>")

              'Password
              Send("<tr" & IIf(bDisplay(iPassword), "", " style=""display:none;""") & ">")
              Send("  <td><label for=""Password"">" & Copient.PhraseLib.Lookup("term.password", LanguageID) & ":</label> </td>")
              Dim tmpPass As String
              tmpPass = IIf(SaveFailed, Request.QueryString("Password"), MyCommon.NZ(Cust.GetPassword, ""))
              If (tmpPass <> "" AndAlso Not SaveFailed) Then
                tmpPass = MyCryptLib.SQL_StringDecrypt(tmpPass)
              End If
              sDisabled = IIf(bEdit(iPassword), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""Password"" name=""Password"" maxlength=""" & MAX_CUST_PASSWORD_CLEARTEXT_LEN & """ value=""" & tmpPass & """ " & sDisabled & " /></td>")
              Send("</tr>")
              
              'Airmile member ID
              TempValue = IIf(bDisplay(iAirMile) AndAlso Not SaveFailed, MyCommon.NZ(Cust.GetGeneralInfo.GetAirmileMemberID, ""), Request.QueryString("AirmileMemberID"))
              Send("<tr" & IIf(bDisplay(iAirMile) AndAlso HasAirmileMemberID, "", " style=""display:none;""") & ">")
              Send("  <td><label for=""AirmileMemberID"">" & Copient.PhraseLib.Lookup("term.airmilememberid", LanguageID) & GetDesignatorText("AirmileMemberID", AltIDCol, AltIDVerCol) & ":</label> </td>")
              sDisabled = IIf(bEdit(iAirMile), "", " disabled=""disabled""")
              Send("  <td><input type=""text"" class=""medium"" id=""AirmileMemberID"" name=""AirmileMemberID"" maxlength=""50"" value=""" & TempValue & """ " & sDisabled & " /></td>")
              Send("</tr>")
              Send("</table>")
              
              'Controls
              Send("<br clear=""left"" />")
              Send("<br/>")
              Send("<input type=""submit"" accesskey=""s"" class=""regular"" id=""save"" name=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
              Send("<input type=""button"" class=""regular"" id=""cancel"" name=""cancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""window.location.href='CM-cashier-inquiry.aspx?CustomerPK=" & CustomerPK & "&ExtCardID=" & ExtCardID & "&GUID=" & sGUID & "&CashierID=" & sCashierID & "&StoreID=" & sStoreId & "';"" />")
              'Send("<input type=""button"" class=""regular"" id=""exit"" name=""exit"" value=""" & Copient.PhraseLib.Lookup("term.exit", LanguageID) & """ onclick=""window.location.href='CM-cashier-inquiry.aspx?exit=Exit';"" />")
            End If
          End If
          Send("<hr class=""hidden"" />")
          Send("</div>")
        End If
      %>
    </div>
    <br clear="all" />
  </div>
</form>

<script runat="server">  
  Private CardStatusTable As Hashtable = Nothing
  
  Function ParseDate(ByVal sDate As String) As String()
    Dim DateParts() As String = {"", "", ""}
    
    If (sDate IsNot Nothing) Then
      Select Case sDate.Length
        Case 4
          DateParts(0) = sDate.Substring(0, 2).PadLeft(2, "0")
          DateParts(1) = sDate.Substring(2, 2).PadLeft(2, "0")
          DateParts(2) = ""
        Case 8
          DateParts(0) = sDate.Substring(0, 2).PadLeft(2, "0")
          DateParts(1) = sDate.Substring(2, 2).PadLeft(2, "0")
          DateParts(2) = sDate.Substring(4).PadLeft(4, "0")
      End Select
    End If
    
    Return DateParts
    
  End Function
  
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
  
  Function ParseTable(ByVal TableCol As String) As String
    Dim Table As String = ""
    Dim Fields() As String
    
    If (TableCol IsNot Nothing) Then
      Fields = TableCol.Split(".")
      Table = Fields(0)
    End If
    
    Return Table
  End Function
  
  Function GetDesignatorText(ByVal Field As String, ByVal AltID As String, ByVal Verifier As String) As String
    Dim Tag As String = ""
    Dim LanguageId As Integer = Me.LanguageID
    
    If (AltID.ToUpper = Field.ToUpper) Then
      Tag = " <span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.alternateid", LanguageId) & ")</span>"
    ElseIf (Verifier.ToUpper = Field.ToUpper) Then
      Tag = " <span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.alternate-id-verifier", LanguageId) & ")</span>"
    End If
    
    Return Tag
  End Function
  
  Function GetNewAltID(ByVal AltIDColumn As String, ByRef FieldName As String) As String
    Dim NewAltID As String = ""
    
    Select Case AltIDColumn.ToUpper
      Case "PHONE"
        NewAltID = Request.QueryString("Phone1")
        FieldName = Copient.PhraseLib.Lookup("customer.phone", LanguageID)
      Case "LASTNAME"
        NewAltID = Request.QueryString("LastName")
        FieldName = Copient.PhraseLib.Lookup("term.lastname", LanguageID)
      Case "FIRSTNAME"
        NewAltID = Request.QueryString("FirstName")
        FieldName = Copient.PhraseLib.Lookup("term.firstname", LanguageID)
      Case "ALTID"
        NewAltID = Request.QueryString("AltIDValue")
        FieldName = Copient.PhraseLib.Lookup("term.alternateid", LanguageID)
      Case "EMAIL"
        NewAltID = Request.QueryString("Email")
        FieldName = Copient.PhraseLib.Lookup("term.email", LanguageID)
      Case "DOB"
        NewAltID = Request.QueryString("dob1") & Request.QueryString("dob2") & Request.QueryString("dob3")
        FieldName = Copient.PhraseLib.Lookup("customer.dateofbirth", LanguageID)
      Case Else
        NewAltID = ""
        FieldName = ""
    End Select
    
    Return NewAltID
  End Function
  
  'Checks to see if the given month is a number between 1 and 12
  Function ValidateMonth(ByVal sMonth As String) As Boolean
    Dim Month As String = sMonth.Trim
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
  
  'Checks that the give day is a number between 1 and 31
  Function ValidateDay(ByVal sDay As String) As Boolean
    Dim Day As String = sDay.Trim
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
  
  Function ValidateYear(ByVal sYear As String) As Boolean
    Dim Year As String = sYear.Trim
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
  
  Private Function IsValidGUID(ByVal GUID As String, ByRef MyCommon As Copient.CommonInc) As Boolean
    Dim IsValid As Boolean = False
    Dim ConnInc As New Copient.ConnectorInc
    
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      ' verfiy using GUIDs defined for the CustomerInquiry Connector (2)
      IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 2, GUID)
    Catch ex As Exception
      IsValid = False
    End Try
      
    Return IsValid
  End Function
  
  Private Function GetStyleFileNamesCashier(ByRef MyCommon As Copient.CommonInc) As String()
    Dim FileName As String() = {"logix-screen.css", ""}
    Dim dst As System.Data.DataTable
    Dim AdminUserID As Long = 1

    MyCommon.QueryStr = "select UIS.FileName, UIS.BaseStyle from UIStyles as UIS inner join AdminUsers as AU on AU.StyleID=UIS.StyleID where AdminUserID=" & AdminUserID & ";"
    dst = MyCommon.LRT_Select

    If dst.Rows.Count > 0 Then
      If Not MyCommon.NZ(dst.Rows(0).Item("BaseStyle"), False) Then
        FileName(0) = "logix-screen.css"
        FileName(1) = MyCommon.NZ(dst.Rows(0).Item("FileName"), "logix-screen.css")
      Else
        ReDim FileName(0)
        FileName(0) = MyCommon.NZ(dst.Rows(0).Item("FileName"), "logix-screen.css")
      End If
    Else
      FileName(0) = "logix-screen.css"
    End If

    Return FileName
  End Function
  
  Private Function BuildLogDescription(ByVal sField As String, ByVal sOld As String, ByVal sNew As String) As String
    Dim sDescription As String
    
    sDescription = sField & " modifed - Old: '" & sOld & "' New: '" & sNew & "'"
    Return sDescription
  End Function

  Private Sub Send_Links_Cashier(ByRef MyCommon As Copient.CommonInc, Optional ByVal Handheld As Boolean = False, Optional ByVal Restricted As Boolean = False)
    Dim Logix As New Copient.LogixInc
    Dim dt As System.Data.DataTable
    Dim myUrl As String = ""
    Dim FileNames As String() = Nothing
    Dim i As Integer = 0

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
                Send("<link rel=""stylesheet"" href=""/css/logix-screen.css"" type=""text/css"" media=""screen"" />")
            End If
        Else
            FileNames = GetStyleFileNamesCashier(MyCommon)
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

  End Sub

</script>

<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
