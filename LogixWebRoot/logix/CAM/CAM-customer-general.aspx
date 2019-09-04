<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-customer-general.aspx 
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
  Dim MyCryptLib As New Copient.CryptLib
  Dim MyLookup As New Copient.CustomerLookup
  Dim Logix As New Copient.LogixInc
  Dim MyAltID As New Copient.AlternateID
  Dim AltIDResponse As New Copient.AlternateID.CreateUpdateResponse
  Dim MyCAM As New Copient.CAM
  Dim rstResults As DataTable = Nothing
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim row3 As DataRow
  Dim rst4 As DataTable
  Dim CurrentOffers As String = ""
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim CardTypeID As Integer = 0
  Dim ExtID As String = ""
  Dim Prefix As String = ""
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim Suffix As String = ""
  Dim FullName As String = ""
  Dim ExtCustomerID As String = ""
  Dim Edit As Boolean
  Dim TotalRedeemCt As Integer = 0
  Dim TotalRedeemAmt As Double = 0.0
  Dim CustExtID As String = ""
  Dim i As Integer = 0
  Dim offerCt As Integer = 0
  Dim transCt As Integer = 0
  Dim ElemStyle As String = ""
  Dim OfferName As String = ""
  Dim XID As String = ""
  Dim IsPtsOffer As Boolean = False
  Dim IsAccumOffer As Boolean = False
  Dim UnknownPhrase As String = ""
  Dim DisabledPtsAdj As String = ""
  Dim DisabledAccumAdj As String = ""
  Dim CgGroupIDs As String = ""
  Dim SortText As String = "O.Name"
  Dim SortDirection As String = "ASC"
  Dim OfferCMWhere As String = ""
  Dim OfferCPEWhere As String = ""
  Dim OfferTerms As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim HHPK As Integer = 0
  Dim HouseholdID As String = ""
  Dim HHCustIdList As New ArrayList(5)
  Dim PaddedExtID As String = StrDup(25, "0")
  Dim CustExtIdList As String = ""
  Dim Shaded As String = " class=""shaded"""
  Dim HasSearchResults As Boolean = False
  Dim CustomerTypeID As Integer = 0
  Dim Employee As Integer
  Dim TestCard As Integer
  Dim OfferID As Integer = 0
  Dim ClientUserID1 As String = ""
  Dim IDLength As Integer = 0
  Dim CustomerGroupIDs As String() = Nothing
  Dim loopCtr As Integer = 0
  Dim searchterms As String = ""
  Dim restrictLinks As Boolean = False
  Dim PointsIDBuf As New StringBuilder()
  Dim PointsNameBuf As New StringBuilder()
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim SavedCardStatus As Integer = 0
  Dim AltIDField As Object = Nothing
  Dim IDVerifierField As Object = Nothing
  Dim AltIDCol As String = ""
  Dim AltIDVerCol As String = ""
  Dim bUpdateAlt As Boolean = True
  Dim NewAltID As String = ""
  Dim AltIDFailed As Boolean = False
  Dim BannerID As Integer = 0
  Dim DateValid As Boolean = False
  
  Dim Address As String = ""
  Dim City As String = ""
  Dim State As String = ""
  Dim Zip As String = ""
  Dim Country As String = ""
  Dim FullAddress As String = ""
  Dim Phone1 As String = ""
  Dim Phone2 As String = ""
  Dim Phone3 As String = ""
  Dim FullPhone As String = ""
  Dim MobilePhone1 As String = ""
  Dim MobilePhone2 As String = ""
  Dim MobilePhone3 As String = ""
  Dim Email As String = ""
  Dim Household As String = ""
  Dim Password As String = ""
  Dim DOB_month As String = ""
  Dim DOB_day As String = ""
  Dim DOB_year As String = ""
  Dim DOB As String = ""
  Dim AltIDValue As String = ""
  Dim AltIDVerifier As String = ""
  Dim CAMSErrorMessage As String = ""
  Dim CustomerStatusID As Integer = 0

  ' default urls for links from this page
  Dim URLOfferSum As String = "../offer-sum.aspx"
  Dim URLCPEOfferSum As String = "../CPEoffer-sum.aspx"
  Dim URLcgroupedit As String = "../cgroup-edit.aspx"
  Dim URLpointedit As String = "../point-edit.aspx"
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  
  Dim Customers(-1) As Copient.Customer
  Dim HHCustomers(-1) As Copient.Customer
  Dim Cust As New Copient.Customer
  Dim CustExt As New Copient.CustomerExt
  Dim CustNotes(-1) As Copient.CustomerNote
  Dim CustNote As New Copient.CustomerNote
  Dim Note As String = ""
  Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-customer-general.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  
  AltIDCol = ParseTableCol(MyCommon.Fetch_SystemOption(60))
  AltIDVerCol = ParseTableCol(MyCommon.Fetch_SystemOption(61))

  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If (MyCommon.NZ(rst.Rows(0).Item("prestrict"), False) = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      restrictLinks = True
    End If
  End If
  
  CustomerPK = 0
  Edit = False
  
  ' special handling for customer inquery direct link in 
  If (restrictLinks) Then
    URLOfferSum = ""
    URLCPEOfferSum = ""
    URLcgroupedit = ""
    URLpointedit = ""
  End If
  
  ' set session to nothing just to be sure
  Session.Add("extraLink", "")
  
  If (Request.QueryString("mode") = "summary") Then
    URLtrackBack = Request.QueryString("exiturl")
    inCardNumber = Request.QueryString("cardnumber")
    extraLink = "&mode=summary&exiturl=" & URLtrackBack & "&cardnumber=" & inCardNumber
    Session.Add("extraLink", extraLink)
  End If
  
  ' hack for popups check session for extralink
  If (Session("extraLink").ToString = "") Then
    extraLink = Session("extraLink")
  End If
  
  If (MyCommon.Extract_Val(Request.QueryString("CustPK")) = 0 AndAlso (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) = 0)) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "../customer-inquiry.aspx")
  End If
  
  If (MyCommon.Extract_Val(Request.QueryString("CardPK")) > 0) Then
    CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  ElseIf (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0) Then
    CardPK = MyLookup.FindCardPK(Long.Parse(Request.QueryString("CustPK")), 2)
  End If

  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
    CardTypeID = MyLookup.FindCardTypeFromCardPK(CardPK)
  End If

  If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0 Or (Request.QueryString("searchterms") <> "" And _
      (Request.QueryString("Search") <> "" Or Request.QueryString("searchPressed") <> "")) Or _
      inCardNumber <> "" _
      ) Then
    ' Someone wants to search for a customer.  First, let's get their primary key from our database
    If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0 Or (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0)) Then
      If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0) Then
        CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
      ElseIf (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0) Then
        CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      End If
      ReDim Customers(0)
      Customers(0) = MyLookup.FindCustomerInfo(CustomerPK, ReturnCode)
      MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.Prefix, C.FirstName, C.MiddleName, C.LastName, C.Suffix," & _
                          "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email, CE.MobilePhone " & _
                          "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                          "left join Customers C2 with (NoLock) on C2.CustomerPK=C.HHPK " & _
                          "where C.CustomerPK=" & CustomerPK
    Else
      ' IF the page was called from an outside application, set ClientUserID1 to the outside passed in value
      If (inCardNumber <> "" And Request.QueryString("mode") = "summary") Then
        'leave as is without new common.pad_extcardid 
        ClientUserID1 = MyCommon.Pad_ExtCardID(inCardNumber, Copient.commonShared.CardTypes.CAM)
        searchterms = Request.QueryString("searchterms")
        Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.ALL_CUSTOMER_TYPES, -1, ClientUserID1, ReturnCode)
      End If
      
    End If
    
    If (Customers.Length > 1) Then
      HasSearchResults = True
    ElseIf (Customers.Length = 1) Then
      ' ok we found a primary key for the external id provided
      CustomerPK = Customers(0).GetCustomerPK
      IsHouseholdID = Customers(0).GetCustomerTypeID = 1
      Prefix = Customers(0).GetPrefix
      FirstName = Customers(0).GetFirstName
      MiddleName = Customers(0).GetMiddleName
      LastName = Customers(0).GetLastName
      Suffix = Customers(0).GetSuffix
      HasSearchResults = False
    Else
      infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
      infoMessage = infoMessage & " <a href=""customer-general.aspx?mode=add&Search=Search" & extraLink & "&searchterms=" & Request.QueryString("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
      HasSearchResults = False
    End If
    
    If (Request.QueryString("Edit") <> "") Then
      Edit = True
    End If
    
  ElseIf (Request.QueryString("editterms") <> "" And Request.QueryString("Edit") = Copient.PhraseLib.Lookup("term.edit", LanguageID)) Then
    ' someone wants to search for a customer.  First lets get their primary key from our database
    ExtCardID = MyCommon.Pad_ExtCardID(Request.QueryString("editterms"), 0)
    MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where CardTypeID=0 and ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "'"
    rst = MyCommon.LXS_Select
    If (rst.Rows.Count > 0) Then
      ' ok we found a primary key for the external id provided
      CustomerPK = rst.Rows(0).Item("CustomerPK")
      IsHouseholdID = False
      ExtID = Request.QueryString("editterms")
    Else
      infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
    End If
    Edit = True
    
  ElseIf (Request.QueryString("save") <> "") Then
    ExtCardID = MyCommon.Pad_ExtCardID(Request.QueryString("editterms"), 2)
    ' setup the page so the thing draws like it should
    If (Request.QueryString("CustomerPK") <> "") Then
      MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where CustomerPK='" & MyCommon.Extract_Val(Request.QueryString("CustomerPK")) & "'"
    Else
      MyCommon.QueryStr = "select top 1 CustomerPK from CardIDs with (NoLock) where CardTypeID=2 and ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "'"
    End If
    
    rst = MyCommon.LXS_Select
    If (rst.Rows.Count > 0) Then
      ' ok we found a primary key for the external id provided
      CustomerPK = rst.Rows(0).Item("CustomerPK")
      
      MyCommon.QueryStr = "select CustomerPK, CustomerTypeID, HHPK, CustomerStatusID from Customers with (NoLock) where CustomerPK=" & MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
      rst2 = MyCommon.LXS_Select
      If rst2.Rows.Count > 0 Then
        IsHouseholdID = (MyCommon.NZ(rst2.Rows(0).Item("CustomerTypeID"), 0) = 1)
        
        MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CardTypeID=2 and CustomerPK=" & CustomerPK
        rst2 = MyCommon.LXS_Select
        If rst2.Rows.Count > 0 Then
                    ExtCardID = MyCryptLib.SQL_StringDecrypt(rst2.Rows(0).Item("ExtCardID").ToString())
        End If
      
      End If
      
      ExtID = Request.QueryString("editterms")
    End If
    
    ' ok someone wants to save the users information lets put it in the db
    Prefix = MyCommon.Parse_Quotes(Request.QueryString("Prefix"))
    FirstName = MyCommon.Parse_Quotes(Request.QueryString("FirstName"))
    MiddleName = MyCommon.Parse_Quotes(Request.QueryString("MiddleName"))
    LastName = MyCommon.Parse_Quotes(Request.QueryString("LastName"))
    Suffix = MyCommon.Parse_Quotes(Request.QueryString("Suffix"))
    CustomerStatusID = MyCommon.Extract_Val(Request.QueryString("CustomerStatusID"))
    Address = MyCommon.Parse_Quotes(Request.QueryString("Address"))
    City = MyCommon.Parse_Quotes(Request.QueryString("City"))
    State = MyCommon.Parse_Quotes(Request.QueryString("State"))
    Zip = MyCommon.Parse_Quotes(Request.QueryString("Zip"))
    Country = MyCommon.Parse_Quotes(Request.QueryString("Country"))
    Phone1 = MyCommon.Parse_Quotes(Request.QueryString("Phone1"))
    Phone2 = MyCommon.Parse_Quotes(Request.QueryString("Phone2"))
    Phone3 = MyCommon.Parse_Quotes(Request.QueryString("Phone3"))
    MobilePhone1 = MyCommon.Parse_Quotes(Request.QueryString("MobilePhone1"))
    MobilePhone2 = MyCommon.Parse_Quotes(Request.QueryString("MobilePhone2"))
    MobilePhone3 = MyCommon.Parse_Quotes(Request.QueryString("MobilePhone3"))
    Email = MyCommon.Parse_Quotes(Request.QueryString("Email"))
    Household = MyCommon.Parse_Quotes(Request.QueryString("Household"))
    Password = Request.QueryString("Password")
    DOB_month = Request.QueryString("dob1")
    DOB_day = Request.QueryString("dob2")
    DOB_year = Request.QueryString("dob3")
    DOB = ""
    AltIDValue = Request.QueryString("AltIDValue")
    AltIDVerifier = Request.QueryString("AltIDVerifier")
    BannerID = Request.QueryString("BannerID")
    If (Request.QueryString("Employee") = "on") Then
      Employee = 1
    Else
      Employee = 0
    End If
    If (Request.QueryString("TestCard") = "on") Then
      TestCard = 1
    Else
      TestCard = 0
    End If
    
    ' handle updates to Alternate Identifier 
    If (AltIDCol.Trim <> "") Then
      NewAltID = GetNewAltID(AltIDCol, AltIDField)
      AltIDResponse = MyAltID.UpdateCustomerAltID(CustomerPK, NewAltID, BannerID)
      Select Case AltIDResponse
        Case Copient.AlternateID.CreateUpdateResponse.ALTIDINUSE
          If NewAltID = "" AndAlso MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(81)) <= 1 Then
            infoMessage = Copient.PhraseLib.Lookup("error.noaltid", LanguageID)
          Else
            infoMessage = Copient.PhraseLib.Detokenize("customer-general.AltIDInUse", LanguageID, AltIDField, NewAltID)
          End If
          AltIDFailed = True
        Case Copient.AlternateID.CreateUpdateResponse.MEMBERNOTFOUND
          infoMessage = Copient.PhraseLib.Detokenize("customer-general.CustomerNotFound", LanguageID, ExtCardID)
          AltIDFailed = True
        Case Copient.AlternateID.CreateUpdateResponse.ERROR_APPLICATION
          infoMessage = Copient.PhraseLib.Detokenize("customer-general.AltIDError", LanguageID, MyAltID.ErrorMessage)
          AltIDFailed = True
      End Select
      Edit = True
    End If

    If (infoMessage.Trim = "") Then
      MyCommon.QueryStr = "update Customers with (RowLock) set CPEStoreSendFlag = 1, FirstName='" & FirstName & "', MiddleName='" & MiddleName & "', LastName='" & LastName & "', Employee=" & Employee & ", " & _
                          " Prefix='" & Prefix & "', Suffix='" & Suffix & "', CustomerStatusID=" & CustomerStatusID & ", CustomerTypeID=2, password=N'" & MyCryptLib.SQL_StringEncrypt(Password) & "', " & _
                          " BannerID = " & IIf(BannerID > 0, BannerID.ToString, "NULL") & ", TestCard=" & TestCard & " " & _
                          " where CustomerPK=" & CustomerPK & ""
      MyCommon.LXS_Execute()
      ' format the DOB
      If (DOB_month = "" And DOB_day = "" And DOB_year = "") Then  'Allows a null to be set for the DOB when there is nothing saved in the DOB fields
        DateValid = True
        DOB = "NULL"
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
        infoMessage = TempMessage
      Else
                DOB = "" & DOB_month.Trim.PadLeft(2, "0") & DOB_day.Trim.PadLeft(2, "0") & DOB_year.Trim.PadLeft(4, "0") & ""
        DateValid = True
      End If
      
      ' alright time to deal with the extension data
      ' first check if there is a record in the extension table for this user
      MyCommon.QueryStr = "select CustomerPK from CustomerExt with (NoLock) where CustomerPK =" & CustomerPK
      rst = MyCommon.LXS_Select()
      If (DateValid) Then
        If rst.Rows.Count = 0 Then
          ' nothing found we need to do an insert
                    MyCommon.QueryStr = "insert into CustomerExt with (RowLock) (Address,City,State,Zip,Country,PhoneAsEntered,PhoneDigitsOnly,MobilePhoneAsEntered,MobilePhoneDigitsOnly,Email,DOB,CustomerPK) values (" & _
                                        " '" & Logix.TrimAll(Address) & "', " & _
                                        " '" & Logix.TrimAll(City) & "', " & _
                                        " '" & Logix.TrimAll(State) & "', " & _
                                        " '" & Logix.TrimAll(Zip) & "', " & _
                                        " '" & Logix.TrimAll(Country) & "', " & _
                                        " '" & MyCryptLib.SQL_StringEncrypt(Phone1.Trim) & "', " & _
                                        " '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Phone1.Trim)) & "', " & _
                                        " '" & MyCryptLib.SQL_StringEncrypt(MobilePhone1.Trim) & "', " & _
                                        " '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(MobilePhone1.Trim)) & "', " & _
                                        " '" & MyCryptLib.SQL_StringEncrypt(Logix.TrimAll(Email)) & "'," & _
                                        " '" & MyCryptLib.SQL_StringEncrypt(DOB) & "'," & _
                                        " " & CustomerPK & ") "
          MyCommon.LXS_Execute()
        Else
          ' found a row we can do an update
          ' now update the other extention table
                    MyCommon.QueryStr = "update CustomerExt with (RowLock) set " & _
                                        " Address='" & Logix.TrimAll(Address) & "'," & _
                                        " City='" & Logix.TrimAll(City) & "'," & _
                                        " State='" & Logix.TrimAll(State) & "'," & _
                                        " Zip='" & Logix.TrimAll(Zip) & "'," & _
                                        " Country='" & Logix.TrimAll(Country) & "', " & _
                                        " PhoneAsEntered='" & MyCryptLib.SQL_StringEncrypt(Phone1.Trim) & "'," & _
                                        " PhoneDigitsOnly='" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Phone1.Trim)) & "'," & _
                                        " MobilePhoneAsEntered='" & MyCryptLib.SQL_StringEncrypt(MobilePhone1.Trim) & "'," & _
                                        " MobilePhoneDigitsOnly='" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(MobilePhone1.Trim)) & "'," & _
                                        " Email='" & MyCryptLib.SQL_StringEncrypt(Logix.TrimAll(Email)) & "'," & _
                                        " DOB='" & MyCryptLib.SQL_StringEncrypt(DOB) & "' " & _
                                        " where CustomerPK =" & CustomerPK
          MyCommon.LXS_Execute()
        End If
        
        MyCommon.Activity_Log2(25, 11, GetCustomerPK(MyCommon, Request.QueryString("editterms")), AdminUserID, Copient.PhraseLib.Lookup("history.customer-edited-info", LanguageID))
      End If
      Edit = True
    End If
  End If
  
  If CardPK > 0 Then
    Send_HeadBegin("term.customer", "term.general", MyCommon.Extract_Val(ExtCardID))
  Else
    Send_HeadBegin("term.customer", "term.general")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""/css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
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
      Send_Subtabs(Logix, 33, 3, LanguageID, CustomerPK, , CardPK)
    Else
      Send_Subtabs(Logix, 33, 3, LanguageID, CustomerPK)
    End If
  Else
    If CardPK > 0 Then
      Send_Subtabs(Logix, 94, 3, LanguageID, CustomerPK, extraLink, CardPK)
    Else
      Send_Subtabs(Logix, 94, 3, LanguageID, CustomerPK, extraLink)
    End If
  End If
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
%>
<script type="text/javascript">
<!--
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
    // }
    //retVal = retVal && isValidPhoneCombo("MobilePhone");

    // validate date of birth
    for (var i = 1; i <= 3 && retVal; i++) {
      retVal = retVal && isValidDobPart(i);
    }

    return retVal;
  }

  //function isValidPhonePart(prefix, partNum) {
  //  var retVal = true;
  //  var elemPart = document.getElementById(prefix + partNum);

  //  if (elemPart != null) {
  //    if (partNum ==1 && elemPart.value!="" && elemPart.value.length != 3) { 
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
  //  
  //  if (!retVal) {
  //    alert("Phone number should be in either 7 or 10 digit phone number format.");
  //  }
  //  return retVal;
  //}

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

//-->
</script>
<script type="text/javascript" src="/javascript/jquery.min.js"></script>
<script type="text/javascript" src="/javascript/thickbox.js"></script>
<form id="mainform" name="mainform" action="CAM-customer-general.aspx" onsubmit="return isValidEntry();">
<input type="hidden" name="altid" value="<%Sendb(AltIDCol) %>" />
<input type="hidden" name="verifier" value="<%Sendb(AltIDVerCol) %>" />
<div id="intro">
  <h1 id="title">
    <%
      If CardPK = 0 Then
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
        End If
      Else
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & ExtCardID)
        Else
          Sendb(Copient.PhraseLib.Lookup("term.camcard", LanguageID) & " #" & ExtCardID)
        End If
      End If
      If AltIDFailed Then
        MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
        rst2 = MyCommon.LXS_Select
        If rst2.Rows.Count > 0 Then
          FullName = IIf(MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, rst2.Rows(0).Item("Prefix") & " ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, rst2.Rows(0).Item("FirstName") & " ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, rst2.Rows(0).Item("LastName"), "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & rst2.Rows(0).Item("Suffix"), "")
        Else
          FullName = IIf(Prefix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, Prefix & " ", "")
          FullName &= IIf(FirstName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, FirstName & " ", "")
          FullName &= IIf(MiddleName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MiddleName, 1) & ". ", "")
          FullName &= IIf(LastName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, LastName, "")
          FullName &= IIf(Suffix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & Suffix, "")
        End If
      Else
        FullName = IIf(Prefix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, Prefix & " ", "")
        FullName &= IIf(FirstName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, FirstName & " ", "")
        FullName &= IIf(MiddleName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MiddleName, 1) & ". ", "")
        FullName &= IIf(LastName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, LastName, "")
        FullName &= IIf(Suffix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & Suffix, "")
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
  <%If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br />")%>
  <%--
    <%If (Logix.UserRoles.ViewCustomerIdData = False) Then
        Send(Copient.PhraseLib.Lookup("error.forbidden", LanguageID))
        Send("</div>")
        Send("</form>")
        GoTo done
      End If
    %>
  --%>
  <input type="hidden" id="CustomerPK" name="CustomerPK" value="<%Sendb(CustomerPK)%>" />
  <%
    If (Request.QueryString("mode") = "summary") Then
      Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
      Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
    End If
  %>
  <div id="column">
    <% Send("<input type=""hidden"" id=""editterms"" name=""editterms"" value=""" & ExtID & """ />")%>
    <% If (CustomerPK > 0) Then%>
    <div class="box" id="identity" <%if (customerpk = 0) then sendb(" style=""display: none;""")%>>
      <%
        MyCommon.QueryStr = "select C.CustomerPK, Prefix, FirstName, MiddleName, LastName, Suffix, Password, Employee, " & _
                            "TestCard, CS.CustomerStatusID, CS.PhraseID, Address, City, State, Zip, Country, PhoneAsEntered As Phone, " & _
                            "Email, DOB, C.BannerID, C.AltID, C.Verifier, CE.MobilePhoneAsEntered as MobilePhone " & _
                            "from Customers as C left join CustomerExt as CE on CE.CustomerPK=C.CustomerPK " & _
                            "left join CustomerStatus as CS on CS.CustomerStatusID=C.CustomerStatusID " & _
                            "where C.CustomerPK='" & CustomerPK & "';"
        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
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
            If CustomerPK > 0 Then
              If MyCommon.NZ(rst.Rows(0).Item("address"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Address Then
                FullAddress &= MyCommon.TruncateString(MyCommon.NZ(rst.Rows(0).Item("address"), ""), 25) & "<br />"
              End If
              If MyCommon.NZ(rst.Rows(0).Item("city"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_City Then
                FullAddress &= MyCommon.TruncateString(MyCommon.NZ(rst.Rows(0).Item("city"), ""), 25) & ", "
              End If
              If MyCommon.NZ(rst.Rows(0).Item("state"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_State Then
                FullAddress &= MyCommon.TruncateString(MyCommon.NZ(rst.Rows(0).Item("state"), ""), 25) & "&nbsp;"
              End If
              If MyCommon.NZ(rst.Rows(0).Item("zip"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_ZIP Then
                FullAddress &= MyCommon.NZ(rst.Rows(0).Item("zip"), "")
              End If
            End If
            'Build up the full phone
            FullPhone = ""
            If CustomerPK > 0 Then
                      If MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("phone"), "")) <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Phone Then
                          FullPhone &= MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("phone"), "&nbsp;")) & "<br />"
                      End If
                      If MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("mobilephone"), "")) <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MobilePhone Then
                          FullPhone &= Copient.PhraseLib.Lookup("term.mobile", LanguageID) & ": " & MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("mobilephone"), "&nbsp;"))
                      End If
            End If
            'Get email address
            If CustomerPK > 0 Then
                      If MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("email"), "")) <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Email Then
                          Email = MyCryptLib.SQL_StringDecrypt(MyCommon.TruncateString(MyCommon.NZ(rst.Rows(0).Item("email"), "&nbsp;"), 256))
                      End If
            End If
            'Get date of birth
            If CustomerPK > 0 Then
                      If MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("DOB"), "")) <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_DOB Then
                          DOB = MyCryptLib.SQL_StringDecrypt(rst.Rows(0).Item("DOB"))
                      End If
            End If
              
            If FullName <> "" OrElse FullAddress <> "" OrElse FullPhone <> "" OrElse Email <> "" OrElse DOB <> "" Then
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
              
            ' Test card enabled?
            If MyCommon.Fetch_SystemOption(88) Then
              If (MyCommon.NZ(rst.Rows(0).Item("TestCard"), 0) <> 0) Then
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
              
            ' Report card status
            If Logix.UserRoles.AccessCustomerIdData_Status Then
              If (MyCommon.NZ(rst.Rows(0).Item("CustomerStatusID"), 0) = 1) Then
                Send("<span class=""darkgreen""><b>")
              ElseIf (MyCommon.NZ(rst.Rows(0).Item("CustomerStatusID"), 0) = 2) Then
                Send("<span class=""darkgray""><b>")
              Else
                Send("<span class=""red""><b>")
              End If
              Send(Copient.PhraseLib.Lookup("term.cardis", LanguageID) & StrConv(Copient.PhraseLib.Lookup(MyCommon.NZ(rst.Rows(0).Item("PhraseID"), 0), LanguageID), VbStrConv.Lowercase) & ".<br />")
              Send("</b></span>")
            End If
              
            MyCommon.QueryStr = "select CustomerPK from CustomerLocks where CustomerPK=" & CustomerPK & ";"
            rst2 = MyCommon.LXS_Select
            If rst2.Rows.Count > 0 Then
              Send("<span class=""red"">" & Copient.PhraseLib.Lookup("customer.locked", LanguageID) & "</span><br />")
            End If
            Send("<br class=""half"" />")
            ' Buttons
            If (Logix.UserRoles.EditCustomerIdData) Then
              Send("<input type=""button"" class=""regular"" id=""edit"" name=""edit"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""window.location.href='CAM-customer-general.aspx?edit=Edit" & extraLink & "&amp;editterms=" & ExtCustomerID & "&amp;CustPK=" & rst.Rows(0).Item("CustomerPK") & "';"" />")
            Else
              Send("&nbsp;<br />")
            End If
          Else
            Dim TempValue As String = ""
            Dim DOBParts() As String = {"", "", ""}
              
            Send("<table summary=""" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & """>")
            'Prefix for name
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Prefix, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Prefix"">" & Copient.PhraseLib.Lookup("term.prefix", LanguageID) & GetDesignatorText("Prefix", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Prefix = IIf(AltIDFailed, Prefix, MyCommon.NZ(rst.Rows(0).Item("Prefix"), "").Replace("""", "&quot;"))
            Send("  <td><input type=""text"" class=""medium"" id=""Prefix"" name=""Prefix"" maxlength=""20"" value=""" & Prefix & """ /></td>")
            Send("</tr>")
            ' First Name
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_FirstName, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""FirstName"">" & Copient.PhraseLib.Lookup("term.firstname", LanguageID) & GetDesignatorText("FirstName", AltIDCol, AltIDVerCol) & ":</label> </td>")
            FirstName = IIf(AltIDFailed, FirstName, MyCommon.NZ(rst.Rows(0).Item("FirstName"), UnknownPhrase).Replace("""", "&quot;"))
            Send("  <td><input type=""text"" class=""medium"" id=""FirstName"" name=""FirstName"" maxlength=""50"" value=""" & FirstName & """ /></td>")
            Send("</tr>")
            ' Middle Name
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_MiddleName, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""MiddleName"">" & Copient.PhraseLib.Lookup("term.middlename", LanguageID) & GetDesignatorText("MiddleName", AltIDCol, AltIDVerCol) & ":</label> </td>")
            MiddleName = IIf(AltIDFailed, MiddleName, MyCommon.NZ(rst.Rows(0).Item("MiddleName"), "").Replace("""", "&quot;"))
            Send("  <td><input type=""text"" class=""medium"" id=""MiddleName"" name=""MiddleName"" maxlength=""50"" value=""" & MiddleName & """ /></td>")
            Send("</tr>")
            'Last Name
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_LastName, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""LastName"">" & Copient.PhraseLib.Lookup("term.lastname", LanguageID) & GetDesignatorText("LastName", AltIDCol, AltIDVerCol) & ":</label> </td>")
            LastName = IIf(AltIDFailed, LastName, MyCommon.NZ(rst.Rows(0).Item("LastName"), UnknownPhrase).Replace("""", "&quot;"))
            Send("  <td><input type=""text"" class=""medium"" id=""LastName"" name=""LastName"" maxlength=""50"" value=""" & LastName & """ /></td>")
            Send("</tr>")
            ' Suffix for name
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Suffix, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Suffix"">" & Copient.PhraseLib.Lookup("term.suffix", LanguageID) & GetDesignatorText("Suffix", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Suffix = IIf(AltIDFailed, Suffix, MyCommon.NZ(rst.Rows(0).Item("Suffix"), "").Replace("""", "&quot;"))
            Send("  <td><input type=""text"" class=""medium"" id=""Suffix"" name=""Suffix"" maxlength=""20"" value=""" & Suffix & """ /></td>")
            Send("</tr>")
            'Alt ID
            If (AltIDCol.ToUpper = "ALTID") Then
              Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_AltID, "", " style=""display:none;""") & ">")
              Send("  <td><label for=""AltIDValue"">" & Copient.PhraseLib.Lookup("term.alternateid", LanguageID) & ":</label> </td>")
              AltIDValue = IIf(AltIDFailed, AltIDValue, MyCommon.NZ(rst.Rows(0).Item("AltID"), "").Replace("""", "&quot;"))
              Send("  <td><input type=""text"" class=""medium"" id=""AltIDValue"" name=""AltIDValue"" maxlength=""20"" value=""" & AltIDValue & """ /></td>")
              Send("</tr>")
            End If
            If (AltIDVerCol.ToUpper = "VERIFIER") Then
              Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_AltID, "", " style=""display:none;""") & ">")
              Send("  <td><label for=""AltIDVerifier"">" & Copient.PhraseLib.Lookup("term.alternate-id-verifier", LanguageID) & ":</label> </td>")
              AltIDVerifier = IIf(AltIDFailed, AltIDVerifier, MyCommon.NZ(rst.Rows(0).Item("Verifier"), "").Replace("""", "&quot;"))
              Send("  <td><input type=""text"" class=""medium"" id=""AltIDVerifier"" name=""AltIDVerifier"" maxlength=""20"" value=""" & AltIDVerifier & """ /></td>")
              Send("</tr>")
            End If
              
            ' Test card enabled?
            If MyCommon.Fetch_SystemOption(88) Then
              Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Test, "", " style=""display:none;""") & ">")
              Send("  <td><label for=""TestCard"">" & Copient.PhraseLib.Lookup("term.testcard", LanguageID) & ":</label> </td>")
              Sendb("  <td><input type=""checkbox"" id=""TestCard"" name=""TestCard""")
              TestCard = IIf(MyCommon.NZ(rst.Rows(0).Item("TestCard"), 0), 1, 0)
              If (TestCard) Then
                Sendb(" checked=""checked""")
              End If
              If (Not Logix.UserRoles.AccessCustomerIdData_Test) Then
                Sendb(" disabled=""disabled""")
              End If
              Sendb(" />")
              Send("</tr>")
            End If
              
            If MyCommon.Fetch_SystemOption(66) Then
              Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Banner, "", " style=""display:none;""") & ">")
              Send("  <td><label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banner", LanguageID) & ":</label> </td>")
              Send("  <td><select class=""medium"" id=""BannerID"" name=""BannerID"">")
              Send("<option value=""0"">** None Selected **</option>")
              MyCommon.QueryStr = "select BannerID, Name, Description from Banners with (NoLock) where Deleted=0;"
              rst2 = MyCommon.LRT_Select
              For Each row2 In rst2.Rows
                BannerID = IIf(AltIDFailed, BannerID, MyCommon.NZ(rst.Rows(0).Item("BannerID"), 0))
                If (BannerID = MyCommon.NZ(row2.Item("BannerID"), 0)) Then
                  Send("<option value=""" & MyCommon.NZ(row2.Item("BannerID"), 0) & """ selected=""selected"">" & MyCommon.TruncateString(MyCommon.NZ(row2.Item("Name"), "&nbsp;"), 25) & "</option>")
                Else
                  Send("<option value=""" & MyCommon.NZ(row2.Item("BannerID"), 0) & """>" & MyCommon.TruncateString(MyCommon.NZ(row2.Item("Name"), "&nbsp;"), 25) & "</option>")
                End If
              Next
              Send("  </select></td>")
              Send("</tr>")
            End If
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Status, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""CustomerStatusID"">" & Copient.PhraseLib.Lookup("customer.customerstatus", LanguageID) & ":</label> </td>")
            Send("  <td><select class=""medium"" id=""CustomerStatusID"" name=""CustomerStatusID"">")
            MyCommon.QueryStr = "select CustomerStatusID, PhraseID, Description from CustomerStatus with (NoLock);"
            rst2 = MyCommon.LXS_Select
            For Each row2 In rst2.Rows
              CustomerStatusID = IIf(AltIDFailed, CustomerStatusID, MyCommon.NZ(rst.Rows(0).Item("CustomerStatusID"), 0))
              If (CustomerStatusID = MyCommon.NZ(row2.Item("CustomerStatusID"), 0)) Then
                Send("<option value=""" & MyCommon.NZ(row2.Item("CustomerStatusID"), 0) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</option>")
              Else
                Send("<option value=""" & MyCommon.NZ(row2.Item("CustomerStatusID"), 0) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID) & "</option>")
              End If
            Next
            Send("  </select></td>")
            Send("</tr>")
            'Address
            TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Address AndAlso Not AltIDFailed, MyCommon.NZ(rst.Rows(0).Item("Address"), ""), Request.QueryString("Address"))
            If TempValue <> "" Then
              TempValue.Replace("""", "&quot;")
            End If
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Address, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Address"">" & Copient.PhraseLib.Lookup("customer.address", LanguageID) & GetDesignatorText("Address", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td><input type=""text"" class=""medium"" id=""Address"" name=""Address"" maxlength=""200"" value=""" & TempValue & """ /></td>")
            Send("</tr>")
            'City
            TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_City AndAlso Not AltIDFailed, MyCommon.NZ(rst.Rows(0).Item("City"), ""), Request.QueryString("City"))
            If TempValue <> "" Then
              TempValue.Replace("""", "&quot;")
            End If
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_City, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""City"">" & Copient.PhraseLib.Lookup("customer.city", LanguageID) & GetDesignatorText("City", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td><input type=""text"" class=""medium"" id=""City"" name=""City"" maxlength=""100"" value=""" & TempValue & """ /></td>")
            Send("</tr>")
            'State
            TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_State AndAlso Not AltIDFailed, MyCommon.NZ(rst.Rows(0).Item("State"), ""), Request.QueryString("State"))
            If TempValue <> "" Then
              TempValue.Replace("""", "&quot;")
            End If
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_State, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""State"">" & Copient.PhraseLib.Lookup("customer.state", LanguageID) & GetDesignatorText("State", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td><input type=""text"" class=""medium"" id=""State"" name=""State"" maxlength=""50"" value=""" & TempValue & """ /></td>")
            Send("</tr>")
            'ZIP
            TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_ZIP AndAlso Not AltIDFailed, MyCommon.NZ(rst.Rows(0).Item("Zip"), ""), Request.QueryString("Zip"))
            If TempValue <> "" Then
              TempValue.Replace("""", "&quot;")
            End If
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_ZIP, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Zip"">" & Copient.PhraseLib.Lookup("customer.zip", LanguageID) & GetDesignatorText("Zip", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td><input type=""text"" class=""medium"" id=""Zip"" name=""Zip"" maxlength=""20"" value=""" & TempValue & """ /></td>")
            Send("</tr>")
            'Country
            TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Country AndAlso Not AltIDFailed, MyCommon.NZ(rst.Rows(0).Item("Country"), ""), Request.QueryString("Country"))
            If TempValue <> "" Then
              TempValue.Replace("""", "&quot;")
            End If
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Country, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Country"">" & Copient.PhraseLib.Lookup("term.country", LanguageID) & GetDesignatorText("Country", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td><input type=""text"" class=""medium"" id=""Country"" name=""Country"" maxlength=""50"" value=""" & TempValue & """ /></td>")
            Send("</tr>")
            'Phone
                  TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Phone AndAlso Not AltIDFailed, MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("Phone"), "")), Request.QueryString("Phone1"))
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Phone, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Phone1"">" & Copient.PhraseLib.Lookup("customer.phone", LanguageID) & GetDesignatorText("Phone", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td style=""font-size:18px;"" >")
            Send("<input type=""text"" class=""medium"" id=""Phone1"" name=""Phone1"" maxlength=""50"" value=""" & TempValue & """ />&nbsp;")
            Send("  </td>")
            Send("</tr>")
            ' Mobile phone number
                  TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_MobilePhone AndAlso Not AltIDFailed, MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("MobilePhone"), "")), Request.QueryString("MobilePhone1"))
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_MobilePhone, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Phone1"">" & Copient.PhraseLib.Lookup("customer.mobilephone", LanguageID) & GetDesignatorText("Phone", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td style=""font-size:18px;"" >")
            Send("<input type=""text"" class=""medium"" id=""MobilePhone1"" name=""MobilePhone1"" maxlength=""50"" value=""" & TempValue & """ />&nbsp;")
            Send("  </td>")
            Send("</tr>")
            TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_Email AndAlso Not AltIDFailed, MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("Email"), "")), Request.QueryString("Email"))
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Email, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Email"">" & Copient.PhraseLib.Lookup("customer.email", LanguageID) & GetDesignatorText("email", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td><input type=""text"" class=""medium"" id=""Email"" name=""Email"" maxlength=""200"" value=""" & TempValue & """ /></td>")
            Send("</tr>")
            TempValue = IIf(Logix.UserRoles.AccessCustomerIdData_DOB AndAlso Not AltIDFailed,MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(rst.Rows(0).Item("dob"), "")), Request.QueryString("dob1") & Request.QueryString("dob2") & Request.QueryString("dob3"))
            DOBParts = ParseDateOfBirth(TempValue)
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_DOB, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""dob1"">" & Copient.PhraseLib.Lookup("customer.dateofbirth", LanguageID) & GetDesignatorText("DOB", AltIDCol, AltIDVerCol) & ":</label> </td>")
            Send("  <td style=""font-size:18px;"" >")
            Send("    <input type=""text"" style=""width:39px;"" id=""dob1"" name=""dob1"" maxlength=""2"" value=""" & DOBParts(0) & """ />&nbsp;/&nbsp;")
            Send("    <input type=""text"" style=""width:40px;"" id=""dob2"" name=""dob2"" maxlength=""2"" value=""" & DOBParts(1) & """ />&nbsp;/&nbsp;")
            Send("    <input type=""text"" class=""shorter"" id=""dob3"" name=""dob3"" maxlength=""4"" value=""" & DOBParts(2) & """ />")
            Send("  </td>")
            Send("</tr>")
            Dim tmpPass As String
            tmpPass = IIf(AltIDFailed, Password, MyCommon.NZ(rst.Rows(0).Item("Password"), ""))
            If (tmpPass <> "" AndAlso Not AltIDFailed) Then
              tmpPass = MyCryptLib.SQL_StringDecrypt(tmpPass)
            End If
            Send("<tr" & IIf(Logix.UserRoles.AccessCustomerIdData_Password, "", " style=""display:none;""") & ">")
            Send("  <td><label for=""Password"">" & Copient.PhraseLib.Lookup("term.password", LanguageID) & ":</label> </td>")
            Send("  <td><input type=""text"" class=""medium"" id=""Password"" name=""Password"" maxlength=""" & MAX_CUST_PASSWORD_CLEARTEXT_LEN & """ value=""" & tmpPass & """ /></td>")
            Send("</tr>")
            Send("</table>")
            If (Logix.UserRoles.EditCustomerIdData) Then
              Send_Save()
              Send("<input type=""button"" class=""regular"" id=""cancel"" name=""cancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""window.location.href='CAM-customer-general.aspx?CustPK=" & CustomerPK & extraLink & "';"" />")
            End If
            'MyCommon.NZ(rst.Rows(0).Item("Employee"), UnknownPhrase) & "<br />")
          End If
        End If
        Send("<hr class=""hidden"" />")
      %>
    </div>
    <% End If%>
    <div class="box" id="recentnotes">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.notes", LanguageID))%>
        </span>
      </h2>
      <%
        MyCommon.QueryStr = "select NoteID, CustomerPK, AdminUserID, CreatedDate, Note, FirstName, LastName, Private, Important " & _
                            "from CustomerNotes CN with (NoLock) " & _
                            "where CustomerPK = " & CustomerPK & " order by CreatedDate DESC;"
        rst3 = MyCommon.LXS_Select()
        If rst3.Rows.Count > 0 Then
          If rst3.Rows.Count > 10 Then
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
          For Each row3 In rst3.Rows
            If i > 10 Then
              GoTo closenotes
            Else
              Send("    <tr" & Shaded & ">")
              Send("      <td>" & MyCommon.NZ(row3.Item("CreatedDate"), "") & "</td>")
              If IsDBNull(row3.Item("FirstName")) AndAlso IsDBNull(row3.Item("LastName")) Then
                MyCommon.QueryStr = "select FirstName, LastName from AdminUsers where AdminUserID=" & row3.Item("AdminUserID") & ";"
                rst4 = MyCommon.LRT_Select()
                Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst4.Rows(0).Item("FirstName"), ""), 25) & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst4.Rows(0).Item("LastName"), ""), 25) & "</td>")
              Else
                Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("FirstName"), ""), 25) & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("LastName"), ""), 25) & "</td>")
              End If
              Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Note"), ""), 40) & "</td>")
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
          Send("  </tbody>")
          Send("</table>")
        Else
          Send(Copient.PhraseLib.Lookup("customer.nonotesposted", LanguageID) & "<br />")
        End If
      %>
    </div>
  </div>
  <br clear="all" />
</div>
</form>
<script runat="server">
  Function GetCustomerPK(ByRef MyCommon As Copient.CommonInc, ByVal ExtCardID As String) As Integer
    Dim dt As DataTable
        Dim CustomerPK As Integer = 0
        Dim MyCryptLib As New Copient.CryptLib
    
    If (Not MyCommon.LXSadoConn.State = ConnectionState.Open) Then MyCommon.Open_LogixXS()
    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, 2)
    MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where CardTypeID=2 and ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "';"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
    End If
    
    Return CustomerPK
  End Function
  
  Function GetIDValue(ByRef MyCommon As Copient.CommonInc, ByVal CustomerPK As Long, ByVal SystemOptionID As Integer) As Object
    Dim IDValue As Object = Nothing
    Dim IDField As String = ""
    Dim SysOptValue As String = ""
    Dim TableName As String = ""
    Dim FieldName As String = ""
    Dim DelimPos As Integer = -1
    Dim dt As DataTable
    
    If (Not MyCommon.LXSadoConn.State = ConnectionState.Open) Then MyCommon.Open_LogixXS()
    If (Not MyCommon.LRTadoConn.State = ConnectionState.Open) Then MyCommon.Open_LogixRT()
    
    SysOptValue = MyCommon.Fetch_SystemOption(SystemOptionID)
    DelimPos = SysOptValue.IndexOf(".")
    
    If (SysOptValue <> "" AndAlso DelimPos > 0) Then
      TableName = SysOptValue.Substring(0, DelimPos)
      FieldName = SysOptValue.Substring(DelimPos + 1)
      ' determine if the Alt ID system option value and ID Verifier are valid fields in the CustomerExt table
      MyCommon.QueryStr = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS " & _
                          "WHERE TABLE_NAME ='" & TableName & "'and COLUMN_NAME = '" & FieldName & "';"
      dt = MyCommon.LXS_Select
      If (dt.Rows.Count > 0) Then
        MyCommon.QueryStr = "select " & FieldName & " from " & TableName & " with (NoLock) where CustomerPK=" & CustomerPK & ";"
        dt = MyCommon.LXS_Select
        If (dt.Rows.Count > 0) Then
          IDValue = MyCommon.NZ(dt.Rows(0).Item(0), Nothing)
        End If
      End If
    End If
    
    Return IDValue
  End Function
  
  Function FormatPhoneNumber(ByVal Phone As String) As String
    Dim PhoneNumber As String = Phone
    
    If (Phone IsNot Nothing) Then
      Select Case Phone.Length
        Case 7
          PhoneNumber = String.Format("{0}-{1}", Phone.Substring(0, 3), Phone.Substring(3))
        Case 10 To 1000
          PhoneNumber = String.Format("({0}) {1}-{2}", Phone.Substring(0, 3), Phone.Substring(3, 3), Phone.Substring(6))
        Case Else
          PhoneNumber = Phone
      End Select
    End If
    
    Return PhoneNumber
  End Function
  
  Function ParsePhoneNumber(ByVal PhoneNumber As String) As String()
    Dim PhoneParts() As String = {"", "", ""}
    
    If (PhoneNumber IsNot Nothing) Then
      Select Case PhoneNumber.Length
        Case 1 To 3
          PhoneParts(0) = PhoneNumber
          PhoneParts(1) = ""
          PhoneParts(2) = ""
        Case 4
          PhoneParts(0) = ""
          PhoneParts(1) = ""
          PhoneParts(2) = PhoneNumber
        Case 5 To 7
          PhoneParts(0) = ""
          PhoneParts(1) = PhoneNumber.Substring(0, 3)
          PhoneParts(2) = PhoneNumber.Substring(3)
        Case 8 To 1000
          PhoneParts(0) = PhoneNumber.Substring(0, 3)
          PhoneParts(1) = PhoneNumber.Substring(3, 3)
          PhoneParts(2) = PhoneNumber.Substring(6)
      End Select
    End If
    
    Return PhoneParts
  End Function
  
  Function ParseDateOfBirth(ByVal DateOfBirth As String) As String()
    Dim DOBParts() As String = {"", "", ""}
    
    If (DateOfBirth IsNot Nothing) Then
      Select Case DateOfBirth.Length
        Case 4
          DOBParts(0) = ""
          DOBParts(1) = ""
          DOBParts(2) = DateOfBirth
        Case 8
          DOBParts(0) = DateOfBirth.Substring(0, 2)
          DOBParts(1) = DateOfBirth.Substring(2, 2)
          DOBParts(2) = DateOfBirth.Substring(4)
      End Select
    End If
    
    Return DOBParts
    
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
  
  Function GetDesignatorText(ByVal Field As String, ByVal AltID As String, ByVal Verifier As String) As String
    Dim Tag As String = ""
    
    If (AltID.ToUpper = Field.ToUpper) Then
      Tag = " <span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.alternateid", LanguageID) & ")</span>"
    ElseIf (Verifier.ToUpper = Field.ToUpper) Then
      Tag = " <span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.alternate-id-verifier", LanguageID) & ")</span>"
    End If
    
    Return Tag
  End Function
  
  Function GetNewAltID(ByVal AltIDColumn As String, ByRef FieldName As String) As String
    Dim NewAltID As String = ""
    
    Select Case AltIDColumn.ToUpper
      Case "PHONE"
        NewAltID = Request.QueryString("Phone1") & Request.QueryString("Phone2") & Request.QueryString("Phone3")
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
</script>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
