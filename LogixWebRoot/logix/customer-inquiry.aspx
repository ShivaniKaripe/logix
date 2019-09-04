<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="Copient.commonShared" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS" %>

<script runat="server">
  '
  ' $Id: customer-inquiry.aspx 131150 2018-11-15 20:35:16Z ac185320 $
  '    
 
  
  Function shouldDoCustomizedCustomerInquiry(ByRef MyCommon As Copient.CommonInc) As Boolean
    Const USE_CUSTOMIZED_CUSTOMER_INQUIRY As Integer = 107
    Return (MyCommon.Fetch_SystemOption(USE_CUSTOMIZED_CUSTOMER_INQUIRY) = 1)
  End Function
  
  Function validateCard(ByVal card As String, ByRef MyCommon As Copient.CommonInc) As String
    Const PHYSICAL_CARD_LENGTH As Integer = 12
    Const MEMBER_ID_LENGTH As Integer = 15
    card = Trim(card)
        
    Dim cardConverter As New Copient.CustomizedCustomerInquiry(MyCommon.Get_Install_Path() & "/AgentFiles/CustomizedCustomerInquiryCard.config")
    If (card.Length = PHYSICAL_CARD_LENGTH And IsNumeric(card)) Then
      card = cardConverter.getMemberIdFromCardNumber(card)
    ElseIf (card.Length = MEMBER_ID_LENGTH And IsNumeric(card)) Then
      Dim physical_card As String = cardConverter.getCardNumberFromMemberId(Long.Parse(card))
    Else
      Throw New ArgumentException(String.Format("{0} ({1})", Copient.PhraseLib.Lookup("term.invalid-cust-specific-card-number", LanguageID), card))
    End If
    Return card
  End Function
    
    
  Function transformCard(ByVal card As String, ByVal cardType As String, ByRef MyCommon As Copient.CommonInc) As String
    Const STANDARD_CUSTOMER_CARD_TYPEID As String = "0"
    
    If (cardType = STANDARD_CUSTOMER_CARD_TYPEID AndAlso Not isEmpty(card) AndAlso shouldDoCustomizedCustomerInquiry(MyCommon)) Then
      Try
        
        Return validateCard(card, MyCommon)
        
      Catch ex As Exception ' ensure any exceptions get logged
        MyCommon.Write_Log(MyCommon.LogPath & "/customerinquiry.txt", ex.Message, True)
        Throw
      End Try
    End If ' card is not empty and should do the transform/validation 
    
    Return card
  End Function
  
  Function ReplaceSpecialChar(ByVal inputString As String)
    'Replacing % with [%] will allow sql to search for % when using a like.
    'Replacing _ with [_] will allow sql to search for _ when using a like.
    Return inputString.Trim().Replace("%", "[%]").Replace("_", "[_]")
  End Function

  Function DetermineSearchTypeID(ByVal SearchTypeID As Integer, ByRef MyCommon As Copient.CommonInc, Logix As Copient.LogixInc) As Integer
    MyCommon.QueryStr = "select SearchTypeID, Name, PhraseID, Enabled from CustomerSearchTypes with (NoLock) "
    If Not Logix.UserRoles.AccessCustomerIdData_LastName Then
      'No permission to see last name, so disallow searches on it
      MyCommon.QueryStr &= "where SearchTypeID<>4 "
    End If
    MyCommon.QueryStr &= "order by SearchTypeID;"
    Dim rst As DataTable = MyCommon.LXS_Select
    For Each row In rst.Rows
      If MyCommon.NZ(row.Item("SearchTypeID"), 0) = SearchTypeID AND MyCommon.NZ(row.Item("Enabled"), false) Then
         Return SearchTypeID
      End If
    Next

    Return 1
  End Function

</script>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-inquiry.aspx 
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
  Dim Logix As New Copient.LogixInc
  Dim MyCAM As New Copient.CAM
    Dim CAMCustomer As New Copient.Customer
  Dim dt As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim row3 As DataRow
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim ExtID As String
  Dim ExtCustomerID As String = ""
  Dim Edit As Boolean
  Dim i As Integer = 0
  Dim IsHouseholdID As Boolean = False
  Dim HHPK As Integer = 0
  Dim HouseholdID As String = ""
  Dim Shaded As String = " class=""shaded"""
  Dim HasSearchResults As Boolean = False
  Dim FullName As String = ""
  Dim FullAddress As String = ""
  Dim CustomerTypeID As Integer = 0
  Dim CardTypeID As Integer = 0
  Dim OfferID As Integer = 0
  Dim ClientUserID1 As String = ""
  Dim IDLength As Integer = 0
  Dim CustomerGroupIDs As String() = Nothing
  Dim loopCtr As Integer = 0
  Dim searchterms As String = ""
  Dim restrictLinks As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim SearchTypeID As Integer = 0
  Dim ShowAll As Boolean = False
  Dim ShowBannerCol As Boolean = False
  Dim Coupon As Boolean = False
  Dim PageName As String = ""
  
  ' default urls for links from this page
  Dim URLOfferSum As String = "offer-sum.aspx"
  Dim URLCPEOfferSum As String = "CPEoffer-sum.aspx"
  Dim URLcgroupedit As String = "cgroup-edit.aspx"
  Dim URLpointedit As String = "point-edit.aspx"
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  Dim inSearchTypeID As Integer = 0
  Dim DefaultSearchTypeID As Integer = 0
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  
  Dim WhereBuf As New StringBuilder
  Dim CAMErrorMessage As String = ""
  
  Response.Expires = 0
  MyCommon.AppName = "customer-inquiry.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  Dim MyLookup As New Copient.CustomerLookup(MyCommon)
  Dim Cust As Copient.Customer
  Dim CustExt As Copient.CustomerExt
  Dim Customers(-1) As Copient.Customer
  Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
  Dim EmptyCriteria As Boolean
  Dim SessionID As String = ""
  Dim HasCustomerID As Boolean = True
  Dim ExtraCardTypeID As Integer = Nothing
  Dim ValidSearchType As Boolean = False
  Dim CardType As Integer = 0
  Dim CustSearchType As Copient.CustomerAbstract.SEARCH_TYPE = Copient.CustomerAbstract.SEARCH_TYPE.ALL_CUSTOMER_TYPES
  Dim FoundRecordCount As Int64 = 0
  Dim CustRow As Integer = 0
    Const CustPerRow As Integer = 20
    
    Dim SystemCacheData As ICacheData
    Dim isEncryptedPIData As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  

  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    CurrentRequest.Resolver.AppName = "customer-inquiry.aspx"
    SystemCacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()
    
  MyLookup.SetAdminUserID(AdminUserID)
  MyLookup.SetLanguageID(LanguageID)
  restrictLinks = MyLookup.IsRestrictedUser(AdminUserID)
  
  DefaultSearchTypeID = MyCommon.Fetch_SystemOption(30) + 1
    
  CustomerPK = 0
    If Server.HtmlEncode(Request.QueryString("searchby")) = "6" Then
    Coupon = True
  End If
  
  Send_HeadBegin("term.customerinquiry")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
  Send_Scripts()
%>
<script type="text/javascript">
  var phone_field_length = 0;

  function openCoupons(CouponID, ProgramID) {
    openPopup('sv-adjust-program.aspx?CouponID=' + CouponID + '&ProgramID=' + ProgramID + '&CustomerPK=0&Opener=customer-inquiry.aspx&Coupon=coupon');
  }

  function searchNew() {
    var elem = document.getElementById("CustomerPK");

    if (elem != null) {
      elem.value = ""
    }
  }

  function showCriteriaFor(type, doFocus) {
    var elemInput = null;
    var elemSelector = null;

    elemInput = document.getElementById("trCH");
    if (elemInput != null) elemInput.style.display = (type == 1) ? "" : "none";
    elemSelector = document.getElementById("CardTypeID0");
    if (elemSelector != null) {
      if (type == 1) {
        elemSelector.removeAttribute("disabled");
      } else {
        elemSelector.disabled = true;
      }
      elemSelector.style.display = (type == 1) ? "" : "none";
    }

    elemInput = document.getElementById("trHH");
    if (elemInput != null) elemInput.style.display = (type == 2) ? "" : "none";
    elemSelector = document.getElementById("CardTypeID1");
    if (elemSelector != null) {
      if (type == 2) {
        elemSelector.removeAttribute("disabled");
      } else {
        elemSelector.disabled = true;
      }
      elemSelector.style.display = (type == 2) ? "" : "none";
    }

    elemInput = document.getElementById("trAltID");
    if (elemInput != null) elemInput.style.display = (type == 3) ? "" : "none";

    elemInput = document.getElementById("trName");
    if (elemInput != null) elemInput.style.display = (type == 4) ? "" : "none";

    elemInput = document.getElementById("trPhone");
    if (elemInput != null) elemInput.style.display = (type == 5) ? "" : "none";

    elemInput = document.getElementById("trCoupon");
    if (elemInput != null) elemInput.style.display = (type == 6) ? "" : "none";

    elemInput = document.getElementById("trCAM");
    if (elemInput != null) elemInput.style.display = (type == 7) ? "" : "none";
    
    elemInput = document.getElementById("trLastNamePartial");
    if (elemInput != null) elemInput.style.display = (type == 8) ? "" : "none";
    
    elemSelector = document.getElementById("CardTypeID2");
    if (elemSelector != null) {
      if (type == 7) {
        elemSelector.removeAttribute("disabled");
      } else {
        elemSelector.disabled = true;
      }
      elemSelector.style.display = (type == 7) ? "" : "none";
    }
        
    if (doFocus) {
      setFocusCtrl(type);
    }
  }

  function setFocusCtrl(type) {
    if (type == 1) {
      elemName = "cardID";
    } else if (type == 2) {
      elemName = "hhID";
    } else if (type == 3) {
      elemName = "altID";
    } else if (type == 4) {
      elemName = "lastname";
    } else if (type == 5) {
      elemName = "phone1";
    } else if (type == 6) {
      elemName = "couponID";
    } else if (type == 7) {
      elemName = "camID"
    } else {
      elemName = "cardID";
    }
    focusElem = document.getElementById(elemName);
    if (focusElem != null) {
      focusElem.focus();
      focusElem.select();
    }
  }

  function openOffers() {
    window.location = 'offer-list.aspx?CustomerInquiry=1';
  }

  function openReports() {
    window.location = 'customer-reports.aspx';
  }

  function openPending() {
    window.location = 'point-adjust-pending.aspx';
  }

  function openManual() {
    window.location = '/logix/CAM/CAM-customer-manual.aspx';
  }

  
  function launchAdvSearch() {
    self.name = "CustomerInquiryAdvSearchWin";
    <%
        Send("openPopup(""customer-inquiry-adv-search.aspx"");")
    %>
  }

</script>
<%
  If Coupon Then
        MyCommon.QueryStr = "select SVProgramID from StoredValue where CustomerPK=0 and ExternalID='" & Server.HtmlEncode(Request.QueryString("couponID")) & "';"
    rst2 = MyCommon.LXS_Select
    If rst2.Rows.Count > 0 Then
      Send("<script type=""text/javascript"">")
            Send("  openCoupons('" & Server.HtmlEncode(Request.QueryString("couponID")) & "', '" & MyCommon.NZ(rst2.Rows(0).Item("SVProgramID"), 0) & "');")
      Send("</script>")
    Else
      infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.CouldNotFindCoupon", LanguageID)
    End If
  End If
  
  Send_HeadEnd()
  
  ' special handling for customer inquery direct link in 
  If (restrictLinks) Then
    URLOfferSum = ""
    URLCPEOfferSum = ""
    URLcgroupedit = ""
    URLpointedit = ""
  End If
  
  ' set session to nothing just to be sure
  Session.Add("extraLink", "")
  
    If (Server.HtmlEncode(Request.QueryString("mode")) = "summary") Then
        
        URLtrackBack = Server.HtmlEncode(Request.QueryString("exiturl"))
        inCardNumber = Server.HtmlEncode(Request.QueryString("cardnumber"))
        inSearchTypeID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("searchtype")))
    If inSearchTypeID = 0 Then inSearchTypeID = 1
        If (Server.HtmlEncode(Request.QueryString("CardTypeID")) <> "") Then
            CardType = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID")))
    End If
    extraLink = "&mode=summary&exiturl=" & URLtrackBack & "&cardnumber=" & inCardNumber & "&searchtype=" & inSearchTypeID & "&CardTypeID=" & CardType
    Session.Add("extraLink", extraLink)
    If Session("SessionID") IsNot Nothing Then Session.Remove("SessionID")

    ElseIf (Server.HtmlEncode(Request.QueryString("mode")) = "addCAM") Then
        ClientUserID1 = MyCommon.Pad_ExtCardID(Server.HtmlEncode(Request.QueryString("number")), 2)
    CustomerPK = MyCAM.AddCustomer(ClientUserID1, AdminUserID, CAMErrorMessage)
    If CustomerPK > 0 Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "CAM/CAM-customer-general.aspx?CustPK=" & CustomerPK & "&CardPK=" & CardPK)
    Else
      infoMessage = CAMErrorMessage
    End If

  End If
  
    If (Server.HtmlEncode(Request.QueryString("showall")) <> "") Then
    ShowAll = True
  End If
  
  ' Hack for popups check session for extralink
  If (Session("extraLink").ToString = "") Then
    extraLink = Session("extraLink")
  End If
  
  ' SEARCHING FOR A CUSTOMER ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    If (Server.HtmlEncode(Request.QueryString("Search")) <> "" Or Server.HtmlEncode(Request.QueryString("searchPressed")) <> "") Then
  
    If Session("SessionID") IsNot Nothing Then
      Session.Remove("SessionID")
    End If

  End If
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  
    If (MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CustPK"))) > 0 Or (Server.HtmlEncode(Request.QueryString("Search")) <> "" Or Server.HtmlEncode(Request.QueryString("searchPressed")) <> "") Or inCardNumber <> "") Then
        
        If (Server.HtmlEncode(Request.QueryString("mode")) = "add") Then

      Cust = New Copient.Customer()
      Dim convertedSearchType As Integer = 0
      
      If inSearchTypeID = 1 Then 'customer
        convertedSearchType = 0
      ElseIf inSearchTypeID = 2 Then 'household
        convertedSearchType = 1
      ElseIf inSearchTypeID = 7 Then 'CAM
        convertedSearchType = 2
      End If
      
            If (Server.HtmlEncode(Request.QueryString("CardTypeID")) <> "") Then
                CardType = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID"))) ' defaulted to 0 above 
      End If

      Try

        Dim newCardNum As String = transformCard(IIf(Server.HtmlEncode(Request.QueryString("cardnumber")) <> "", Server.HtmlEncode(Request.QueryString("cardnumber")), Server.HtmlEncode(Request.QueryString("number"))), CardType, MyCommon)
        If MyCommon.AllowToProcessCustomerCard(newCardNum, CardType, Nothing) Then
          Cust.AddCard(New Copient.Card(newCardNum, CardType, 0))
          If MyLookup.AddCustomer(Cust, ReturnCode) Then
             CustomerPK = Cust.GetCustomerPK
          Else
             If ReturnCode = Copient.CustomerAbstract.RETURN_CODE.INVALID_REWARDCARD Then
                infoMessage = Copient.PhraseLib.Lookup("customer.invalidcard", LanguageID) & " (" & newCardNum & ")"
             End If
          End If
        Else
          infoMessage = Copient.PhraseLib.Lookup("term.invalidnumericcard", LanguageID)
        End If

      Catch e As ArgumentException
        infoMessage = e.Message
        CustomerPK = 0 ' ensure that's set to 0 so we don't do anything weird with it
      End Try
                
    End If
    
    
    
    ' Someone wants to search for a customer.  First let's get their primary key from our database
    If (CustomerPK > 0 Or MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CustPK"))) > 0 Or (MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CustomerPK"))) > 0)) Then
            
      If CustomerPK = 0 Then
                If (MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CustPK"))) > 0) Then
                    CustomerPK = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CustPK")))
                ElseIf (MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CustomerPK"))) > 0) Then
                    CustomerPK = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CustomerPK")))
        End If
      End If
      ReDim Customers(0)
      Customers(0) = MyLookup.FindCustomerInfo(CustomerPK, ReturnCode)
      If ReturnCode <> Copient.CustomerAbstract.RETURN_CODE.OK Then
        ReDim Customers(-1)
      End If
    Else
      
            If (inCardNumber <> "" AndAlso Server.HtmlEncode(Request.QueryString("mode")) = "summary") Then

        ' register the session ID used for grouping customer inquiry activities
                If Server.HtmlEncode(Request.QueryString("SessionID")) <> "" Then
                    SessionID = Server.HtmlEncode(Request.QueryString("SessionID")).Trim
          If Session("SessionID") IsNot Nothing Then
            Session("SessionID") = SessionID
          Else
            Session.Add("SessionID", SessionID)
          End If
        End If
        ' The page was called from an outside application, so set ClientUserID1 to the outside passed-in value
        ClientUserID1 = MyCommon.Pad_ExtCardID(inCardNumber, Copient.commonShared.CardTypes.CUSTOMER)
                searchterms = Server.HtmlEncode(Request.QueryString("searchterms"))
        ' Assign the requested search type if it's valid
        If (inSearchTypeID < 0 Or (inSearchTypeID > 2 AndAlso Not inSearchTypeID = 7)) Then
          infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.InvalidSearchType", LanguageID)
        ElseIf inSearchTypeID = 2 Then
          ExtraCardTypeID = 1
          ValidSearchType = True
        ElseIf inSearchTypeID = 1 Then
                    If (Server.HtmlEncode(Request.QueryString("CardTypeID")) <> "") Then
                        ExtraCardTypeID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID")))
            ValidSearchType = True
          Else
            ExtraCardTypeID = 0
            ValidSearchType = True
          End If
        ElseIf inSearchTypeID = 7 Then
          ExtraCardTypeID = 2
          ValidSearchType = True
        End If
        If ValidSearchType Then
          If (inSearchTypeID >= 1 AndAlso inSearchTypeID <= 7) OrElse inSearchTypeID = 99 Then
            CustSearchType = inSearchTypeID
            'ExtraCardTypeID is defaulted to 0 if not set to maintain the original value
            Customers = MyLookup.FindCustomerMatches(CustSearchType, IIf(ExtraCardTypeID = Nothing, 0, ExtraCardTypeID), ClientUserID1, ReturnCode)
          Else
            'ExtraCardTypeID is defaulted to 0 if not set to maintain the original value
            Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.ALL_CUSTOMER_TYPES, IIf(ExtraCardTypeID = Nothing, 0, ExtraCardTypeID), ClientUserID1, ReturnCode)
          End If
        End If

            End If ' ( inCardNumber <> "" AndAlso Server.HtmlEncode(Request.QueryString("mode") = "summary")

      
            
      If (Server.HtmlEncode(Request.QueryString("search")) <> "" And ClientUserID1 = "" And Not Coupon) Then
                
                If (Server.HtmlEncode(Request.QueryString("mode")) = "custominquiryadvancedsearch") Then
          Try
        
                        Dim cardIDType As String = ifEmptyString(Server.HtmlEncode(Request.QueryString("CardTypeID")), "0")
            Dim tmpCardID As String = transformCard(ifEmptyString(Server.HtmlEncode(Request.QueryString("cardID")), ""), cardIDType, MyCommon)
            tmpCardID = MyCommon.Pad_ExtCardID(tmpCardID, Convert.ToInt32(cardIDType))

                        Dim tmpEmail As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("Email")), ""))
                        Dim tmpAddress As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("Address")), ""))
                        Dim tmpCity As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("City")), ""))
                        Dim tmpState As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("State")), ""))
                        Dim tmpZip As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("Zip")), ""))
                        Dim tmpPhone As String = ifEmptyString(Server.HtmlEncode(Request.QueryString("phone1")), "")
                        Dim tmpFirstName As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("firstname")), ""))
                        Dim tmpLastName As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("lastname")), ""))
                        Dim tmpLastNamePartial As String = ReplaceSpecialChar(ifEmptyString(Server.HtmlEncode(Request.QueryString("lastnamepartial")), ""))

            EmptyCriteria = Not (tmpCardID <> "" OrElse tmpEmail <> "" OrElse tmpAddress <> "" OrElse tmpCity <> "" OrElse tmpState <> "" OrElse tmpZip <> "" OrElse tmpPhone <> "" OrElse tmpFirstName <> "" OrElse tmpLastName <> "" OrElse tmpLastNamePartial <> "")
            
            If (Not EmptyCriteria) Then
              FoundRecordCount = MyLookup.FindCustomerMatchesCount(isEncryptedPIData,"CardID", tmpCardID, "Email", tmpEmail, "Address", tmpAddress, "City", tmpCity, "State", tmpState, "Zip", tmpZip, "Phone", tmpPhone, "FirstName", tmpFirstName, "LastName", tmpLastName, "LastNamePartial", tmpLastNamePartial)
                            Dim NextRow As Integer = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("row")))
              If NextRow = 0 Then NextRow = 1
              CustRow = NextRow + CustPerRow 'Move the row number to the next page
              Customers = MyLookup.FindCustomerMatchesLimited(Copient.CustomerAbstract.SEARCH_TYPE.MULTIFIELD, cardIDType, ReturnCode, NextRow, CustPerRow,isEncryptedPIData, _
              "CardID", tmpCardID, "Email", tmpEmail, "Address", tmpAddress, "City", tmpCity, "State", tmpState, "Zip", tmpZip, "Phone", tmpPhone, "FirstName", tmpFirstName, "LastName", tmpLastName, "LastNamePartial", tmpLastNamePartial)
            Else
              If (infoMessage = "") Then
                infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.enter-criteria", LanguageID)
              End If
            End If

          Catch e As ArgumentException
            infoMessage = e.Message
          End Try
        
                ElseIf (Server.HtmlEncode(Request.QueryString("searchby")) <> "") Then
        
                    SearchTypeID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("searchby")))
                    isEncryptedPIData = SystemCacheData.GetCustomerSearchType(SearchTypeID).IsEncrypted
          Select Case SearchTypeID

            Case 1 ' card number
                            EmptyCriteria = isEmpty(Server.HtmlEncode(Request.QueryString("cardID")))
              If Not EmptyCriteria Then
                
                Try
                                    ClientUserID1 = transformCard(Request.QueryString("cardID").Trim, MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID"))), MyCommon)
                                    Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.CARDHOLDER_ID, ClientUserID1, 0, MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID"))), ReturnCode, isEncryptedPIData)
                Catch e As ArgumentException
                  infoMessage = e.Message
                End Try

              End If

            Case 2 ' Household id
                            EmptyCriteria = isEmpty(Server.HtmlEncode(Request.QueryString("hhID")))
              If Not EmptyCriteria Then
                ClientUserID1 = Server.HtmlEncode(Request.QueryString("hhID")).Trim
                                Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.HOUSEHOLD_ID, ClientUserID1, 1, MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID"))), ReturnCode, isEncryptedPIData)
              End If

            Case 3 ' Alternate ID
                            EmptyCriteria = isEmpty(Server.HtmlEncode(Request.QueryString("altID")))
              If Not EmptyCriteria Then
                                Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.ALT_ID, Server.HtmlEncode(Request.QueryString("altID")).Trim, -1, -1, ReturnCode,isEncryptedPIData)
              End If

            Case 4 ' Last name
                            FoundRecordCount = MyLookup.FindCustomerMatchesCount(isEncryptedPIData,"LastName", Server.HtmlEncode(Request.QueryString("lastname")).Trim)
                            EmptyCriteria = isEmpty(Server.HtmlEncode(Request.QueryString("lastname")))
              If Not EmptyCriteria Then
                If Not ShowAll Then
                                    Dim NextRow As Integer = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("row")))
                  If NextRow = 0 Then NextRow = 1
                  CustRow = NextRow + CustPerRow 'Move the row number to the next page
                                    Customers = MyLookup.FindCustomerMatchesLimited(Copient.CustomerAbstract.SEARCH_TYPE.LAST_NAME, -1, ReturnCode, NextRow, CustPerRow,isEncryptedPIData, "LastName", Server.HtmlEncode(Request.QueryString("lastname")).Trim)
                Else
                                    Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.LAST_NAME, Server.HtmlEncode(Request.QueryString("lastname")).Trim, -1, -1, ReturnCode, isEncryptedPIData)
                End If
              End If

            Case 5 ' Phone number
                            FoundRecordCount = MyLookup.FindCustomerMatchesCount(isEncryptedPIData,"Phone", Server.HtmlEncode(Request.QueryString("phone1")).Trim)
                            EmptyCriteria = isEmpty(Server.HtmlEncode(Request.QueryString("phone1")).Trim)
              If Not EmptyCriteria Then
                If Not ShowAll Then
                                    Dim NextRow As Integer = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("row")))
                  If NextRow = 0 Then NextRow = 1
                  CustRow = NextRow + CustPerRow 'Move the row number to the next page
                  Customers = MyLookup.FindCustomerMatchesLimited(Copient.CustomerAbstract.SEARCH_TYPE.PHONE, -1, ReturnCode, NextRow, CustPerRow,isEncryptedPIData, "Phone", Server.HtmlEncode(Request.QueryString("phone1")).Trim)
                Else
                                    Customers = MyLookup.FindCustomerMatches(Copient.CustomerAbstract.SEARCH_TYPE.PHONE, Server.HtmlEncode(Request.QueryString("phone1")).Trim, -1, -1, ReturnCode, isEncryptedPIData)
                End If
              End If

            Case 6 ' ????
              EmptyCriteria = False
              Coupon = True

            Case 7 ' CAM Card ID
                            If Server.HtmlEncode(Request.QueryString("camID")) = "" Then
                infoMessage = CAMErrorMessage
                            ElseIf MyCAM.VerifyCardNumber(Server.HtmlEncode(Request.QueryString("camID")), CAMErrorMessage) Then
                                CAMCustomer = MyCAM.GetCustomer(Server.HtmlEncode(Request.QueryString("camID")))
                If CAMCustomer IsNot Nothing AndAlso CAMCustomer.GetCustomerPK > 0 Then
                                    ClientUserID1 = Server.HtmlEncode(Request.QueryString("camID"))
                  CardPK = MyLookup.FindCardPK(ClientUserID1, 2)
                  ' redirect CAM customer to the CAM customer inquiry page
                  Response.Status = "301 Moved Permanently"
                  Response.AddHeader("Location", "CAM/CAM-customer-general.aspx?CustPK=" & CAMCustomer.GetCustomerPK & "&CardPK=" & CardPK)
                  GoTo done
                Else
                                    infoMessage = Copient.PhraseLib.Lookup("cam.unable-to-find-customer", LanguageID) & " """ & Server.HtmlEncode(Request.QueryString("camID")) & """. &nbsp;&nbsp;&nbsp;" & Copient.PhraseLib.Lookup("term.searchagain", LanguageID)
                  If Logix.UserRoles.CreateCustomer AndAlso MyCommon.Fetch_SystemOption(103) Then
                                        infoMessage &= " " & Copient.PhraseLib.Lookup("term.or", LanguageID) & " <a style=""color:#0088ff;"" href=""customer-inquiry.aspx?mode=addCAM&number=" & Server.HtmlEncode(Request.QueryString("camID")) & "&CardTypeID=" & CardTypeID & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
                  End If
                End If
              Else
                infoMessage = CAMErrorMessage
              End If
              
            Case 8 ' Last name(partial)
                            FoundRecordCount = MyLookup.FindCustomerMatchesCount(isEncryptedPIData,"LastNamePartial", Server.HtmlEncode(Request.QueryString("lastnamepartial")).Trim)
                            Dim NextRow As Integer = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("row")))
              If NextRow = 0 Then NextRow = 1
              CustRow = NextRow + CustPerRow 'Move the row number to the next page
                            EmptyCriteria = isEmpty(Server.HtmlEncode(Request.QueryString("lastnamepartial")))
              If Not EmptyCriteria Then
                                Customers = MyLookup.FindCustomerMatchesLimited(Copient.CustomerAbstract.SEARCH_TYPE.LAST_NAME_PARTIAL, -1, ReturnCode, NextRow, CustPerRow,isEncryptedPIData, "LastNamePartial", Server.HtmlEncode(Request.QueryString("lastnamepartial")).Trim)
              End If

          End Select
                    
        Else
          ' IIf is added to prevent any Existing Error Message from getting modified
          infoMessage = IIf(String.IsNullOrEmpty(infoMessage), Copient.PhraseLib.Lookup("customer-inquiry.select-type", LanguageID), infoMessage)
        End If
                
        If (EmptyCriteria And infoMessage = "") Then
          infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.enter-criteria", LanguageID)
        End If

      Else

                If MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("searchtype"))) > 0 Then
                    SearchTypeID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("searchtype")))
        Else
          SearchTypeID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("searchby")))
        End If

      End If

    End If ' 

    
        
    If (infoMessage = "" AndAlso Not Coupon) Then
      
      If (Customers.Length > 1) Then ' more than one match was found 
      
        HasSearchResults = True
      
      ElseIf (Customers.Length = 1) Then ' only one match was found
      
        CustomerPK = Customers(0).GetCustomerPK
        CustomerTypeID = Customers(0).GetCustomerTypeID
                CardTypeID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID")))
        'MyCommon.Write_Log("custinq.txt", "CustomerPK=" & CustomerPK & ", CustomerTypeID=" & CustomerTypeID & ", CardTypeID=" & CardTypeID & vbCrLf)
        If ClientUserID1 <> "" Then
          'If the ExtraCardTypeID is set, use this value but otherwise use the card type ID
          CardPK = MyLookup.FindCardPK(ClientUserID1, IIf(ExtraCardTypeID = Nothing, CardTypeID, ExtraCardTypeID))
        Else
          CardPK = MyLookup.FindCardPK(CustomerPK, CardTypeID)
        End If
        Response.Status = "301 Moved Permanently"
        If (MyCommon.Fetch_SystemOption(79) = 0) Then
          If CustomerTypeID = 2 Then
            PageName = "/logix/CAM/CAM-customer-general.aspx"
          Else
            PageName = "/logix/customer-general.aspx"
          End If
          If CardPK = 0 Then
            MyCommon.QueryStr = "select top 1 CardPK from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by CardTypeID;"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
              CardPK = MyCommon.NZ(dt.Rows(0).Item("CardPK"), 0)
            End If
          End If
          If (CustomerTypeID <> 2 And MyCommon.Fetch_SystemOption(107) = 1) Then
            Response.AddHeader("Location", PageName & "?edit=Edit&editterms=" & CustomerPK & "&CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, ""))
          Else
            Response.AddHeader("Location", PageName & "?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, ""))
          End If
        ElseIf (MyCommon.Fetch_SystemOption(79) = 1) Then
          If CustomerTypeID = 2 Then
            PageName = "/logix/CAM/CAM-customer-offers.aspx"
          Else
            PageName = "/logix/customer-offers.aspx"
          End If
          If CardPK = 0 Then
            MyCommon.QueryStr = "select top 1 CardPK from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by CardTypeID;"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
              CardPK = MyCommon.NZ(dt.Rows(0).Item("CardPK"), 0)
            End If
          End If
          Response.AddHeader("Location", PageName & "?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, ""))
        End If
        GoTo done
      
      Else ' no matches were found 
        
        If (SearchTypeID = 1 OrElse SearchTypeID = 2) Then ' if it was a card number or household id search
                    Dim IDToAdd As String = IIf(SearchTypeID = 1, Server.HtmlEncode(Request.QueryString("CardID")), Server.HtmlEncode(Request.QueryString("HHID")))
                    If (Server.HtmlEncode(Request.QueryString("CardTypeID")) <> "") Then
                        CardTypeID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID")))
          End If
          If MyCommon.AllowToProcessCustomerCard(IDToAdd, CardTypeID, Nothing) Then
                        infoMessage = "" & Copient.PhraseLib.Lookup("customer.Customerfoundcriteria", LanguageID) & ""
            If MyCommon.Fetch_SystemOption(103) AndAlso Logix.UserRoles.CreateCustomer Then ' the user has the right to create customers and it's turned on              
              If (IsNumeric(IDToAdd)) Then
                If Int(MyCommon.Extract_Val(IDToAdd)) > 0 Then ' the card has a numeric value > 0
                  IDToAdd = transformCard(IDToAdd, CardTypeID, MyCommon)
                  IDToAdd = MyCommon.Pad_ExtCardID(IDToAdd, CardTypeID)
                  infoMessage &= " " & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & _
                  " <a href=""customer-inquiry.aspx?mode=add&Search=Search" & extraLink & "&number=" & IDToAdd & "&CardTypeID=" & CardTypeID & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
                End If ' the card has a numeric value > 0
              Else
                IDToAdd = transformCard(IDToAdd, CardTypeID, MyCommon)
                IDToAdd = MyCommon.Pad_ExtCardID(IDToAdd, CardTypeID)
                infoMessage &= " " & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & _
                     " <a href=""customer-inquiry.aspx?mode=add&Search=Search" & extraLink & "&number=" & IDToAdd & "&CardTypeID=" & CardTypeID & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
              End If
            End If
          Else
            infoMessage = Copient.PhraseLib.Lookup("term.invalidnumericcard", LanguageID)
          End If

        Else
                    If isEncryptedPIData Then
                        infoMessage = "" & Copient.PhraseLib.Lookup("customer.Customerfoundcriteria", LanguageID) & ""
                    Else
                        infoMessage = "" & Copient.PhraseLib.Lookup("customer.cardnotfoundcriteria", LanguageID) & ""
                        
                    End If
                End If
        
                HasSearchResults = False
    
      End If '
    
    End If ' If ( infoMessage = "" And Not Coupon) Then
    
  End If
  
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If (Not restrictLinks) Then
    Send_Tabs(Logix, 3)
    Send_Subtabs(Logix, 32, 2, LanguageID, CustomerPK)
  Else
    Send_Subtabs(Logix, 91, 1, LanguageID, CustomerPK, extraLink)
  End If
  
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
%>
<form id="mainform" name="mainform" action="customer-inquiry.aspx">
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.customerinquiry", LanguageID))
      If (restrictLinks AndAlso URLtrackBack <> "") Then
        Send(" - <a href=""" & URLtrackBack & """>" & Copient.PhraseLib.Lookup("customer-inquiry.return", LanguageID) & "</a>")
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) AndAlso Logix.UserRoles.AccessCAMAdjustment Then
        Send("<input type=""button"" class=""regular"" id=""manual"" name=""manual"" value=""" & Copient.PhraseLib.Lookup("term.manual", LanguageID) & """ onclick=""openManual();"" />")
      End If
      If Logix.UserRoles.FavoriteOffersForOthers Then
        Send("<input type=""button"" class=""regular"" id=""favorites"" name=""favorites"" value=""" & Copient.PhraseLib.Lookup("term.favorites", LanguageID) & """ onclick=""openOffers();"" />")
      End If
      If Logix.UserRoles.AccessCustomerInquiryReporting Then
        Send("<input type=""button"" class=""regular"" id=""reports"" name=""reports"" value=""" & Copient.PhraseLib.Lookup("term.reports", LanguageID) & """ onclick=""openReports();"" />")
      End If
    %>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Sendb(Copient.PhraseLib.Lookup("customer.main", LanguageID))%>
  <br />
  <br class="half" />
  <input type="hidden" id="CustomerPK" name="CustomerPK" value="<%Sendb(CustomerPK)%>" />
  <%
      If (Server.HtmlEncode(Request.QueryString("mode")) = "summary") Then
          Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
          Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
      End If
  %>
  <div id="column">
    <% If (Not Edit) Then%>
    <table width="100%" summary="<% Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID))%>">
      <thead>
        <tr style="visibility: hidden;">
          <th style="width: 85px;">
          </th>
          <th>
          </th>
        </tr>
      </thead>
      <tbody>
        <tr id="trSearchBy" style="display: <% Sendb(IIf(MyCommon.Fetch_SystemOption(107) = "1", "none", ""))%>">
          <td>
            <label for="searchby">
              <% Sendb(Copient.PhraseLib.Lookup("term.searchby", LanguageID))%>
              :</label>
          </td>
          <td>
            <select name="searchby" id="searchby" onchange="javascript:showCriteriaFor(this.value, true);">
              <%
                If SearchTypeID = 0 Then
                  SearchTypeID = DefaultSearchTypeID
                  SearchTypeID = DetermineSearchTypeID(SearchTypeID, MyCommon, Logix)
                End If
                MyCommon.QueryStr = "select SearchTypeID, Name, PhraseID, Enabled from CustomerSearchTypes with (NoLock) "
                If Not Logix.UserRoles.AccessCustomerIdData_LastName Then
                  'No permission to see last name, so disallow searches on it
                  MyCommon.QueryStr &= "where SearchTypeID<>4 "
                End If
                MyCommon.QueryStr &= "order by SearchTypeID;"
                rst2 = MyCommon.LXS_Select
                For Each row In rst2.Rows
                  If MyCommon.NZ(row.Item("Enabled"), False) Then
                    Send("<option value=""" & MyCommon.NZ(row.Item("SearchTypeID"), 0) & """" & IIf(SearchTypeID = MyCommon.NZ(row.Item("SearchTypeID"), 0), " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
                  End If
                Next
              %>
            </select>
            <%
              MyCommon.QueryStr = "select TypeID from CustomerTypes with (NoLock) order by TypeID;"
              rst2 = MyCommon.LXS_Select
              If rst2.Rows.Count > 0 Then
                For Each row In rst2.Rows
                  MyCommon.QueryStr = "select CardTypeID, Description, PhraseID from CardTypes with (NoLock) where CustTypeID=" & row.Item("TypeID") & " and CardTypeID <> 8 order by CardTypeID;" 'Consumer Account Number shouldn't be displayed on UI
                  rst3 = MyCommon.LXS_Select
                  If rst3.Rows.Count > 1 Then
                    Send("<select id=""CardTypeID" & row.Item("TypeID") & """ name=""CardTypeID"" style=""display:none;"" disabled=""disabled"">")
                    For Each row3 In rst3.Rows
                                Sendb("<option value=""" & MyCommon.NZ(row3.Item("CardTypeID"), 0) & """" & IIf(MyCommon.NZ(row3.Item("CardTypeID"), 0) = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("CardTypeID"))), " selected=""selected""", "") & ">")
                      If Not IsDBNull(row3.Item("PhraseID")) Then
                        Sendb(Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID))
                      Else
                        Sendb(MyCommon.NZ(row3.Item("Description"), ""))
                      End If
                      Send("</option>")
                    Next
                    Send("</select>")
                  ElseIf rst3.Rows.Count = 1 Then
                    Send("<input type=""hidden"" id=""CardTypeID" & row.Item("TypeID") & """ name=""CardTypeID"" value=""" & MyCommon.NZ(rst3.Rows(0).Item("CardTypeID"), 0) & """ disabled=""disabled"" />")
                  End If
                Next
              End If
            %>
          </td>
        </tr>
        <tr id="trCH">
          <td style="width: 95px;">
            <label for="cardID">
              <%If (MyCommon.Fetch_SystemOption(107) = "1") Then
                  Sendb(Copient.PhraseLib.Lookup("term.cust-specific-card-number", LanguageID))
                Else
                  Sendb(Copient.PhraseLib.Lookup("term.cardnumber", LanguageID))
                End If
              %>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="cardID" name="cardID" maxlength="256" value="<%Sendb(Server.HtmlEncode(Request.QueryString("cardID")))%>" />
          </td>
        </tr>
        <tr id="trHH" style="display: none;">
          <td style="width: 85px;">
            <label for="hhID">
              <% Sendb(Copient.PhraseLib.Lookup("term.householdid", LanguageID))%>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="hhID" name="hhID" maxlength="256" value="<%Sendb(Server.HtmlEncode(Request.QueryString("hhID")))%>" />
          </td>
        </tr>
        <tr id="trName" style="display: none;">
          <td style="width: 85px;">
            <label for="lastname">
              <% Sendb(Copient.PhraseLib.Lookup("term.lastname", LanguageID))%>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="lastname" name="lastname" value="<%Sendb(Server.HtmlEncode(Request.QueryString("lastname")))%>" />
          </td>
        </tr>
        <tr id="trPhone" style="display: none;">
          <td style="width: 85px;">
            <label for="phone1">
              <% Sendb(Copient.PhraseLib.Lookup("term.phone", LanguageID))%>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="phone1" name="phone1" maxlength="50" value="<%Sendb(Server.HtmlEncode(Request.QueryString("phone1")))%>" />
          </td>
        </tr>
        <tr id="trAltID" style="display: none;">
          <td style="width: 85px;">
            <label for="altID">
              <% Sendb(Copient.PhraseLib.Lookup("term.alternateid", LanguageID))%>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="altID" name="altID" maxlength="<% Sendb(IIf(MyCommon.GetCardMaxLength(CardTypes.ALTERNATEID) > 0, MyCommon.GetCardMaxLength(CardTypes.ALTERNATEID), 20))%>"
              value="<%Sendb(Server.HtmlEncode(Request.QueryString("altID")))%>" />
          </td>
        </tr>
        <tr id="trCoupon" style="display: none;">
          <td style="width: 85px;">
            <label for="couponID">
              <% Sendb(Copient.PhraseLib.Lookup("term.coupon", LanguageID))%>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="couponID" name="couponID" maxlength="30" value="<%Sendb(Server.HtmlEncode(Request.QueryString("couponID")))%>" />
          </td>
        </tr>
        <tr id="trCAM" style="display: none;">
          <td style="width: 85px;">
            <label for="hhID">
              <% Sendb(Copient.PhraseLib.Lookup("term.cam-cardholderid", LanguageID))%>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="camID" name="camID" maxlength="256" value="<%Sendb(Server.HtmlEncode(Request.QueryString("camID")))%>" />
          </td>
        </tr>
        <tr id="trLastNamePartial" style="display: none;">
          <td style="width: 85px;">
            <label for="lastnamepartial">
              <% Sendb(Copient.PhraseLib.Lookup("term.lastname", LanguageID))%>
              :</label>
          </td>
          <td>
            <input type="text" class="long" id="lastnamepartial" name="lastnamepartial" value="<%Sendb(Server.HtmlEncode(Request.QueryString("lastnamepartial")))%>" />
          </td>
        </tr>
        <tr id="trSubmit">
          <td>
            &nbsp;
          </td>
          <td>
            <input type="submit" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>"
              onclick="searchNew();" />
            <input type="button" class="regular" value="<% Sendb(Copient.PhraseLib.Lookup("term.advanced", LanguageID))%>"
              onclick="launchAdvSearch();" style="display: <% Sendb(IIf(MyCommon.Fetch_SystemOption(106) = 1, "", "none"))%>" />
            <input type="submit" class="regular" id="create" name="create" value="<% Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID)) %>"
              onclick="searchNew();" style="display: none;" />
          </td>
        </tr>
      </tbody>
    </table>
    <div style="display: none;">
      <input type="text" class="longer" id="firstname" name="firstname" value="<%Sendb(Server.HtmlEncode(Request.QueryString("firstname")))%>" />
      <input type="text" class="longer" id="email" name="email" value="<%Sendb(Server.HtmlEncode(Request.QueryString("email")))%>" />
      <input type="text" class="longer" id="address" name="address" value="<%Sendb(Server.HtmlEncode(Request.QueryString("address"))) %>" />
    </div>
    <br />
    <% End If%>
    <% Send("<input type=""hidden"" id=""editterms"" name=""editterms"" value=""" & ExtID & """ />")%>
    <div class="box" id="searchresults" style="display: <% Sendb(IIf(HasSearchResults, "block", "none")) %>;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.results", LanguageID))%>
        </span>
      </h2>
      <%
        ' determine whether any of the results have a customer associated to a banner.
        ' if so, then show the banner name column
        If (Customers IsNot Nothing) Then
          loopCtr = 0
          For Each Cust In Customers
            If (Cust IsNot Nothing) Then
              loopCtr += 1
              If Cust.GetBannerID > 0 Then
                ShowBannerCol = True
              End If
              If (ShowBannerCol OrElse loopCtr >= 20) Then
                Exit For
              End If
            End If
          Next
        End If
      %>
      <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.customers", LanguageID)) %>">
        <thead>
          <tr>
            <%
              Send("<th align=""left"" class=""th-cardholder"" scope=""col"">")
              Send("  " & Copient.PhraseLib.Lookup("term.cardholder", LanguageID))
              Send("</th>")
              Send("<th align=""left"" class=""th-household"" scope=""col"">")
              Send("  " & Copient.PhraseLib.Lookup("term.householdid", LanguageID))
              Send("</th>")
              Send("<th align=""left"" class=""th-yesno"" scope=""col"">")
              Send("  " & Copient.PhraseLib.Lookup("customer-inquiry.hh", LanguageID))
              Send("</th>")
              If (ShowBannerCol AndAlso MyCommon.Fetch_SystemOption(66) = "1") Then
                Send("<th align=""left"" class=""th-banner"" scope=""col"">")
                Send("  " & Copient.PhraseLib.Lookup("term.banner", LanguageID))
                Send("</th>")
              End If
              Send("<th align=""left"" class=""th-name"" scope=""col"">")
              Send("  " & Copient.PhraseLib.Lookup("term.name", LanguageID))
              Send("</th>")
              If Logix.UserRoles.AccessCustomerIdData_Address OrElse _
                 Logix.UserRoles.AccessCustomerIdData_City OrElse _
                 Logix.UserRoles.AccessCustomerIdData_State OrElse _
                 Logix.UserRoles.AccessCustomerIdData_ZIP Then
                Send("<th align=""left"" class=""th-address"" scope=""col"">")
                Send("  " & Copient.PhraseLib.Lookup("term.address", LanguageID))
                Send("</th>")
              End If
              If Logix.UserRoles.AccessCustomerIdData_Phone Then
                Send("<th align=""left"" class=""th-phone"" scope=""col"">")
                Send("  " & Copient.PhraseLib.Lookup("term.phone", LanguageID))
                Send("</th>")
              End If
            %>
          </tr>
        </thead>
        <tbody>
          <%
            If (Customers IsNot Nothing) AndAlso (Customers.Length > 0) Then
              loopCtr = 0
              For Each Cust In Customers
                If (Cust IsNot Nothing) Then
                  loopCtr += 1
                  If (ShowAll = False) AndAlso (loopCtr >= 21) Then
                    Exit For
                    'Make sure each CustomerPK is only sent once
                  ElseIf (loopCtr <> 1) AndAlso (Cust.GetCustomerPK = Customers(IIf((loopCtr - 2) > -1, loopCtr - 2, loopCtr - 1)).GetCustomerPK) Then
                  Else
                    HasCustomerID = False
                    FullName = ""
                    If Cust.GetPrefix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix Then
                      FullName &= MyCommon.SplitNonSpacedString(Cust.GetPrefix, 20) & " "
                    End If
                    If Cust.GetFirstName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName Then
                      FullName &= MyCommon.SplitNonSpacedString(Cust.GetFirstName, 20) & " "
                    End If
                    If Cust.GetMiddleName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName Then
                      FullName &= MyCommon.SplitNonSpacedString(Cust.GetMiddleName, 1) & ". "
                    End If
                    If Cust.GetLastName <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName Then
                      FullName &= MyCommon.SplitNonSpacedString(Cust.GetLastName, 20)
                    End If
                    If Cust.GetSuffix <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix Then
                      FullName &= " " & MyCommon.SplitNonSpacedString(Cust.GetSuffix, 20)
                    End If
                    CustExt = Cust.GetGeneralInfo
                    FullAddress = ""
                    If CustExt.GetAddress <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Address Then
                      FullAddress &= CustExt.GetAddress & "<br />"
                    End If
                    If CustExt.GetAddress <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_City Then
                      FullAddress &= CustExt.GetCity & " "
                    End If
                    If CustExt.GetAddress <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_State Then
                      FullAddress &= CustExt.GetState & " "
                    End If
                    If CustExt.GetAddress <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_ZIP Then
                      FullAddress &= CustExt.GetZip
                    End If
                    Send("<tr" & Shaded & ">")
                    Send("<td>")
                    For i = 0 To Cust.GetCards.GetUpperBound(0)
                      If (Cust.GetCustomerTypeID = 2) Then
                        If (MyCommon.Fetch_SystemOption(79) = 0) Then
                          Sendb("  <a href=""/logix/CAM/CAM-customer-general.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;CardPK=" & Cust.GetCards(i).GetCardPK)
                        ElseIf (MyCommon.Fetch_SystemOption(79) = 1) Then
                          Sendb("  <a href=""/logix/CAM/CAM-customer-offers.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;CardPK=" & Cust.GetCards(i).GetCardPK)
                        End If
                        Send("&amp;search=Search" & extraLink & """>" & Cust.GetCards(i).GetExtCardID & "</a><br />")
                        HasCustomerID = True
                      ElseIf (Cust.GetCustomerTypeID = 0) Then
                        If (MyCommon.Fetch_SystemOption(79) = 0) Then
                          Sendb("  <a href=""customer-general.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;CardPK=" & Cust.GetCards(i).GetCardPK)
                        ElseIf (MyCommon.Fetch_SystemOption(79) = 1) Then
                          Sendb("  <a href=""customer-offers.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;CardPK=" & Cust.GetCards(i).GetCardPK)
                        End If
          		        If (MyCommon.Fetch_SystemOption(144) = 1) AndAlso (Cust.GetCards(i).GetCardTypeID=3) AndAlso (Not Logix.UserRoles.EditCustomerIdData) Then 'mask last 4 digits of Altid, when SystemOption(144) is turned on
                            Send("&amp;search=Search" & extraLink & """>" & Cust.GetCards(i).GetExtCardID.ToString.Substring(0,Cust.GetCards(i).GetExtCardID.Length-4) & "****" &  "</a><br />")
		                Else
                           Send("&amp;search=Search" & extraLink & """>" & Cust.GetCards(i).GetExtCardID & "</a><br />")
        				End If						
                        HasCustomerID = True
                      End If
                    Next
                    Send("</td>")
                    Send("<td>")
                    If (Logix.UserRoles.ViewHHCardholders = False) AndAlso (HasCustomerID = True) Then
                      'In this case the user has no permission to see household cardholders.
                      'However, the user needs to be able to access household records themselves.
                      'Therefor: If the page has determined (above) that there's a customer ID for this record,
                      'then don't show the associated household; otherwise, if the record is *just* a household,
                      'then go ahead and show it (below).
                    Else
                                  If (Cust.GetCustomerTypeID = 1 And Cust.GetCustomerPK > 0) Then
                        For i = 0 To Cust.GetCards.GetUpperBound(0)
                          If (MyCommon.Fetch_SystemOption(79) = 0) Then
                            Sendb("<a href=""customer-general.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;CardPK=" & Cust.GetCards(i).GetCardPK)
                          ElseIf (MyCommon.Fetch_SystemOption(79) = 1) Then
                            Sendb("<a href=""customer-offers.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;CardPK=" & Cust.GetCards(i).GetCardPK)
                          End If
                          Send("&amp;search=Search" & extraLink & """>" & Cust.GetCards(i).GetExtCardID & "</a>")
                        Next
                      Else
                                      If(Cust.GetHHPK > 0)
                        MyCommon.QueryStr = "select CardPK, ExtCardID from CardIDs with (NoLock) where CustomerPK=" & Cust.GetHHPK & ";"
                        dt = MyCommon.LXS_Select
                        If dt.Rows.Count > 0 Then
                          For Each row In dt.Rows
                            If (MyCommon.Fetch_SystemOption(79) = 0) Then
                              Sendb("<a href=""customer-general.aspx?CustPK=" & Cust.GetHHPK & "&amp;CardPK=" & MyCommon.NZ(row.Item("CardPK"), 0))
                            ElseIf (MyCommon.Fetch_SystemOption(79) = 1) Then
                              Sendb("<a href=""customer-offers.aspx?CustPK=" & Cust.GetHHPK & "&amp;CardPK=" & MyCommon.NZ(row.Item("CardPK"), 0))
                            End If
                            Send("&amp;search=Search" & extraLink & """>" & MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID").ToString()) & "</a>")
                          Next
                        End If
                      End If
                    End If
                              End If
                    Send("</td>")
                    Send("  <td>" & IIf(Cust.GetCustomerTypeID = 1, Copient.PhraseLib.Lookup("term.yes", LanguageID), Copient.PhraseLib.Lookup("term.no", LanguageID)) & "</td>")
                    If (ShowBannerCol AndAlso MyCommon.Fetch_SystemOption(66) = "1") Then
                      MyCommon.QueryStr = "select Name from Banners with (NoLock) where Deleted=0 and BannerID=" & Cust.GetBannerID
                      rst2 = MyCommon.LRT_Select
                      If (rst2.Rows.Count > 0) Then
                        Send("<td>" & MyCommon.NZ(rst2.Rows(0).Item("Name"), "&nbsp;") & "</td>")
                      Else
                        Send("<td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                      End If
                    End If
                    Sendb("  <td>")
                    'If (Cust.GetCustomerTypeID = 2) Then
                    '  If (MyCommon.Fetch_SystemOption(79) = 0) Then
                    '    Sendb("<a href=""/logix/CAM/CAM-customer-general.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;CardPK=" & CardPK & "&amp;search=Search" & extraLink & """>")
                    '  ElseIf (MyCommon.Fetch_SystemOption(79) = 1) Then
                    '    Sendb("<a href=""/logix/CAM/CAM-customer-offers.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;search=Search" & extraLink & """>")
                    '  End If
                    'ElseIf (Cust.GetCustomerTypeID = 0) Then
                    '  If (MyCommon.Fetch_SystemOption(79) = 0) Then
                    '    Sendb("<a href=""customer-general.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;search=Search" & extraLink & """>")
                    '  ElseIf (MyCommon.Fetch_SystemOption(79) = 1) Then
                    '    Sendb("<a href=""customer-offers.aspx?CustPK=" & Cust.GetCustomerPK & "&amp;search=Search" & extraLink & """>")
                    '  End If
                    'End If
                    If FullName <> "" Then
                      Sendb(FullName)
                    Else
                      Sendb("&nbsp;&nbsp;&nbsp;")
                    End If
                    'Send("</a>")
                    Send("</td>")
                    If Logix.UserRoles.AccessCustomerIdData_Address OrElse _
                       Logix.UserRoles.AccessCustomerIdData_City OrElse _
                       Logix.UserRoles.AccessCustomerIdData_State OrElse _
                       Logix.UserRoles.AccessCustomerIdData_ZIP Then
                      Send("  <td>" & MyCommon.SplitNonSpacedString(FullAddress, 20) & "</td>")
                    End If
                    If (Logix.UserRoles.AccessCustomerIdData_Phone) Then
                      Send("  <td>" & MyLookup.FormatPhoneNumber(CustExt.GetPhone) & "</td>")
                    End If
                    Send("</tr>")
                    Shaded = IIf(Shaded = " class=""shaded""", "", " class=""shaded""")
                  End If
                End If
              Next
            Else
              Send("<tr>")
              Send("  <td colspan=""8""></td>")
              Send("</tr>")
            End If
          %>
        </tbody>
      </table>
      <%
          
        'Limited search
        If FoundRecordCount > 0 AndAlso Not ShowAll Then
          Send("<center><i>")
              Dim RowNum As Integer = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("row")))
          If RowNum = 0 Then RowNum = 1
          If RowNum > 1 Then
                  Sendb("<a href=""customer-inquiry.aspx?search=" & Server.HtmlEncode(Request.QueryString("search") & "&mode=" & Server.HtmlEncode(Request.QueryString("mode")) & "&CustomerPK=" & Server.HtmlEncode(Request.QueryString("CustomerPK"))))
                  Sendb("&searchby=" & Server.HtmlEncode(Request.QueryString("searchby")) & "&cardID=" & Server.HtmlEncode(Request.QueryString("cardID")) & "&hhID=" & Server.HtmlEncode(Request.QueryString("hhID")))
                  Sendb("&lastname=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("lastname"))) & "&phone1=" & Server.HtmlEncode(Request.QueryString("phone1")))
            Sendb("&firstname=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("firstname"))) & "&email=" & Server.HtmlEncode(Request.QueryString("email")))
                  Sendb("&address=" & Server.HtmlEncode(Request.QueryString("address")) & "&editterms=" & Server.HtmlEncode(Request.QueryString("editterms")) & "&lastnamepartial=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("lastnamepartial"))))
                  Sendb("&city=" & Server.HtmlEncode(Request.QueryString("city")) & "&state=" & Server.HtmlEncode(Request.QueryString("state")) & "&zip=" & Server.HtmlEncode(Request.QueryString("zip")))
            Dim LastRow As Integer = (RowNum - CustPerRow)
            If LastRow <= 0 Then LastRow = 1
            Sendb("&row=" & LastRow & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "&nbsp;&nbsp;</a>")
          End If
          If ((RowNum + CustPerRow) < FoundRecordCount) Then
                  Sendb("<a href=""customer-inquiry.aspx?search=" & Server.HtmlEncode(Request.QueryString("search")) & "&mode=" & Server.HtmlEncode(Request.QueryString("mode")) & "&CustomerPK=" & Server.HtmlEncode(Request.QueryString("CustomerPK")))
                  Sendb("&searchby=" & Server.HtmlEncode(Request.QueryString("searchby")) & "&cardID=" & Server.HtmlEncode(Request.QueryString("cardID")) & "&hhID=" & Server.HtmlEncode(Request.QueryString("hhID")))
                  Sendb("&lastname=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("lastname"))) & "&phone1=" & Server.HtmlEncode(Request.QueryString("phone1")))
                  Sendb("&firstname=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("firstname"))) & "&email=" & Server.HtmlEncode(Request.QueryString("email")))
                  Sendb("&address=" & Server.HtmlEncode(Request.QueryString("address")) & "&editterms=" & Server.HtmlEncode(Request.QueryString("editterms")) & "&lastnamepartial=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("lastnamepartial"))))
                  Sendb("&city=" & Server.HtmlEncode(Request.QueryString("city")) & "&state=" & Server.HtmlEncode(Request.QueryString("state")) & "&zip=" & Server.HtmlEncode(Request.QueryString("zip")))
            Sendb("&row=" & (RowNum + CustPerRow) & """>&nbsp;&nbsp;" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a>")
          End If
          Send("</i></center>")
        End If
        'Top 20 - Show all search
        If (FoundRecordCount >= 21) Then
          Send("<center style='' ><i>")
          If (ShowAll = False) Then
                  Dim RowNum As Integer = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("row")))
            Dim firstRecord As Integer
            Dim lastRecord As Integer
            If RowNum = 0 Then
              firstRecord = 1
              lastRecord = CustPerRow
            Else
              firstRecord = RowNum
              If FoundRecordCount > (RowNum + CustPerRow) Then
                lastRecord = RowNum + CustPerRow - 1
              Else
                lastRecord = (FoundRecordCount - RowNum) + RowNum
              End If
            End If
            Dim displayText As String
            displayText = firstRecord & " - " & lastRecord
            Sendb(Copient.PhraseLib.Detokenize("customer-inquiry.AllResultsShown", LanguageID, displayText))
                  Sendb("<a href=""customer-inquiry.aspx?search=" & Server.HtmlEncode(Request.QueryString("search")) & "&mode=" & Server.HtmlEncode(Request.QueryString("mode")) & "&CustomerPK=" & Server.HtmlEncode(Request.QueryString("CustomerPK")))
                  Sendb("&searchby=" & Server.HtmlEncode(Request.QueryString("searchby")) & "&cardID=" & Server.HtmlEncode(Request.QueryString("cardID")) & "&hhID=" & Server.HtmlEncode(Request.QueryString("hhID")))
                  Sendb("&lastname=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("lastname"))) & "&phone1=" & Server.HtmlEncode(Request.QueryString("phone1")))
                  Sendb("&firstname=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("firstname"))) & "&email=" & Server.HtmlEncode(Request.QueryString("email")))
                  Sendb("&address=" & Server.HtmlEncode(Request.QueryString("address")) & "&editterms=" & Server.HtmlEncode(Request.QueryString("editterms")) & "&lastnamepartial=" & HttpUtility.UrlEncode(Server.HtmlEncode(Request.QueryString("lastnamepartial"))))
                  Sendb("&city=" & Server.HtmlEncode(Request.QueryString("city")) & "&state=" & Server.HtmlEncode(Request.QueryString("state")) & "&zip=" & Server.HtmlEncode(Request.QueryString("zip")))
            Sendb("&showall=1"">" & Copient.PhraseLib.Detokenize("customer-inquiry.ShowAllResultCount", LanguageID, FoundRecordCount) & "</a>")
            ' Sendb(Copient.PhraseLib.Detokenize("customer-inquiry.AllResultsShown", LanguageID, FoundRecordCount)) 'All {0} results 
          Else
            Sendb(Copient.PhraseLib.Detokenize("customer-inquiry.AllResultsShown", LanguageID, FoundRecordCount)) 'All {0} results shown.           
          End If
          Send("</i></center>")
          'Else
          'End If
        End If

      %>
      <hr class="hidden" />
    </div>
  </div>
</div>
</form>
<%
  Send("<script type=""text/javascript"">")
  If (SearchTypeID > 0) Then
    If (SearchTypeID = 6 And infoMessage = "") Then
      Send("  showCriteriaFor(" & SearchTypeID & ",false);")
    Else
      Send("  showCriteriaFor(" & SearchTypeID & ",true);")
    End If
  Else
    Send("  showCriteriaFor(1,true);")
    Send("  setFocusCtrl(1);")
  End If
  Send("</script>")
done:
  Send_BodyEnd("mainform")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
