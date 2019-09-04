<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.Odbc" %>
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
  Dim MyLookup As New Copient.CustomerLookup
  Dim Logix As New Copient.LogixInc
  Dim rstResults As DataTable = Nothing
  Dim rst As DataTable
  Dim row As DataRow
  Dim rstCurrentOffers As DataTable
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim row3 As DataRow
  Dim dtPP As DataTable
  Dim dtPP_Points As DataTable
  Dim dtSV_Points As DataTable
  Dim dt As DataTable
  
  Dim rstStoredValue As DataTable
  Dim rowCount As Integer
  Dim CurrentOffers As String = ""
  Dim CustomerPK As Long
  Dim CardTypeId As Integer
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim FullName As String = ""
  Dim ProgramID As String
  Dim ProgramName As String
  Dim ProgramDesc As String
  Dim PointsBalance As String
  Dim sPointsText As String
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
  Dim FullAddress As String = ""
  Dim CustomerTypeID As Integer = 0
  Dim Employee As Integer = 0
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
  Dim SVQuantity As Integer = 0
  Dim SVIDBuf As New StringBuilder()
  Dim SVNameBuf As New StringBuilder()
  Dim ProgName As String
  Dim cgXml As String = ""
  Dim offerXml As String = ""
  Dim reader As SqlDataReader = Nothing
  Dim rows() As DataRow
  Dim SessionID As String = ""
  Dim ExtHostTypeID As Integer = 0
  Dim ExternalProgram As Boolean = False

  ' default urls for links from this page
  Dim URLOfferSum As String = "offer-sum.aspx"
  Dim URLCPEOfferSum As String = "CPEoffer-sum.aspx"
  Dim URLcgroupedit As String = "cgroup-edit.aspx"
  Dim URLpointedit As String = "point-edit.aspx"
  Dim URLtrackBack As String = ""
  Dim inCardNumber As String = ""
  ' tack on the customercare remote links if needed
  Dim extraLink As String = ""
  Dim bEME As Boolean = False
  Dim sErrorMsg As String = ""
  
  Dim objTemp As Object
  Dim intNumDecimalPlaces As Integer
  Dim decFactor As Decimal
  Dim decTemp As Decimal
  Dim sTemp1 As String
  Dim bNeedToFormat As Boolean
  
  'variables used in Offer Eligibility Conditions
  Dim dtEligibleOffers As New DataTable
  Dim dtEligiblePP As New DataTable
  Dim dtEligibleSVP As New DataTable
  Dim eligibleOfferXML As String = String.Empty
  
  'DB2 connection String
  Dim user As String
  Dim euser As String
  Dim pwd As String
  Dim epwd As String
  Dim upos As Integer
  Dim ppos As Integer
  Dim pend As Integer
  Dim tempstr As String
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-inquiry.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If MyCommon.Fetch_SystemOption(80) = "1" Then
    bEME = True
  End If

  objTemp = MyCommon.Fetch_CM_SystemOption(41)
  If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
    intNumDecimalPlaces = 0
  End If
  decFactor = (10 ^ intNumDecimalPlaces)
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If

  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If (rst.Rows(0).Item("prestrict") = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      restrictLinks = True
    End If
  End If
  
  Edit = False
  ProgramID = ""
  ProgramName = ""
  If (MyCommon.Extract_Val(Request.QueryString("CustPK")) > 0) Then
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
  ElseIf (MyCommon.Extract_Val(Request.QueryString("CustomerPK")) > 0) Then
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  End If
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
  End If

  'get the CardType for this customer
  MyCommon.QueryStr = "select CardTypeID from CardIDs where CustomerPK=" & CustomerPK
  If CardPK > 0 Then
    MyCommon.QueryStr += " and CardPK=" & CardPK & ";"
  End If
  rst = MyCommon.LXS_Select()
  If (rst.Rows.Count = 1) Then
    CardTypeId = MyCommon.NZ(rst.Rows(0).Item("CardTypeID"), 0)
  Else
    CardTypeId = 0
    CardPK = 0
    'infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
  End If
  
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
  
  If Not (MyCommon.Extract_Val(Request.QueryString("CustPK")) <> 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "customer-inquiry.aspx")
  End If
  
  If (CustomerPK > 0) Then
    MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, C.Employee, " & _
                        "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                        "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                        "left join Customers C2 with (NoLock) on C2.CustomerPK = C.HHPK " & _
                        "where C.CustomerPK=" & CustomerPK
  Else
    ' IF the page was called from an outside application set ClientUserID1 to the outside passed in value
    If (inCardNumber <> "" And Request.QueryString("mode") = "summary") Then
      'Redirect to customer-inquiry.aspx
    End If
    If (Request.QueryString("searchterms") <> "" And ClientUserID1 = "") Then
      ClientUserID1 = MyCommon.Pad_ExtCardID(Request.QueryString("searchterms"), CardTypeId)
            MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.HHPK, C.FirstName, C.MiddleName, C.LastName, C.Employee, " & _
                                "CE.Address, CE.City, CE.State, CE.Zip, CE.PhoneAsEntered, CE.email " & _
                                "from Customers C with (NoLock) left join CustomerExt CE with (NoLock) on CE.CustomerPK = C.CustomerPK " & _
                                "left join Customers C2 with (NoLock) on C2.CustomerPK = C.HHPK " & _
                                "where C.PrimaryExtID='" & MyCryptLib.SQL_StringEncrypt(ClientUserID1) & "' or CE.PhoneDigitsOnly = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.DigitsOnly(Request.QueryString("searchterms"))) & _
                                "' or CE.email = '" & MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Request.QueryString("searchterms"))) & "' or C.LastName like '%" & MyCommon.Parse_Quotes(Request.QueryString("searchterms")) & "%';"
    End If
  End If
  rstResults = MyCommon.LXS_Select
  If (rstResults.Rows.Count = 1) Then
    CustomerPK = rstResults.Rows(0).Item("CustomerPK")
    IsHouseholdID = (MyCommon.NZ(rstResults.Rows(0).Item("CustomerTypeID"), 0) = 1)
    Employee = rstResults.Rows(0).Item("Employee")
  Else
    infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
    infoMessage = infoMessage & " <a href=""customer-inquiry.aspx?mode=add&Search=Search" & extraLink & "&searchterms=" & Request.QueryString("searchterms") & """>" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</a>"
  End If
  
  If (Request.QueryString("Adj") = "Adj" And Request.QueryString("editterms") <> "") Then
    CustomerPK = Request.QueryString("CustomerPK")
    ' ok we got the program IDs
    MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
    rst = MyCommon.LXS_Select
    If (rst.Rows.Count = 0) Then
      infoMessage = "" & Copient.PhraseLib.Lookup("customer.notfound", LanguageID) & ""
    End If
    
    Dim progIDArray() As String
    Dim y As Int16
    Dim progIDString As String
    
    progIDString = Request.QueryString("ProgramIDs")
    progIDArray = progIDString.Split(",")
    
    For y = 0 To progIDArray.GetUpperBound(0)
      MyCommon.QueryStr = "select PromoVarID from PointsPrograms with (NoLock) where ProgramID=" & progIDArray(y) & ";"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0 And MyCommon.Extract_Val(Request.QueryString("pointadjust-" & progIDArray(y))) <> 0) Then
        ' when submitted call this dbo.pc_PromoVar_Update @CustomerPK bigint, @PromoVarID bigint, @VariableTypeID int, @AdjAmount decimal(12,3)
        'Send("updating customerpk: " & CustomerPK & "points for program var : " & MyCommon.Extract_Val(rst.Rows.Item(0).Item("PromoVarID")) & " by " & MyCommon.Extract_Val(Request.QueryString("pointadjust-" & progIDArray(y))))
        MyCommon.QueryStr = "dbo.pc_PromoVar_Update"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
        MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(rst.Rows.Item(0).Item("PromoVarID"))
        MyCommon.LXSsp.Parameters.Add("@VariableTypeID", SqlDbType.BigInt).Value = 3
        MyCommon.LXSsp.Parameters.Add("@AdjAmount", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("pointadjust-" & progIDArray(y)))
        MyCommon.LXSsp.ExecuteNonQuery()
        MyCommon.Close_LXSsp()
      End If
    Next
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "customer-inquiry.aspx?searchterms=" & CustomerPK & "&search=Search" & extraLink)
    GoTo done
    
  End If
  
  UnknownPhrase = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
  DisabledPtsAdj = IIf(Logix.UserRoles.EditPointsBalances, "", " disabled=""disabled"" ")
  DisabledAccumAdj = IIf(Logix.UserRoles.EditAccumBalances, "", " disabled=""disabled"" ")
  
  If CardPK > 0 Then
    Send_HeadBegin("term.customer", "term.adjustments", MyCommon.Extract_Val(ExtCardID))
  Else
    Send_HeadBegin("term.customer", "term.adjustments")
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
      Send_Subtabs(Logix, 32, 5, LanguageID, CustomerPK, , CardPK)
    Else
      Send_Subtabs(Logix, 32, 5, LanguageID, CustomerPK)
    End If
  Else
    If CardPK > 0 Then
      Send_Subtabs(Logix, 91, 6, LanguageID, CustomerPK, extraLink, CardPK)
    Else
      Send_Subtabs(Logix, 91, 6, LanguageID, CustomerPK, extraLink)
    End If
  End If
  
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
%>
<script type="text/javascript">

  function adjustPoints(custPK, cardPK, enableSetTo) {
    var elemProg = document.getElementById('functionselectpt');
    var elemSel = null;
    var progID = -1;
    var progName = "";
    var qryStr = "";
    var substringprogName = progName
  
    if (elemProg != null && elemProg.selectedIndex > -1) {
      elemSel = elemProg.options[elemProg.selectedIndex];
      progID = elemSel.value;
      progName = elemSel.text;
      if (enableSetTo == true) {
        substringprogName = progName.substr(progName.length - 2, 2);
        if (substringprogName != "s)") {
          alert(Sendb(Copient.PhraseLib.Lookup("term.invalidpointsprogram", LanguageID)));
        }
        else {
          qryStr = "point-adjust-program.aspx?ProgramID=" + progID + "&CustomerPK=" + custPK + "&CardPK=" + cardPK + "&Opener=<%Sendb(CopientFileName)%>&SetTo=" + enableSetTo + ""
	    openPopup(qryStr);
	  }
  }
  else {
    qryStr = "point-adjust-program.aspx?ProgramID=" + progID + "&CustomerPK=" + custPK + "&CardPK=" + cardPK + "&Opener=<%Sendb(CopientFileName)%>&SetTo=" + enableSetTo + ""
	   openPopup(qryStr);
	 }
 }
}

function handleKeyUpPt(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselectpt").size = "10";
  
  // Set references to the form elements
  selectObj = document.forms['mainform'].functionselectpt;
  textObj = document.forms['mainform'].functioninputpt;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlistpt.length;
  
  // Set the search pattern depending
  if(document.forms['mainform'].functionradiopt[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlistpt[i].search(re) != -1 && vallistpt[i] != "") {
      selectObj[numShown] = new Option(functionlistpt[i],vallistpt[i]);
      numShown++;
    }
    // Stop when the number to show is reached
    if(numShown == maxNumToShow) {
      break;
    }
  }
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClickPt() {
  selectObj = document.forms['mainform'].functionselectpt;
  textObj = document.forms['mainform'].functioninputpt;
  
  if (selectObj != null && selectObj.selectedIndex > -1) {
    selectedValue = selectObj.options[selectObj.selectedIndex].text;        
    if(selectedValue != "") {   
      textObj.value = selectedValue;
    }
  }
}

function handleKeyDownPt(e) {
  var key = e.which ? e.which : e.keyCode;
  var elem = document.getElementById('functionselectpt');
  
  if (elem != null && key == 13) {
    if (e && e.preventDefault) {
      e.preventDefault(); // DOM style
    } else {
      e.keyCode = 9;
    }
    if (elem.selectedIndex > -1) {
      adjustPoints(<% Sendb(CustomerPK)%>, <% Sendb(CardPK)%>,false);
    }
    return false;
  }
  return true;
}

function adjustStoredValue(custPK, cardPK) {
  var elemProg = document.getElementById('functionselectsv');
  var elemSel = null;
  var progID = -1;
  var progName = "";
  var qryStr = "";
  
  if (elemProg != null && elemProg.selectedIndex > -1) {
    elemSel = elemProg.options[elemProg.selectedIndex];
    progID = elemSel.value;
    progName = elemSel.text;
    qryStr = "sv-adjust-program.aspx?ProgramID=" + progID + "&CustomerPK=" + custPK + "&CardPK=" + cardPK + "&Opener=<%Sendb(CopientFileName)%>"
    openPopup(qryStr);
  }
}

function handleKeyUpSv(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  
  document.getElementById("functionselectsv").size = "10";
  
  // Set references to the form elements
  selectObj = document.forms['mainform'].functionselectsv;
  textObj = document.forms['mainform'].functioninputsv;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlistsv.length;
  
  // Set the search pattern depending
  if(document.forms['mainform'].functionradiosv[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  // Clear the options list
  selectObj.length = 0;
  
  // Loop through the array and re-add matching options
  numShown = 0;
  for(i = 0; i < functionListLength; i++) {
    if(functionlistsv[i].search(re) != -1 && vallistsv[i] != "") {
      selectObj[numShown] = new Option(functionlistsv[i],vallistsv[i]);
      numShown++;
    }
    // Stop when the number to show is reached
    if(numShown == maxNumToShow) {
      break;
    }
  }
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClickSv() {
  selectObj = document.forms['mainform'].functionselectsv;
  textObj = document.forms['mainform'].functioninputsv;
  
  if (selectObj != null && selectObj.selectedIndex > -1) {
    selectedValue = selectObj.options[selectObj.selectedIndex].text;        
    if(selectedValue != "") {   
      textObj.value = selectedValue;
    }
  }
}

function handleKeyDownSv(e) {
  var key = e.which ? e.which : e.keyCode;
  var elem = document.getElementById('functionselectsv');
  
  if (elem != null && key == 13) {
    if (e && e.preventDefault) {
      e.preventDefault(); // DOM style
    } else {
      e.keyCode = 9;
    }
    if (elem.selectedIndex > -1) {
      adjustStoredValue(<% Sendb(CustomerPK)%>, <% Sendb(CardPK)%>);
    }
    return false;
  }
  return true;
}

</script>
<script type="text/javascript" src="../javascript/jquery.min.js"></script>
<script type="text/javascript" src="../javascript/thickbox.js"></script>
<form id="mainform" name="mainform" action="customer-inquiry.aspx">
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
            Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
          End If
        End If
        MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
        rst2 = MyCommon.LXS_Select
        If rst2.Rows.Count > 0 Then
          FullName = IIf(MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") & " ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") & " ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, MyCommon.NZ(rst2.Rows(0).Item("LastName"), ""), "")
          FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & MyCommon.NZ(rst2.Rows(0).Item("Suffix"), ""), "")
        End If
        If FullName <> "" Then
          Sendb(": " & MyCommon.TruncateString(FullName, 30))
        End If
        If (restrictLinks AndAlso URLtrackBack <> "") Then
          Send(" - <a href=""" & URLtrackBack & """>" & Copient.PhraseLib.Lookup("customer-inquiry.return", LanguageID) & "</a>")
        End If
      %>
    </h1>
    <div id="controls" <% Sendb(IIf(restrictLinks AndAlso URLtrackBack <> "", " style=""width:115px;""", ""))%>>
      <%
        If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
          Send_CustomerNotes(CustomerPK, CardPK)
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div><br />")%>
    <input type="hidden" id="CustomerPK" name="CustomerPK" value="<%Sendb(CustomerPK)%>" />
    <%
      If (Request.QueryString("mode") = "summary") Then
        Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
        Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=""" & URLtrackBack & """ />")
      End If

      'If this is a household card and system option 24 is on, get all groups 
      '  that this household and members of this household belong to
      If (MyCommon.Fetch_CM_SystemOption(24) = "1" AndAlso CardTypeId = 1) Then
        MyCommon.QueryStr = "create table #CustomerList ([CustomerPK] bigint NULL);" & _
        "insert into #CustomerList (CustomerPK)" & _
        "select CustomerPK from Customers with (NoLock) where CustomerPK=" & CustomerPK & " or HHPK=" & CustomerPK & ";" & _
        "select distinct CustomerGroupID from GroupMembership with (NoLock)" & _
        "where Deleted=0 and CustomerPK in (select CustomerPK from #CustomerList with (NoLock));" & _
        "drop table #CustomerList;"
        rst = MyCommon.LXS_Select()
      Else
        MyCommon.QueryStr = "select CustomerGroupID from groupmembership with (NoLock) where customerpk=" & CustomerPK & " and deleted=0"
        rst = MyCommon.LXS_Select()
      End If
          
      cgXml = "<customergroups><id>1</id><id>2</id>"
      rowCount = rst.Rows.Count
      If rowCount > 0 Then
        For Each row In rst.Rows
          cgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
        Next
      End If
      cgXml &= "</customergroups>"
          
      MyCommon.QueryStr = "dbo.pa_CustomerOffersCurrentAndAdd"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = cgXml
      MyCommon.LRTsp.Parameters.Add("@IsEmployee", SqlDbType.Bit).Value = Employee
      If (Request.QueryString("offerterms") <> "") Then
        MyCommon.LRTsp.Parameters.Add("@Filter", SqlDbType.NVarChar, 50).Value = Request.QueryString("offerterms")
      End If
      MyCommon.LRTsp.Parameters.Add("@ShowAdd", SqlDbType.Bit).Value = 0
      reader = MyCommon.LRTsp.ExecuteReader
          
      rstCurrentOffers = New DataTable
      rstCurrentOffers.Load(reader)
          
      MyCommon.Close_LRTsp()
      reader.Close()

      'Load Eligible Offers linked to the Customer Groups that the Customer is a Part Of
      MyCommon.QueryStr = "dbo.pa_CustomerEligibleOffers"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = cgXml
      reader = MyCommon.LRTsp.ExecuteReader
      dtEligibleOffers.Load(reader)
      MyCommon.Close_LRTsp()
      reader.Close()
      eligibleOfferXML = "<offers>"
      rowCount = dtEligibleOffers.Rows.Count
      If rowCount > 0 Then
        For Each row In dtEligibleOffers.Rows
          eligibleOfferXML &= "<id>" & MyCommon.NZ(row.Item("OfferID"), "0") & "</id>"
        Next
      End If
      eligibleOfferXML &= "</offers>"
    
      ' load up points program for all the customers offers enmasse
      offerXml = "<offers>"
      rowCount = rstCurrentOffers.Rows.Count
      If rowCount > 0 Then
        For Each row In rstCurrentOffers.Rows
          offerXml &= "<id>" & MyCommon.NZ(row.Item("OfferID"), "0") & "</id>"
        Next
      End If
      offerXml &= "</offers>"
          
      MyCommon.QueryStr = "dbo.pa_CustomerOffersWithPointsPrograms"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@offerXml", SqlDbType.Xml).Value = offerXml
      dtPP = MyCommon.LRTsp_select
    
      If Not (String.IsNullOrWhiteSpace(eligibleOfferXML)) Then
        MyCommon.QueryStr = "dbo.pa_CustomerEligibleOffersWithPointsPrograms"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@offerXml", SqlDbType.Xml).Value = eligibleOfferXML
        dtEligiblePP = MyCommon.LRTsp_select
        If dtEligiblePP.Rows.Count > 0 Then
          dtPP.Merge(dtEligiblePP, True, MissingSchemaAction.Ignore)
          dtPP = dtPP.DefaultView.ToTable(True)
        End If
      End If
          
      If bEME Then
        ' load up points balances for customers points programs enmasse.
        MyCommon.QueryStr = "select P.ProgramID, P.Amount as PointsBal, PV.ExternalID" & _
                            " from Points as P with (NoLock)" & _
                            " inner join PromoVariables as PV with (NoLock) on PV.PromoVarID = P.PromoVarID" & _
                            " where P.CustomerPK = " & CustomerPK & " and PV.ExternalID is null" & _
                            " order by ProgramID;"
        dtPP_Points = MyCommon.LXS_Select
            
        'get the external card ID for the customer/household
            MyCommon.QueryStr = "select isnull(ExtCardIDOriginal, '') as ExtCardID from CardIDs where CustomerPK=" & CustomerPK & " and CardTypeID=0;"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
          ExtCardID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString())
        End If
        If ExtCardID.Length > 0 Then
          Dim ExternalPP As Copient.ExternalRewards
          Dim sDb2Connection As String
          Dim iDb2Connection As Integer = 5
            
          sDb2Connection = MyCommon.Fetch_CM_SystemOption(iDb2Connection)
          tempstr = sDb2Connection
                If tempstr IsNot Nothing AndAlso tempstr <> "" Then
          upos = InStr(tempstr, "UID=", CompareMethod.Text)
          ppos = InStr(tempstr, ";PWD=", CompareMethod.Text)
          pend = InStr(tempstr, ";host", CompareMethod.Text)
          euser = tempstr.Substring(upos + 3, ppos - upos - 4)
          epwd = tempstr.Substring(ppos + 4, pend - ppos - 5)
          user = MyCryptLib.SQL_StringDecrypt(euser)
          pwd = MyCryptLib.SQL_StringDecrypt(epwd)
          tempstr = tempstr.Replace(euser, user)
          sDb2Connection = tempstr.Replace(epwd, pwd)
                End If
          Try
            ' add ProgramIds for external points programs for which the customer has no current balance.
            ExternalPP = New Copient.ExternalRewards("", "", "", sDb2Connection)
            ExternalPP.appendExtProgramBalances(ExtCardID, False, dtPP_Points, MyCommon)
          Catch ex As Exception
            sErrorMsg = "EME: " & ex.Message
          End Try
        End If
      Else
        ' load up points balances for customers points programs enmasse.
        MyCommon.QueryStr = "select ProgramID, Sum(Amount) as PointsBal " & _
                            "from Points where CustomerPK=" & CustomerPK & " " & _
                            "group by ProgramID order by ProgramID;"
        dtPP_Points = MyCommon.LXS_Select
      End If
      
      If sErrorMsg <> "" Then
        Send("    <div id=""infobar"" class=""red-background"">")
        Send("        " & sErrorMsg)
        Send("    </div>")
      End If

    %>
    <div id="column1">
      <div class="box" id="pointsbalances">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.points", LanguageID))%>
          </span>
        </h2>
        <%
          
          ' merge the points with their respective points programs
          For Each row In dtPP_Points.Rows
            rows = dtPP.Select("ProgramID=" & MyCommon.NZ(row.Item("ProgramID"), "-1"))
            If (rows IsNot Nothing AndAlso rows.Length > 0) Then
              For Each row2 In rows
                row2.Item("PointsBal") = MyCommon.NZ(row.Item("PointsBal"), 0)
              Next
            Else
              MyCommon.QueryStr = "select ProgramName, ExternalProgram, ExtHostTypeID from PointsPrograms with (NoLock) where ProgramID=" & MyCommon.NZ(row.Item("ProgramID"), -1) & " AND Deleted = 0"
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                row2 = dtPP.NewRow()
                row2.Item("ProgramID") = MyCommon.NZ(row.Item("ProgramID"), -1)
                row2.Item("ProgramName") = MyCommon.NZ(rst.Rows(0).Item("ProgramName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                row2.Item("PointsBal") = MyCommon.NZ(row.Item("PointsBal"), 0)
                row2.Item("ExternalProgram") = MyCommon.NZ(rst.Rows(0).Item("ExternalProgram"), False)
                row2.Item("ExtHostTypeID") = MyCommon.NZ(rst.Rows(0).Item("ExtHostTypeID"), 0)
                dtPP.Rows.Add(row2)
              End If

            End If
          Next
          
          If (Logix.UserRoles.AccessPointsBalances AndAlso CustomerPK > 0) Then
            Send("<center>")
            Send("<b>" & Copient.PhraseLib.Lookup("customer-inquiry.adjustpoints", LanguageID) & "</b><br />")
        %>
        <br class="half" />
        <input type="radio" id="functionradiopt1" name="functionradiopt" <% If (MyCommon.Fetch_SystemOption(175) = "1") Then Sendb(" checked=""checked""")%> /><label
          for="functionradiopt1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradiopt2" name="functionradiopt" <% If (MyCommon.Fetch_SystemOption(175) = "2") Then Sendb(" checked=""checked""")%> /><label for="functionradiopt2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input type="text" class="longer" onkeyup="handleKeyUpPt(200);" onkeydown="handleKeyDownPt(event);"
          id="functioninputpt" name="functioninputpt" maxlength="100" value="" /><br />
        <select class="longer" id="functionselectpt" name="functionselectpt" onclick="handleSelectClickPt();"
          onkeydown="handleKeyDownPt(event);" ondblclick="if (this.selectedIndex > -1) { adjustPoints(<% Sendb(CustomerPK)%>, <% Sendb(CardPK)%>,false); }"
          size="10">
          <%
            rows = dtPP.Select("", "ProgramName")
            For Each row In rows
              ProgramID = MyCommon.NZ(row.Item("ProgramID"), "-1")
              ProgramName = MyCommon.NZ(row.Item("ProgramName"), "")
              If MyCommon.NZ(row.Item("ExternalProgram"), False) = False Then
                PointsBalance = " (" & Decimal.ToInt32(MyCommon.NZ(row.Item("PointsBal"), 0)) & " " & Copient.PhraseLib.Lookup("term.points", LanguageID).ToLower & ")"
                sPointsText = ProgramName & PointsBalance
              Else
                If bEME Then
                  PointsBalance = " (" & MyCommon.NZ(row.Item("PointsBal"), 0) & " " & Copient.PhraseLib.Lookup("term.points", LanguageID).ToLower & "*)"
                  MyCommon.QueryStr = "select Description from PointsPrograms with (NoLock) where ProgramID=" & MyCommon.NZ(row.Item("ProgramID"), -1)
                  rst = MyCommon.LRT_Select
                  If rst.Rows.Count > 0 Then
                    ProgramDesc = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
                  Else
                    ProgramDesc = ""
                  End If
                  If ProgramDesc <> "" Then
                    Dim ExtraCharCount As Integer
                    If (ProgramName.Length + PointsBalance.Length) <= 42 Then
                      ProgramDesc = " - " & Left(ProgramDesc, 42 - ProgramName.Length - PointsBalance.Length)
                    Else
                      ExtraCharCount = (ProgramName.Length + PointsBalance.Length) - 42
                      ProgramDesc = " - " & Left(ProgramDesc, ExtraCharCount)
                    End If
                  End If
                  sPointsText = ProgramName & ProgramDesc & PointsBalance
                Else
                  PointsBalance = " (" & MyCommon.NZ(row.Item("PointsBal"), 0) & " " & Copient.PhraseLib.Lookup("term.points", LanguageID).ToLower & ")"
                  sPointsText = ProgramName & "(" & Copient.PhraseLib.Lookup("term.external", LanguageID) & ")"
                End If
              End If
              Sendb("<option value=""" & ProgramID & """>" & sPointsText)
              Send("</option>")
              ' store this for use later when creating the searchable javascript array
              ProgName = sPointsText.Replace("'", "\'")
              ProgName = ProgName.Replace("""", "\""")
              If (PointsIDBuf.Length > 0) Then
                PointsIDBuf.Append(",")
                PointsNameBuf.Append(",")
              End If
              PointsIDBuf.Append("""" & ProgramID & """")
              PointsNameBuf.Append("""" & sPointsText & """")
            Next
          %>
        </select>
        <br />
        <br class="half" />
        <% If (Logix.UserRoles.EditPointsBalances AndAlso rows.Length > 0) Then%>
        <input type="button" class="regular" id="ptselect" name="ptselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.adjust", LanguageID))%>"
          onclick="if (document.mainform.functionselectpt.selectedIndex > -1) { <% Sendb("adjustPoints('" & CustomerPK & "', '" & CardPK & "',false); }")%>" />
        <% If MyCommon.NZ(MyCommon.Fetch_SystemOption(189), "0") = "1" Then%>
	     &nbsp;&nbsp; &nbsp;&nbsp;
	     <input type="button" class="regular" id="ptSet" name="ptSet" value="<% Sendb(Copient.PhraseLib.Lookup("term.set", LanguageID))%>"
         onclick="if (document.mainform.functionselectpt.selectedIndex > -1) { <% Sendb("adjustPoints('" & CustomerPK & "', '" & CardPK & "',true); }")%>" />
        <%End If%>
        <%End If%>
      </center>
      <%
      Else
        Send_Denied(0, "perm.customers-ptbalaccess")
      End If
      %>
        <hr class="hidden" />
      </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="storedvalue">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID))%>
          </span>
        </h2>
        <% If (Logix.UserRoles.AccessStoredValue AndAlso CustomerPK > 0) Then%>
        <center>
          <% If (CustomerPK <> 0 AndAlso rstCurrentOffers IsNot Nothing AndAlso offerXml <> "") Then%>
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("customer-inquiry.adjuststoredvalue", LanguageID))%>:</b><br />
          <br class="half" />
          <input type="radio" id="functionradiosv1" name="functionradiosv" <% If (MyCommon.Fetch_SystemOption(175) = "1") Then Sendb(" checked=""checked""")%> /><label
            for="functionradiosv1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
          <input type="radio" id="functionradiosv2" name="functionradiosv" <% If (MyCommon.Fetch_SystemOption(175) = "2") Then Sendb(" checked=""checked""")%> /><label for="functionradiosv2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
          <input type="text" class="longer" id="functioninputsv" name="functioninputsv" onkeyup="handleKeyUpSv(200);"
            onkeydown="handleKeyDownSv(event);" maxlength="100" value="" /><br />
          <select class="longer" id="functionselectsv" name="functionselectsv" onclick="handleSelectClickSv();"
            onkeydown="handleKeyDownSv(event);" ondblclick="if (this.selectedIndex > -1) { adjustStoredValue(<% Sendb(CustomerPK)%>, <% Sendb(CardPK)%>); }"
            size="10">
            <%
              MyCommon.QueryStr = "dbo.pa_CustomerOffersWithSVPrograms"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@offerXml", SqlDbType.Xml).Value = offerXml
              rstStoredValue = MyCommon.LRTsp_select
            
              If Not (String.IsNullOrWhiteSpace(eligibleOfferXML)) Then
                MyCommon.QueryStr = "dbo.pa_CustomerEligibleOffersWithSVPrograms"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@offerXml", SqlDbType.Xml).Value = eligibleOfferXML
                dtEligibleSVP = MyCommon.LRTsp_select
                If dtEligibleSVP.Rows.Count > 0 Then
                  rstStoredValue.Merge(dtEligibleSVP)
                  rstStoredValue = rstStoredValue.DefaultView.ToTable(True)
                End If
              End If
            
              ' load up points balances for customers stored value programs enmasse.
              If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(91) = 1 Then
                MyCommon.QueryStr = "select SVProgramID, Sum(QtyEarned) - Sum(QtyUsed) as PointsBal " & _
                                    "from StoredValue with (NoLock) " & _
                                    "where CustomerPK=" & CustomerPK & " and ExpireDate >= getdate() and Deleted=0" & _
                                    "group by SVProgramID order by SVProgramID;"
              Else
                MyCommon.QueryStr = "select SVProgramID, Sum(QtyEarned) - Sum(QtyUsed) as PointsBal " & _
                                    "from StoredValue with (NoLock) " & _
                                    "where StatusFlag=1 and CustomerPK=" & CustomerPK & " and ExpireDate >= getdate() and Deleted=0" & _
                                    "group by SVProgramID order by SVProgramID;"
              End If
              dtSV_Points = MyCommon.LXS_Select
              ' merge the points with their respective points programs
              For Each row In dtSV_Points.Rows
                rows = rstStoredValue.Select("SVProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), "-1"))
                If rows.Length > 0 Then
                  rows(0).Item("PointsBal") = MyCommon.NZ(row.Item("PointsBal"), 0)
                Else
                  Dim dtSV_Programs As DataTable
                  MyCommon.QueryStr = "select Name,Description from StoredValuePrograms with (NoLock) where deleted=0 and SVProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), "-1") & ";"
                  dtSV_Programs = MyCommon.LRT_Select
                  If dtSV_Programs.Rows.Count > 0 Then
                    Dim dr As DataRow
                    dr = rstStoredValue.NewRow
                    dr.Item("SVProgramID") = row.Item("SVProgramID")
                    dr.Item("PointsBal") = row.Item("PointsBal")
                    dr.Item("Name") = dtSV_Programs.Rows(0).Item("Name")
                    dr.Item("Description") = dtSV_Programs.Rows(0).Item("Description")
                    rstStoredValue.Rows.Add(dr)
                  End If
                End If
              Next
              rows = rstStoredValue.Select("", "Name")
              For Each row3 In rows
                If (MyCommon.NZ(row3.Item("SVProgramID"), -1) > 0) Then
                  ' store this for use later when creating the searchable javascript array
                  ProgName = MyCommon.NZ(row3.Item("Name"), "").Replace("'", "\'")
                  ProgName = ProgName.Replace("""", "\""")
                  If (SVIDBuf.Length > 0) Then
                    SVIDBuf.Append(",")
                    SVNameBuf.Append(",")
                  End If
                  ' If System option 41 is a valid integer value and this stored value program is a
                  '  points stored value program, format displayed points
                  ' Check this stored value program type
                  Dim TempTable As DataTable
                  Dim TempRow As DataRow
                  MyCommon.QueryStr = "select SVTypeID from StoredValuePrograms " & _
                                    "where SVProgramID = " & row3.Item("SVProgramID")
                  TempTable = MyCommon.LRT_Select
                  TempRow = TempTable.Rows(0)
                  bNeedToFormat = False
                  If intNumDecimalPlaces > 0 Then
                    If Int(MyCommon.NZ(TempRow.Item("SVTypeID"), 0)) = 1 Then
                      bNeedToFormat = True
                      decTemp = (Int(MyCommon.NZ(row3.Item("PointsBal"), 0)) * 1.0) / decFactor
                      sTemp1 = FormatNumber(decTemp, intNumDecimalPlaces)
                    End If
                  End If
                  SVIDBuf.Append("""" & MyCommon.NZ(row3.Item("SVProgramID"), "-1") & """")
                  SVNameBuf.Append("""" & ProgName & " (" & MyCommon.NZ(row3.Item("PointsBal"), "0") & " " & Copient.PhraseLib.Lookup("term.units", LanguageID) & ")" & """")
                  'Send("<option value=""" & MyCommon.NZ(row3.Item("SVProgramID"), "-1") & """>" & MyCommon.NZ(row3.Item("Name"), "") & " (" & MyCommon.NZ(row3.Item("PointsBal"), "0") & " " & IIf(bNeedToFormat,sTemp1,"") & " " & Copient.PhraseLib.Lookup("term.units", LanguageID) & ")</option>")
                  Send("<option value=""" & MyCommon.NZ(row3.Item("SVProgramID"), "-1") & """>" & MyCommon.NZ(row3.Item("Name"), "") & " (" & IIf(bNeedToFormat, sTemp1, MyCommon.NZ(row3.Item("PointsBal"), "0")) & " " & Copient.PhraseLib.Lookup("term.units", LanguageID) & ")</option>")
                End If
              Next
            %>
          </select>
          <br />
          <br class="half" />
          <% If (Logix.UserRoles.ModifyStoredValue AndAlso rstStoredValue.Rows.Count > 0) Then%>
          <input type="button" class="regular" id="svselect" name="svselect" value="<% Sendb(Copient.PhraseLib.Lookup("term.adjust", LanguageID))%>"
            onclick="if (document.mainform.functionselectsv.selectedIndex > -1) { <% Sendb("adjustStoredValue('" & CustomerPK & "', '" & CardPK & "'); }")%>" />
          <% End If%>
          <% End If%>
        </center>
        <% 
        Else
          Send_Denied(0, "perm.customer-svaccess")
        End If
        %>
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<script type="text/javascript" language="javascript">
  // This is the javascript array holding the function list
  // The PrintJavascriptArray ASP function can be used to print this array.
  <%
  If (PointsIDBuf.Length > 0) Then
    Send("")
    Sendb("var functionlistpt = Array(")
    Sendb(PointsNameBuf.ToString())
    Send(");")
    Sendb("var vallistpt = Array(")
    Sendb(PointsIDBuf.ToString())
    Send(");")
  Else
    Sendb("var functionlistpt = Array(")
    Send("""" & "" & """);")
    Sendb("var vallistpt = Array(")
    Send("""" & "" & """);")
  End If
                
  If (SVIDBuf.Length > 0) Then
    Send("")
    Sendb("var functionlistsv = Array(")
    Sendb(SVNameBuf.ToString())
    Send(");")
    Sendb("var vallistsv = Array(")
    Sendb(SVIDBuf.ToString())
    Send(");")
  Else
    Sendb("var functionlistsv = Array(")
    Send("""" & "" & """);")
    Sendb("var vallistsv = Array(")
    Send("""" & "" & """);")
  End If
%>
</script>
<script runat="server">
</script>
<%
  'If MyCommon.Fetch_SystemOption(75) Then
  '  If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
  '    Send_Notes(4, CustomerPK, AdminUserID)
  '  End If
  'End If
done:
  If (Request.QueryString("adjWin") = "1") Then
    Send_BodyEnd()
  Else
    Send_BodyEnd("mainform", "functioninputpt")
  End If
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
