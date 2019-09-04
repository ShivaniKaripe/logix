<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-addhousehold.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2013.  All rights reserved by:
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
  
  Dim MyCommon As New Copient.CommonInc
  Dim AdminUserID As Long
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable = Nothing
  Dim HHPK As Long = 0
  Dim HHID As String = ""
  Dim CardPK As Long = 0
  Dim SearchText As String = ""
  Dim CustomerPKs() As String = Nothing
  Dim i As Integer
  Dim InfoMsg As String = ""
  Dim ErrorMsg As String = ""
  Dim Handheld As Boolean = False
  Dim DisplayText As String = ""
  
  Dim CustHH As New Copient.Customer
    Dim Cust As New Copient.Customer
    Dim MyCryptLib As New Copient.CryptLib
  Dim CustAvailable() As Copient.Customer = Nothing
  Dim MyLookup As New Copient.CustomerLookup
  Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
  Dim Added As Boolean = False
  Dim RulesEngine As New Copient.HouseholdRules(MyCommon)
  Dim HHOptions(-1) As Copient.HouseholdRules.InterfaceOption
  Dim HHQueueData As Copient.HouseholdRules.QUEUE_DATA
  Dim HHQueuePKID As Long
  Dim HHQueueCardID As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-addhousehold"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  MyLookup.SetAdminUserID(AdminUserID)
  MyLookup.SetLanguageID(LanguageID)
  
  SearchText = MyCommon.Parse_Quotes(Request.QueryString("custid"))
  DisplayText = Request.QueryString("custid")
  HHPK = MyCommon.Extract_Val(Request.QueryString("HHPK"))
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  
  If HHPK > 0 Then
    CustHH = MyLookup.FindCustomerInfo(HHPK, ReturnCode)
    If CustHH.GetCards.Length > 0 Then
      HHID = CustHH.GetCards(0).GetExtCardID
    End If
  Else
    CustHH = New Copient.Customer
  End If
  
  If (Request.QueryString("add") <> "") Then
    CustomerPKs = Request.QueryString.GetValues("addCustId")
    If CustomerPKs IsNot Nothing AndAlso CustomerPKs.Length > 0 Then
      RulesEngine.SetLogFile("Householding.txt")
      HHOptions = RulesEngine.GetHouseholdingOptions
      HHQueueData = GetQueueEntryForAdd(0, 0, HHOptions)
      For i = 0 To CustomerPKs.GetUpperBound(0)
        HHQueueData.CustomerPK = Long.Parse(CustomerPKs(i))
        HHQueueData.HHPK = CustHH.GetCustomerPK
        HHQueueData.AdminUserID = CType(AdminUserID, Integer)
        Added = RulesEngine.SendToQueue(HHQueueData, HHQueuePKID)
        If Not Added Then
          ErrorMsg = Copient.PhraseLib.Lookup("customer-inquiry.add-to-HH-failed", LanguageID)
        End If
      Next
      
      SearchText = Request.QueryString("searchvalue")
    Else
      ErrorMsg = Copient.PhraseLib.Lookup("customer-inquiry.no-card-selected", LanguageID)
    End If
  End If
  
  If (Request.QueryString("search") <> "" OrElse SearchText <> "") Then
    If SearchText.Trim <> "" Then
      CustAvailable = MyLookup.FindAddToHouseholdCustomers(CustHH.GetCustomerPK, SearchText, ReturnCode)
    Else
      ErrorMsg = Copient.PhraseLib.Lookup("customer-inquiry.enter-criteria", LanguageID)
    End If
  End If
  
  Send_HeadBegin("term.customer", "term.customerinquiry")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.AddHHCardholders = False) Then
    Send_Denied(2, "perm.customers-add-to-hh")
    GoTo done
  End If
  
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  Send("  opener.location = 'customer-general.aspx?CustPK=" & HHPK & "&CardPK=" & CardPK & "'; ")
  Send("} ")
  Send("</script>")
  
%>
<form id="mainform" name="mainform" action="customer-addhousehold.aspx">
<input type="hidden" id="HHID" name="HHID" value="<% Sendb(HHID)%>" />
<input type="hidden" id="HHPK" name="HHPK" value="<% Sendb(HHPK)%>" />
<%
  If CardPK > 0 Then
    Send("<input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
  End If
%>
<input type="hidden" id="searchvalue" name="searchvalue" value="<%Sendb(SearchText)%>" />
<div id="intro">
  <h1 id="title">
    <% Send(Copient.PhraseLib.Lookup("term.addcardholder", LanguageID) & " #" & HHID)%>
  </h1>
</div>
<div id="main">
  <%
    If (ErrorMsg <> "") Then
      Send("<div id=""infobar"" class=""red-background"">")
      Send("  " & ErrorMsg)
      Send("</div>")
    End If
    If (InfoMsg <> "") Then
      Send("<div id=""infobar"" class=""green-background"">")
      Send("  " & InfoMsg)
      Send("</div>")
    End If
  %>
  <div id="column">
    <div class="box" id="searchinput">
      <h2>
        <span>
          <%Sendb(Copient.PhraseLib.Lookup("term.cardholder", LanguageID))%>
        </span>
      </h2>
      <br class="half" />
      <% Sendb(Copient.PhraseLib.Lookup("customer.addtohousehold", LanguageID))%>
      <br />
      <br class="half" />
      <input type="text" id="custid" name="custid" maxlength="50" value="<%Sendb(DisplayText)%>" />
      <input type="submit" class="regular" id="search" name="search" value="<%Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
    </div>
    <% If (CustAvailable IsNot Nothing) Then%>
    <div class="box" id="results">
      <h2>
        <span>
          <%Sendb(Copient.PhraseLib.Lookup("term.results", LanguageID))%>
        </span>
      </h2>
      <span style="font-size: 8pt; font-weight: bold; position: relative; float: right;
        top: -23px;"><a href="customer-addhousehold.aspx?HHPK=<%Sendb(HHPK)%>&CardPK=<%Sendb(CardPK)%>&custid=<%Sendb(SearchText)%>">
          <%Sendb(Copient.PhraseLib.Lookup("term.refresh", LanguageID))%></a> </span>
      <%
        If (CustAvailable.Length > 0) Then
          Send("<center>")
          Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """>")
          Send("    <thead>")
          Send("      <tr>")
          Send("        <th class=""th-add""          scope=""col"">" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "</th>")
          Send("        <th class=""th-cardholder""   scope=""col"">" & Copient.PhraseLib.Lookup("term.cardnumber", LanguageID) & " / " & Copient.PhraseLib.Lookup("term.cardtype", LanguageID) & "</th>")
          Send("        <th class=""th-firstname""    scope=""col"">" & Copient.PhraseLib.Lookup("term.firstname", LanguageID) & "</th>")
          Send("        <th style=""width:25px;""     scope=""col"">" & Copient.PhraseLib.Lookup("term.middleinitialM", LanguageID) & "</th>")
          Send("        <th class=""th-lastname""     scope=""col"">" & Copient.PhraseLib.Lookup("term.lastname", LanguageID) & "</th>")
          Send("        <th class=""th-household""    scope=""col"">" & Copient.PhraseLib.Lookup("term.householdid", LanguageID) & "</th>")
          Send("      </tr>")
          Send("    </thead>")
          Send("    <tbody>")
          For Each Cust In CustAvailable
            ' determine if a pending household request exists for the customer
            If (Cust IsNot Nothing) Then
                  
              Dim a_pk As Long = Cust.GetCustomerPK
              MyCommon.QueryStr = String.Format( _
                "select C.ExtCardID as HouseholdID from HouseholdQueue as HQ with (NoLock) inner join CardIDs as C with (NoLock) on C.CustomerPK = HQ.HHPK and C.CardTypeID = 1  where HQ.CustomerPK = {0} and StatusCode >= 0;", _
                a_pk)
                    
              dt = MyCommon.LXS_Select
              If dt.Rows.Count > 0 Then
                          HHQueueCardID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("HouseholdID").ToString())
              Else
                HHQueueCardID = ""
              End If
                  
              Send("      <tr>")
              If HHQueueCardID = "" Then
                Send("        <td><input type=""checkbox"" name=""addCustId"" id=""addCustId" & MyCommon.NZ(Cust.GetCustomerPK, 0).ToString & """ value=""" & MyCommon.NZ(Cust.GetCustomerPK, 0).ToString & """ /></td>")
              Else
                Send("        <td>" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & "</td>")
              End If
              Sendb("        <td><label for=""addCustId" & MyCommon.NZ(Cust.GetCustomerPK, 0).ToString & """>")
              For i = 0 To (Cust.GetCards.Length - 1)
                If i > 0 Then
                  Sendb("<br />")
                End If
                'If customer card type = alternate ID and Enable AltID PIN Masking is enabled, don't display the last four digits of the alternate ID card number
                Dim cardIdLength As Integer = If(Cust.GetCards(i).GetExtCardID.Length >= 4, Cust.GetCards(i).GetExtCardID.Length - 4, Cust.GetCards(i).GetExtCardID.Length)
                Sendb(String.Format("{0}<br /> {1}", IIf(Cust.GetCards(i).GetCardTypeID = 3 AndAlso MyCommon.NZ(Integer.Parse(MyCommon.Fetch_SystemOption(144)), -1) = 1 AndAlso MyCommon.NZ(Cust.GetCards(i).GetExtCardID, "").Length >= 14, Cust.GetCards(i).GetExtCardID.Substring(0, cardIdLength), MyCommon.NZ(Cust.GetCards(i).GetExtCardID, "")), getCardTypeDescriptionFromCardTypeID(Cust.GetCards(i).GetCardTypeID(), MyCommon)))
              Next
              Send("</label></td>")
              Send("        <td>" & MyCommon.NZ(Cust.GetFirstName, "&nbsp;") & "</td>")
              Send("        <td>" & IIf(Cust.GetMiddleName <> "", Left(Cust.GetMiddleName, 1) & ".", "") & "</td>")
              Send("        <td>" & MyCommon.NZ(Cust.GetLastName, "&nbsp;") & "</td>")
              Send("        <td style=""font-style:italic;"">" & IIf(HHQueueCardID <> "", HHQueueCardID, MyCommon.NZ(Cust.GetHouseHoldID, "")) & "</td>")
              Send("      </tr>")
              
            End If

          Next
          Send("    </tbody>")
          Send("  </table>")
          Send("  <br />")
          Send("  <input type=""submit"" class=""regular"" id=""add"" name=""add"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ />")
          Send("</center>")
        Else
          Send("<i>" & Copient.PhraseLib.Lookup("customer.nocardholdersfound", LanguageID) & "</i>")
        End If
      %>
    </div>
    <% End If%>
  </div>
</div>
</form>
<%
  If (Request.QueryString("add") <> "" AndAlso HHID <> "") Then
    Send("<script type=""text/javascript"">")
    Send("  opener.location = 'customer-general.aspx?CustPK=" & CustHH.GetCustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "&refresh=1&search=Search&searchterms=" & CustHH.GetCustomerPK & "'; ")
    Send("  window.focus();")
    Send("</script>")
  End If

%>
<script runat="server">


  Function getCardTypeDescriptionFromCardTypeID(ByVal cardtypeid As Integer, ByRef commonlib As Copient.CommonInc) As String
        
    Dim q As String = String.Format("SELECT TOP 1 [description], [phraseid] FROM [CardTypes] WHERE [CardTypeID] = {0};", cardtypeid)
    commonlib.QueryStr = q
    Using dt As DataTable = commonlib.LXS_Select
      If dt.Rows.Count > 0 Then
        Return Copient.PhraseLib.Lookup(dt.Rows(0).Item("phraseid"), LanguageID, dt.Rows(0).Item("description"))
      End If
    End Using
    Return Copient.PhraseLib.Lookup("term.unknown", LanguageID)
        
  End Function
    


  Function GetQueueEntryForAdd(ByVal CustomerPK As Long, ByVal HHPK As Long, _
                                ByVal HHOptions As Copient.HouseholdRules.InterfaceOption()) As Copient.HouseholdRules.QUEUE_DATA
    Dim qData As New Copient.HouseholdRules.QUEUE_DATA
    
    qData.ActionTypeID = Copient.HouseholdRules.ACTION_TYPES.ADD
    qData.SourceTypeID = Copient.HouseholdRules.SOURCE_TYPES.LOGIX
    qData.CustomerPK = CustomerPK
    qData.HHPK = HHPK

    If HHOptions IsNot Nothing AndAlso HHOptions.Length >= 5 Then
      qData.Option5Value = HHOptions(0).Value
      qData.Option6Value = HHOptions(1).Value
      qData.Option7Value = HHOptions(2).Value
      qData.Option8Value = HHOptions(3).Value
      qData.Option9Value = HHOptions(4).Value
    End If
    
    Return qData
  End Function
</script>
<%  
done:
  Send_BodyEnd("mainform", "custid")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
