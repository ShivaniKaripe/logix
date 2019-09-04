<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Web" %>
<%
  ' *****************************************************************************
  ' * FILENAME: advanced-search.aspx 
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
  Dim dt As DataTable
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim row3 As DataRow
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim TokenString As String = ""
  Dim XID As String = ""
  Dim OfferID As String = ""
  Dim OfferName As String = ""
  Dim Desc As String = ""
  Dim ROID As String = ""
  Dim CreatedBy As String = ""
  Dim LastUpdatedBy As String = ""
  Dim Engine As String = ""
  Dim Banner As String = ""
  Dim Category As String = ""
  Dim Created1 As String = ""
  Dim Created2 As String = ""
  Dim Starts1 As String = ""
  Dim Starts2 As String = ""
  Dim Ends1 As String = ""
  Dim Ends2 As String = ""
  Dim BannersEnabled As Boolean = False
  Dim CustomerInquiry As Boolean = False
  Dim EnginesInstalled(-1) As Integer
    Dim SearchTypeID As Integer = 0
    Dim MyCryptLib As New Copient.CryptLib
	Dim maxLength As Integer = 256
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  If (Request.QueryString("CustomerInquiry") <> "") Then
    CustomerInquiry = True
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-inquiry-adva-search.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  EnginesInstalled = MyCommon.GetInstalledEngines()
  
  Send_HeadBegin("term.offers", "term.search")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<script type="text/javascript" language="javascript">
  function submitForm() {
      document.mainform.submit();
      if (window.opener != null && !window.opener.closed) {
        window.opener.focus();
        window.close();
      }
  }
  function onEnter(e )
  {
    var keynum=0;
    if(window.event) // IE8 and earlier
      keynum = e.keyCode;
    else if(e.which) // IE9/Firefox/Chrome/Opera/Safari
      keynum = e.which;

    if(keynum == 13)
      return submitForm();
  }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(3)
%>
<form id="mainform" name="mainform" action="customer-inquiry.aspx" target="CustomerInquiryAdvSearchWin">
  <input type="hidden" id="Search" name="Search" value="Search" />
  <input type="hidden" name="mode" id="mode" value="custominquiryadvancedsearch" />
  <div id="intro">
    <h1 id="H1_1">
      <% Sendb(Copient.PhraseLib.Lookup("term.advancedsearchcriteria", LanguageID))%>
    </h1>
    <div id="controls">
    </div>
  </div>
  <div id="main">    
    <div id="column">
      <div class="box" id="criteria">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID))%>
          </span>
        </h2>
        <center>
          <table style="width: 65%;" summary="<% Sendb(Copient.PhraseLib.Lookup("term.advancedsearchcriteria", LanguageID))%>">
            <%--
            <tr id="trSearchBy" style="display:<% Sendb(IIf(MyCommon.Fetch_SystemOption(107) = "1", "none", ""))%>">
              <td>
                <label for="searchby">
                  <% Sendb(Copient.PhraseLib.Lookup("term.searchby", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select name="searchby" id="searchby" >
                  <%
                    MyCommon.QueryStr = "select SearchTypeID, Name, PhraseID, Enabled from CustomerSearchTypes with (NoLock) "
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
                      MyCommon.QueryStr = "select CardTypeID, Description, PhraseID from CardTypes with (NoLock) where CustTypeID=" & row.Item("TypeID") & " order by CardTypeID;"
                      rst3 = MyCommon.LXS_Select
                      If rst3.Rows.Count > 1 Then
                        Send("<select id=""CardTypeID" & row.Item("TypeID") & """ name=""CardTypeID"" style=""display:none;"" disabled=""disabled"">")
                        For Each row3 In rst3.Rows
                          Sendb("<option value=""" & MyCommon.NZ(row3.Item("CardTypeID"), 0) & """" & IIf(MyCommon.NZ(row3.Item("CardTypeID"), 0) = MyCommon.Extract_Val(Request.QueryString("CardTypeID")), " selected=""selected""", "") & ">")
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
            --%>
            <tr id="trCH">
              <td style="width: 95px;">
                <label for="cardID"><% Sendb(IIf(MyCommon.Fetch_SystemOption(107) = "1", Copient.PhraseLib.Lookup("term.cust-specific-card-number", LanguageID), Copient.PhraseLib.Lookup("term.cardnumber", LanguageID)))%>:</label>
              </td>
              <td>
               	<input type="text" class="long" id="cardID" name="cardID" maxlength= """  & maxLength & """ value="<%Sendb(Request.QueryString("cardID")) %>" onkeypress="onEnter(event);"/>
              </td>
            </tr>
            <tr id="trHH" style="display:none">
              <td style="width: 85px;">
                <label for="hhID"><% Sendb(Copient.PhraseLib.Lookup("term.householdid", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" class="long" id="hhID" name="hhID" maxlength="""  & maxLength & """ value="<%Sendb(Request.QueryString("hhID")) %>" onkeypress="onEnter(event);" />
              </td>
            </tr>
            <tr id="tr1">
              <td style="width: 85px;">
                <label for="firstname"><% Sendb(Copient.PhraseLib.Lookup("term.firstname", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="firstname" name="firstname" value="<%Sendb(Request.QueryString("firstname")) %>" onkeypress="onEnter(event);"/>
              </td>
            </tr>
            <%--
            <tr id="trName">
              <td style="width: 85px;">
                <label for="lastname"><% Sendb(Copient.PhraseLib.Lookup("term.lastname", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="lastname" name="lastname" value="<%Sendb(Request.QueryString("lastname")) %>" onkeypress="onEnter(event);" />
              </td>
            </tr>
            --%>
            <%-- Last Name Partial --%>
            <tr id="trLastNamePartial">
              <td style="width: 85px;">
                <label for="lastnamepartial"><% Sendb(Copient.PhraseLib.Lookup("term.lastname", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="lastnamepartial" name="lastnamepartial" value="<%Sendb(Request.QueryString("lastnamepartial")) %>" onkeypress="onEnter(event);"/>
              </td>
            </tr>
			<% If MyCommon.Fetch_SystemOption(196) = "0" Then %>
            <tr id="trPhone">
              <td style="width: 85px;">
                <label for="phone1"><% Sendb(Copient.PhraseLib.Lookup("term.phone", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="phone1" name="phone1" maxlength="50" value="<%Sendb(Request.QueryString("phone")) %>" onkeypress="onEnter(event);"/>
              </td>
            </tr>
			<% End If %>
            <tr>
              <td style="width: 85px;">
                <label for="Address"><% Sendb(Copient.PhraseLib.Lookup("term.Address", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="Address" name="Address" maxlength="256" value="<%Sendb(Request.QueryString("Address")) %>" onkeypress="onEnter(event);"/>
              </td>
            </tr>
            <tr>
              <td style="width: 85px;">
                <label for="city"><% Sendb(Copient.PhraseLib.Lookup("term.city", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="city" name="city" maxlength="256" value="<%Sendb(Request.QueryString("city")) %>" onkeypress="onEnter(event);" />
              </td>
            </tr>
            <tr>
              <td style="width: 85px;">
                <label for="state"><% Sendb(Copient.PhraseLib.Lookup("term.state", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="state" name="state" maxlength="256" value="<%Sendb(Request.QueryString("state")) %>" onkeypress="onEnter(event);" />
              </td>
            </tr>
            <tr>
              <td style="width: 85px;">
                <label for="zip"><% Sendb(Copient.PhraseLib.Lookup("term.postalcode", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="zip" name="zip" maxlength="256" value="<%Sendb(Request.QueryString("zip")) %>"  onkeypress="onEnter(event);"/>
              </td>
            </tr>
			<% If MyCommon.Fetch_SystemOption(196) = "0" Then %>
            <tr>
              <td style="width: 85px;">
                <label for="Email"><% Sendb(Copient.PhraseLib.Lookup("term.Email", LanguageID))%>:</label>
              </td>
              <td>
                <input type="text" id="Email" name="Email" maxlength="256" value="<%Sendb(Request.QueryString("Email")) %>" onkeypress="onEnter(event);"/>
              </td>
            </tr>
			<% End If %>
          </table>
          <br />
          <input type="button" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" onclick="return submitForm();" />
        </center>
        <hr class="hidden" />
      </div>

      <div id="datepicker" class="dpDiv">
      </div>
      <%
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>
    </div>
  </div>
</form>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform")
  Logix = Nothing
  MyCommon = Nothing
%>