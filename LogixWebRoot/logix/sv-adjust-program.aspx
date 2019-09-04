<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="System.Globalization" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
  ' *****************************************************************************
  ' * FILENAME: sv-adjust-program.aspx 
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
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
</script>

<%
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim SvAdjust As New Copient.StoredValue
  Dim dt As DataTable = Nothing
  Dim dt2 As DataTable = Nothing
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim ProgramID As Long
  Dim ProgramName As String = ""
  Dim AdjustPermitted As Boolean = False
  Dim OffersLinked As Boolean = False
  Dim EarnedROID As Long = 0
  Dim EarnedCMOffer As Long = 0
  Dim OfferID As Long
  Dim Opener As String = ""
  Dim OfferIDs As String()
  Dim Adjust As String = ""
  Dim Note As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim LocalID As Long = 0
  Dim IsExpired As Boolean = False
  Dim HHPK As Integer = 0
  Dim HHCardPK As Long = 0
  Dim i As Integer = 0
  Dim KeyCt As Integer = 0
  Dim RefreshParent As Boolean = False
  Dim CouponID As String = ""
  Dim CouponNotFound As Boolean = False
  Dim SessionID As String = ""
  Dim FirstName As String = ""
  Dim LastName As String = ""
  Dim objTemp As Object
  Dim intNumDecimalPlaces As Integer = 0
  Dim decFactor As Decimal = 1.0
  Dim decTemp As Decimal
  Dim IntTemp As Integer
  Dim EnableFuelPartner As Boolean = False
  Dim ExtLocationID As String = ""
  Dim LocationID As Integer = 0
  Dim ReasonID As Integer = 0
  Dim ReasonText As String = ""
  Dim Localization As Copient.Localization

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "sv-adjust-program.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Localization = New Copient.Localization(MyCommon)  
  AdjustPermitted = Logix.UserRoles.EditPointsBalances
  OffersLinked = Logix.UserRoles.AccessOffers
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If

  ' CM engine only
  If MyCommon.IsEngineInstalled(0) Then
    EnableFuelPartner = MyCommon.Fetch_CM_SystemOption(56)
  
    objTemp = MyCommon.Fetch_CM_SystemOption(41)
    If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
      intNumDecimalPlaces = 0
    End If
    decFactor = (10 ^ intNumDecimalPlaces)
  End If
  
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
  OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
  CouponID = Request.QueryString("CouponID")
  Opener = Request.QueryString("Opener")
  ReasonID = MyCommon.Extract_Val(Request.QueryString("ReasonID"))
  ReasonText = Request.QueryString("ReasonText")
  If Opener = "" Then
    Opener = "customer-adjustments.aspx"
  End If
  
  If (CouponID <> "") Then
    MyCommon.QueryStr = "select SVProgramID, OfferID from StoredValue with (NoLock) where CustomerPK = 0 and ExternalID = '" & CouponID & "';"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      ProgramID = MyCommon.NZ(dt.Rows(0).Item("SVProgramID"), 0)
      OfferID = MyCommon.NZ(dt.Rows(0).Item("OfferID"), 0)
      CustomerPK = 0
    Else
      InfoMessage = Copient.PhraseLib.Detokenize("sv-adjust-program.CouponNotFound", LanguageID, CouponID)
      CouponNotFound = True
    End If
  End If
  
  If (CustomerPK > 0) Then
    MyCommon.QueryStr = "select HHPK from Customers with (NoLock) where CustomerPK=" & CustomerPK
    dt = MyCommon.LXS_Select()
    If (dt.Rows.Count > 0) Then
      HHPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
    End If
    If HHPK > 0 Then InfoMessage = Copient.PhraseLib.Lookup("sv.hh-adjust-note", LanguageID)
  End If
  
  'Get the CardPK of the Household card if there is a household PK
  If (HHPK > 0) Then
    MyCommon.QueryStr = "select CardPK from CardIDs where CustomerPK=" & HHPK
    dt = MyCommon.LXS_Select()
    If (dt.Rows.Count > 0) Then
      HHCardPK = MyCommon.NZ(dt.Rows(0).Item("CardPK"), 0)
    End If
  End If
  
  MyCommon.QueryStr = "select Name, SVTypeID from StoredValuePrograms with (NoLock) where SVProgramID = " & ProgramID
  dt = MyCommon.LRT_Select()
  If (dt.Rows.Count > 0) Then
    ProgramName = MyCommon.NZ(dt.Rows(0).Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
  End If
  
  If (Request.QueryString("save") <> "" OrElse Request.QueryString("EnterPressed") <> "" OrElse MyCommon.Extract_Val(Request.QueryString("LocalID")) > 0) Then
    Adjust = Request.QueryString("adjustUnits")
    'handle system option for decimals
    If (MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) = 1) Then
      If Not (Decimal.TryParse(Adjust, decTemp)) Then
        decTemp = 0
      End If
      decTemp = decTemp * decFactor
      'cast as int, any decimal places past decFactor are truncated
      IntTemp = decTemp
      Adjust = IntTemp
    End If

    Note = Request.QueryString("note")
    Note = MyCommon.Parse_Quotes(Note)
    Note = Logix.TrimAll(Note)
    LocalID = MyCommon.Extract_Val(Request.QueryString("LocalID"))
    If MyCommon.Fetch_SystemOption(99) Then
      If (Note <> "") And (Note.Length <= 1000) Then
        MyCommon.QueryStr = "select FirstName, LastName from AdminUsers where AdminUserID=" & AdminUserID & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
          FirstName = MyCommon.NZ(dt.Rows(0).Item("FirstName"), "")
          LastName = MyCommon.NZ(dt.Rows(0).Item("LastName"), "")
        End If
        MyCommon.QueryStr = "insert into CustomerNotes with (RowLock) (CustomerPK, AdminUserID, CreatedDate, NoteTypeID, Note, FirstName, LastName) " & _
                            " values (" & CustomerPK & ", " & AdminUserID & ", getDate(), 1, '" & Note & "', '" & FirstName & "', '" & LastName & "');"
        MyCommon.LXS_Execute()
      Else
        InfoMessage = Copient.PhraseLib.Lookup("sv-adjust-program.NoteRequired", LanguageID)
      End If
    End If
    If (MyCommon.Fetch_SystemOption(108) = "1") Then
      ExtLocationID = MyCommon.NZ(Request.QueryString("location"), "-10")
      MyCommon.QueryStr = "select LocationID from Locations where ExtLocationCode='" & ExtLocationID & "';"
      dt2 = MyCommon.LRT_Select
      If dt2.Rows.Count > 0 Then
        LocationID = MyCommon.NZ(dt2.Rows(0).Item("LocationID"), 0)
      Else If ExtLocationID = "-9" Then
        LocationID = -9
      Else
        InfoMessage &= Copient.PhraseLib.Lookup("customer-edit.UnableToValidateLocation", LanguageID) & " Entered Location: " & ExtLocationID
      End If
    End If
    If (MyCommon.Fetch_SystemOption(224) = "1") Then
	   Dim StoredValueAdjustmentLimit As Integer = 0
	   Dim LimitPeriod As Integer = 1
	   MyCommon.QueryStr = "SELECT StoredValueAdjustmentLimit, LimitPeriod FROM AdminRoleAdjustmentLimits WITH (NoLock) INNER JOIN AdminUserRoles WITH (NoLock) ON AdminRoleAdjustmentLimits.RoleID = AdminUserRoles.RoleID " & _
	                       " WHERE AdminUserRoles.AdminUserID = " & AdminUserID & " ;"
       dt2 = MyCommon.LRT_Select
	   If (dt2.Rows.Count > 0) Then
	     StoredValueAdjustmentLimit = MyCommon.NZ(dt2.Rows(0).Item("StoredValueAdjustmentLimit"), 0)
		 LimitPeriod = MyCommon.NZ(dt2.Rows(0).Item("LimitPeriod"), 1)
	   End If	 
	   MyCommon.QueryStr = "SELECT SUM(CAST(ActivityValue AS Float)) AS CurrentDayAdjustmentsTotal FROM ActivityLog WITH (NoLock) WHERE ActivityTypeID=25 AND ActivitySubTypeID=13 AND " & _
                    "LinkID= " & CustomerPK & " AND LinkID2= " & ProgramID & " AND AdminID= " & AdminUserID & " AND DATEDIFF(DD,ActivityDate,GETDATE())<= ( " & LimitPeriod & " - 1) ;"		  
	   dt = MyCommon.LRT_Select	
	   If (dt.Rows.Count > 0) Then
  	     If ((MyCommon.NZ(dt.Rows(0).Item("CurrentDayAdjustmentsTotal"), 0) + Adjust) > StoredValueAdjustmentLimit ) Then
		   infoMessage &= Copient.PhraseLib.Lookup("points-adjust.AdjustmentExceedsUserRoleLimit", LanguageID)
		 End If  
	   End If	
	End If	
    If InfoMessage = "" Then
      InfoMessage = SvAdjust.AdjustStoredValue(AdminUserID, ProgramID, CustomerPK, Adjust, LocalID, SessionID, OfferID, LocationID, 0, 0, ReasonID, ReasonText)
    End If
    If InfoMessage = "" Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "sv-adjust-program.aspx?ProgramID=" & ProgramID & "&CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & _
                         "&Opener=" & Opener & "&CouponID=" & CouponID & "&OfferID=" & OfferID & _
                         "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                         "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")))
      GoTo done
    End If
  ElseIf (Request.QueryString("HistoryEnabled.x") <> "" OrElse Request.QueryString("HistoryDisabled.x") <> "") Then
    ' Write a cookie and then reload the page
    Response.Cookies("SVHistoryEnabled").Expires = "10/08/2100"
    Response.Cookies("SVHistoryEnabled").Value = IIf(Request.QueryString("HistoryEnabled.x") <> "", "1", "0")
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "sv-adjust-program.aspx?ProgramID=" & ProgramID & "&CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & _
                       "&Opener=" & Opener & "&CouponID=" & CouponID & "&OfferID=" & OfferID & _
                       "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                       "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")))
    GoTo done
  End If
  
  'If (Request.QueryString("RefreshParent") = "true") Then RefreshParent = True
  
  Send_HeadBegin("term.storedvalue", "term.storedvalueadjustment", ProgramID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
#svtabs {
  z-index: 1;
  }
#svtabs span {
  background-color: white;
  border: 1px solid #306080;
  border-bottom: 0;
  -moz-border-radius: 4px 4px 0 0;
  cursor: hand;
  position: relative;
  top: 3px;
  }
.svtabon {
  font-weight: bold;
  padding: 1px 5px 5px 5px;
  }
.svtaboff {
  font-weight: normal;
  padding: 1px 5px 3px 5px;
  }
</style>
<%
  Send_Scripts(New String() {"datePicker.js"})
%>

<script type="text/javascript" language="javascript">
    var datePickerDivID = "datepicker";
    var bSkipUnload = false;
    var bSkipValidate = false;
    var linkToHH = false;

    <% Send_Calendar_Overrides(MyCommon) %>

    function isValidEntry() {
      var retVal = true;
      var elem = document.getElementById("adjust");
      var elemBal = document.getElementById("Balance");
      var unitVal = document.getElementById("UnitValue");
      var unitType = document.getElementById("UnitType");
      var adjustUnits = document.getElementById("adjustUnits");
      var expireType = document.getElementById("ExpireType");
      var x = 0;
      var unitValNum = 1.00;
      var elemValue = elem.value.replace(",",".");
      var decFactor = parseFloat(document.getElementById("decFactor").value);
      var valThousand = elemValue * 1000;

      valThousand = valThousand.toFixed(3);
      valThousand = valThousand.replace(",",".");
      

      if (bSkipValidate) { return true; }
        
      if (elem != null) {
        if (!isNaN(unitVal.value) || parseFloat(unitVal.value) > 0) {
          unitValNum = parseFloat(unitVal.value);
        }
        // check for empty string entry
        if (elem.value == "") {
          retVal = false;
          alert('<%Sendb(Copient.PhraseLib.Lookup("sv-adjust.nozeroadjustments", LanguageID)) %>');
        }
        
        if (isNaN(elemValue)) {
          retVal = false;
          alert('<%Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID)) %>');
          elem.focus();
          elem.select();
        } else {
          if (parseFloat(elemValue) == 0.00) {
            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("sv-adjust.nozeroadjustments", LanguageID)) %>');
          <% If MyCommon.Fetch_SystemOption(100) <> "1" Then %>
          } else if (((unitType.value != "1") || (expireType.value != "5")) && (parseFloat(elemValue) < 0.00)){
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("sv-adjust.mustrevoke", LanguageID)) %>');
          <% End If %>
          } else if (elemBal != null && parseFloat(elemValue) < 0.00) {
            if (parseFloat(elemBal.value.replace(/,/g, '')) * unitValNum + parseFloat(elemValue.replace(/,/g, '')) < 0.00) {
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("sv-adjust.excessiverevoke", LanguageID)) %>');
            }
          } else if ( (unitType.value == "1") &&  (( (parseInt(parseFloat(valThousand) * decFactor)) % (parseInt(unitValNum) * 1000) ) != 0)) {
            alert("unitType.value:" + unitType.value);
            alert("ValThousand:" + valThousand);
            alert("decFactor:" + decFactor);
            alert("unitValNum:" + unitValNum);

            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("sv-adjust-program.InvalidUnitValue", LanguageID)) %>');
          } else if ((unitType.value != "1") && (((parseFloat(valThousand)) % (parseFloat(unitValNum) * 1000)) != 0)) {
            retVal = false;
            alert("unitType.value:" + unitType.value);
            alert("ValThousand:" + valThousand);
            alert("decFactor:" + decFactor);
            alert("unitValNum:" + unitValNum);

            alert('<%Sendb(Copient.PhraseLib.Lookup("sv-adjust-program.InvalidUnitValue", LanguageID)) %>');
          }
          if (!retVal) {
            elem.focus();
            elem.select();
          } else {
            var elemEnter = document.getElementById("EnterPressed");
            if (elemEnter != null) { elemEnter.value = "true"; }
           x = ((parseFloat(elem.value))/ unitValNum).toFixed(3);
            //adjustUnits.value = Math.round(x);
            adjustUnits.value = parseFloat(x);
          }
        }
      }
      bSkipUnload = true;
      return retVal;
    }
    
    function ChangeParentDocument() {
      var refreshElem = document.getElementById("RefreshParent");

      if (opener != null && !opener.closed) {
        if (refreshElem != null && refreshElem.value == 'true') {
          if (linkToHH) {
            opener.location = '<%Sendb(Opener)%>?CustPK=<%Sendb(HHPK)%><%Sendb(IIf(HHCardPK > 0, "&CardPK=" & HHCardPK, ""))%>&adjWin=1';
          } else {
            opener.location = '<%Sendb(Opener)%>?CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&adjWin=1';
          }
        }
      }
    }
    
    function revokeByID(lid, xid) {
      var msg = '<%Sendb(Copient.PhraseLib.Lookup("sv-adjust-program.ConfirmRevoke", LanguageID)) %>';
      var tokenValues = [];

      tokenValues[0] = xid;
      msg = detokenizeString(msg, tokenValues);
      if (confirm(msg)) {
        document.getElementById('LocalID').value = lid;        
        document.frmSvAdj.submit();
      }
    }
    
    function Refresher(ProgramID, CustomerPK) {
      var refreshElem = document.getElementById("RefreshParent");
      var couponElem = document.getElementById("CouponID");
      var couponToken = "";
      
      var RefreshPrt = "false";
      bSkipUnload = true;
      
      if (refreshElem != null && refreshElem.value == 'true') {
        RefreshPrt = "true";
      }

      if (couponElem != null) {
        couponToken = "&CouponID=" + couponElem.value;
      }
      
      location = ('sv-adjust-program.aspx?ProgramID=' + ProgramID + '&CustomerPK=' + CustomerPK + '&Opener=<%Sendb(Opener)%>'+ '&RefreshParent=' + RefreshPrt + couponToken + '&historyTo=<%Sendb(Server.UrlEncode(Request.QueryString("historyTo")))%>&historyFrom=<%Sendb(Server.UrlEncode(Request.QueryString("historyFrom")))%>&OfferID=<%Sendb(OfferID)%>')
    }
    
    function showTab(tabNbr) {
      var elemTab1 = document.getElementById("tab1");
      var elemTab2 = document.getElementById("tab2");
      var elemTab1Body = document.getElementById("tab1body");
      var elemTab2Body = document.getElementById("tab2body");
      
      if (tabNbr == 1) {
        if (elemTab1 != null) {
          elemTab1.setAttribute("class", "svtabon");
          elemTab1.setAttribute("className", "svtabon");
        }
        if (elemTab2 != null) {
          elemTab2.setAttribute("class", "svtaboff");
          elemTab2.setAttribute("className", "svtaboff");
        }
      } else if (tabNbr == 2) {
        if (elemTab1 != null) {
          elemTab1.setAttribute("class", "svtaboff");
          elemTab1.setAttribute("className", "svtaboff");
        }
        if (elemTab2 != null) {
          elemTab2.setAttribute("class", "svtabon");
          elemTab2.setAttribute("className", "svtabon");
        }
      }
      
      if (elemTab1Body != null && elemTab2Body != null) {
        elemTab1Body.style.display = (tabNbr == 1) ? "" : "none";          
        elemTab2Body.style.display = (tabNbr == 2) ? "" : "none";          
      } 
      
    }
    function HandleSwitchToHH() {
      var refreshElem = document.getElementById("RefreshParent");

      linkToHH = true;
      if (refreshElem != null) {
        refreshElem.value = "true";
      }      
    }
    
    function showDetail(row, btn) {
      var elemTr = document.getElementById("histdetail" + row);
      
      if (elemTr != null && btn != null) {
        elemTr.style.display = (btn.value == "+") ? "" : "none";
        btn.value = (btn.value == "+") ? "-" : "+";  
      }
    }
</script>

<% 
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.AccessStoredValue = False) Then
    Send_Denied(2, "perm.customer-svaccess")
    GoTo done
  End If
  
  If SvAdjust.GetExpireDate(ProgramID) = "" OrElse SvAdjust.GetExpireDate(ProgramID) <= Today() Then
    IsExpired = True
  End If
%>
<form id="frmSvAdj" name="frmSvAdj" action="" onsubmit="return isValidEntry();">
  <input type="hidden" id="ProgramID" name="ProgramID" value="<% Sendb(ProgramID)%>" />
  <input type="hidden" id="ProgramName" name="ProgramName" value="<% Sendb(ProgramName)%>" />
  <input type="hidden" id="CustomerPK" name="CustomerPK" value="<% Sendb(CustomerPK)%>" />
  <input type="hidden" id="decFactor" name="decFactor" value="<% Sendb(decFactor)%>" />
  <%
    If (CardPK > 0) Then
      Send("  <input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
    End If
  %>
  <input type="hidden" id="Opener" name="Opener" value="<% Sendb(Opener)%>" />
  <input type="hidden" id="EnterPressed" name="EnterPressed" value="" />
  <input type="hidden" id="LocalID" name="LocalID" value="" />
  <input type="hidden" id="CouponID" name="CouponID" value="<% Sendb(CouponID)%>" />
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID)%>" />
  <input type="hidden" id="RefreshParent" name="RefreshParent" value="<% Sendb(RefreshParent.ToString.ToLower) %>" />
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID) & " #" & ProgramID & ": " & MyCommon.TruncateString(ProgramName, 40))%>
    </h1>
    <div id="controls">
      <%--
      <input type="submit" class="regular" id="save" name="save" value="<% Sendb(Copient.PhraseLib.Lookup("term.save", LanguageID))%>"<% IIf(Logix.UserRoles.EditPointsBalances, "", " disabled=""disabled""") %> />
      --%>
    </div>
    <hr class="hidden" />
  </div>
  <div id="main" style="width: 100%;">
    <%If (InfoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")
        If CouponNotFound AndAlso CouponID <> "" Then GoTo done
      End If
    %>
    <div id="column">
      <%
        If HHPK > 0 AndAlso CustomerPK > 0 Then
          Sendb("<br /><a href=""sv-adjust-program.aspx?")
          KeyCt = Request.QueryString.Keys.Count
          For i = 0 To KeyCt - 1
            If (Request.QueryString.Keys(i) = "CustomerPK") Then
              Sendb("CustomerPK=" & HHPK)
            ElseIf (Request.QueryString.Keys(i) = "CardPK") Then
              If HHCardPK > 0 Then
                Sendb("CardPK=" & HHCardPK)
              Else
                Sendb(Request.QueryString.Keys(i) & "=" & Request.QueryString.Item(i))
              End If
            Else
              Sendb(Request.QueryString.Keys(i) & "=" & Request.QueryString.Item(i))
            End If
            If (i < KeyCt - 1) Then Sendb("&")
          Next
          Sendb("&RefreshParent=true"" onclick=""javascript:HandleSwitchToHH();"" >")
          Send(Copient.PhraseLib.Lookup("sv.hh-adjust-linktext", LanguageID) & "</a><br/><br/>")
        End If
      %>
      <div class="box" id="svAdj" <% Sendb(IIf((HHPK > 0), " style=""display:none;""", "")) %>>
        <h2>
          <span>
            <%Sendb(Copient.PhraseLib.Lookup("term.storedvalueadjustment", LanguageID))%>
          </span>
        </h2>
        <% 
          If (ProgramID > 0) Then
            ShowStoredValue(CustomerPK, ProgramID, ProgramName, Logix.UserRoles.EditPointsBalances, OffersLinked, CouponID)
          Else
            Send(Copient.PhraseLib.Lookup("sv-adjust.nosvprograms", LanguageID) & "<br />")
          End If
        %>
        <hr class="hidden" />
      </div>
	  <%
	If (MyCommon.Fetch_SystemOption(193) = "1") Then
          'Require a reason
          Send("<div class=""box"" id=""reason"">")
          Send("  <h2>")
          Send("    <span>")
          Send("      " & Copient.PhraseLib.Lookup("term.reason", LanguageID))
          Send("    </span>")
          Send("  </h2>")
          Send("  <select id=""reasonID"" name=""reasonID"" style=""width:220px;"">")
          Send("    <option value=""0"">(" & Copient.PhraseLib.Lookup("term.SelectAReason", LanguageID) & ")</option>")
          Dim RC As DataTable
          Dim row As DataRow
          MyCommon.QueryStr = "select ReasonID, Description from AdjustmentReasons with (NoLock) where Enabled=1 and (Program like '%Value%' OR Program like '%All%' OR Program is NULL);"
          RC = MyCommon.LXS_Select
          If RC.Rows.Count > 0 Then
            For Each row In RC.Rows
              Send("    <option value=""" & MyCommon.NZ(row.Item("ReasonID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "&nbsp;") & "</option>")
            Next
          End If
          Send("  </select>")
          Send("  <input id=""reasonText"" name=""reasonText"" type=""text"" value="""" maxlength=""255"" style=""width:400px;"" />")
          Send("  <hr class=""hidden"" />")
          Send("</div>")
          Send("")
	End If
%>
      <% If Not IsExpired Then%>
      <div class="box" id="pending">
        <h2>
          <span>
            <%Sendb(Copient.PhraseLib.Lookup("term.pending", LanguageID))%>
          </span>
        </h2>
        <span style="float: right; font-size: 9px; position: relative; top: -22px;"><a href="javascript:Refresher(<%Sendb(ProgramID)%>,<%Sendb(CustomerPK)%>)">
          <%Sendb(Copient.PhraseLib.Lookup("term.refresh", LanguageID))%>
        </a></span>
        <% 
          If (ProgramID > 0) Then
            ShowPending(CustomerPK, ProgramID)
          Else
            Send(Copient.PhraseLib.Lookup("sv-adjust.nosvprograms", LanguageID) & "<br />")
          End If
        %>
        <hr class="hidden" />
      </div>
      <% End If%>
      <div class="box" id="history">
        <h2>
          <span>
            <%Sendb(Copient.PhraseLib.Lookup("term.history", LanguageID))%>
          </span>
        </h2>
        <div class="boxscrollfull">
          <% 
            If (ProgramID > 0) Then
              ShowHistory(CustomerPK, CardPK, ProgramID, False, CouponID)
            Else
              Send(Copient.PhraseLib.Lookup("sv-adjust.nosvprograms", LanguageID) & "<br />")
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>



<script runat="server">
    Dim dtUsers As DataTable = Nothing
    Dim Localization As New Copient.Localization(MyCommon)
    Dim MyCryptLib As New Copient.CryptLib

    Sub ShowStoredValue(ByVal CustomerPK As Long, ByVal ProgramID As Integer, ByVal ProgramName As String, ByVal Editable As Boolean, ByVal OffersLinked As Boolean, ByVal CouponID As String)
        Dim UpdateAccum As Boolean = False
        Dim dt, dtPrograms As DataTable
        Dim rowProgram As DataRow = Nothing
        Dim row As DataRow = Nothing
        Dim OfferName As String = ""
        Dim Quantity As Integer = 0
        Dim TotalQuantity As Integer = 0
        Dim Value As Double = 0D
        Dim OfferID As Long = 0
        Dim i As Integer = 0
        Dim PendingAdj As Integer = 0
        Dim TotalPending As Integer = 0
        Dim OfferIDs As ArrayList = Nothing
        Dim OfferNames As ArrayList = Nothing
        Dim IsExpired As Boolean = False
        Dim SvAdjust As New Copient.StoredValue
        Dim ValueString As String = ""
        Dim ExpiredString As String = ""

        Dim objTemp As Object
        Dim intNumDecimalPlaces As Integer = 0
        Dim decFactor As Decimal = 1.0
        Dim decTemp As Decimal
        Dim EnableFuelPartner As Boolean = False
        Dim DisableAdjust As Boolean = False
        'Dim Localizer As Copient.Localization
        MyCommon.Open_LogixRT()
        MyCommon.Open_LogixXS()

        If MyCommon.IsEngineInstalled(0) Then
            EnableFuelPartner = MyCommon.Fetch_CM_SystemOption(56)

            objTemp = MyCommon.Fetch_CM_SystemOption(41)
            If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
                intNumDecimalPlaces = 0
            End If
            decFactor = (10 ^ intNumDecimalPlaces)
        End If

        MyCommon.QueryStr = "select distinct I.IncentiveID, I.IncentiveName, SVP.Name, SVP.Value from CPE_Deliverables D with (NoLock)  " &
                            "inner join CPE_DeliverableStoredValue DSV with (NoLock) on DSV.PKID = D.OutputID and DSV.Deleted=0 and D.Deleted=0 and D.DeliverableTypeID=11  " &
                            "inner join StoredValuePrograms SVP with (NoLock) on SVP.SVProgramID = DSV.SVProgramID and SVP.Deleted=0  " &
                            "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = DSV.RewardOptionID and RO.Deleted=0  " &
                            "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID and I.Deleted = 0 and I.IsTemplate=0  " &
                            "where DSV.Quantity >= 0 and SVP.SVProgramID= " & ProgramID &
                            " union " &
                            "select distinct I.IncentiveID, I.IncentiveName, SVP.Name, SVP.Value from CPE_IncentiveStoredValuePrograms ISVP with (NoLock) " &
                            "inner join StoredValuePrograms SVP with (NoLock) on SVP.SVProgramID = ISVP.SVProgramID and SVP.Deleted=0 and ISVP.Deleted=0 " &
                            "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ISVP.RewardOptionID and RO.Deleted=0 " &
                            "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID and I.Deleted = 0 and I.IsTemplate=0 " &
                            "where SVP.SVProgramID= " & ProgramID & " " &
                            "UNION  " &
                             "SELECT DISTINCT I.IncentiveID, I.IncentiveName, SVP.Name, SVP.Value from CPE_DeliverableMonStoredValue DMSV with (NoLock) " &
                             "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DMSV.RewardOptionID  " &
                             "INNER JOIN StoredValuePrograms SVP with (NoLock) on SVP.SVProgramID = DMSV.SVProgramID and SVP.Deleted=0 " &
                             "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID  " &
                             "WHERE DMSV.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and DMSV.SVProgramID= " & ProgramID & " " &
                            " union " &
                            "select distinct I.IncentiveID, I.IncentiveName, SVP.Name, SVP.Value from CPE_Deliverables D with (NoLock)  " &
                            "inner join CPE_Discounts CD with (NoLock) on CD.DiscountID = D.OutputID and CD.Deleted=0 and D.Deleted=0 and D.DeliverableTypeID=2  " &
                            "inner join StoredValuePrograms SVP with (NoLock) on SVP.SVProgramID = CD.SVProgramID and SVP.Deleted=0  " &
                            "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = D.RewardOptionID and RO.Deleted=0  " &
                            "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID and I.Deleted = 0 and I.IsTemplate=0  " &
                            "where SVP.SVProgramID= " & ProgramID &
                             " union " &
                            "select distinct O.OfferID as IncentiveID, O.Name as IncentiveName, SVP.Name, SVP.Value " &
                            "from Offers O with (NoLock) " &
                            "inner join OfferConditions OC with (NoLock) on OC.OfferID = O.OfferID and O.Deleted= 0 and OC.Deleted=0 " &
                            "inner join StoredValuePrograms SVP with (NoLock) on SVP.SVProgramID = OC.LinkID and SVP.Deleted=0 " &
                            "where OC.ConditionTypeId = 6 and SVP.SVProgramID=" & ProgramID & " " &
                            " union " &
                            "select distinct O.OfferID as IncentiveID, O.Name as IncentiveName, SVP.Name, SVP.Value " &
                            "from Offers O with (NoLock) " &
                            "inner join OfferRewards ORew with (NoLock) on ORew.OfferID = O.OfferID and O.Deleted=0 and ORew.Deleted = 0 " &
                            "left join CM_RewardStoredValues as RSV with (NoLock) on RSV.RewardStoredValuesID=ORew.LinkID " &
                            "left join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=RSV.ProgramID " &
                            "where ORew.RewardTypeID=10 and SVP.SVProgramID = " & ProgramID & " " &
                            " UNION " &
                            "Select DISTINCT I.IncentiveID, I.IncentiveName, SVP.Name, SVP.Value  " &
                            "From CPE_DeliverableMonStoredValue DMSV with (NoLock)  " &
                            "inner Join StoredValuePrograms SVP with (NoLock) on SVP.SVProgramID = DMSV.SVProgramID And SVP.Deleted=0 And DMSV.Deleted=0  " &
                            "INNER Join  " &
                            "CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DMSV.RewardOptionID  INNER JOIN CPE_Incentives I with  " &
                            "(NoLock) on I.IncentiveID = RO.IncentiveID  left outer join Buyers as buy with (nolock) on buy.BuyerId=  " &
                            "i.BuyerId WHERE DMSV.Deleted=0 And RO.Deleted=0 And I.Deleted=0 And I.IsTemplate=0 And SVP.SVProgramID=" & ProgramID

        dtPrograms = MyCommon.LRT_Select

        Send("<br class=""half"" />")

        MyCommon.QueryStr = "select SVP.Name, SVP.SVTypeID, SVP.SVExpireType, SVP.FuelPartner, SVP.AllowAdjustments " & _
                            "from StoredValuePrograms as SVP with (NoLock) " & _
                            "where SVP.SVProgramID=" & ProgramID & ";"
        dt = MyCommon.LRT_Select

        If (dt.Rows.Count > 0 AndAlso dt.Rows(0).Item("SVTypeID") = 1 AndAlso dt.Rows(0).Item("SVExpireType") = 5) Then
            Send(Copient.PhraseLib.Lookup("sv-adjust.negativenote", LanguageID))
        Else
            If (EnableFuelPartner And MyCommon.NZ(dt.Rows(0).Item("FuelPartner"), False)) Then
                If Not MyCommon.NZ(dt.Rows(0).Item("AllowAdjustments"), True) Then
                    DisableAdjust = True
                End If
            End If
            If DisableAdjust Then
                Send(Copient.PhraseLib.Lookup("term.adjustmntsdisabled", LanguageID))
            Else
                Send(Copient.PhraseLib.Lookup("sv-adjust.programnote", LanguageID))
            End If
        End If
        Send("<br /><br class=""half"" />")

        If SvAdjust.GetExpireDate(ProgramID) = "" OrElse SvAdjust.GetExpireDate(ProgramID) <= Today() Then
            IsExpired = True
        End If

        If (dtPrograms.Rows.Count > 0) Then
            OfferIDs = New ArrayList(dtPrograms.Rows.Count)
            OfferNames = New ArrayList(dtPrograms.Rows.Count)

            ' get the list of all offers using this stored value program
            For Each rowProgram In dtPrograms.Rows
                OfferID = MyCommon.NZ(rowProgram.Item("IncentiveID"), 0)
                OfferName = MyCommon.NZ(rowProgram.Item("IncentiveName"), "")
                OfferIDs.Add(OfferID)
                OfferNames.Add(OfferName)
            Next

        End If

        ' Find the stored value balance
        TotalQuantity = SvAdjust.GetQuantityBalance(CustomerPK, ProgramID, CouponID)

        ' Draw the UI
        MyCommon.QueryStr = "select SVP.Value, SVP.UnitOfMeasureLimit, SVP.SVTypeID, SVT.PhraseID, SVT.ValuePrecision, SVP.SVExpireType " & _
                            "from StoredValuePrograms as SVP with (NoLock) " & _
                            "inner join SVTypes as SVT on SVP.SVTypeID=SVT.SVTypeID " & _
                            "where SVProgramID=" & ProgramID & ";"
        dt = MyCommon.LRT_Select
        Send("<table summary="""">")
        Send("  <tr>")
        Send("    <td style=""width:85px;""></td>")
        Send("    <td style=""width:150px;""></td>")
        Send("    <td id=""svtabs"">")
        Send("      <span id=""tab1"" class=""svtabon"" onclick=""showTab(1);"">" & Copient.PhraseLib.Lookup("term.programdetails", LanguageID) & "</span>")
        Send("      <span id=""tab2"" class=""svtaboff"" onclick=""showTab(2);"">" & Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID) & "</span>")
        Send("    </td>")
        Send("  </tr>")
        Send("  <tr>")
        Send("    <td><b>" & Copient.PhraseLib.Lookup("term.balance", LanguageID) & ":</b></td>")
        Sendb("    <td>")
        If (dt.Rows(0).Item("SVTypeID") > 1) Then
            ValueString = Math.Round(TotalQuantity * dt.Rows(0).Item("Value"), Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
            Sendb(ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator))
        Else
            decTemp = (Int(TotalQuantity * dt.Rows(0).Item("Value")) * 1.0) / decFactor
            ValueString = CStr(FormatNumber(decTemp, intNumDecimalPlaces) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase))
            Sendb(ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator))
        End If
        decTemp = (TotalQuantity * 1.0) / decFactor
        Sendb(" (" & FormatNumber(decTemp, intNumDecimalPlaces) & " " & StrConv(Copient.PhraseLib.Lookup("term.units", LanguageID), VbStrConv.Lowercase) & ")")
        Sendb("<input type=""hidden"" name=""Balance"" id=""Balance"" value=""" & FormatNumber(decTemp, intNumDecimalPlaces) & """ />")
        Sendb("<input type=""hidden"" name=""UnitValue"" id=""UnitValue"" value=""" & MyCommon.NZ(dt.Rows(0).Item("Value"), 0) & """ />")
        Sendb("<input type=""hidden"" name=""UnitType"" id=""UnitType"" value=""" & MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) & """ />")
        Sendb("<input type=""hidden"" name=""ExpireType"" id=""ExpireType"" value=""" & MyCommon.NZ(dt.Rows(0).Item("SVExpireType"), 0) & """ />")
        Send("</td>")
        Send("    <td rowspan=""3"">")
        Send("      <div id=""tab1body"" class=""boxscroll"" style=""height:80px;"">")
        ExpiredString = SvAdjust.GetExpireDate(ProgramID)
        If ExpiredString <> "" Then
            ExpiredString = Logix.ToShortDateTimeString(Date.Parse(ExpiredString), MyCommon)
        Else
            ExpiredString = Copient.PhraseLib.Lookup("term.not-available", LanguageID)
        End If
        Send("      <span style=""font-weight:bold;width:90px;"">" & Copient.PhraseLib.Lookup("storedvalue.expiredate", LanguageID) & ": </span>" & ExpiredString & "<br />")
        If (dt.Rows.Count > 0) Then
            If (MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) = 3) Then
                Send("<span style=""font-weight:bold;width:90px;"">" & Copient.PhraseLib.Lookup("term.unitlimit", LanguageID) & ": </span>" & MyCommon.NZ(dt.Rows(0).Item("UnitOfMeasureLimit"), ""))
                Send("<br />")
            End If
            Sendb("<span style=""font-weight:bold;width:90px;"">" & Copient.PhraseLib.Lookup("term.value", LanguageID) & ": </span>")
            If MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) > 1 Then
                'Localization = New Copient.Localization(MyCommon)
                ValueString = Localization.Get_Default_Currency_Symbol() & MyCommon.ConvertToCurrentCultureDecimalSymbol(MyCommon.NZ(dt.Rows(0).Item("Value"), 0))
                If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                    ValueString = Left(ValueString, Len(ValueString) - 1)
                End If
                ValueString = ValueString & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase)
                If (MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) = 3 OrElse MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) = 5) Then
                    ValueString = ValueString & " " & StrConv(Copient.PhraseLib.Lookup("term.per", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.unitofmeasure", LanguageID), VbStrConv.Lowercase)
                End If
            Else
                ValueString = CStr(Int(MyCommon.NZ(dt.Rows(0).Item("Value"), 0))) & " " & StrConv(Copient.PhraseLib.Lookup(MyCommon.NZ(dt.Rows(0).Item("PhraseID"), 0), LanguageID), VbStrConv.Lowercase)
            End If
            Sendb(ValueString)
        End If
        Send("      </div>")
        Send("      <div id=""tab2body"" class=""boxscroll"" style=""height:80px;display:none;"">")
        If (OfferIDs IsNot Nothing AndAlso OfferNames IsNot Nothing AndAlso OfferIDs.Count = OfferNames.Count) Then
            For i = 0 To OfferIDs.Count - 1
                If OffersLinked Then
                    Send("  <a href=""offer-redirect.aspx?OfferID=" & OfferIDs(i) & """ target=""main"">" & OfferNames(i) & "</a><br />")
                Else
                    Send("  " & OfferNames(i) & "<br />")
                End If
            Next
        End If
        Send("      </div>")
        Send("    </td>")
        'Send("    <td>" & TotalQuantity & "<input type=""hidden"" id=""Balance"" name=""Balance"" value=""" & TotalQuantity & """ /></td>")
        'Send("    <td><input type=""text"" class=""short"" id=""adjust"" name=""adjust"" value=""""" & IIf(Editable, "", " disabled=""disabled""") & " /></td>")
        Send("  </tr>")
        Send("  <tr>")
        Send("    <td><b><label for=""adjust"">" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & ":</label></b></td>")
        Dim CurrencySymbol As String = String.Empty
        If MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) > 1 Then
            CurrencySymbol = Localization.Get_Default_Currency_Symbol()
        End If
        Sendb("    <td>" & CurrencySymbol & "<input type=""text"" class=""short"" name=""adjust"" id=""adjust"" maxlength=""6"" value=""""" & IIf((Editable = False) OrElse (IsExpired = True) OrElse CustomerPK = 0, " disabled=""disabled""", "") & " />")
        Sendb("    " & IIf(MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0) = 1, StrConv(Copient.PhraseLib.Lookup(MyCommon.NZ(dt.Rows(0).Item("PhraseID"), 0), LanguageID), VbStrConv.Lowercase), ""))
        Sendb("<input type=""hidden"" id=""adjustUnits"" name=""adjustUnits"" value="""" />")
        Send("</td>")
        Send("  </tr>")
        If (MyCommon.Fetch_SystemOption(108) = "1") Then
            Send("  <tr>")
            Send("    <td><b><label for=""location"">" & Copient.PhraseLib.Lookup("term.issuingcostcenter", LanguageID) & ":</label></b></td>")
            Sendb("    <td><input type=""text"" class=""short"" name=""location"" id=""location"" maxlength=""25"" value=""""" & IIf((Editable = False) OrElse (IsExpired = True) OrElse CustomerPK = 0, " disabled=""disabled""", "") & " />")
            Send("</td>")
            Send("  </tr>")
        End If

        If MyCommon.Fetch_SystemOption(99) AndAlso ((Editable = True) And (IsExpired = False) And (CustomerPK > 0)) Then
            Send("  <tr>")
            Send("    <td></td>")
            Send("    <td></td>")
            Send("  </tr>")
            Send("  <tr>")
            Send("    <td style=""vertical-align:top;""><b><label for=""note"">" & Copient.PhraseLib.Lookup("term.explanation", LanguageID) & ":</label></b></td>")
            Send("    <td colspan=""2""><textarea id=""note"" name=""note"" style=""width:506px;font-size:12px;""></textarea></td>")
            Send("  </tr>")
            Send("  <tr>")
            Send("    <td></td>")
            Send("    <td><input type=""submit"" class=""short"" name=""save"" id=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """" & IIf((Editable = False) OrElse (IsExpired = True) OrElse CustomerPK = 0 OrElse DisableAdjust, " disabled=""disabled""", "") & " /></td>")
            Send("  </tr>")
        Else
            Send("  <tr>")
            Send("    <td></td>")
            Send("    <td><input type=""submit"" class=""short"" name=""save"" id=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """" & IIf((Editable = False) OrElse (IsExpired = True) OrElse CustomerPK = 0 OrElse DisableAdjust, " disabled=""disabled""", "") & " /></td>")
            Send("  </tr>")
        End If
        Send("</table>")
    End Sub

    Sub ShowPending(ByVal CustomerPK As Integer, ByVal ProgramID As Integer)
        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim row As DataRow
        Dim dtTime As DateTime
        Dim TotalEarned As Integer = 0
        Dim TotalUsed As Integer = 0
        Dim ValueString As String = ""
        Dim Value As Decimal = 0
        Dim objTemp As Object
        Dim intNumDecimalPlaces As Integer = 0
        Dim decFactor As Decimal = 1.0
        Dim decTemp As Decimal

        If MyCommon.IsEngineInstalled(0) Then
            objTemp = MyCommon.Fetch_CM_SystemOption(41)
            If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
                intNumDecimalPlaces = 0
            End If
            decFactor = (10 ^ intNumDecimalPlaces)
        End If

        MyCommon.QueryStr = "select Operation, Col11 as ExternalID, Col6 as QtyEarned, Col7 as QtyUsed, Col8 as [Value], Col9 as EarnedDate " & _
                            "from CPE_UploadTemp_SV with (NoLock) " & _
                            "where TableNum = '10' and Operation ='1' and Col3 = " & ProgramID & " and Col5 = " & CustomerPK & " " & _
                            "union " & _
                            "select Operation, Col7 as ExternalID, 0 as QtyEarned, Col6 as QtyUsed, Col10 as [Value], Col9 as EarnedDate " & _
                            "from CPE_UploadTemp_SV with (NoLock) " & _
                            "where TableNum = '10' and Operation ='3' and Col11 = " & ProgramID & " and Col3 = " & CustomerPK & " " & _
                            "order by EarnedDate DESC;"
        dt = MyCommon.LXS_Select

        If (dt.Rows.Count > 0) Then
            Value = MyCommon.NZ(dt.Rows(0).Item("Value"), 0)
            MyCommon.QueryStr = "select SVP.SVTypeID, SVT.PhraseID, SVT.ValuePrecision " & _
                                "from StoredValuePrograms as SVP with (NoLock) " & _
                                "inner join SVTypes as SVT with (NoLock) on SVT.SVTypeID=SVP.SVTypeID " & _
                                "where SVProgramID=" & ProgramID & ";"
            dt2 = MyCommon.LRT_Select
            Send("        <table summary=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """>")
            Send("         <thead>")
            Send("          <tr>")
            Send("            <th scope=""col"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "</th>")
            Send("            <th scope=""col"">" & Copient.PhraseLib.Lookup("term.earneddate", LanguageID) & "</th>")
            Send("            <th scope=""col"" style=""text-align:right"">" & Copient.PhraseLib.Lookup("term.AmountEarned", LanguageID) & "</th>")
            Send("            <th scope=""col"" style=""text-align:right"">" & Copient.PhraseLib.Lookup("term.AmountUsed", LanguageID) & "</th>")
            Send("          </tr>")
            Send("         </thead>")
            Send("         <tbody>")
            For Each row In dt.Rows
                Send("           <tr>")
                Send("             <td>" & row.Item("ExternalID").ToString() & "</td>")
                dtTime = (System.ComponentModel.TypeDescriptor.GetConverter(New DateTime(1990, 5, 6)).ConvertFrom(MyCommon.NZ(row.Item("EarnedDate"), Now)))
                Send("             <td>" & Logix.ToShortDateTimeString(dtTime, MyCommon) & "</td>")
                If dt2.Rows(0).Item("SVTypeID") > 1 Then
                    ValueString = (MyCommon.NZ(row.Item("QtyEarned"), 0) * Value)
                    If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                        ValueString = Left(ValueString, Len(ValueString) - 1)
                    End If
                    Send("             <td style=""text-align:right"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
                    ValueString = Math.Round(MyCommon.NZ(row.Item("QtyUsed"), 0) * Value, Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
                    If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                        ValueString = Left(ValueString, Len(ValueString) - 1)
                    End If
                    Send("             <td style=""text-align:right"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
                Else
                    decTemp = (Int(MyCommon.NZ(row.Item("QtyEarned"), 0) * Value) * 1.0) / decFactor
                    Send("             <td style=""text-align:right"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
                    decTemp = (Int(MyCommon.NZ(row.Item("QtyUsed"), 0) * Value) * 1.0) / decFactor
                    Send("             <td style=""text-align:right"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
                End If
                Send("           </tr>")
                TotalEarned += MyCommon.NZ(row.Item("QtyEarned"), 0)
                TotalUsed += MyCommon.NZ(row.Item("QtyUsed"), 0)
            Next
            Send("           <tr>")
            Send("             <td colspan=""2""></td>")
            If dt2.Rows(0).Item("SVTypeID") > 1 Then
                ValueString = Math.Round(TotalEarned * Value, Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
                If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                    ValueString = Left(ValueString, Len(ValueString) - 1)
                End If
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
                ValueString = Math.Round(TotalUsed * Value, Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
                If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                    ValueString = Left(ValueString, Len(ValueString) - 1)
                End If
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
            Else
                decTemp = (Int(TotalEarned * Value) * 1.0) / decFactor
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
                decTemp = (Int(TotalUsed * Value) * 1.0) / decFactor
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
            End If
            Send("          </tr>")
            Send("         </tbody>")
            Send("       </table>")
        End If
    End Sub

    Sub ShowHistory(ByVal CustomerPK As Integer, ByVal CardPK As Integer, ByVal ProgramID As Integer, ByVal ShowExpired As Boolean, _
                    Optional ByVal CouponID As String = "")
        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim dt3 As DataTable
        Dim row As DataRow
        Dim TotalEarned As Integer = 0
        Dim TotalUsed As Integer = 0
        Dim PageNum As Integer = 0
        Dim MorePages As Boolean
        'Dim linesPerPage As Integer = 15
        Dim linesPerPage As Integer = 50
        Dim Shaded As String = ""
        Dim PrevLocalID As Long = 0
        Dim ValueString As String = ""
        Dim Value As Decimal = 0
        Dim StartDate, EndDate As Date
        Dim StartDateStr As String = ""
        Dim EndDateStr As String = ""
        Dim Cookie As HttpCookie = Nothing
        Dim HistoryEnabled As Boolean = True
        Dim AltText As String = ""
        Dim SortText As String = "ActivityDate"
        Dim SortDirection As String = ""
        Dim HistoryRecordsTotalCount As Integer = 0
        Dim PageCount As Integer = 0
        Dim HistoryCountShown As Integer = 0
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim OrderText As String = ""
        Dim objTemp As Object
        Dim intNumDecimalPlaces As Integer = 0
        Dim decFactor As Decimal = 1.0
        Dim decTemp As Decimal
        Dim PresentedCustomerID As String = ""
        Dim ResolvedCustomerID As String = ""
        Dim EnableFuelPartner As Boolean = False
        Dim DisableAdjust As Boolean = False

        If MyCommon.IsEngineInstalled(0) Then
            EnableFuelPartner = MyCommon.Fetch_CM_SystemOption(56)

            objTemp = MyCommon.Fetch_CM_SystemOption(41)
            If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
                intNumDecimalPlaces = 0
            End If
            decFactor = (10 ^ intNumDecimalPlaces)
        End If

        'PageNum = Request.QueryString("pagenum")
        If Not Integer.TryParse(Request.QueryString("pagenum"), PageNum) Then
            PageNum = 1
        End If

        If PageNum < 0 Then
            PageNum = 1
            MorePages = False
        End If

        If (Request.QueryString("pagenum") = "") Then
            If (Request.QueryString("SortDirection") = "ASC") Then
                SortDirection = "DESC"
            ElseIf (Request.QueryString("SortDirection") = "DESC") Then
                SortDirection = "ASC"
            Else
                SortDirection = "DESC"
            End If
        Else
            SortDirection = Request.QueryString("SortDirection")
        End If
        If Request.QueryString("SortText") <> "" Then
            SortText = Request.QueryString("SortText")
        Else
            SortText = ""
        End If

        ' determine if history should be loaded per the users request
        Cookie = Request.Cookies("SVHistoryEnabled")
        If Not (Cookie Is Nothing) Then
            HistoryEnabled = IIf(Cookie.Value = "0", False, True)
        End If

        ' if user is attempting a search when history display is disabled, then enable it for them
        If (Not HistoryEnabled AndAlso Request.QueryString("SearchHistory") <> "") Then
            Response.Cookies("SVHistoryEnabled").Expires = "10/08/2100"
            Response.Cookies("SVHistoryEnabled").Value = "1"
            HistoryEnabled = True
        End If

        If Not Date.TryParse(Request.QueryString("historyFrom"), StartDate) Then
            StartDate = Date.Now.AddDays(-30)
        End If
        StartDate = New Date(StartDate.Year, StartDate.Month, StartDate.Day, 0, 0, 0)

        If Date.TryParse(Request.QueryString("historyTo"), EndDate) Then
            If EndDate < StartDate Then
                EndDate = Date.Now
            End If
        Else
            EndDate = Date.Now
        End If
        EndDate = New Date(EndDate.Year, EndDate.Month, EndDate.Day, 23, 59, 59)

        If HistoryEnabled Then
            StartDateStr = Logix.ToShortDateString(StartDate, MyCommon)
            EndDateStr = Logix.ToShortDateString(EndDate, MyCommon)
            AltText = Copient.PhraseLib.Lookup("customer-inquiry.noloadhistory", LanguageID)
        Else
            StartDateStr = ""
            EndDateStr = ""
            AltText = Copient.PhraseLib.Lookup("customer-inquiry.loadhistory", LanguageID)
        End If

        Send("<div style=""width:100%;background-color:#e0e0e0;text-align:center;border:1px solid #808080;"">")
        Sendb("<input type=""image"" name=""" & IIf(HistoryEnabled, "HistoryDisabled", "HistoryEnabled") & """ src=""" & IIf(HistoryEnabled, "/images/history-on.png", "/images/history-off.png") & """")
        Send("  alt=""" & AltText & """ title=""" & AltText & """ style=""position:absolute;left:30px;margin-top:2px;"" onclick=""javascript:bSkipValidate=true;"" />")
        Send("<label for=""historyFrom""><b>" & Copient.PhraseLib.Lookup("term.startdate", LanguageID) & "</b>:</label>")
        Send("<input type=""text"" id=""historyFrom"" name=""historyFrom"" class=""short"" value=""" & StartDateStr & """ />")
        Send("<img src=""/images/calendar.png"" class=""calendar"" id=""start-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('historyFrom', event);"" />")
        Send("&nbsp;&nbsp;")
        Send("<label for=""historyTo""><b>" & Copient.PhraseLib.Lookup("term.enddate", LanguageID) & "</b>:</label>")
        Send("<input type=""text"" id=""historyTo"" name=""historyTo"" class=""short"" value=""" & EndDateStr & """ />")
        Send("<img src=""/images/calendar.png"" class=""calendar"" id=""end-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('historyTo', event);"" />")
        Send("&nbsp;&nbsp;")
        Send("<input type=""submit"" id=""SearchHistory"" name=""SearchHistory"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""javascript:bSkipValidate=true;"" />")
        Send("</div>")
        Send("<br class=""half"" />")

        If Not HistoryEnabled Then Exit Sub

        If SortText <> "" Then
            OrderText = " order by " & SortText & " " & SortDirection & ";"
        Else
            OrderText = " order by EarnedDate DESC, LocalID, ExternalID, Status;"
        End If

        'check the system option 117 to get the count of records to be fetched from database
        If MyCommon.Fetch_SystemOption(117) = "ALL" Then
            MyCommon.QueryStr = "select CustomerPK, LocalID, ExternalID, EarnedDate, QtyEarned, QtyUsed, Value, ServerSerial, LastLocationID, " & _
                                  "  SVS.Description as Status, SVS.PhraseID as StatusPhraseID, " & _
                                  "  IsNull(SUM(QtyEarned) OVER(PARTITION BY CustomerPK,LocalID,ServerSerial),0) - IsNull(SUM(QtyUsed) OVER(PARTITION BY CustomerPK,LocalID,ServerSerial),0) as QtyAvail, " & _
                                  "  AdminUserID, ExpireDate, PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate " & _
                                  "  from SVHistory SVH with (NoLock) " & _
                                  "inner join StoredValueStatus SVS with (NoLock) on SVS.StatusID=SVH.StatusFlag " & _
                                  "where LastLocationID <> -8 and CustomerPK=" & CustomerPK & " and SVProgramID=" & ProgramID & IIf(ShowExpired, " and StatusFlag = 3 ", " ") & " " & _
                                  "and LastUpdate between '" & StartDate.ToString(CultureInfo.InvariantCulture) & "' and '" & EndDate.ToString(CultureInfo.InvariantCulture) & "' " & _
                                  IIf(CouponID = "", "", " and ExternalID='" & CouponID & "' ") & " " & _
                                  "and Deleted=0" & OrderText
        Else
            MyCommon.QueryStr = "select Top " & MyCommon.Fetch_SystemOption(117) & " CustomerPK, LocalID, ExternalID, EarnedDate, QtyEarned, QtyUsed, Value, ServerSerial, LastLocationID, " & _
                                  "  SVS.Description as Status, SVS.PhraseID as StatusPhraseID, " & _
                                  "  IsNull(SUM(QtyEarned) OVER(PARTITION BY CustomerPK,LocalID,ServerSerial),0) - IsNull(SUM(QtyUsed) OVER(PARTITION BY CustomerPK,LocalID,ServerSerial),0) as QtyAvail, " & _
                                  "  AdminUserID, ExpireDate, PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate " & _
                                  "  from SVHistory SVH with (NoLock) " & _
                                  "inner join StoredValueStatus SVS with (NoLock) on SVS.StatusID=SVH.StatusFlag " & _
                                  "where LastLocationID <> -8 and CustomerPK=" & CustomerPK & " and SVProgramID=" & ProgramID & IIf(ShowExpired, " and StatusFlag = 3 ", " ") & " " & _
                                  "and LastUpdate between '" & StartDate.ToString(CultureInfo.InvariantCulture) & "' and '" & EndDate.ToString(CultureInfo.InvariantCulture) & "' " & _
                                  IIf(CouponID = "", "", " and ExternalID='" & CouponID & "' ") & " " & _
                                  "and Deleted=0" & OrderText
        End If

        'this is used to get count of total records
        dt = MyCommon.LXS_Select

        If (dt.Rows.Count > 0) Then
            HistoryRecordsTotalCount = dt.Rows.Count
            PageCount = Math.Ceiling(HistoryRecordsTotalCount / linesPerPage)

            'clear the existing contents
            dt.Clear()

            MyCommon.QueryStr = "SELECT TOP " & linesPerPage & " * FROM (select CustomerPK, LocalID, ExternalID, EarnedDate, QtyEarned, QtyUsed, Value, ServerSerial, LastLocationID, " & _
                                 "  SVS.Description as Status, SVS.PhraseID as StatusPhraseID, " & _
                                 "  IsNull(SUM(QtyEarned) OVER(PARTITION BY CustomerPK,LocalID,ServerSerial),0) - IsNull(SUM(QtyUsed) OVER(PARTITION BY CustomerPK,LocalID,ServerSerial),0) as QtyAvail, " & _
                                 "  AdminUserID, ExpireDate, PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate,ROW_NUMBER() OVER(ORDER BY SVH.EarnedDate DESC) AS RowID " & _
                                 "  from SVHistory SVH with (NoLock) " & _
                                 "inner join StoredValueStatus SVS with (NoLock) on SVS.StatusID=SVH.StatusFlag " & _
                                 "where LastLocationID <> -8 and CustomerPK=" & CustomerPK & " and SVProgramID=" & ProgramID & IIf(ShowExpired, " and StatusFlag = 3 ", " ") & " " & _
                                 "and LastUpdate between '" & StartDate.ToString(CultureInfo.InvariantCulture) & "' and '" & EndDate.ToString(CultureInfo.InvariantCulture) & "' " & _
                                 IIf(CouponID = "", "", " and ExternalID='" & CouponID & "' ") & " " & _
                                 "and Deleted=0  ) AS SVHistoryResult WHERE RowId >" & ((PageNum - 1) * linesPerPage) & " and RowId<=" & (PageNum * linesPerPage) & " " & OrderText

            'this is used to fetch current set of records to display
            dt = MyCommon.LXS_Select

            If (ShowExpired) Then
                Send("<br />")
                Send("<h2>" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & "</h2>")
            End If
            Value = MyCommon.NZ(dt.Rows(0).Item("Value"), 0)
            MyCommon.QueryStr = "select SVP.SVTypeID, SVP.SVExpireType, SVT.PhraseID, SVT.ValuePrecision, SVP.FuelPartner, SVP.AllowAdjustments " & _
                                "from StoredValuePrograms as SVP with (NoLock) " & _
                                "inner join SVTypes as SVT with (NoLock) on SVT.SVTypeID=SVP.SVTypeID " & _
                                "where SVProgramID=" & ProgramID & ";"
            dt2 = MyCommon.LRT_Select
            Send("        <table summary=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """>")
            Send("         <thead>")
            Send("          <tr>")
            Send("            <th style=""width:30px;"">&nbsp;</th>")
            Send("            <th scope=""col"">" & Copient.PhraseLib.Lookup("term.revoke", LanguageID) & "</th>")
            Sendb("            <th scope=""col""><a id=""extidlink"" href=""sv-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SortText=ExternalID&amp;SortDirection=" & SortDirection & """> " & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "</a>")
            If SortText = "ExternalID" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            Else
            End If
            Send("</th>")
            Send("            <th scope=""col"">" & Copient.PhraseLib.Lookup("term.status", LanguageID) & "</th>")
            Send("            <th scope=""col"">" & Copient.PhraseLib.Lookup("term.adjusted", LanguageID) & "</th>")
            Sendb("            <th scope=""col""><a id=""earneddatelink"" href=""sv-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SortText=EarnedDate&amp;SortDirection=" & SortDirection & """> " & Copient.PhraseLib.Lookup("term.earneddate", LanguageID) & "</a>")
            If SortText = "EarnedDate" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            Else
            End If
            Send("</th>")
            Send("            <th scope=""col""><a id=""expdatelink"" href=""sv-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SortText=ExpireDate&amp;SortDirection=" & SortDirection & """> " & Copient.PhraseLib.Lookup("storedvalue.expired-date", LanguageID) & "</a>")
            If SortText = "ExpireDate" Then
                If SortDirection = "ASC" Then
                    Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                    Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
            Else
            End If
            Sendb("           </th>")
            Send("            <th class="""" scope=""col"" style=""text-align:center;width:20px !important;"" title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """>" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</th>")
            Send("            <th scope=""col"" style=""text-align:right"">" & Copient.PhraseLib.Lookup("term.AmountEarned", LanguageID) & "</th>")
            Send("            <th scope=""col"" style=""text-align:right"">" & Copient.PhraseLib.Lookup("term.AmountUsed", LanguageID) & "</th>")
            Send("          </tr>")
            Send("         </thead>")
            Send("         <tbody>")
            j = 0
            For Each row In dt.Rows
                j += 1
                If (PrevLocalID <> MyCommon.NZ(row.Item("LocalID"), 0)) Then
                    Shaded = IIf(Shaded = " class=""shaded""", "", " class=""shaded""")
                End If
                Send("           <tr id=""hist" & j & """" & Shaded & ">")
                Send("             <td><input class=""ex more"" type=""button"" value=""+"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """ onclick=""javascript:showDetail(" & j & ", this);"" /></td>")
                If (Not ShowExpired AndAlso MyCommon.NZ(row.Item("QtyAvail"), 0) > 0 AndAlso MyCommon.NZ(row.Item("Status"), "") = "Earned") Then
                    If dt2.Rows(0).Item("SVTypeID") = 1 AndAlso dt2.Rows(0).Item("SVExpireType") = 5 Then
                        Sendb("             <td>")
                        Sendb("<input class=""ex more"" type=""button"" value=""+"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """ onclick=""javascript:showDetail(" & j & ", this);"" />")
                        Send("</td>")
                    Else
                        Sendb("             <td>")
                        If (EnableFuelPartner And MyCommon.NZ(dt2.Rows(0).Item("FuelPartner"), False)) Then
                            If Not MyCommon.NZ(dt2.Rows(0).Item("AllowAdjustments"), True) Then
                                DisableAdjust = True
                            End If
                        End If
                        If DisableAdjust Then
                            Sendb("<input class=""ex"" type=""button"" disabled=""disabled"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""revokeByID(" & MyCommon.NZ(row.Item("LocalID"), -1) & ", '" & MyCommon.NZ(row.Item("ExternalID"), -1) & "');"" />")
                        Else
                            Sendb("<input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""revokeByID(" & MyCommon.NZ(row.Item("LocalID"), -1) & ", '" & IIf(IsDBNull(row.Item("ExternalID")), -1, row.Item("ExternalID").ToString()) & "');"" />")
                        End If
                        Send("</td>")
                    End If
                Else
                    Send("             <td><input class=""ex more"" type=""button"" value=""+"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """ onclick=""javascript:showDetail(" & j & ", this);"" /></td>")
                End If
                Send("             <td>" & row.Item("ExternalID").ToString() & "</td>")
                If (ShowExpired) Then
                    Send("             <td>" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & "</td>")
                Else
                    Send("             <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("StatusPhraseID"), ""), LanguageID) & "</td>")
                End If
                If (MyCommon.NZ(row.Item("LastLocationID"), 0) = -9) Then
                    Sendb("             <td>" & Copient.PhraseLib.Lookup("term.yes", LanguageID))
                    If (MyCommon.NZ(row.Item("AdminUserID"), -1) > 0) Then
                        Sendb("<br />(<i>" & GetUserName(row.Item("AdminUserID")) & "</i>)")
                    End If
                    Send("</td>")
                ElseIf (MyCommon.NZ(row.Item("LastLocationID"), 0) = -99) Then
                    Send("             <td>" & Copient.PhraseLib.Lookup("term.autoofferdist", LanguageID) & "</td>")
                Else
                    Send("             <td>" & Copient.PhraseLib.Lookup("term.no", LanguageID) & "</td>")
                End If
                If MyCommon.NZ(row.Item("EarnedDate"), "1/1/1980") = "1/1/1980" Then
                    Send("             <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                Else
                    Send("             <td>" & Logix.ToShortDateTimeString(row.Item("EarnedDate"), MyCommon) & "</td>")
                End If
                If MyCommon.NZ(row.Item("ExpireDate"), "1/1/1980") = "1/1/1980" Then
                    Send("             <td>&nbsp;</td>")
                Else
                    Send("             <td>" & Logix.ToShortDateTimeString(row.Item("ExpireDate"), MyCommon) & "</td>")
                End If
                Send("             <td style=""text-align:center;"">" & IIf(MyCommon.NZ(row.Item("Replayed"), False), "<span title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """ style=""color:#dd0000;cursor:default;font-size:10px;font-weight:bold;"">" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</span>", "") & "</td>")
                If dt2.Rows(0).Item("SVTypeID") > 1 Then
                    ValueString = Math.Round(MyCommon.NZ(row.Item("QtyEarned"), 0) * Value, Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
                    If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                        ValueString = Left(ValueString, Len(ValueString) - 1)
                    End If
                    Sendb("             <td style=""text-align:right"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
                    ValueString = Math.Round(MyCommon.NZ(row.Item("QtyUsed"), 0) * Value, Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
                    If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                        ValueString = Left(ValueString, Len(ValueString) - 1)
                    End If
                    Sendb("             <td style=""text-align:right"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
                Else
                    decTemp = (Int(MyCommon.NZ(row.Item("QtyEarned"), 0) * Value) * 1.0) / decFactor
                    Sendb("             <td style=""text-align:right"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
                    decTemp = (Int(MyCommon.NZ(row.Item("QtyUsed"), 0) * Value) * 1.0) / decFactor
                    Sendb("             <td style=""text-align:right"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
                End If
                Send("           </tr>")
                If IsDBNull(row.Item("PresentedCustomerID")) Then
                    PresentedCustomerID = "Unknown"
                Else
                    PresentedCustomerID = If(row.Item("PresentedCustomerID") = "0", "0", MyCryptLib.SQL_StringDecrypt(row.Item("PresentedCustomerID").ToString()))
                End If
                If IsDBNull(row.Item("ResolvedCustomerID")) Then
                    ResolvedCustomerID = "Unknown"
                Else
                    ResolvedCustomerID = If(row.Item("ResolvedCustomerID") = "0", "0", MyCryptLib.SQL_StringDecrypt(row.Item("ResolvedCustomerID").ToString()))
                End If
                If (ResolvedCustomerID = "0" OrElse ResolvedCustomerID = "Unknown") Then
                    MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK=" & MyCommon.NZ(row.Item("CustomerPK"), 0) & ";"
                    dt3 = MyCommon.LXS_Select
                    If dt3.Rows.Count > 0 Then
                        Dim tmpExtCID As String = MyCryptLib.SQL_StringDecrypt(dt3.Rows(0).Item("ExtCardID").ToString())
                        ResolvedCustomerID = IIf(tmpExtCID = "", Copient.PhraseLib.Lookup("term.unknown", LanguageID), tmpExtCID)
                    End If
                End If
                Send("           <tr id=""histdetail" & j & """" & Shaded & " style=""display:none;color:#777777;"">")
                Send("             <td></td>")
                Send("             <td colspan=""9"">")
                Send("               " & Copient.PhraseLib.Lookup("term.presented", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & PresentedCustomerID & " &nbsp;|&nbsp; ")
                Send("               " & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & ResolvedCustomerID & " &nbsp;|&nbsp; ")
                Send("               " & Copient.PhraseLib.Lookup("term.household", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & IIf(IsDBNull(row.Item("HHID")), "Unknown", row.Item("HHID").ToString()))
                Send("             </td>")
                Send("           </tr>")
                TotalEarned += MyCommon.NZ(row.Item("QtyEarned"), 0)
                TotalUsed += MyCommon.NZ(row.Item("QtyUsed"), 0)
                PrevLocalID = MyCommon.NZ(row.Item("LocalID"), 0)
            Next
            Send("           <tr>")
            Send("             <td colspan=""8""></td>")
            If dt2.Rows(0).Item("SVTypeID") > 1 Then
                ValueString = Math.Round(TotalEarned * Value, Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
                If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                    ValueString = Left(ValueString, Len(ValueString) - 1)
                End If
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
                ValueString = Math.Round(TotalUsed * Value, Localization.Get_Default_Currency_Precision()) & " " & Localization.Get_Default_Currency_Symbol()
                If Right(ValueString, 1) = "0" AndAlso (MyCommon.NZ(dt2.Rows(0).Item("ValuePrecision"), 0) = 2) Then
                    ValueString = Left(ValueString, Len(ValueString) - 1)
                End If
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & ValueString.Replace(".", MyCommon.GetAdminUser.Culture.NumberFormat.CurrencyDecimalSeparator) & "</td>")
            Else
                decTemp = (Int(TotalEarned * Value) * 1.0) / decFactor
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
                decTemp = (Int(TotalUsed * Value) * 1.0) / decFactor
                Send("             <td style=""border-top: solid 1px #888888;text-align:right;"">" & FormatNumber(decTemp, intNumDecimalPlaces) & "</td>")
            End If
            Send("          </tr>")
            If PageCount > 1 Then
                'this row added for navigation directions
                Send("           <tr>")
                Send("             <td colspan=""9"">")
                If PageNum <= 1 Then
                    Send("   <span id=""previous"">◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</span>")
                    Send("   <span id=""next""><a id=""nextPageLink"" href=""sv-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;pagenum= " & PageNum+1 & " &amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SortText=ExternalID&amp;historyTo=" & Request.QueryString("historyTo") & "&amp;historyFrom=" & Request.QueryString("historyFrom") & "&amp;SortDirection=" & SortDirection & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a></span>&nbsp;")
                ElseIf PageNum = PageCount Then
                    Send("   <span id=""previous""><a id=""previousPageLink"" href=""sv-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;pagenum= " & PageNum-1 & " &amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SortText=ExternalID&amp;historyTo=" & Request.QueryString("historyTo") & "&amp;historyFrom=" & Request.QueryString("historyFrom") & "&amp;SortDirection=" & SortDirection & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a></span>")
                    Send("   <span id=""next"">" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</span>&nbsp;")
                Else
                    Send("   <span id=""previous""><a id=""previousPageLink"" href=""sv-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;pagenum= " & PageNum-1 & " &amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SortText=ExternalID&amp;historyTo=" & Request.QueryString("historyTo") & "&amp;historyFrom=" & Request.QueryString("historyFrom") & "&amp;SortDirection=" & SortDirection & """>◄" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & "</a></span>")
                    Send("   <span id=""next""><a id=""nextPageLink"" href=""sv-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;pagenum= " & PageNum+1 & " &amp;CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SortText=ExternalID&amp;historyTo=" & Request.QueryString("historyTo") & "&amp;historyFrom=" & Request.QueryString("historyFrom") & "&amp;SortDirection=" & SortDirection & """>" & Copient.PhraseLib.Lookup("term.next", LanguageID) & "►</a></span>&nbsp;")
                End If    'closing of If PageNum < = 1 Then
                Send("             </td>")
                Send("          </tr>")
            End If     'closing of If PageCount > 1 Then
            Send("         </tbody>")
            Send("       </table>")
        End If
    End Sub

    Function GetUserName(ByVal AdminUserID As Integer) As String
        Dim UserName As String = ""
        Dim rows() As DataRow = Nothing

        If (dtUsers Is Nothing) Then
            MyCommon.QueryStr = "select AdminUserID, UserName from AdminUsers with (NoLock) order by AdminUserID;"
            dtUsers = MyCommon.LRT_Select()
        End If
        If (AdminUserID > 0) Then
            rows = dtUsers.Select("AdminUserID =" & AdminUserID)
            If (rows.Length > 0) Then
                UserName = MyCommon.NZ(rows(0).Item("UserName"), "")
            End If
        End If

        Return UserName
    End Function
</script>

<%
done:
  If (HHPK > 0 OrElse Not IsExpired OrElse CustomerPK = 0) Then
    Send_BodyEnd()
  Else
    Send_BodyEnd("frmSvAdj", "adjust")
  End If
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
