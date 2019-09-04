<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: point-adjust.aspx 
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
  
  Dim MyCommon As New Copient.CommonInc
  Dim AdminUserID As Long
  Dim Logix As New Copient.LogixInc
  Dim ExcentusRetailerID As String = ""
  Dim ExcentusSiteID As String = ""
  Dim ExcentusPassword As String = ""
  Dim ExcentusURL As String = ""
  Dim dt As DataTable = Nothing
  Dim rst2 As DataTable
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim HHCardPK As Long = 0
  Dim OfferID As Long
  
  Dim OfferExpired As Boolean = False
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  
  Dim PageTitle As String = ""
  Dim OfferName As String = ""
  Dim ProgramID As Long
  Dim ProgramName As String = ""
  Dim AdjustPermitted As Boolean = False
  Dim EarnedROID As Integer = 0
  Dim EarnedCMOffer As Integer = 0
  Dim OfferDesc As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HHPK As Integer = 0
  Dim i As Integer = 0
  Dim KeyCt As Integer = 0
  Dim RefreshParent As Boolean = False
  Dim SessionID As String = ""
  Dim ProgramIDs(), PointsAdjs() As String
  Dim ValidAdj As Boolean = False
  Dim ValidAdjCount As Integer = 0
  Dim Note As String = ""
  Dim FirstName As String = ""
  Dim LastName As String = ""
  Dim IsUSAirMiles As Boolean = False
  Dim IsHousehold As Boolean = False
  Dim HHMembershipCount As Integer = 0
  Dim ExternalProgram As Boolean = False
  Dim ExtHostTypeID As Integer = 0
  Dim DecimalValues As Boolean = False
  Dim ER As Copient.ExternalRewards
  Dim ReasonID As Integer = 0
  Dim ReasonText As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "point-adjust.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If MyCommon.Fetch_SystemOption(80) = "1" Then
    Try
      Dim sDb2Connection As String
      Dim iDb2Connection As Integer = 5
            
      sDb2Connection = MyCommon.Fetch_CM_SystemOption(iDb2Connection)
      ER = New Copient.ExternalRewards("", "", "", sDb2Connection)
    Catch
      Throw
    End Try
  Else
    ExcentusRetailerID = MyCommon.Fetch_InterfaceOption(39)
    ExcentusSiteID = MyCommon.Fetch_InterfaceOption(40)
    ExcentusPassword = MyCommon.Fetch_InterfaceOption(41)
    ExcentusURL = MyCommon.Fetch_InterfaceOption(42)
    ER = New Copient.ExternalRewards(ExcentusRetailerID, ExcentusSiteID, ExcentusPassword, ExcentusURL)
  End If
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If
  
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
  If (OfferID > 0) Then
    MyCommon.QueryStr = "select IncentiveName as Name, OID.EngineID, OID.EngineSubTypeID from CPE_Incentives as I with (NoLock) " & _
                        "inner join OfferIDs as OID on OID.OfferID=I.IncentiveID " & _
                        "where I.IncentiveID=" & OfferID & " " & _
                        " union " & _
                        "select Name, OID.EngineID, OID.EngineSubTypeID from Offers as O with (NoLock) " & _
                        "inner join OfferIDs as OID on OID.OfferID=O.OfferID " & _
                        "where O.OfferID=" & OfferID & ";"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      OfferName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
      If (MyCommon.NZ(dt.Rows(0).Item("EngineID"), -1) = 2) AndAlso (MyCommon.NZ(dt.Rows(0).Item("EngineSubTypeID"), -1) = 2) Then
        IsUSAirMiles = True
      End If
    End If
  End If
  
  If (Logix.UserRoles.EditPointsBalances) AndAlso (IsUSAirMiles = False) Then
    AdjustPermitted = True
  ElseIf (Logix.UserRoles.EditAirmilesPointsBalances) AndAlso (IsUSAirMiles = True) Then
    AdjustPermitted = True
  End If
  
  MyCommon.QueryStr = "select HHPK, CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
  dt = MyCommon.LXS_Select()
  If (dt.Rows.Count > 0) Then
    HHPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
    If MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0) = 1 Then
      IsHousehold = True
    End If
  End If
  If HHPK > 0 Then
    infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.hh-adjust-note", LanguageID)
    'Get the CardPK of the Household card if there is a household PK
    MyCommon.QueryStr = "select CardPK from CardIDs where CustomerPK=" & HHPK
    dt = MyCommon.LXS_Select()
    If (dt.Rows.Count > 0) Then
      HHCardPK = MyCommon.NZ(dt.Rows(0).Item("CardPK"), 0)
    End If
  End If
  
  If IsHousehold Then
    MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where HHPK=" & CustomerPK & ";"
    dt = MyCommon.LXS_Select
    HHMembershipCount = dt.Rows.Count
  End If
  
  If (OfferID > 0) Then
    PageTitle = OfferName
    
    ' Grab the offer description from the appropriate table
    MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID = " & OfferID
    rst2 = MyCommon.LRT_Select
    If (MyCommon.NZ(rst2.Rows(0).Item("EngineID"), 0) <> 2) Then
      MyCommon.QueryStr = "select Description from Offers with (NoLock) where OfferID = " & OfferID
      rst2 = MyCommon.LRT_Select
    Else
      MyCommon.QueryStr = "select Description from CPE_Incentives with (NoLock) where IncentiveID = " & OfferID
      rst2 = MyCommon.LRT_Select
    End If
    If (rst2.Rows.Count > 0) Then
      OfferDesc = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
    End If
    
    ' Check if offer is CM or CPE
    MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where OfferID = " & OfferID & " and Deleted=0;"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      EarnedCMOffer = OfferID
      EarnedROID = 0
    Else
      MyCommon.QueryStr = "select RewardOptionID from CPE_Incentives I with (NoLock) " & _
                          "inner join CPE_RewardOptions RO with (NoLock) on RO.IncentiveID = I.IncentiveID " & _
                          "where RO.Deleted=0 and RO.TouchResponse=0 and I.Deleted=0 and I.IncentiveID = " & OfferID & ";"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        EarnedROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
        EarnedCMOffer = 0
      End If
    End If
    
    ' Get offer status
    StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
  End If
  
  If (Request.QueryString("save") <> "") Then
    i = 0
    ProgramIDs = Request.QueryString.GetValues("programID")
    PointsAdjs = Request.QueryString.GetValues("adjust")
    If Request.QueryString("reasonID") <> "" Then
      ReasonID = MyCommon.Extract_Val(Request.QueryString("reasonID"))
    End If
    If Request.QueryString("reasonText") <> "" Then
      ReasonText = Request.QueryString("reasonText")
    End If
    
    If Not (ProgramIDs Is Nothing Or PointsAdjs Is Nothing) Then
      For i = ProgramIDs.GetLowerBound(0) To ProgramIDs.GetUpperBound(0)
        If PointsAdjs(i) <> "" Then
          MyCommon.QueryStr = "select ProgramID, ExternalProgram, ExtHostTypeID, DecimalValues from PointsPrograms with (NoLock) where ProgramID=" & MyCommon.Extract_Val(ProgramIDs(i)) & ";"
          dt = MyCommon.LRT_Select
          If dt.Rows.Count > 0 Then
            ExternalProgram = IIf(MyCommon.NZ(dt.Rows(0).Item("ExternalProgram"), False), True, False)
            ExtHostTypeID = MyCommon.NZ(dt.Rows(0).Item("ExtHostTypeID"), 0)
            DecimalValues = IIf(MyCommon.NZ(dt.Rows(0).Item("DecimalValues"), False), True, False)
          End If
          If ExternalProgram AndAlso ExtHostTypeID = 2 Then
            ValidAdj = IsValidAdjustment(OfferID, CustomerPK, infoMessage, OfferID, MyCommon.Extract_Val(ProgramIDs(i)), (PointsAdjs(i) * 100), ExternalProgram, ExtHostTypeID, DecimalValues, ReasonID, ReasonText)
          Else
            ValidAdj = IsValidAdjustment(OfferID, CustomerPK, infoMessage, OfferID, MyCommon.Extract_Val(ProgramIDs(i)), MyCommon.Extract_Val(PointsAdjs(i)), ExternalProgram, ExtHostTypeID, DecimalValues, ReasonID, ReasonText)
          End If
        Else
          ValidAdj = True
        End If
        If ValidAdj Then
          ValidAdjCount += 1
	  If (IsUSAirMiles = True) Then ProcessIssuance(OfferID, CustomerPK, MyCommon.Extract_Val(ProgramIDs(i)), MyCommon.Extract_Val(PointsAdjs(i)), infoMessage)
        End If
        If Not ValidAdj Then
          Exit For
        End If
      Next
      
      If ValidAdjCount > 0 Then
        Note = Request.QueryString("note")
        Note = MyCommon.Parse_Quotes(Note)
        Note = Logix.TrimAll(Note)
        If MyCommon.Fetch_SystemOption(99) Then ' 99 = 'Require note when adjusting balance'
          
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
            infoMessage = Copient.PhraseLib.Lookup("sv-adjust-program.NoteRequired", LanguageID)
          End If

          End If

      End If ' ValidAdjCount > 0 Then

      If infoMessage = "" Then
        AdjustPoint(AdminUserID, SessionID, OfferID, ER, ReasonID, ReasonText)
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "point-adjust.aspx?CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & _
                           "&OfferID=" & OfferID & _
                           "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                           "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")))
        GoTo done
      End If
    End If
    ' NB: The above line was originally AdjustPoint(AdminUserID, EarnedROID, EarnedCMOffer).
    ' I hard-coded these values to zero since that is what needs to be recorded in order to
    ' indicate that the points were awarded via a manual adjustment (and *not* via a
    ' particular offer or reward). --Huw
    
  ElseIf (Request.QueryString("HistoryEnabled.x") <> "" OrElse Request.QueryString("HistoryDisabled.x") <> "") Then
    ' Write a cookie and then reload the page
    Response.Cookies("HistoryEnabled").Expires = "10/08/2100"
    Response.Cookies("HistoryEnabled").Value = IIf(Request.QueryString("HistoryEnabled.x") <> "", "1", "0")
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "point-adjust.aspx?CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & _
                       "&OfferID=" & OfferID & _
                       "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                       "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")))
    GoTo done
  End If
  
  'If (Request.QueryString("RefreshParent") = "true") Then RefreshParent = True
  
  Send_HeadBegin("term.offer", "term.pointsadjustment", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts(New String() {"datePicker.js"})
%>
<script type="text/javascript">
  var datePickerDivID = "datepicker";
  var linkToHH = false;
  var IsSaveButtonClicked = 1;

  function OtherButtonsClicked() {
    IsSaveButtonClicked = 0;
}

  <% Send_Calendar_Overrides(MyCommon) %>
  
  function isValidEntry() {

  if (IsSaveButtonClicked == 1) {
    

    var retVal = true;
    var elems = document.getElementsByName("adjust");
    var elem = document.getElementById("adjust");
    var maxadjusts = document.getElementsByName("maxadjust");
    var maxadjust = document.getElementById("maxadjust");
    var externalprograms = document.getElementsByName("externalprogram");
    var externalprogram = document.getElementById("externalprogram");
    var exthosttypeids = document.getElementsByName("exthosttypeid");
    var exthosttypeid = document.getElementById("exthosttypeid");
    
    // Check the validity of the value
    if (elems != null) {
      for (var i=0; i < elems.length; i++) {
        if (elems[i].value == "") {
          retVal = false;
          if (externalprogram.value == 1 && exthosttypeid.value == 2) {
            alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerrordecimal", LanguageID)) %>');
          } else {
            alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerror", LanguageID)) %>');
          }
          elems[i].focus();
          elems[i].select();
          break;
        } else {
          if (externalprograms[i].value == 1 && exthosttypeids[i].value == 2) {
            if (elems[i].value != "" && (isNaN(elems[i].value))) {
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerrordecimal", LanguageID)) %>');
              elems[i].focus();
              elems[i].select();
              break;
            }
          } else {
            if (elems[i].value != "" && (isNaN(elems[i].value) || (isInt(elems[i].value) == false && exthosttypeids[i].value != "2"))) {
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerror", LanguageID)) %>');
              elems[i].focus();
              elems[i].select();
              break;
            }
          }
        }
      }
    } else {
      if (isNaN(elem.value) || (isInt(elem.value) == false && exthosttypeid.value != "2")) {
        retVal = false;
        alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerror", LanguageID)) %>');
        elem.focus();
      }
    }
    // Check that the value doesn't exceed a MaxAdjustment
    if (retVal) {
      if (maxadjusts != null) {
        for (var i=0; i < maxadjusts.length; i++) {
          if (maxadjusts[i].value != '') {
            if (((parseInt(elems[i].value) > 0) && (parseInt(elems[i].value) > parseInt(maxadjusts[i].value))) || ((parseInt(elems[i].value) < 0) && (parseInt(elems[i].value) < -parseInt(maxadjusts[i].value)))) {
              if (confirm('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.maxadjustexceeded", LanguageID)) %>')) {
              } else {
                retVal = false;
                elems[i].focus();
                elems[i].select();
              }
              break;
            }
          }
        }
      } else {
        if (maxadjust.value != '') {
          if (((parseInt(elem.value) > 0) && (parseInt(elem.value) > parseInt(maxadjust.value))) || ((parseInt(elem.value) < 0) && (parseInt(elem.value) < -parseInt(maxadjust.value)))) {
            if (confirm('<% Sendb(Copient.PhraseLib.Lookup("points-adjust.MaxAdjustmentExceeded", LanguageID)) %>')) {
            } else {
              retVal = false;
              elem.focus();
            }
          }
        }
      }
    }
    return retVal;
    }
  }
  
  function ChangeParentDocument() {
    var refreshElem = document.getElementById("RefreshParent");
    
    <% If HHPK > 0 Then  %>
    if (opener != null && !opener.closed) {
      if (refreshElem != null && refreshElem.value == 'true') {
        if (linkToHH) {
          opener.location = 'customer-offers.aspx?CustPK=<%Sendb(HHPK)%><%Sendb(IIf(HHCardPK > 0, "&CardPK=" & HHCardPK, ""))%>';
        } else {
          opener.location = 'customer-offers.aspx?CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>';
        }
      }
    }
    <% End If %>
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
  
  If (Logix.UserRoles.AccessPointsBalances = False) AndAlso (IsUSAirMiles = False) Then
    Send_Denied(2, "perm.customers-ptbalaccess")
    GoTo done
  ElseIf (Logix.UserRoles.AccessAirmilesPointsBalances = False) AndAlso (IsUSAirMiles = True) Then
    Send_Denied(2, "perm.customers-ptbalaccess-airmiles")
    GoTo done
  End If
%>
<form id="mainform" name="mainform" action="" onsubmit="return isValidEntry();">
  <input type="hidden" id="OfferId" name="OfferId" value="<% Sendb(OfferID)%>" />
  <input type="hidden" id="OfferName" name="OfferName" value="<% Sendb(OfferName)%>" />
  <input type="hidden" id="CustomerPK" name="CustomerPK" value="<% Sendb(CustomerPK)%>" />
  <input type="hidden" id="CardPK" name="CardPK" value="<% Sendb(CardPK)%>" />
  <input type="hidden" id="RefreshParent" name="RefreshParent" value="<% Sendb(RefreshParent.ToString.ToLower) %>" />
  
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(PageTitle, 40))%>
    </h1>
    <div id="controls">
      <%
        If (AdjustPermitted) Then
          Send_Save()
        End If
      %>
    </div>
  </div>
  <div id="main" style="width: 100%;">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column">
      <%
        If (OfferDesc <> "") Then
          Send("<p id=""description"">" & MyCommon.SplitNonSpacedString(OfferDesc, 50) & "</p>")
        End If
        Send("<p id=""status"">" & Copient.PhraseLib.Lookup("term.status", LanguageID) & ": " & StatusText & "</p>")
        If HHPK > 0 Then
          Sendb("<br /><a href=""point-adjust.aspx?")
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
          Send(Copient.PhraseLib.Lookup("customer-inquiry.hh-adjust-linktext", LanguageID) & "</a><br/><br/>")
        End If
        
        Send("<div class=""box"" id=""ptAdj""" & IIf((AdjustPermitted = False) OrElse ((StatusCode = 5) AND (IsUSAirMiles = False)) OrElse (HHPK > 0), " style=""display:none;""", "") & ">")
        Send("  <h2>")
        Send("    <span>")
        Send("      " & Copient.PhraseLib.Lookup("term.pointsadjustment", LanguageID))
        Send("    </span>")
        Send("  </h2>")
        If (OfferID > 0) Then
          ShowPoints(CustomerPK, OfferID, AdjustPermitted, ER, Logix)
        Else
          Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
        End If
        Send("  <hr class=""hidden"" />")
        Send("</div>")
        Send("")
        
        If (MyCommon.Fetch_SystemOption(108) = "1") Then
          'Require a reason
          Send("<div class=""box"" id=""reason""" & IIf((AdjustPermitted = False) OrElse ((StatusCode = 5) AND (IsUSAirMiles = False)) OrElse (HHPK > 0), " style=""display:none;""", "") & ">")
          Send("  <h2>")
          Send("    <span>")
          Send("      " & Copient.PhraseLib.Lookup("term.reason", LanguageID))
          Send("    </span>")
          Send("  </h2>")
          Send("  <select id=""reasonID"" name=""reasonID"" style=""width:220px;"">")
          Send("    <option value=""0"">(" & Copient.PhraseLib.Lookup("term.SelectAReason", LanguageID) & ")</option>")
          Dim RC As DataTable
          Dim row As DataRow
          MyCommon.QueryStr = "select ReasonID, Description from AdjustmentReasons with (NoLock) where Enabled=1;"
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
        
        Send("<div class=""box"" id=""pending""" & IIf(NonExtProgramIdList.Count = 0, " style=""display:none;""", "") & ">")
        Send("  <h2>")
        Send("    <span>")
        Send("      " & Copient.PhraseLib.Lookup("term.pending", LanguageID))
        Send("    </span>")
        Send("  </h2>")
        Sendb("  <span style=""float:right;font-size:9px;position:relative;top:-22px;"">")
        Sendb("<a href=""point-adjust.aspx?OfferID=" & OfferID & "&amp;OfferName=" & Server.UrlEncode(OfferName) & "&amp;CustomerPK=" & CustomerPK & "&amp;historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & "&amp;historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")) & """>" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "</a>")
        Send("</span>")
        If (OfferID > 0) Then
          ShowPendingOrHistory(CustomerPK, EarnedROID, EarnedCMOffer, False, IsHousehold, HHMembershipCount, Logix)
        Else
          Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
        End If
        Send("  <hr class=""hidden"" />")
        Send("</div>")
        Send("")
        
        Send("<div class=""box"" id=""history""" & IIf(NonExtProgramIdList.Count = 0, " style=""display:none;""", "") & ">")
        Send("  <h2>")
        Send("    <span>")
        Send("      " & Copient.PhraseLib.Lookup("term.history", LanguageID))
        Send("    </span>")
        Send("  </h2>")
        If (OfferID > 0) Then
          ShowPendingOrHistory(CustomerPK, EarnedROID, EarnedCMOffer, True, IsHousehold, HHMembershipCount, Logix)
        Else
          Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
        End If
        Send("  <hr class=""hidden"" />")
        Send("</div>")
        Send("")
      %>
    </div>
  </div>
</form>

<script runat="server">
  Dim MyCommon As New Copient.CommonInc
  Dim ProgramIdList As New ArrayList
    Dim NonExtProgramIdList As New ArrayList
    Dim MyCryptLib As New Copient.CryptLib
  
  Sub ShowPoints(ByVal CustomerPK As Long, ByVal OfferID As String, ByVal AdjustPermitted As Boolean, ByVal ER As Copient.ExternalRewards, _
                 ByRef Logix As Copient.LogixInc)
    Dim UpdateAccum As Boolean = False
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim dtPrograms As DataTable
    Dim dtPoints As DataTable
    Dim rowProgram As DataRow = Nothing
    Dim PromoVarID As Integer
    Dim ProgramID As Integer
    Dim ExternalProgram As Boolean = False
    Dim ExtHostTypeID As Integer = 0
    Dim DecimalValues As Boolean = False
    Dim Points As Integer = 0
    Dim PointsAsDecimal As Decimal = 0
    Dim PendingAdj As Integer = 0
    Dim MaxAdjustment As Integer = 0
    Dim i As Integer = 1
    Dim ExtCardID As String = ""
    Dim ErrorThrown As Boolean = False
    Dim RewardStatus As Copient.ExternalRewards.RewardStatusInfo
    Dim ExpirationAmount As New Decimal(0)
    Dim DaysTillExpiration As Integer = 0
    Dim bEme As Boolean = False
    Dim dtExtBalances As DataTable = Nothing
    Dim ExternalID As String
    Dim rows() As DataRow
    Dim sErrorMsg As String = ""
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    MyCommon.SetAdminUser(Logix.Fetch_AdminUser(MyCommon, AdminUserID))
    
    MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by CardTypeID;"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
            ExtCardID = IIf(IsDBNull(dt.Rows(0).Item("ExtCardID")), "0", MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString()))
    End If
    
    If MyCommon.Fetch_SystemOption(80) = "1" Then
      bEme = True
      Try
        dtExtBalances = ER.getExternalBalances(ExtCardID)
      Catch ex As Exception
        sErrorMsg = "EME: " & ex.Message
        bEme = False
        AdjustPermitted = False
      End Try
    End If

    MyCommon.QueryStr = "dbo.pa_OfferPointsPrograms"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
    dtPrograms = MyCommon.LRTsp_select
    MyCommon.Close_LRTsp()

    If (sErrorMsg <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & sErrorMsg & "</div>")
    End If

    Send("      <br class=""half"" />")
    Send("      " & Copient.PhraseLib.Lookup("term.note", LanguageID) & ": " & Copient.PhraseLib.Lookup("point-adjust.offerid", LanguageID) & " #" & OfferID & ".<br />")
    For Each rowProgram In dtPrograms.Rows
      If MyCommon.NZ(rowProgram.Item("ExternalProgram"), False) AndAlso MyCommon.NZ(rowProgram.Item("ExtHostTypeID"), 0) = 2 Then
        Send("      <span style=""color:#cc0000;"">" & Copient.PhraseLib.Lookup("point-adjust.programnoteext", LanguageID) & "</span><br />")
      End If
    Next
    Send("      <br class=""half"" />")
    Send("        <table summary=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """>")
    Send("         <thead>")
    Send("          <tr>")
    Send("            <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.program", LanguageID) & "</th>")
    Send("            <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
    Send("            <th class=""th-var"" scope=""col"">" & Copient.PhraseLib.Lookup("term.promovarid", LanguageID) & "</th>")
    Send("            <th class=""th-quantity"" scope=""col"" style=""text-align:center"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
    Send("            <th class=""th-pending"" scope=""col"" style=""text-align:center"">" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & "</th>")
    Send("            <th class=""th-adjustment"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & "</th>")
    Send("          </tr>")
    Send("         </thead>")
    Send("         <tbody>")
    
    For Each rowProgram In dtPrograms.Rows
      PromoVarID = MyCommon.NZ(rowProgram.Item("PromoVarID"), 0)
      ProgramID = MyCommon.NZ(rowProgram.Item("ProgramID"), 0)
      ExternalProgram = MyCommon.NZ(rowProgram.Item("ExternalProgram"), False)
      ExtHostTypeID = MyCommon.NZ(rowProgram.Item("ExtHostTypeID"), 0)
      DecimalValues = MyCommon.NZ(rowProgram.Item("DecimalValues"), False)
      ErrorThrown = False
      PointsAsDecimal = 0
      ExpirationAmount = 0
      DaysTillExpiration = 0
      
      ProgramIdList.Add(ProgramID)
      If ExtHostTypeID = 0 Or ExtHostTypeID = 1 Then
        NonExtProgramIdList.Add(ProgramID)
      End If
      
      If ExternalProgram AndAlso ExtHostTypeID = 2 Then
        'use the ExternalRewards DLL
        Try
          RewardStatus = ER.getUserRewardStatus(ExtCardID, 0)
          PointsAsDecimal = RewardStatus.totalRewardAmount
          ExpirationAmount = RewardStatus.rewardExpirationAmount
          DaysTillExpiration = RewardStatus.rewardExpirationDays
        Catch ex As Copient.ExternalRewards.RewardsException
          ErrorThrown = True
        End Try
      ElseIf ExternalProgram AndAlso ExtHostTypeID = 1 Then
        'use the ExternalRewards DLL (EME)
        Points = 0
        If bEme Then
          MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where Deleted=0 and VarTypeID=3 and LinkID=" & ProgramID & ";"
          dt = MyCommon.LXS_Select()
          If dt.Rows.Count > 0 Then
            ExternalID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExternalID"))
            If dtExtBalances.Rows.Count > 0 Then
              rows = dtExtBalances.Select("ExternalID='" & ExternalID & "'")
              If (rows.Length > 0) Then
                Points = MyCommon.NZ(rows(0).Item("Balance"), 0)
              End If
            End If
          End If
        End If
      Else
        MyCommon.QueryStr = "select PKID, Amount from Points with (NoLock) " & _
                            "where CustomerPK=" & CustomerPK & " and (PromoVarID=" & PromoVarID & " or ProgramID=" & ProgramID & ");"
        dtPoints = MyCommon.LXS_Select()
        If (dtPoints.Rows.Count > 0) Then
          Points = MyCommon.NZ(dtPoints.Rows(0).Item("Amount"), 0)
        Else
          Points = 0
        End If
      End If
      
      If ExternalProgram AndAlso ExtHostTypeID = 1 Then
        PendingAdj = 0
      Else
        MyCommon.QueryStr = "select sum( Convert( int, cast( IsNull( Col3, 0 )  AS float ) ) ) as Pending from CPE_UploadTemp_PA with (NoLock) where Col1 = '" & ProgramID & "' and Col2 = '" & CustomerPK & "';"
        dt = MyCommon.LXS_Select
        If (dt.Rows.Count > 0) Then
          PendingAdj = MyCommon.NZ(dt.Rows(0).Item("Pending"), 0)
        Else
          PendingAdj = 0
        End If
      End If
      
      Send("          <tr>")
      Send("            <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rowProgram.Item("ProgramName"), "nbsp;"), 25) & "</td>")
      Send("            <td>")
      Send("              <input type=""hidden"" name=""programID"" value=""" & ProgramID & """ />" & ProgramID)
      Send("              <input type=""hidden"" id=""externalprogram"" name=""externalprogram"" value=""" & IIf(ExternalProgram, "1", "0") & """ />")
      Send("              <input type=""hidden"" id=""exthosttypeid"" name=""exthosttypeid"" value=""" & ExtHostTypeID & """ />")
      Send("            </td>")
      Send("            <td>" & PromoVarID & "</td>")
      If ExternalProgram And ExtHostTypeID = 2 Then
        If ErrorThrown Then
          Send("            <td colspan=""3"" style=""font-size:10px;color:#cc0000;"">" & Copient.PhraseLib.Lookup("points-adjust.Unresolvable", LanguageID) & "<input type=""hidden"" id=""adjust"" name=""adjust"" value=""0"" /></td>")
        Else
          Send("            <td style=""text-align:center"">$" & Format(PointsAsDecimal, "0.00") & "</td>")
          Send("            <td style=""text-align:center"">" & "-" & "</td>")
          Sendb("           <td>$<input type=""text"" class=""short"" name=""adjust"" style=""text-align:right;"" maxlength=""7"" value=""""" & IIf(AdjustPermitted = False, " disabled=""disabled""", "") & " />")
          Sendb("<input type=""hidden"" name=""maxadjust"" value="""" />")
          Send("</td>")
        End If
      Else
        Send("            <td style=""text-align:center"">" & Points & "</td>")
        Send("            <td style=""text-align:center"">" & PendingAdj & "</td>")
        Sendb("            <td><input type=""text"" class=""short"" name=""adjust"" style=""text-align:right;"" maxlength=""7"" value=""""" & IIf((ExternalProgram AndAlso ExtHostTypeID > 1) OrElse AdjustPermitted = False, " disabled=""disabled""", "") & " />")
        MyCommon.QueryStr = "select MaxAdjustment from CPE_DeliverablePoints as DP with (NoLock) " & _
                            "inner join CPE_RewardOptions as RO on RO.RewardOptionID=DP.RewardOptionID " & _
                            "where RO.IncentiveID=" & OfferID & " and DP.ProgramID=" & ProgramID & ";"
        dt2 = MyCommon.LRT_Select
        If dt2.Rows.Count > 0 Then
          If Not IsDBNull(dt2.Rows(0).Item("MaxAdjustment")) Then
            Sendb("<input type=""hidden"" name=""maxadjust"" value=""" & dt2.Rows(0).Item("MaxAdjustment") & """ />")
          Else
            Sendb("<input type=""hidden"" name=""maxadjust"" value="""" />")
          End If
        Else
          Sendb("<input type=""hidden"" name=""maxadjust"" value="""" />")
        End If
        Send("</td>")
      End If
      Send("          </tr>")
      
      If ExternalProgram AndAlso ExtHostTypeID = 2 AndAlso Not ErrorThrown AndAlso ExpirationAmount > 0 Then
        Send("          <tr>")
        Send("            <td colspan=""6"" style=""font-size:10px;color:#cc0000;padding-left:15px;"">Note: " & FormatCurrency(ExpirationAmount, 2) & " will expire in " & DaysTillExpiration & " days (" & Logix.ToShortDateString(DateAdd("d", DaysTillExpiration, Now()), MyCommon) & ").</td>")
        Send("          </tr>")
      End If
      
      If MyCommon.Fetch_SystemOption(99) Then
        Send("          <tr>")
        Send("            <td style=""vertical-align:top;""><label for=""note"">" & Copient.PhraseLib.Lookup("point-adjust.AdjusmentExplanation", LanguageID) & ":</note></td>")
        Send("            <td colspan=""5"">")
        Send("              <textarea id=""note"" name=""note"" style=""height:38px;width:448px;font-size:12px;""></textarea>")
        Send("            <td>")
        Send("          </tr>")
      End If
      i += 1
    Next
    Send("         </tbody>")
    Send("        </table>")
    
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
  End Sub
  
  Sub ShowPendingOrHistory(ByVal CustomerPK As Long, ByVal EarnedROID As Integer, _
                           ByVal EarnedCMOfferID As Integer, ByVal ShowHistory As Boolean, _
                           ByVal IsHousehold As Boolean, ByVal HHMembershipCount As Integer, _
                           ByRef Logix As Copient.LogixInc)
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim dt3 As DataTable
    Dim row As DataRow
    Dim sQueryBuilder As New StringBuilder()
    Dim ProgramName As String = ""
    Dim ProgramList As String = ""
    Dim ProgramID As Integer
    Dim OfferNumber As Integer
    Dim LocationID As Long
    Dim iLocalServerID As Integer
    Dim iEngineId As Integer = -1
    Dim bManualEntry As Boolean
    Dim bAutoTransfer As Boolean
    Dim ExtLocationCode As String = ""
    Dim i As Integer = 0
    Dim j As Integer = 0
    Dim StartDate, EndDate As Date
    Dim StartDateStr As String = ""
    Dim EndDateStr As String = ""
    Dim Cookie As HttpCookie = Nothing
    Dim HistoryEnabled As Boolean = True
    Dim AltText As String = ""
    Dim PresentedCustomerID As String = ""
    Dim ResolvedCustomerID As String = ""
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    Dim ShowPOSTimeStamp As Boolean = IIf(MyCommon.Fetch_CPE_SystemOption(131) = "1", True, False)

    If (NonExtProgramIdList.Count > 0) Then
      For i = 0 To NonExtProgramIdList.Count - 1
        If (i > 0) Then ProgramList &= ", "
        If ShowHistory Then ProgramList &= "'"
        ProgramList &= NonExtProgramIdList.Item(i).ToString
        If ShowHistory Then ProgramList &= "'"
      Next
    Else
      If ShowHistory Then ProgramList &= "'"
      ProgramList &= "-1"
      If ShowHistory Then ProgramList &= "'"
    End If
    If NonExtProgramIdList.Count > 0 Then
      If (ShowHistory) Then
        ' determine if history should be loaded per the users request
        Cookie = Request.Cookies("HistoryEnabled")
        If Not (Cookie Is Nothing) Then
          HistoryEnabled = IIf(Cookie.Value = "0", False, True)
        End If
      
        ' if user is attempting a search when history display is disabled, then enable it for them
        If (Not HistoryEnabled AndAlso Request.QueryString("SearchHistory") <> "") Then
          Response.Cookies("HistoryEnabled").Expires = "10/08/2100"
          Response.Cookies("HistoryEnabled").Value = "1"
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
      
        Send("<div style=""width:100%;background-color:#e0e0e0;text-align:center;border:1px solid #808080;position:relative;"">")
        Send("<input type=""image"" name=""" & IIf(HistoryEnabled, "HistoryDisabled", "HistoryEnabled") & """ onclick=""OtherButtonsClicked()""  src=""" & IIf(HistoryEnabled, "/images/history-on.png", "/images/history-off.png") & """ alt=""" & AltText & """ title=""" & AltText & """ style=""position:absolute;left:15px;margin-top:2px;"" />")
        Send("<label for=""historyFrom""><b>" & Copient.PhraseLib.Lookup("term.startdate", LanguageID) & "</b>:</label>")
        Send("<input type=""text"" id=""historyFrom"" name=""historyFrom"" class=""short"" value=""" & StartDateStr & """ />")
        Send("<img src=""/images/calendar.png"" class=""calendar"" id=""start-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('historyFrom', event);"" />")
        Send("&nbsp;&nbsp;")
        Send("<label for=""historyTo""><b>" & Copient.PhraseLib.Lookup("term.enddate", LanguageID) & "</b>:</label>")
        Send("<input type=""text"" id=""historyTo"" name=""historyTo"" class=""short"" value=""" & EndDateStr & """ />")
        Send("<img src=""/images/calendar.png"" class=""calendar"" id=""end-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('historyTo', event);"" />")
        Send("&nbsp;&nbsp;")
        Send("<input type=""submit"" id=""SearchHistory"" name=""SearchHistory"" onclick=""OtherButtonsClicked()"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
        Sendb("<span style=""position:absolute;right:15px;margin-top:3px;"">")
        Sendb("<span style=""color:#000000;cursor:help;"" alt=""" & Copient.PhraseLib.Lookup("points-adjust.BlackNormalAdjustments", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("points-adjust.BlackNormalAdjustments", LanguageID) & """>█</span>&nbsp;")
        Sendb("<span style=""color:#cc0000;cursor:help;"" alt=""" & Copient.PhraseLib.Lookup("points-adjust.RedManualAdjustments", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("points-adjust.RedManualAdjustments", LanguageID) & """>█</span>&nbsp;")
        Sendb("<span style=""color:#0000cc;cursor:help;"" alt=""" & Copient.PhraseLib.Lookup("points-adjust.BlueExpiredAdjustments", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("points-adjust.BlueExpiredAdjustments", LanguageID) & """>█</span>")
        Send("</span>")
        Send("</div>")
        Send("<br class=""half"" />")
      
        If Not HistoryEnabled Then Exit Sub
      
        ' if this is a household, then add all the cardholders in the household to the query
        Dim CardholderClause As String = ""
        MyCommon.QueryStr = "select top 20 CustomerPK from Customers with (NoLock) where HHPK=" & CustomerPK
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
          CardholderClause = " and CustomerPK in (" & CustomerPK
          For i = 0 To dt.Rows.Count - 1
            CardholderClause &= "," & MyCommon.NZ(dt.Rows(i).Item("CustomerPK"), 0)
          Next
          CardholderClause &= ") order by " & IIf(ShowPOSTimeStamp, "POSTimeStamp", "LastUpdate") & " desc"
        Else
          CardholderClause = " and CustomerPK = " & CustomerPK & " order by " & IIf(ShowPOSTimeStamp, "POSTimeStamp", "LastUpdate") & " desc"
        End If
        
        sQueryBuilder.Append("select top 50 * from ( ")
        sQueryBuilder.Append("select Top 50 CustomerPK, ProgramID, AdjAmount, EarnedUnderROID, EarnedUnderCMOfferID, LastUpdate, LastServerID, LocationID, ")
        sQueryBuilder.Append("PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate, 0 as Inactive, POSTimeStamp ")
        sQueryBuilder.Append("from PointsHistoryView with (NoLock) ")
        sQueryBuilder.Append("where LocationID <> -8 and ProgramID In (" & ProgramList & ") and LastUpdate between '" & StartDate.ToString("yyyy-MM-dd HH:mm:ss") & "' and '" & EndDate.ToString("yyyy-MM-dd HH:mm:ss") & "' ")
        sQueryBuilder.Append(CardholderClause)
        sQueryBuilder.Append(" UNION ALL ")
        sQueryBuilder.Append("select Top 50 CustomerPK, ProgramID, AdjAmount, EarnedUnderROID, EarnedUnderCMOfferID, LastUpdate, LastServerID, LocationID, ")
        sQueryBuilder.Append("PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate, 1 as Inactive, POSTimeStamp ")
        sQueryBuilder.Append("from InActivePointsHistoryView with (NoLock) ")
        sQueryBuilder.Append("where LocationID <> -8 and ProgramID In (" & ProgramList & ") and LastUpdate between '" & StartDate.ToString("yyyy-MM-dd HH:mm:ss") & "' and '" & EndDate.ToString("yyyy-MM-dd HH:mm:ss") & "' ")
        sQueryBuilder.Append(CardholderClause)
        sQueryBuilder.Append(") as PH order by " & IIf(ShowPOSTimeStamp, "POSTimeStamp", "LastUpdate") & " desc;")
        
      Else
        sQueryBuilder.Append("select Col2 as CustomerPK, Convert(int, Col1) as ProgramID, Convert( int, cast( IsNull( Col3, 0 )  AS float ) ) as AdjAmount, ")
        sQueryBuilder.Append("Convert(int,Col4) as EarnedUnderROID, 0 as EarnedUnderCMOfferID, getDate() as LastUpdate, Replayed, ReplayedDate, 0 as Inactive, POSTimeStamp ")
        sQueryBuilder.Append("from CPE_UploadTemp_PA with (NoLock) ")
        sQueryBuilder.Append("where Col1 in (" & ProgramList & ") ")
        sQueryBuilder.Append("and Col2='" & CustomerPK & "';")
      End If
    
      MyCommon.QueryStr = sQueryBuilder.ToString
      dt = MyCommon.LXS_Select
      If (dt.Rows.Count > 0) Then
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.history", LanguageID) & """>")
        Send("  <tr>")
        If (ShowHistory) Then
          Send("    <th style=""width:30px;"">&nbsp;</th>")
        End If
        Send("    <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & "</th>")
        Send("    <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
        Send("    <th class=""th-adjustment"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & "</th>")
        Send("    <th class=""th-author"" scope=""col"">" & Copient.PhraseLib.Lookup("term.from", LanguageID) & "</th>")
        Send("    <th class="""" scope=""col"" style=""text-align:center;width:20px;"" title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """>" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</th>")
        If ShowPOSTimeStamp Then
          Send("    <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
        Else
          Send("    <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.lastupdated", LanguageID) & "</th>")
        End If
        Send("  </tr>")
        j = 0
        For Each row In dt.Rows
          j += 1
          ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
          If (ShowHistory) Then
            LocationID = MyCommon.NZ(row.Item("LocationID"), 0)
            iLocalServerID = MyCommon.NZ(row.Item("LastServerID"), 0)
            If LocationID > 0 Then
              MyCommon.QueryStr = "select ExtLocationCode, EngineID from Locations with (NoLock) where LocationID=" & LocationID & ";"
            Else
              MyCommon.QueryStr = "select L.ExtLocationCode, L.EngineID from LocalServers LS with (NoLock) inner join Locations L with (NoLock) on L.LocationID=LS.LocationID where LS.LocalServerID=" & iLocalServerID & ";"
            End If
            dt2 = MyCommon.LRT_Select
            If (dt2.Rows.Count > 0) Then
              ExtLocationCode = MyCommon.NZ(dt2.Rows(0).Item("ExtLocationCode"), "")
              iEngineId = MyCommon.NZ(dt2.Rows(0).Item("EngineID"), -1)
            End If
          End If
          MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & ProgramID
          dt2 = MyCommon.LRT_Select
          If (dt2.Rows.Count > 0) Then
            ProgramName = MyCommon.NZ(dt2.Rows(0).Item("ProgramName"), "")
          End If
          If MyCommon.NZ(row.Item("EarnedUnderROID"), 0) = 0 AndAlso MyCommon.NZ(row.Item("EarnedUnderCMOfferID"), 0) = 0 Then
            OfferNumber = 0
          Else
            If MyCommon.NZ(row.Item("EarnedUnderROID"), 0) <> 0 Then
              MyCommon.QueryStr = "select IncentiveID as OfferID from CPE_RewardOptions as RO where RewardOptionID=" & MyCommon.NZ(row.Item("EarnedUnderROID"), 0)
              dt3 = MyCommon.LRT_Select
              If (dt3.Rows.Count > 0) Then
                OfferNumber = MyCommon.NZ(dt3.Rows(0).Item("OfferID"), 0)
              End If
            End If
            If MyCommon.NZ(row.Item("EarnedUnderCMOfferID"), 0) <> 0 Then
              OfferNumber = MyCommon.NZ(row.Item("EarnedUnderCMOfferID"), 0)
            End If
          End If
          'Determine if it's ManualEntry or an AutoTransfer
          bManualEntry = False
          bAutoTransfer = False
          If iEngineId = 0 Then
            'CM location
            If iLocalServerID < 1 Then
              bManualEntry = True
            End If
          Else
            If MyCommon.NZ(row.Item("EarnedUnderROID"), 0) = 0 Then
              bManualEntry = True
            End If
          End If
          If LocationID = -99 Then
            bAutoTransfer = True
          End If
          Sendb("  <tr id=""hist" & j & """" & IIf(IsHousehold AndAlso CustomerPK = MyCommon.NZ(row.Item("CustomerPK"), 0), " class=""shaded""", ""))
          If MyCommon.NZ(row.Item("Inactive"), 0) = 1 Then
            Sendb(" style=""color:#0000cc;""")
          ElseIf bManualEntry Then
            Sendb(" style=""color:#cc0000;""")
          End If
          Send(">")
          If (ShowHistory) Then
            Send("    <td><input class=""ex more"" type=""button"" value=""+"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """ onclick=""javascript:showDetail(" & j & ", this);"" /></td>")
          End If
          Send("    <td>" & MyCommon.SplitNonSpacedString(ProgramName, 25) & "</td>")
          Send("    <td>" & ProgramID & "</td>")
          Send("    <td>" & MyCommon.NZ(row.Item("AdjAmount"), "&nbsp;") & "</td>")
          If bAutoTransfer Then
            Send("    <td>" & Copient.PhraseLib.Lookup("term.autoofferdist", LanguageID) & "</td>")
          ElseIf bManualEntry Then
            Send("    <td>" & Copient.PhraseLib.Lookup("term.logix-manual-entry", LanguageID) & "</td>")
          Else
            Sendb("    <td>")
            If OfferNumber > 0 Then
              Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferNumber & " ")
            End If
            If ExtLocationCode <> "" Then
              If OfferNumber > 0 Then
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.at", LanguageID), VbStrConv.Lowercase) & " ")
              End If
              Sendb(ExtLocationCode)
            End If
            Send("</td>")
          End If
          Send("    <td style=""text-align:center;"">" & IIf(MyCommon.NZ(row.Item("Replayed"), False), "<span title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """ style=""color:#dd0000;cursor:default;font-size:10px;font-weight:bold;"">" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</span>", "") & "</td>")
          If ShowPOSTimeStamp Then
            If (Not IsDBNull(row.Item("POSTimeStamp"))) Then
              Send("    <td>" & Logix.ToShortDateTimeString(row.Item("POSTimeStamp"), MyCommon) & "</td>")
            Else
              Send("    <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
            End If
          Else
            If (Not IsDBNull(row.Item("LastUpdate"))) Then
              Send("    <td>" & Logix.ToShortDateTimeString(row.Item("LastUpdate"), MyCommon) & "</td>")
            Else
              Send("    <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
            End If
          End If
          Send("  </tr>")
          If (ShowHistory) Then
                        PresentedCustomerID = IIf(IsDBNull(row.Item("PresentedCustomerID")), "Unknown", MyCryptLib.SQL_StringDecrypt(row.Item("PresentedCustomerID").ToString()))
                        ResolvedCustomerID = IIf(IsDBNull(row.Item("ResolvedCustomerID")), "Unknown", MyCryptLib.SQL_StringDecrypt(row.Item("ResolvedCustomerID").ToString()))
            If (ResolvedCustomerID = "0" OrElse ResolvedCustomerID = "Unknown") Then
              MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK=" & MyCommon.NZ(row.Item("CustomerPK"), 0) & ";"
              dt3 = MyCommon.LXS_Select
              If dt3.Rows.Count > 0 Then
                ResolvedCustomerID = IIf(IsDBNull(dt3.Rows(0).Item("ExtCardID")), "Unknown",MyCryptLib.SQL_StringDecrypt(dt3.Rows(0).Item("ExtCardID").ToString()))
              End If
            End If
            Send("  <tr id=""histdetail" & j & """" & IIf(IsHousehold AndAlso CustomerPK = MyCommon.NZ(row.Item("CustomerPK"), 0), " class=""shaded""", "") & " style=""display:none;color:#777777;"">")
            Send("    <td></td>")
            Send("    <td colspan=""5"">")
            Send("      " & Copient.PhraseLib.Lookup("term.presented", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & PresentedCustomerID & " &nbsp;|&nbsp; ")
            Send("      " & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & ResolvedCustomerID & " &nbsp;|&nbsp; ")
            Send("      " & Copient.PhraseLib.Lookup("term.household", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & IIf(IsDBNull(row.Item("HHID")), Copient.PhraseLib.Lookup("term.unknown", LanguageID), row.Item("HHID").ToString()))
            Send("    </td>")
            Send("  </tr>")
          End If
        Next
        Send("</table>")
      Else
        If (ShowHistory) Then
          Send("<i>" & Copient.PhraseLib.Lookup("point-adjust.nopointshistory", LanguageID) & "</i>")
        Else
          Send("<i>" & Copient.PhraseLib.Lookup("point-adjust.nopending", LanguageID) & "</i>")
        End If
      End If
    End If
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
  End Sub
  
  Function IsValidAdjustment(ByVal OfferID As Long, ByVal CustomerPK As Long, ByRef infoMessage As String, ByRef WarningProgramID As Long, ByRef ProgramID As Long, _
                             ByRef AdjustAmt As Long, ByRef ExternalProgram As Boolean, ByRef ExtHostTypeID As Integer, ByRef DecimalValues As Boolean, _
                             ByVal ReasonID As Integer, ByVal ReasonText As String) As Boolean
    Dim ValidAdj As Boolean = False
    Dim MyCAM As New Copient.CAM
    Dim MyPoints As New Copient.Points
    Dim PointsBal, WarningLimit, PendingAdj As Long
    Dim dt As DataTable
    Dim ConfirmMessage As String = ""
    
    PointsBal = MyPoints.GetBalance(CustomerPK, ProgramID)
    WarningLimit = MyCAM.GetMaxAdjustment(OfferID, ProgramID)
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    MyCommon.QueryStr = "select sum(Convert( int, cast( IsNull( Col3, 0 )  AS float ) ) ) as Pending from CPE_UploadTemp_PA with (NoLock) " & _
                        "where Col1='" & ProgramID & "' and Col2='" & CustomerPK & "';"
    dt = MyCommon.LXS_Select
    If (dt.Rows.Count > 0) Then
      PendingAdj = MyCommon.NZ(dt.Rows(0).Item("Pending"), 0)
    End If
    
    'System.IO.File.AppendAllText("D:\Copient\aaa.txt", "ISVALIDADJUSTMENT FUNCTION" & vbCrLf)
    'System.IO.File.AppendAllText("D:\Copient\aaa.txt", " AdjustAmt = " & AdjustAmt & vbCrLf)
    
    If ProgramID <= 0 Then
      infoMessage &= Copient.PhraseLib.Detokenize("customer-manual.InvalidProgramID", LanguageID, ProgramID)
    Else
      If (MyCommon.Fetch_SystemOption(108) = "1" AndAlso (ReasonID = 0 OrElse RTrim(ReasonText) = "")) Then
        infoMessage &= Copient.PhraseLib.Lookup("points-adjust.ReasonRequired", LanguageID)
      Else
        If ExternalProgram And ExtHostTypeID = 2 Then
          'This is the new CRM (Excentus) case, which merely needs a check for negativity.
          If AdjustAmt <= 0 Then
            infoMessage &= Copient.PhraseLib.Lookup("points-adjust.InvalidAdjustment", LanguageID)
          Else
            ValidAdj = True
          End If
        Else
          'This is the normal, pre-existing check routine.
          If ((PointsBal + PendingAdj) + AdjustAmt < 0 AndAlso AdjustAmt < 0) Then
            infoMessage &= Copient.PhraseLib.Detokenize("points-adjust.NegativeBalWarning", LanguageID, (-(PointsBal + PendingAdj)))
          ElseIf (Math.Abs(AdjustAmt) > WarningLimit) Then
            ' if a warning has already been issued then allow the adjustment
            If WarningProgramID = 0 Then
              infoMessage &= Copient.PhraseLib.Lookup("points-adjust.OverMaxLimit", LanguageID)
              WarningProgramID = ProgramID
            Else
              ValidAdj = True
            End If
          Else
            ValidAdj = True
          End If
        End If
      End If
    End If
    
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    
    Return ValidAdj
  End Function
  
  Sub AdjustPoint(ByVal AdminUserID As String, ByVal SessionID As String, ByVal SelectedOfferID As Long, ByRef ER As Copient.ExternalRewards, _
                  ByVal ReasonID As Integer, ByVal ReasonText As String)
    Dim ProgramIDs() As String
    Dim PointsAdjs() As String
    Dim CustomerPK As String = ""
    Dim i As Integer
    Dim MyPoints As New Copient.Points
    Dim ExtCardID As String = ""
    Dim dt As System.Data.DataTable
    Dim ExternalProgram As Boolean
    Dim ExtHostTypeID As Integer
    Dim ExternalID As String
    Dim ProgramName As String
    Dim PromoVarID As String
    
    If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
    If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
    
    ProgramIDs = Request.QueryString.GetValues("programID")
    PointsAdjs = Request.QueryString.GetValues("adjust")
    CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
    
    'System.IO.File.AppendAllText("D:\Copient\aaa.txt", "ADJUSTPOINT SUB" & vbCrLf)
    'For i = ProgramIDs.GetLowerBound(0) To ProgramIDs.GetUpperBound(0)
    '  If (MyCommon.Extract_Val(PointsAdjs(i)) <> 0) Then
    '    'System.IO.File.AppendAllText("D:\Copient\aaa.txt", " PointsAdj (" & ProgramIDs(i).ToString & ") = " & PointsAdjs(i) & vbCrLf)
    '  End If
    'Next
    
    MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by CardTypeID;"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      ExtCardID = MyCommon.NZ(dt.Rows(0).Item("ExtCardID"), "0")
    End If
    
    If (CustomerPK > 0 AndAlso Not ProgramIDs Is Nothing AndAlso Not PointsAdjs Is Nothing) Then
      For i = ProgramIDs.GetLowerBound(0) To ProgramIDs.GetUpperBound(0)
        If (MyCommon.Extract_Val(PointsAdjs(i)) <> 0) Then
          MyCommon.QueryStr = "select ExternalProgram, ExtHostTypeID, ExtHostProgramID, ProgramName, PromoVarID from PointsPrograms with (NoLock) where Deleted=0 and ProgramID=" & ProgramIDs(i) & ";"
          dt = MyCommon.LRT_Select()
          If dt.Rows.Count > 0 Then
            ExternalProgram = MyCommon.NZ(dt.Rows(0).Item("ExternalProgram"), False)
            ExtHostTypeID = MyCommon.NZ(dt.Rows(0).Item("ExtHostTypeID"), 0)
            ExternalID = MyCommon.NZ(dt.Rows(0).Item("ExtHostProgramID"), "")
            ProgramName = MyCommon.NZ(dt.Rows(0).Item("ProgramName"), "")
            PromoVarID = MyCommon.NZ(dt.Rows(0).Item("PromoVarID"), "")
            If (ExternalProgram AndAlso ExtHostTypeID = 2) Then
              'it's an external program for which we must use the rewardUser subroutine in the ExternalRewards DLL
              ER.rewardUser(ProgramIDs(i).ToString, "LogixAdjustment", PointsAdjs(i), ExtCardID, 0)
              'System.IO.File.AppendAllText("D:\Copient\aaa.txt", "Adjustment performed for " & PointsAdjs(i))
              'System.IO.File.AppendAllText("D:\Copient\aaa.txt", vbCrLf & vbCrLf)
            ElseIf (ExternalProgram AndAlso ExtHostTypeID = 1) Then
              'it's an EME external program for which we must use the rewardUser subroutine in the ExternalRewards DLL
              ER.updateExternalBalance(ExternalID, ProgramName, PointsAdjs(i), ExtCardID, CustomerPK, PromoVarID, "0", "-9", "0", MyCommon)
            Else
              MyPoints.AdjustPoints(AdminUserID, MyCommon.Extract_Val(ProgramIDs(i)), CustomerPK, MyCommon.Extract_Val(PointsAdjs(i)), 0, 0, SessionID, SelectedOfferID, -9, ReasonID, ReasonText)
            End If
          End If
        End If
      Next
    End If
  End Sub
  Sub ProcessIssuance(ByVal OfferID as Long, ByVal CustomerPK as Long, ByVal ProgramID as String, ByVal PointsAdj As String, ByRef infoMessage As String)
	Dim CurrentIssuanceTable As String
	Dim CustomerTypeID As String
	Dim CardTypeID As String
	Dim AdjAmount As String
	Dim PrimaryExtID As String
	Dim AirmileMemberID As String
	Dim PromoEngine As String = "CPE"
	Dim dt As DataTable = Nothing
	Dim dt2 As DataTable = Nothing
	
	If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
    If (MyCommon.LEXadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixEX()
	
	MyCommon.QueryStr = "select top(1) TableName from IssuanceTables order by PKID desc;"
	dt = MyCommon.LEX_Select
	If dt.Rows.Count > 0 Then
		CurrentIssuanceTable = MyCommon.NZ(dt.Rows(0).Item("TableName"), "")
	End If
	
	MyCommon.QueryStr = "select CI.ExtCardID, CT.CustTypeID, CT.CardTypeID, CE.AirmileMemberID from CardIDs as CI with (NoLock) " & _
						"join CardTypes as CT with (NoLock) on CT.CardTypeID = CI.CardTypeID " & _
						"join CustomerExt as CE on CI.CustomerPK=CE.CustomerPK where CI.CustomerPK = " & CustomerPK & ";"
	dt2 = MyCommon.LXS_Select
        If dt2.Rows.Count > 0 Then
            'No need to decrypt as we are passting this value into Insert  
            PrimaryExtID =MyCryptLib.SQL_StringDecrypt(dt2.Rows(0).Item("ExtCardID").ToString())
            CustomerTypeID = MyCommon.NZ(dt2.Rows(0).Item("CustTypeID"), "")
            CardTypeID = MyCommon.NZ(dt2.Rows(0).Item("CardTypeID"), "")
            AirmileMemberID = MyCommon.NZ(dt2.Rows(0).Item("AirmileMemberID"), "")
		
            MyCommon.QueryStr = "Insert into " & CurrentIssuanceTable & " with (RowLock)" & _
                        " (ClientLocationCode, LocationID, BoxID, TransactionNumber, PrimaryExtID, OfferID, ROID," & _
                        "  IssuanceDate, DeliverableType, Void, RewardQty, ProgramID," & _
                        "  SourceTypeID, PromoEngine, CustomerTypeID, CardTypeID, RewardValue," & _
                        "  LogixTransNum, AirmileMemberID)" & _
                        " values " & _
                        " (0, 0, 0, 0, '" & PrimaryExtID & "', " & OfferID & ", 0, getdate(), 8, 0," & _
                        " " & PointsAdj & ", " & ProgramID & ", 1, '" & PromoEngine & "', " & CustomerTypeID & ", " & CardTypeID & ", " & PointsAdj & ", 0, '" & AirmileMemberID & "');"
            MyCommon.LEX_Execute()
        End If
  End Sub
</script>

<%
done:
  Send_BodyEnd("mainform", "adjust")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixEX()
  MyCommon = Nothing
  Logix = Nothing
%>
