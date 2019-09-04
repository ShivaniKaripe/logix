<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="Copient" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: point-adjust-program.aspx 
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
    Dim MyCryptLib As New Copient.CryptLib
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
    Dim OfferID As Long
    Dim PageTitle As String = ""
    Dim OfferName As String = ""
    Dim ProgramID As Long
    Dim ProgramName As String = ""
    Dim ProgramDesc As String = ""
    Dim AdjustPermitted As Boolean = False
    Dim EarnedROID As Integer = 0
    Dim EarnedCMOffer As Integer = 0
    Dim OfferDesc As String = ""
    Dim InfoMessage As String = ""
    Dim ConfirmMessage As String = ""
    Dim Handheld As Boolean = False
    Dim HHPK As Integer = 0
    Dim HHCardPK As Long = 0
    Dim i As Integer = 0
    Dim KeyCt As Integer = 0
    Dim RefreshParent As Boolean = False
    Dim SessionID As String = ""
    Dim ProgramIDs(), PointsAdjs() As String
    Dim ValidAdj As Boolean = False
    Dim tempProgram As Long
    Dim tempPoints As Long
    Dim Note As String = ""
    Dim FirstName As String = ""
    Dim LastName As String = ""
    Dim IsUSAirMilesProgram As Boolean = False
    Dim IsHousehold As Boolean = False
    Dim HHMembershipCount As Integer = 0
    Dim ExternalProgram As Boolean = False
    Dim ExtHostTypeID As Integer = 0
    Dim DecimalValues As Boolean = False
    Dim ER As Copient.ExternalRewards
    Dim ReasonID As Integer = 0
    Dim ReasonText As String = ""
    Dim ExtLocationID As String = ""
    Dim ReasonDescription As String = ""
    Dim iAdjustLimit As Integer = 0
    Dim bolSetpointsprogrambalances As Boolean = False
    Dim customerLookup As CustomerLookup = New CustomerLookup()
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "point-adjust-program.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    If Request.QueryString("SetTo") = "true" Then
        bolSetpointsprogrambalances = True
    End If
  
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
    ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
  
    If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then
        If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(66), iAdjustLimit) Then iAdjustLimit = 0
    Else
        iAdjustLimit = 0
    End If
  
    Dim Opener As String = ""
    If Request.QueryString("Opener") <> "" Then Opener = Request.QueryString("Opener")
    MyCommon.QueryStr = "select ProgramName, Description, ExternalProgram, ExtHostTypeID, DecimalValues from PointsPrograms with (NoLock) " & _
                        "where ProgramID=" & ProgramID & ";"
    dt = MyCommon.LRT_Select()
    If (dt.Rows.Count > 0) Then
        ProgramName = MyCommon.NZ(dt.Rows(0).Item("ProgramName"), "")
        ProgramDesc = MyCommon.NZ(dt.Rows(0).Item("Description"), "")
        ExternalProgram = MyCommon.NZ(dt.Rows(0).Item("ExternalProgram"), False)
        ExtHostTypeID = MyCommon.NZ(dt.Rows(0).Item("ExtHostTypeID"), 0)
        DecimalValues = MyCommon.NZ(dt.Rows(0).Item("DecimalValues"), False)
    End If
  
    'See if the program is associated to at least one US Airmiles offer (Engine 2, SubEngine 2)
    MyCommon.QueryStr = "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_IncentivePointsGroups IPG with (NoLock) " & _
                        "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                        "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                        "INNER JOIN PointsPrograms PP with (NoLock) on IPG.ProgramID=PP.ProgramID " & _
                        "INNER JOIN PromoEngineSubTypes PEST with (NoLock) on PEST.PromoEngineID=I.EngineID " & _
                        "WHERE IPG.ProgramID=" & ProgramID & " And EngineSubTypeID=2 And IPG.Deleted=0 And RO.Deleted=0 And i.Deleted=0 And PP.Deleted=0 " & _
                        " UNION " & _
                        "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_DeliverablePoints DP " & _
                        "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DP.RewardOptionID " & _
                        "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                        "INNER JOIN PromoEngineSubTypes PEST with (NoLock) on PEST.PromoEngineID = I.EngineID " & _
                        "WHERE ProgramID=" & ProgramID & " And EngineSubTypeID=2 And DP.Deleted=0 And RO.Deleted=0 And i.Deleted=0;"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
        IsUSAirMilesProgram = True
    End If
  
    If (Logix.UserRoles.EditPointsBalances) AndAlso (IsUSAirMilesProgram = False) Then
        AdjustPermitted = True
    ElseIf (Logix.UserRoles.EditAirmilesPointsBalances) AndAlso (IsUSAirMilesProgram = True) Then
        AdjustPermitted = True
    End If
  
    If (CustomerPK = 0) Then
        InfoMessage = Copient.PhraseLib.Lookup("term.UnableToFindCustomer", LanguageID)
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
        InfoMessage = Copient.PhraseLib.Lookup("customer-inquiry.hh-adjust-note", LanguageID)
    End If
    If (HHPK > 0) Then
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
  
    If (Request.QueryString("confirmmessage") <> "") Then
        ConfirmMessage = Request.QueryString("confirmmessage")
    End If
  
    If (Request.QueryString("save") <> "") Then
        i = 0
        Dim PointsAdjValue As String
        Dim AdjAmt As Long
        ProgramIDs = Request.QueryString.GetValues("programID")
        'PointsAdjs = Request.QueryString.GetValues("adjust")
        If Request.QueryString("reasonID") <> "" Then
            ReasonID = MyCommon.Extract_Val(Request.QueryString("reasonID"))
        End If
        If Request.QueryString("reasonText") <> "" Then
            ReasonText = Request.QueryString("reasonText")
        End If
        ExtLocationID = Request.QueryString("location")
        If Request.QueryString("SetToValue") IsNot Nothing Then
            tempProgram = Request.QueryString("programID")
            PointsAdjValue = Request.QueryString("SetToValue")
            If PointsAdjValue.Trim <> "" Then
                tempPoints = PointsAdjValue
                If ExternalProgram AndAlso ExtHostTypeID = 2 Then
                    If Double.TryParse(PointsAdjValue, AdjAmt) Then
                        tempPoints = Math.Floor(PointsAdjValue * 100)
                    End If
                End If
                
                ValidAdj = IsValidAdjustment(OfferID, CustomerPK, InfoMessage, OfferID, tempProgram, tempPoints, ExternalProgram, ExtHostTypeID, DecimalValues, ReasonID, ReasonText, ExtLocationID, AdminUserID)
            Else
                ValidAdj = False
            End If
        Else
            PointsAdjs = Request.QueryString.GetValues("adjust")

            For i = ProgramIDs.GetLowerBound(0) To ProgramIDs.GetUpperBound(0)
                tempProgram = MyCommon.Extract_Val(ProgramIDs(i))
                tempPoints = MyCommon.Extract_Val(PointsAdjs(i))
                If ExternalProgram AndAlso ExtHostTypeID = 2 Then
                    If Double.TryParse(PointsAdjs(i), AdjAmt) Then
                        tempPoints = Math.Floor(PointsAdjs(i) * 100)
                    End If
                End If
                'System.IO.File.AppendAllText("D:\Copient\aaa.txt", "SAVE ROUTINE" & vbCrLf)
                'System.IO.File.AppendAllText("D:\Copient\aaa.txt", " tempPoints = " & tempPoints & vbCrLf)
                ValidAdj = IsValidAdjustment(OfferID, CustomerPK, InfoMessage, OfferID, tempProgram, tempPoints, ExternalProgram, ExtHostTypeID, DecimalValues, ReasonID, ReasonText, ExtLocationID, AdminUserID)
                If Not ValidAdj Then Exit For
            Next
        End If
        If ValidAdj Then
	        Dim Note1 As String = ""
	        Note1 = "Adjusted Program " & ProgramID & " by " & tempPoints & ""
            MyCommon.QueryStr = "select Description from AdjustmentReasons with (NoLock) where ReasonID=" & ReasonID & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
                ReasonDescription = MyCommon.NZ(dt.Rows(0).Item("Description"), "")
            End If
            If ReasonID <> 0 AndAlso RTrim(ReasonText) <> "" Then
              Note1 &= ". Reason: " & ReasonDescription
              Note1 &= " (" & ReasonText & ") "
            End If
            Note = Request.QueryString("note")
            Note = MyCommon.Parse_Quotes(Note)
            Note = Logix.TrimAll(Note)
            If MyCommon.Fetch_SystemOption(99) Then
                If (Note <> "") And (Note.Length <= 1000) Then
		            Note = Note1 & "." & Note
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
            If InfoMessage = "" Then
                AdjustPoint(AdminUserID, ProgramID, SessionID, ExternalProgram, ExtHostTypeID, ER, ReasonID, ReasonText, ExtLocationID)
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "point-adjust-program.aspx?CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & _
                                   "&ProgramID=" & ProgramID & "&ProgramName=" & Server.UrlEncode(ProgramName) & "&RefreshParent=true" & _
                                   "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                                   "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")) & _
                                   "&confirmMessage=" & ConfirmMessage & "&tempPoints=" & tempPoints & "&SetTo=" & Request.QueryString("SetTo"))
                GoTo done
            End If
        End If
    
    ElseIf (Request.QueryString("HistoryEnabled.x") <> "" OrElse Request.QueryString("HistoryDisabled.x") <> "") Then
        ' Write a cookie and then reload the page
        Response.Cookies("HistoryEnabled").Expires = "10/08/2100"
        Response.Cookies("HistoryEnabled").Value = IIf(Request.QueryString("HistoryEnabled.x") <> "", "1", "0")
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "point-adjust-program.aspx?CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & _
                           "&ProgramID=" & ProgramID & "&ProgramName=" & Server.UrlEncode(ProgramName) & "&RefreshParent=true" & _
                           "&historyTo=" & Server.UrlEncode(Request.QueryString("historyTo")) & _
                           "&historyFrom=" & Server.UrlEncode(Request.QueryString("historyFrom")) & _
                           "&SetTo=" & Request.QueryString("SetTo"))
        GoTo done
    End If
  
    If (Request.QueryString("RefreshParent") = "true") Then RefreshParent = True
  
    Send_HeadBegin("term.point", "term.pointsadjustment", ProgramID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts(New String() {"datePicker.js"})
%>
<script type="text/javascript">
    var datePickerDivID = "datepicker";
    var bSkipUnload = false;
    var linkToHH = false;

  <% Send_Calendar_Overrides(MyCommon)%>

    function isValidEntry(externalProgram, extHostTypeID) {
        var retVal = true;
        var elems = document.getElementsByName("adjust");
        var elemLimit = document.getElementById("AdjustLimit");
        var pointsbalsetvalue = document.getElementById("SetTo").value;
        if (pointsbalsetvalue == "True") {
            elems = document.getElementsByName("SetToValue");
        }

        if (elems != null) {
            for (var i = 0; i < elems.length; i++) {
                if (elems[i].value == "") {
                    // An empty value must be allowed to pass in order to permit the form
                    // to be submitted via the date search button in the history area.
                } else {
                    if (externalProgram == 1 && extHostTypeID == 2) {
                        //CRM(Excentus) checking
                        if (elems[i].value != "" && isNaN(elems[i].value)) {
                            retVal = false;
                            alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerrordecimal", LanguageID))%>');
              elems[i].focus();
              elems[i].select();
              break;
          }
      } else {
              //Regular checking
          if (elems[i].value != "" && (isNaN(elems[i].value) || (isInt(elems[i].value) == false))) {
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerror", LanguageID))%>');
              elems[i].focus();
              elems[i].select();
              break;
          } else {
              if (elemLimit != null && parseFloat(elemLimit.value) > 0.00) {
                  if (Math.abs(parseFloat(elems[i].value)) > parseFloat(elemLimit.value)) {
                      retVal = false;
                      alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.exceededlimit", LanguageID))%>');
                  elems[i].focus();
                  elems[i].select();
              }
          }
      }
  }
}
}
} else {
    var elem = document.getElementById("adjust");
    if (pointsbalsetvalue == "True") {
        elem = document.getElementsByName("SetToValue");
    }
    if (externalProgram == 1 && extHostTypeID == 2) {
        //CRM(Excentus) checking
        if (isNaN(elem.value)) {
            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerrordecimal", LanguageID))%>');
          elem.focus();
      }
  } else {
          //Regular checking
      if (isNaN(elem.value) || (isInt(elem.value) == false)) {
          retVal = false;
          alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.entryerror", LanguageID))%>');
          elem.focus();
      } else {
          if (elemLimit != null && parseFloat(elemLimit.value) > 0.00) {
              if (Math.abs(parseFloat(elem.value)) > parseFloat(elemLimit.value)) {
                  retVal = false;
                  alert('<%Sendb(Copient.PhraseLib.Lookup("points-adjust.exceededlimit", LanguageID))%>');
              elem.focus();
          }
      }
  }
}
}
    bSkipUnload = true;
    if (pointsbalsetvalue == "True") {
        if (retVal == true) {
            var enteredsetvalue = document.getElementById("SetToValue").value;
            if (enteredsetvalue != '') {
                if (parseFloat(enteredsetvalue) < 0) {
                    alert('<%Sendb(Copient.PhraseLib.Lookup("error.requires_valid_positive_integer", LanguageID))%>');
	          retVal = false;
	          document.getElementById("SetToValue").focus();
	      }
      }
  }
}
    return retVal;
}

function ChangeParentDocument() {
    var refreshElem = document.getElementById("RefreshParent");
    var OpenerElem = document.getElementById("Opener");

    if (opener != null && !opener.closed && bSkipUnload != true) {
        if (refreshElem != null && refreshElem.value == "true") {
            if (OpenerElem.value == "customer-adjustments.aspx") {
                if (linkToHH) {
                    opener.location = 'customer-adjustments.aspx?CustPK=<%Sendb(HHPK)%><%Sendb(IIf(HHCardPK > 0, "&CardPK=" & HHCardPK, ""))%>&adjWin=1';
        } else {
            opener.location = 'customer-adjustments.aspx?CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&adjWin=1';
        }
    }
}
}
}

function Refresher(ProgramID, CustomerPK) {
    var refreshElem = document.getElementById("RefreshParent");
    var RefreshPrt = "false";

    bSkipUnload = true;

    if (refreshElem != null && refreshElem.value == 'true') {
        RefreshPrt = "true";
    }
    location = ('point-adjust-program.aspx?ProgramID=' + ProgramID + '&CustomerPK=' + CustomerPK + '<%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>&RefreshParent=' + RefreshPrt + '&historyTo=<%Sendb(Server.UrlEncode(Request.QueryString("historyTo")))%>&historyFrom=<%Sendb(Server.UrlEncode(Request.QueryString("historyFrom")))%>&SetTo=<%Sendb(Request.QueryString("SetTo"))%>')
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
  
    If (Logix.UserRoles.AccessPointsBalances = False) AndAlso (IsUSAirMilesProgram = False) Then
        Send_Denied(2, "perm.customers-ptbalaccess")
        GoTo done
    ElseIf (Logix.UserRoles.AccessAirmilesPointsBalances = False) AndAlso (IsUSAirMilesProgram = True) Then
        Send_Denied(2, "perm.customers-ptbalaccess-airmiles")
        GoTo done
    End If
%>
<form id="mainform" name="mainform" action="" onsubmit="return isValidEntry(<%Sendb(IIf(ExternalProgram, "1", "0"))%>,<% Sendb(ExtHostTypeID)%>);">
    <input type="hidden" id="ProgramID" name="ProgramID" value="<% Sendb(ProgramID)%>" />
    <input type="hidden" id="CustomerPK" name="CustomerPK" value="<% Sendb(CustomerPK)%>" />
    <input type="hidden" id="CardPK" name="CardPK" value="<% Sendb(CardPK)%>" />
    <input type="hidden" id="RefreshParent" name="RefreshParent" value="<% Sendb(RefreshParent.ToString.ToLower)%>" />
    <input type="hidden" id="Opener" name="Opener" value="<% Sendb(Opener)%>" />
    <input type="hidden" id="AdjustLimit" name="AdjustLimit" value="<% Sendb(iAdjustLimit)%>" />
    <input type="hidden" id="SetTo" name="SetTo" value="<% Sendb(bolSetpointsprogrambalances)%>" />
    <div id="intro">
        <h1 id="title">
            <% Sendb(Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & " #" & ProgramID & ": " & MyCommon.TruncateString(ProgramName, 40))%>
        </h1>
        <div id="controls">
            <%
                If (AdjustPermitted) Then
                    Send_Save()
                End If
            %>
        </div>
        <hr class="hidden" />
    </div>
    <div id="main">
        <%
            If (InfoMessage <> "") Then
                Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")
            End If
            If (ConfirmMessage <> "") Then
                Send("<div id=""infobar"" class=""green-background"">" & ConfirmMessage & "</div>")
            End If
        %>
        <div id="column">
            <%
                If ProgramDesc <> "" Then
                    Sendb("<p id=""description"">" & MyCommon.SplitNonSpacedString(ProgramDesc, 50) & "</p>")
                End If
                If HHPK > 0 Then
                    Sendb("<br /><a href=""point-adjust-program.aspx?")
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
        
                Send("<div class=""box"" id=""ptAdj""" & IIf((AdjustPermitted = False) OrElse (HHPK > 0), " style=""display:none;""", "") & ">")
                Send("  <h2>")
                Send("    <span>")
                Send("      " & Copient.PhraseLib.Lookup("term.pointsadjustment", LanguageID))
                Send("    </span>")
                Send("  </h2>")
                If (ProgramID > 0) Then
                    ShowPoints(CustomerPK, ProgramID, ExternalProgram, ExtHostTypeID, ER, Logix)
                Else
                    Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
                End If
                Send("  <hr class=""hidden"" />")
                Send("</div>")
                Send("")
        
                If (MyCommon.Fetch_SystemOption(108) = "1" Or MyCommon.Fetch_SystemOption(193) = "1") Then
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
                    If (MyCommon.Fetch_SystemOption(193) = "1") Then
                        MyCommon.QueryStr = "select ReasonID, Description from AdjustmentReasons with (NoLock) where Enabled=1 and (Program like '%Point%' OR Program like '%All%' OR Program is NULL);"
                    Else
                        MyCommon.QueryStr = "select ReasonID, Description from AdjustmentReasons with (NoLock) where Enabled=1 and UserDefined=0;"
                    End If
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
        
                If (ExternalProgram And ExtHostTypeID <> 1) Then
                    'Do not show history
                Else
                    If (Not ExternalProgram) Then
                        Send("<div class=""box"" id=""pending"">")
                        Send("  <h2>")
                        Send("    <span>")
                        Send("      " & Copient.PhraseLib.Lookup("term.pending", LanguageID))
                        Send("    </span>")
                        Send("  </h2>")
                        Send("  <span style=""float:right; font-size:9px; position:relative; top:-22px;""><a href=""javascript:Refresher(" & ProgramID & "," & CustomerPK & ")"">" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "</a></span>")
                        If (ProgramID > 0) Then
                            ShowPendingOrHistory(CustomerPK, ProgramID, False, IsHousehold, HHMembershipCount, Logix)
                        Else
                            Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
                        End If
                        Send("  <hr class=""hidden"" />")
                        Send("</div>")
                        Send("")
                    End If
                    Send("<div class=""box"" id=""history"">")
                    Send("  <h2>")
                    Send("    <span>")
                    Send("      " & Copient.PhraseLib.Lookup("term.history", LanguageID))
                    Send("    </span>")
                    Send("  </h2>")
                    If (ProgramID > 0) Then
                        ShowPendingOrHistory(CustomerPK, ProgramID, True, IsHousehold, HHMembershipCount, Logix)
                    Else
                        Send(Copient.PhraseLib.Lookup("point-adjust.nopointprograms", LanguageID) & "<br />")
                    End If
                    Send("  <hr class=""hidden"" />")
                    Send("</div>")
                    Send("")
                End If
            %>
        </div>
    </div>
</form>

<script runat="server">
    Dim MyCommon As New Copient.CommonInc
    Dim MyCryptLib As New Copient.CryptLib
  
    Sub ShowPoints(ByVal CustomerPK As Long, ByVal ProgramID As Long, ByVal ExternalProgram As Boolean, ByVal ExtHostTypeID As Integer, ByVal ER As Copient.ExternalRewards, ByRef Logix As Copient.LogixInc)
        Dim UpdateAccum As Boolean = False
        Dim dt As DataTable
        Dim dtPrograms As DataTable
        Dim dtPoints As DataTable
        Dim rowProgram As DataRow = Nothing
        Dim Points As Integer = 0
        Dim PointsAsDecimal As Decimal = 0
        Dim PendingAdj As Long = 0
        Dim i As Integer = 1
        Dim PromoVarID As Long = 0
        Dim ExtCardID As String = ""
        Dim ErrorThrown As Boolean = False
        Dim RewardStatus As Copient.ExternalRewards.RewardStatusInfo
        Dim ExpirationAmount As New Decimal(0)
        Dim DaysTillExpiration As Integer = 0
        Dim sErrorMsg As String = ""
        Dim Setpointsprogrambalances As Boolean = False
    
        MyCommon.Open_LogixRT()
        MyCommon.Open_LogixXS()
        MyCommon.SetAdminUser(Logix.Fetch_AdminUser(MyCommon, AdminUserID))
    
        MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by CardTypeID;"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
            ExtCardID = IIf(IsDBNull(dt.Rows(0).Item("ExtCardID")), "0",MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString()))
        End If
    
        If Request.QueryString("SetTo") = "true" Then
            Setpointsprogrambalances = True
        End If
	
        MyCommon.QueryStr = "select ProgramID, ProgramName, PromoVarID, ExternalProgram, ExtHostTypeID, ExtHostProgramID from PointsPrograms with (NoLock) " & _
                            "where ProgramID=" & ProgramID & ";"
        dtPrograms = MyCommon.LRT_Select
    
        Send("        <br class=""half"" />")
        Send("        " & Copient.PhraseLib.Lookup("point-adjust.programnote", LanguageID) & " " & ProgramID & ".<br />")
        If ExternalProgram And ExtHostTypeID > 1 Then
            Send("        <span style=""color:#cc0000;"">" & Copient.PhraseLib.Lookup("point-adjust.programnoteext", LanguageID) & "</span><br />")
        End If
        Send("        <br class=""half"" />")
        Send("        <table summary=""" & Copient.PhraseLib.Lookup("term.programs", LanguageID) & """>")
        Send("         <thead>")
        Send("          <tr>")
        Send("            <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.program", LanguageID) & "</th>")
        Send("            <th class=""th-id"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
        Send("            <th class=""th-var"" scope=""col"">" & Copient.PhraseLib.Lookup("term.promovarid", LanguageID) & "</th>")
        Send("            <th class=""th-quantity"" scope=""col"" style=""text-align:center"">" & Copient.PhraseLib.Lookup("term.quantity", LanguageID) & "</th>")
        If ExternalProgram Then
            'Pending adjustments are inaccessible
        Else
            Send("            <th class=""th-pending"" scope=""col"" style=""text-align:center"">" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & "</th>")
        End If
        If (MyCommon.Fetch_SystemOption(108) = "1") Then
            Send("            <th class=""th-location"" scope=""col"">" & Copient.PhraseLib.Lookup("term.issuingcostcenter", LanguageID) & "</th>")
        End If
        Send("            <th class=""th-adjustment"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjustment", LanguageID) & "</th>")
        If Setpointsprogrambalances = True Then
            Send("            <th class=""th-adjustment"" scope=""col"">" & Copient.PhraseLib.Lookup("term.setto", LanguageID) & "</th>")
        End If
        Send("          </tr>")
        Send("         </thead>")
        Send("         <tbody>")
    
        For Each rowProgram In dtPrograms.Rows
            ExpirationAmount = 0
            PointsAsDecimal = 0
            DaysTillExpiration = 0
      
            PromoVarID = MyCommon.NZ(rowProgram.Item("PromoVarID"), 0)
            ProgramID = MyCommon.NZ(rowProgram.Item("ProgramID"), 0)
      
            If ExternalProgram And ExtHostTypeID = 2 Then 'use the ExternalRewards DLL
                Try
                    RewardStatus = ER.getUserRewardStatus(ExtCardID, 0)
                    PointsAsDecimal = RewardStatus.totalRewardAmount
                    ExpirationAmount = RewardStatus.rewardExpirationAmount
                    DaysTillExpiration = RewardStatus.rewardExpirationDays
                Catch ex As Copient.ExternalRewards.RewardsException
                    ErrorThrown = True
                End Try
            ElseIf ExternalProgram And ExtHostTypeID = 1 Then 'use the ExternalRewards DLL (EME)
                Points = 0
                If ExtCardID.Length > 0 Then
                    Dim dtExtBalances As DataTable = Nothing
                    Dim ExternalID As String
                    Dim rows() As DataRow
          
                    MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where Deleted=0 and VarTypeID=3 and LinkID=" & ProgramID & ";"
                    dt = MyCommon.LXS_Select()
                    If dt.Rows.Count > 0 Then
                        ExternalID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExternalID").ToString())
                        Try
                            dtExtBalances = ER.getExternalBalances(ExtCardID)
                        Catch ex As Exception
                            sErrorMsg = "EME: " & ex.Message
                            ErrorThrown = True
                        End Try
                        If Not dtExtBalances Is Nothing AndAlso dtExtBalances.Rows.Count > 0 Then
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
      
            MyCommon.QueryStr = "select sum(Convert(bigint, IsNull(Col3,0))) as Pending from CPE_UploadTemp_PA with (NoLock) " & _
                                "where Col1='" & ProgramID & "' and Col2='" & CustomerPK & "';"
            dt = MyCommon.LXS_Select
            If (dt.Rows.Count > 0) Then
                PendingAdj = MyCommon.Extract_Val(MyCommon.NZ(dt.Rows(0).Item("Pending"), 0))
            End If
            Send("          <tr>")
            Send("            <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rowProgram.Item("ProgramName"), "nbsp;"), 25) & "</td>")
            Send("            <td>" & ProgramID & "</td>")
            Send("            <td>" & PromoVarID & "</td>")
      
            If ExternalProgram AndAlso ExtHostTypeID = 2 Then
                'Pending adjustments are inaccessible
                If ErrorThrown Then
                    Send("            <td colspan=""2"" style=""font-size:10px;color:#cc0000;"">" & Copient.PhraseLib.Lookup("points-adjust.Unresolvable", LanguageID) & "<input type=""hidden"" id=""adjust"" name=""adjust"" value=""0"" /></td>")
                Else
                    Send("            <td style=""text-align:right"">$" & Format(PointsAsDecimal, "0.00") & "</td>")
                    Send("            <td>$<input type=""text"" class=""short"" id=""adjust"" name=""adjust"" style=""text-align:right;"" maxlength=""7"" value=""""" & IIf(Setpointsprogrambalances, " disabled=""disabled""", "") & " /></td>")
                    If Setpointsprogrambalances = True Then
                        Send("            <td>$<input type=""text"" class=""short"" id=""SetToValue"" name=""SetToValue"" style=""text-align:right;"" maxlength=""7"" value="""" /></td>")
                    End If
                End If
            ElseIf ExternalProgram AndAlso ExtHostTypeID = 1 Then
                'Pending adjustments are inaccessible
                If ErrorThrown Then
                    Send("            <td colspan=""2"" style=""font-size:10px;color:#cc0000;"">" & Copient.PhraseLib.Lookup("points-adjust.Unresolvable", LanguageID) & "<input type=""hidden"" id=""adjust"" name=""adjust"" value=""0"" /></td>")
                Else
                    Send("            <td style=""text-align:right"">" & Points & "</td>")
                    Send("            <td><input type=""text"" class=""short"" id=""adjust"" name=""adjust"" style=""text-align:right;"" maxlength=""7"" value=""""" & IIf(Setpointsprogrambalances, " disabled=""disabled""", "") & " /></td>")
                    If Setpointsprogrambalances = True Then
                        Send("            <td><input type=""text"" class=""short"" id=""SetToValue"" name=""SetToValue"" style=""text-align:right;"" maxlength=""7"" value="""" /></td>")
                    End If
                End If
            Else
                Send("            <td style=""text-align:right"">" & Points & "</td>")
                Send("            <td style=""text-align:right"">" & PendingAdj & "</td>")
                If (MyCommon.Fetch_SystemOption(108) = "1") Then
                    Send("            <td><input type=""text"" class=""short"" id=""location"" name=""location"" style=""text-align:right;"" maxlength=""50"" value="""" /></td>")
                End If
                Send("            <td><input type=""text"" class=""short"" id=""adjust"" name=""adjust"" style=""text-align:right;"" maxlength=""7"" value=""""" & IIf(Setpointsprogrambalances, " disabled=""disabled""", IIf(ExternalProgram AndAlso ExtHostTypeID < 1, " disabled=""disabled""", "")) & " /></td>")
                If Setpointsprogrambalances = True Then
                    Send("            <td><input type=""text"" class=""short"" id=""SetToValue"" name=""SetToValue"" style=""text-align:right;"" maxlength=""7"" value="""" /></td>")
                End If
            End If
            Send("          </tr>")
      
            
            If ExternalProgram AndAlso ExtHostTypeID = 2 AndAlso Not ErrorThrown AndAlso ExpirationAmount > 0 Then
                Send("          <tr>")
                Send("            <td colspan=""6"" style=""font-size:10px;color:#cc0000;padding-left:5px;"">" & Copient.PhraseLib.Detokenize("point-adjust-program.ExpirationWarning", LanguageID, Math.Round(ExpirationAmount, 2), DaysTillExpiration, Logix.ToShortDateTimeString(DateAdd("d", DaysTillExpiration, Now()), MyCommon)) & "</td>")
                Send("          </tr>")
            End If
      
            If MyCommon.Fetch_SystemOption(99) Then
                Send("          <tr>")
                Send("            <td style=""vertical-align:top;""><label for=""note"">" & Copient.PhraseLib.Lookup("point-adjust.AdjusmentExplanation", LanguageID) & "</note></td>")
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
  
    Sub ShowPendingOrHistory(ByVal CustomerPK As Long, ByVal ProgramID As Long, ByVal ShowHistory As Boolean, ByVal IsHousehold As Boolean, ByVal HHMembershipCount As Integer, ByRef Logix As Copient.LogixInc)
        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim dt3 As DataTable
        Dim row As DataRow
        Dim sQueryBuilder As New StringBuilder()
        Dim ProgramName As String = ""
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
        Dim sizeOfData As Integer
        Dim idNumber As Integer
        Dim idSearch As String
        Dim idSearchText As String
        Dim PageNum As Integer = 0
        Dim MorePages As Boolean
        Dim linesPerPage As Integer = 15
        Dim Shaded As String = "shaded"
        Dim Cookie As HttpCookie = Nothing
        Dim HistoryEnabled As Boolean = True
        Dim AltText As String = ""
        Dim PresentedCustomerID As String = ""
        Dim ResolvedCustomerID As String = ""
    
        PageNum = Request.QueryString("pagenum")
        If PageNum < 0 Then PageNum = 0
        MorePages = False
    
        Dim SortText As String = "ActivityDate"
        Dim SortDirection As String = ""
    
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
    
        '*Search by EarnedUnderROID and EarnedUnderCMofferID in case this is in a CM environment.
        If Request.QueryString("SortText") <> "" Then
            SortText = Request.QueryString("SortText")
            If SortText = "EarnedUnderROID" Then SortText = "EarnedUnderROID,EarnedUnderCMOfferID"
        Else
            SortText = ""
        End If
    
        MyCommon.Open_LogixRT()
        MyCommon.Open_LogixXS()
    
        Dim ShowPOSTimeStamp As Boolean = IIf(MyCommon.Fetch_CPE_SystemOption(131) = "1", True, False)

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
            Send("<input type=""image"" name=""" & IIf(HistoryEnabled, "HistoryDisabled", "HistoryEnabled") & """ src=""" & IIf(HistoryEnabled, "/images/history-on.png", "/images/history-off.png") & """ alt=""" & AltText & """ title=""" & AltText & """ style=""position:absolute;left:15px;margin-top:2px;"" />")
            Send("<label for=""historyFrom""><b>" & Copient.PhraseLib.Lookup("term.startdate", LanguageID) & "</b>:</label>")
            Send("<input type=""text"" id=""historyFrom"" name=""historyFrom"" class=""short"" value=""" & StartDateStr & """ />")
            Send("<img src=""/images/calendar.png"" class=""calendar"" id=""start-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('historyFrom', event);"" />")
            Send("&nbsp;&nbsp;")
            Send("<label for=""historyTo""><b>" & Copient.PhraseLib.Lookup("term.enddate", LanguageID) & "</b>:</label>")
            Send("<input type=""text"" id=""historyTo"" name=""historyTo"" class=""short"" value=""" & EndDateStr & """ />")
            Send("<img src=""/images/calendar.png"" class=""calendar"" id=""end-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('historyTo', event);"" />")
            Send("&nbsp;&nbsp;")
            Send("<input type=""submit"" id=""SearchHistory"" name=""SearchHistory"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ />")
            Sendb("<span style=""position:absolute;right:15px;margin-top:3px;"">")
            Sendb("<span style=""color:#000000;cursor:help;"" alt=""" & Copient.PhraseLib.Lookup("points-adjust.BlackNormalAdjustments", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("points-adjust.BlackNormalAdjustments", LanguageID) & """>█</span>&nbsp;")
            Sendb("<span style=""color:#cc0000;cursor:help;"" alt=""" & Copient.PhraseLib.Lookup("points-adjust.RedManualAdjustments", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("points-adjust.RedManualAdjustments", LanguageID) & """>█</span>&nbsp;")
            Sendb("<span style=""color:#0000cc;cursor:help;"" alt=""" & Copient.PhraseLib.Lookup("points-adjust.BlueExpiredAdjustments", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("points-adjust.BlueExpiredAdjustments", LanguageID) & """>█</span>")
            Send("</span>")
            Send("</div>")
            Send("<br class=""half"" />")
      
            If Not HistoryEnabled Then Exit Sub
            
            Dim CardholderClause As String = ""

            If SortText <> "" Then
                CardholderClause &= " and CustomerPK = " & CustomerPK & " order by " & SortText & " " & SortDirection
            Else
                CardholderClause &= " and CustomerPK = " & CustomerPK & " order by " & IIf(ShowPOSTimeStamp, "POSTimeStamp", "LastUpdate") & " desc"
            End If
      
            sQueryBuilder.Append("select top 50 * from ( ")
            sQueryBuilder.Append("select Top 50 CustomerPK, ProgramID, AdjAmount, EarnedUnderROID, EarnedUnderCMOfferID, LastUpdate, LastServerID, LocationID, ")
            sQueryBuilder.Append("PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate, 0 as Inactive, POSTimeStamp ")
            sQueryBuilder.Append("from PointsHistoryView with (NoLock) ")
            sQueryBuilder.Append("where LocationID <> -8 and ProgramID = " & ProgramID & " and LastUpdate between '" & StartDate.ToString() & "' and '" & EndDate.ToString() & "' ")
            sQueryBuilder.Append(CardholderClause)
            sQueryBuilder.Append(" UNION ALL ")
            sQueryBuilder.Append("select Top 50 CustomerPK, ProgramID, AdjAmount, EarnedUnderROID, EarnedUnderCMOfferID, LastUpdate, LastServerID, LocationID, ")
            sQueryBuilder.Append("PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate, 1 as Inactive, POSTimeStamp ")
            sQueryBuilder.Append("from InactivePointsHistoryView with (NoLock) ")
            sQueryBuilder.Append("where LocationID <> -8 and ProgramID = " & ProgramID & " and LastUpdate between '" & StartDate.ToString() & "' and '" & EndDate.ToString() & "' ")
            sQueryBuilder.Append(CardholderClause)
            sQueryBuilder.Append(") as PH order by " & IIf(ShowPOSTimeStamp, "POSTimeStamp", "LastUpdate") & " " & SortDirection & ";")
      
        Else
            sQueryBuilder.Append("select Col2 as CustomerPK, Convert(int, Col1) as ProgramID, Convert(int,Col3) as AdjAmount, ")
            sQueryBuilder.Append("Convert(int,Col4) as EarnedUnderROID, 0 as EarnedUnderCMOfferID, getDate() as LastUpdate, Replayed, ReplayedDate, 0 as Inactive, POSTimeStamp ")
            sQueryBuilder.Append("from CPE_UploadTemp_PA with (NoLock) ")
            sQueryBuilder.Append("where Col1='" & ProgramID & "' ")
            sQueryBuilder.Append("and Col2='" & CustomerPK & "'")
            sQueryBuilder.Append(" UNION ALL ")
            sQueryBuilder.Append("select CustomerPK, ProgramID, Amount, EarnedUnderROID, 0 as EarnedUnderCMOfferID, ")
            sQueryBuilder.Append("getDate() as LastUpdate, Replayed, ReplayedDate, 0 as Inactive, POSTimeStamp from CPE_PointHistoryMovementTemp with (NoLock) ")
            sQueryBuilder.Append("where ProgramID='" & ProgramID & "' and CustomerPK='" & CustomerPK & "';")
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
      
            If (ShowHistory) Then
                Sendb("    <th class=""th-author"" scope=""col""><a id=""fromlink"" href=""point-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;CustomerPK=" & CustomerPK & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SetTo=" & Request.QueryString("SetTo") & "&amp;SortText=EarnedUnderROID&amp;SortDirection=" & SortDirection & """>" & Copient.PhraseLib.Lookup("term.from", LanguageID) & "</a>")
                If SortText = "EarnedUnderROID" Then
                    If SortDirection = "ASC" Then
                        Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                        Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                Else
                End If
                Send("</th>")
                Send("    <th class="""" scope=""col"" title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """>" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</th>")
                If ShowPOSTimeStamp Then
                    Sendb("    <th class=""th-datetime"" scope=""col""><a id=""POSTimeLink"" href=""point-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;CustomerPK=" & CustomerPK & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SetTo=" & Request.QueryString("SetTo") & "&amp;SortText=POSTimeStamp&amp;SortDirection=" & SortDirection & """>" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</a>")
                Else
                    Sendb("    <th class=""th-datetime"" scope=""col""><a id=""lastupdatelink"" href=""point-adjust-program.aspx?ProgramID=" & ProgramID & "&amp;CustomerPK=" & CustomerPK & "&amp;Opener=" & Request.QueryString("Opener") & "&amp;SetTo=" & Request.QueryString("SetTo") & "&amp;SortText=LastUpdate&amp;SortDirection=" & SortDirection & """>" & Copient.PhraseLib.Lookup("term.lastupdated", LanguageID) & "</a>")
                End If
                If SortText = "LastUpdate" OrElse SortText = "POSTimeStamp" Then
                    If SortDirection = "ASC" Then
                        Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                        Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                End If
                Send("</th>")
            Else
                Send("    <th class=""th-author"" scope=""col"">" & Copient.PhraseLib.Lookup("term.from", LanguageID) & "</th>")
                Send("    <th class="""" scope=""col"" title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """>" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</th>")
                If ShowPOSTimeStamp Then
                    Send("    <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.lastupdated", LanguageID) & "</th>")
                Else
                    Send("    <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
                End If
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
                        MyCommon.QueryStr = "select L.ExtLocationCode, L.EngineID from LocalServers LS with (NoLock) inner join Locations L with (NoLock) on L.LocationID = LS.LocationID where LS.LocalServerID=" & iLocalServerID & ";"
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
                    ' CM location
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
                Send("  <tr id=""hist" & j & """" & IIf(IsHousehold AndAlso CustomerPK = MyCommon.NZ(row.Item("CustomerPK"), 0), " class=""shaded""", ""))
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
                Send("    <td>" & IIf(MyCommon.NZ(row.Item("Replayed"), False), "<span title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """ style=""color:#dd0000;cursor:default;font-size:10px;font-weight:bold;"">R</span>", "") & "</td>")
        
                If ShowPOSTimeStamp Then
                    If IsDBNull(row.Item("POSTimeStamp")) Then
                        Send("    <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    Else
                        Send("    <td>" & Logix.ToShortDateTimeString(row.Item("POSTimeStamp"), MyCommon) & "</td>")
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
                    PresentedCustomerID = IIf(IsDBNull(row.Item("PresentedCustomerID")), "Unknown",MyCryptLib.SQL_StringDecrypt(row.Item("PresentedCustomerID").ToString()))
                    ResolvedCustomerID = IIf(IsDBNull(row.Item("ResolvedCustomerID")), "Unknown",MyCryptLib.SQL_StringDecrypt(row.Item("ResolvedCustomerID").ToString()))
                    If (ResolvedCustomerID = "0" OrElse ResolvedCustomerID = "Unknown") Then
                        MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK=" & MyCommon.NZ(row.Item("CustomerPK"), 0) & ";"
                        dt3 = MyCommon.LXS_Select
                        If dt3.Rows.Count > 0 Then
                            ResolvedCustomerID = IIf(IsDBNull(dt3.Rows(0).Item("ExtCardID")), "Unknown", MyCryptLib.SQL_StringDecrypt(dt3.Rows(0).Item("ExtCardID").ToString()))
                        End If
                    End If
                    Send("  <tr id=""histdetail" & j & """" & IIf(IsHousehold AndAlso CustomerPK = MyCommon.NZ(row.Item("CustomerPK"), 0), " class=""shaded""", "") & " style=""display:none;color:#777777;"">")
                    Send("    <td></td>")
                    Send("    <td colspan=""5"">")
                    Send("      " & Copient.PhraseLib.Lookup("term.presented", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & PresentedCustomerID & " &nbsp;|&nbsp; ")
                    Send("      " & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & ResolvedCustomerID & " &nbsp;|&nbsp; ")
          Send("      " & Copient.PhraseLib.Lookup("term.household", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & IIf(IsDBNull(row.Item("HHID")), "Unknown", row.Item("HHID").ToString()))
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
        MyCommon.Close_LogixRT()
        MyCommon.Close_LogixXS()
    End Sub
  
    Sub AdjustPoint(ByVal AdminUserID As String, ByVal ProgramID As Long, ByVal SessionID As String, ByVal ExternalProgram As Boolean, _
                    ByVal ExtHostTypeID As Integer, ByRef ER As Copient.ExternalRewards, ByVal ReasonID As Integer, ByVal ReasonText As String, _
                    ByVal ExtLocationID As String)
        Dim PointsAdj As Integer = 0
        Dim LocationID As Integer = -9
        Dim PointsAdjDecimal As Decimal = 0
        Dim CustomerPK As String = ""
        Dim i As Integer
        Dim MyPoints As New Copient.Points
        Dim ExtCardID As String = ""
        Dim dt, dt2 As System.Data.DataTable
        Dim ExternalID As String
        Dim ProgramName As String
        Dim PromoVarID As String
        Dim dtPointsTest As DataTable
        Dim dtPointsAdjTest As DataTable
        Dim PendingAdjTest, PointsTest As Long
        Dim dtNames As DataTable = Nothing
        Dim FirstName, LastName As String

        If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
    
        ProgramID = MyCommon.Extract_Val(Request.QueryString("ProgramID"))
        CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
        If Request.QueryString("SetToValue") IsNot Nothing Then
            PointsAdj = MyCommon.Extract_Val(Request.QueryString("SetToValue"))
            PointsAdjDecimal = MyCommon.Extract_Val(Request.QueryString("SetToValue"))
	  
	
            MyCommon.QueryStr = "select PKID, Amount from Points with (NoLock) " & _
                                  "where CustomerPK=" & CustomerPK & " and  ProgramID=" & ProgramID
            dtPointsTest = MyCommon.LXS_Select()
            If (dtPointsTest.Rows.Count > 0) Then
                PointsTest = MyCommon.NZ(dtPointsTest.Rows(0).Item("Amount"), 0)
            Else
                PendingAdjTest = 0
            End If
            MyCommon.QueryStr = "select sum(Convert(bigint, IsNull(Col3,0))) as Pending from CPE_UploadTemp_PA with (NoLock) " & _
                                "where Col1='" & ProgramID & "' and Col2='" & CustomerPK & "';"
            dtPointsAdjTest = MyCommon.LXS_Select
            If (dtPointsAdjTest.Rows.Count > 0) Then
                PendingAdjTest = MyCommon.Extract_Val(MyCommon.NZ(dtPointsAdjTest.Rows(0).Item("Pending"), 0))
            Else
                PendingAdjTest = 0
            End If
            PointsAdjDecimal = -1 * (PointsTest + PendingAdjTest - PointsAdjDecimal)
            PointsAdj = PointsAdjDecimal
	  
            MyCommon.QueryStr = "select FirstName, LastName from AdminUsers where AdminUserID=" & AdminUserID & ";"
            dtNames = MyCommon.LRT_Select
            If dtNames.Rows.Count > 0 Then
                FirstName = MyCommon.NZ(dtNames.Rows(0).Item("FirstName"), "").TOUpper()
                LastName = MyCommon.NZ(dtNames.Rows(0).Item("LastName"), "").ToUpper()
            End If
            ReasonText = "Adjustment of " & PointsAdjDecimal.ToString() & " points due to the action taken by user " & FirstName & " " & LastName & " to set the program balance to " & MyCommon.Extract_Val(Request.QueryString("SetToValue")).ToString()
        Else
            PointsAdj = MyCommon.Extract_Val(Request.QueryString("adjust"))
            PointsAdjDecimal = MyCommon.Extract_Val(Request.QueryString("adjust"))
        End If

        'System.IO.File.AppendAllText("D:\Copient\aaa.txt", "ADJUSTPOINT SUB" & vbCrLf)
        'System.IO.File.AppendAllText("D:\Copient\aaa.txt", " MyCommon.Extract_Val(Request.QueryString(""adjust"")) = " & MyCommon.Extract_Val(Request.QueryString("adjust")) & vbCrLf)
        'System.IO.File.AppendAllText("D:\Copient\aaa.txt", " PointsAdj = " & PointsAdj & vbCrLf)
        'System.IO.File.AppendAllText("D:\Copient\aaa.txt", " PointsAdjDecimal = " & PointsAdjDecimal & vbCrLf)
    
        MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by CardTypeID;"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
            ExtCardID = IIf(IsDBNull(dt.Rows(0).Item("ExtCardID")), "0",MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString()))
        End If
    
        If (MyCommon.Fetch_SystemOption(108) = "1") Then
            MyCommon.QueryStr = "select LocationID from Locations with (NoLock) " & _
                        "where ExtLocationCode='" & ExtLocationID & "';"
            dt2 = MyCommon.LRT_Select
            LocationID = MyCommon.NZ(dt2.Rows(0).Item("LocationID"), -9)
        End If

        If (CustomerPK > 0 AndAlso ProgramID > 0) Then
            If (ExternalProgram AndAlso ExtHostTypeID = 2) Then
                'it's an external program for which we must use the rewardUser subroutine in the ExternalRewards DLL
                If PointsAdjDecimal > 0 Then
                    ER.rewardUser(ProgramID.ToString, "LogixAdjustment", PointsAdjDecimal, ExtCardID, 0)
                    'System.IO.File.AppendAllText("D:\Copient\aaa.txt", "Adjustment performed for " & PointsAdjDecimal)
                    'System.IO.File.AppendAllText("D:\Copient\aaa.txt", vbCrLf & vbCrLf)
                End If
            ElseIf (ExternalProgram AndAlso ExtHostTypeID = 1) Then
                'it's an EME external program for which we must use the rewardUser subroutine in the ExternalRewards DLL
                If PointsAdjDecimal <> 0 Then
                    MyCommon.QueryStr = "select ExtHostProgramID, ProgramName, PromoVarID from PointsPrograms with (NoLock) where Deleted=0 and ProgramID=" & ProgramID & ";"
                    dt = MyCommon.LRT_Select()
                    If dt.Rows.Count > 0 Then
                        ExternalID = dt.Rows(0).Item("ExtHostProgramID")
                        ProgramName = dt.Rows(0).Item("ProgramName")
                        PromoVarID = MyCommon.NZ(dt.Rows(0).Item("PromoVarID"), "")
                        ER.updateExternalBalance(ExternalID, ProgramName, PointsAdjDecimal, ExtCardID, CustomerPK, PromoVarID, "0", "-9", "0", MyCommon)
                    End If
                End If
            Else 'we handle adjustments as normal
                If PointsAdj Then
                    MyPoints.AdjustPoints(AdminUserID, ProgramID, CustomerPK, PointsAdj, 0, 0, SessionID, 0, LocationID, ReasonID, ReasonText)
                End If
            End If
        End If
    End Sub
  
    Function IsValidAdjustment(ByVal OfferID As Long, ByVal CustomerPK As Long, ByRef infoMessage As String, ByRef WarningProgramID As Long, ByRef ProgramID As Long, _
                               ByRef AdjustAmt As Long, ByRef ExternalProgram As Boolean, ByRef ExtHostTypeID As Integer, ByRef DecimalValues As Boolean, _
                               ByVal ReasonID As Integer, ByVal ReasonText As String, ByVal ExtLocationID As String, ByVal LoggedUserID As Long) As Boolean
        Dim ValidAdj As Boolean = False
        Dim MyCAM As New Copient.CAM
        Dim MyPoints As New Copient.Points
        Dim PointsBal, WarningLimit, PendingAdj As Long
        Dim dt, dt2 As DataTable
        Dim ConfirmMessage As String = ""
        Dim customerLookup As CustomerLookup = New CustomerLookup()
        
        Dim IsCustomerLocked = customerLookup.IsCustomerLocked(CustomerPK)
        If IsCustomerLocked AndAlso AdjustAmt < 0 Then
            infoMessage = Copient.PhraseLib.Lookup("Error.pointsAdjustBurn", LanguageID, "Customer is Locked. Cannot burn points.")
            ValidAdj = False
        End If
        
        PointsBal = MyPoints.GetBalance(CustomerPK, ProgramID)
        WarningLimit = MyCAM.GetMaxAdjustment(OfferID, ProgramID)
    
        MyCommon.Open_LogixRT()
        MyCommon.Open_LogixXS()
    
        MyCommon.QueryStr = "select sum(Convert(bigint, IsNull(Col3,0))) as Pending from CPE_UploadTemp_PA with (NoLock) " & _
                            "where Col1='" & ProgramID & "' and Col2='" & CustomerPK & "';"
        dt = MyCommon.LXS_Select
        If (dt.Rows.Count > 0) Then
            PendingAdj = MyCommon.NZ(dt.Rows(0).Item("Pending"), 0)
        End If

        MyCommon.QueryStr = "select LocationID from Locations with (NoLock) " & _
                        "where ExtLocationCode='" & ExtLocationID & "';"
        dt2 = MyCommon.LRT_Select
    
        'System.IO.File.AppendAllText("D:\Copient\aaa.txt", "ISVALIDADJUSTMENT FUNCTION" & vbCrLf)
        'System.IO.File.AppendAllText("D:\Copient\aaa.txt", " AdjustAmt = " & AdjustAmt & vbCrLf)
    
        If ProgramID <= 0 Then
            infoMessage &= Copient.PhraseLib.Detokenize("customer-manual.InvalidProgramID", LanguageID, ProgramID)
        Else
            If (MyCommon.Fetch_SystemOption(108) = "1") Then
                If (ReasonID = 0 OrElse RTrim(ReasonText) = "") Then
                    infoMessage &= Copient.PhraseLib.Lookup("points-adjust.ReasonRequired", LanguageID)
                ElseIf (dt2.Rows.Count <= 0 AndAlso MyCommon.NZ(ExtLocationID, "") <> "-9") Then
                    infoMessage &= Copient.PhraseLib.Lookup("customer-edit.UnableToValidateLocation", LanguageID) & " Entered Location: " & ExtLocationID
                End If
            Else
                If (MyCommon.Fetch_SystemOption(224) = "1") Then
                    Dim PointsAdjustmentLimit As Integer = 0
                    Dim LimitPeriod As Integer = 1
                    MyCommon.QueryStr = "SELECT PointsAdjustmentLimit, LimitPeriod FROM AdminRoleAdjustmentLimits WITH (NoLock) INNER JOIN AdminUserRoles WITH (NoLock) ON AdminRoleAdjustmentLimits.RoleID = AdminUserRoles.RoleID " & _
                             " WHERE AdminUserRoles.AdminUserID = " & AdminUserID & " ;"
                    dt2 = MyCommon.LRT_Select
                    If (dt2.Rows.Count > 0) Then
                        PointsAdjustmentLimit = MyCommon.NZ(dt2.Rows(0).Item("PointsAdjustmentLimit"), 0)
                        LimitPeriod = MyCommon.NZ(dt2.Rows(0).Item("LimitPeriod"), 1)
                    End If
                    MyCommon.QueryStr = "SELECT SUM(CAST(ActivityValue AS Float)) AS CurrentDayAdjustmentsTotal FROM ActivityLog WITH (NoLock) WHERE ActivityTypeID=25 AND ActivitySubTypeID=12 AND " & _
                            "LinkID= " & CustomerPK & " AND LinkID2= " & ProgramID & " AND AdminID= " & AdminUserID & " AND DATEDIFF(DD,ActivityDate,GETDATE())<= ( " & LimitPeriod & " - 1) ;"
                    dt = MyCommon.LRT_Select
                    If (dt.Rows.Count > 0) Then
                        If ((MyCommon.NZ(dt.Rows(0).Item("CurrentDayAdjustmentsTotal"), 0) + AdjustAmt) > PointsAdjustmentLimit) Then
                            infoMessage &= Copient.PhraseLib.Lookup("points-adjust.AdjustmentExceedsUserRoleLimit", LanguageID)
                        End If
                    End If
                End If
		
                If ExternalProgram And ExtHostTypeID = 2 Then
                    'This is the new CRM (Excentus) case, which merely needs a check for negativity.
                    If AdjustAmt <= 0 Then
                        infoMessage &= Copient.PhraseLib.Lookup("points-adjust.InvalidAdjustment", LanguageID)
                    Else
                        ValidAdj = True
                    End If
                ElseIf ExternalProgram And ExtHostTypeID = 1 Then
                    ' EME
                    ValidAdj = True
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
</script>
<%
done:
    If (HHPK = 0) Then
        Send_BodyEnd("mainform", "adjust")
    Else
        Send_BodyEnd()
    End If
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    MyCommon = Nothing
    Logix = Nothing
%>
