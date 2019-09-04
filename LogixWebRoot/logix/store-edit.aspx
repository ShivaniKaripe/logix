<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
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

function launchScReport(locID) {
   openPopup('sanity-check-rpt.aspx?loc=' + locID);
}

  function ShoworHideTerminal(engine) {
    var terminal2 = document.getElementById("divTerminal2");
    var terminal9 = document.getElementById("divTerminal9");

    if (terminal2 != null) {
      if (engine.value == 2)
        terminal2.style.display = "";
      else
        terminal2.style.display = "none";
    }

    if (terminal9 != null) {
      if (engine.value == 9)
        terminal9.style.display = "";
      else
        terminal9.style.display = "none";
    }
  }

  function ActivateBrickAndMortar() {
    var engineElem = document.getElementById("EngineID");
    var brickElem = document.getElementById("Brick");
    var brickCodeElem = document.getElementById("BrickCode");
    var ueCodeElem = document.getElementById("ExtLocationCode");
    var ueNameElem = document.getElementById("LocationName");
    var uehCodeElem = document.getElementById("hExtLocationCode");
    var uehNameElem = document.getElementById("hLocationName");
    var ueAddressElem = document.getElementById("address");
    var ueTimezoneElem = document.getElementById("timezone");
    var ueContactElem = document.getElementById("contact");
    var ueTestingElem = document.getElementById("TestingDiv");

    if ((engineElem != null) && (brickElem != null)) {
      if (engineElem.value == "9") {
        brickElem.style.visibility = "visible";
        brickElem.style.display = "block";
      } else {
        brickElem.style.visibility = "hidden";
        brickElem.style.display = "none";
        brickCodeElem.value = "0";
        ueCodeElem.disabled = false;
        ueNameElem.disabled = false;
        ueCodeElem.value = "";
        ueNameElem.value = "";
        ueTestingElem.style.visibility = "visible";
        ueTestingElem.style.display = "block";
        ueAddressElem.style.visibility = "visible";
        ueAddressElem.style.display = "block";
        ueTimezoneElem.style.visibility = "visible";
        ueTimezoneElem.style.display = "block";
        ueContactElem.style.visibility = "visible";
        ueContactElem.style.display = "block";
      }
    }
  }

  function SetExtCode() {
    var brickCodeElem = document.getElementById("BrickCode");
    var ueCodeElem = document.getElementById("ExtLocationCode");
    var ueNameElem = document.getElementById("LocationName");
    var uehCodeElem = document.getElementById("hExtLocationCode");
    var uehNameElem = document.getElementById("hLocationName");
    var ueAddressElem = document.getElementById("address");
    var ueTimezoneElem = document.getElementById("timezone");
    var ueContactElem = document.getElementById("contact");
    var ueTestingElem = document.getElementById("TestingDiv");
    var ueLocalServer = document.getElementById("localserver");

    if ((brickCodeElem != null) && (ueCodeElem != null) && (ueNameElem != null)) {
      if (brickCodeElem.value != "0") {
        ueCodeElem.disabled = true;
        ueNameElem.disabled = true;
        var code_name1 = brickCodeElem.value.toString();
        var code_name2 = code_name1.split("|");
        ueCodeElem.value = "UE" + code_name2[0];
        ueNameElem.value = code_name2[1] + " (UE)";
        ueContactElem.style.visibility = "hidden";
        ueContactElem.style.display = "none";
        ueTimezoneElem.style.visibility = "hidden";
        ueTimezoneElem.style.display = "none";
        ueAddressElem.style.visibility = "hidden";
        ueAddressElem.style.display = "none";
        ueTestingElem.style.visibility = "hidden";
        ueTestingElem.style.display = "none";
        ueLocalServer.style.visibility = "hidden";
        ueLocalServer.style.display = "none";
      } else {
        ueCodeElem.disabled = false;
        ueNameElem.disabled = false;
        ueCodeElem.value = "";
        ueNameElem.value = "";
        ueTestingElem.style.visibility = "visible";
        ueTestingElem.style.display = "block";
        ueAddressElem.style.visibility = "visible";
        ueAddressElem.style.display = "block";
        ueTimezoneElem.style.visibility = "visible";
        ueTimezoneElem.style.display = "block";
        ueContactElem.style.visibility = "visible";
        ueContactElem.style.display = "block";
        ueLocalServer.style.visibility = "visible";
        ueLocalServer.style.display = "block";
      }
      if ((uehCodeElem != null) && (uehNameElem != null)) {
        uehCodeElem.value = ueCodeElem.value;
        uehNameElem.value = ueNameElem.value;
      }
    }
  }

</script>
<%
    ' *****************************************************************************
    ' * FILENAME: store-edit.aspx
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

    Dim AdminUserID As Long = 0
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim LocationID As Long = 0
    Dim LocationName As String = ""
    Dim ExtLocationCode As String = ""
    Dim Description As String = ""
    Dim ContactName As String = ""
    Dim PhoneNumber As String = ""
    Dim Address1 As String = ""
    Dim Address2 As String = ""
    Dim City As String = ""
    Dim EngineType As Integer = -1
    Dim EngineName As String = ""
    Dim TestingLocation As Boolean = False
    Dim OldTestingLocation As Boolean = False
    Dim State As String = ""
    Dim StateType As String = ""
    Dim Zip As String = ""
    Dim CountryID As Integer = 1
    Dim CountryName As String = ""
    Dim GMapString As String = ""
    Dim LastUpdate As String = ""
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim bSave As Boolean
    Dim bDelete As Boolean
    Dim bCreate As Boolean
    Dim dtGroups As DataTable = Nothing
    Dim dtOffers As DataTable = Nothing
    Dim sQuery As String
    Dim longDate As New DateTime
    Dim longDateString As String
    Dim bGenerateIPL As Boolean
    Dim OptionID As Integer
    Dim LocationOptionValue As String = ""
    Dim tempstr As String = ""
    Dim AllowSave As Boolean = True
    Dim ShowActionButton As Boolean = False
    Dim infoMessage As String = ""
    Dim informationMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannersEnabled As Boolean = False
    Dim BannerID As Integer = 0
    Dim BannerCt As Integer = 0
    Dim BannerName As String = ""
    Dim HasAllBannersAccess As Boolean = False
    Dim LocationTypeID As Integer = 1
    Dim TimeZone As String = ""
    Dim CurrencyID As Integer = 0
    Dim UOMSetID As Integer = 0
    Dim DefaultLanguageID As Integer = 0
    Dim LocLangID As Integer
    Dim LangName As String

    Dim ImageFetchURL As String = ""
    Dim IncentiveFetchURL As String = ""
    Dim PhoneHomeIPOverride As String = ""

    Dim OfflineFTPUser As String = ""
    Dim OfflineFTPPass As String = ""
    Dim OfflineFTPPath As String = ""
    Dim OfflineFTPIP As String = ""
    Dim SanityCheckPassed As Boolean = False
    Dim DefaultEngine As Integer = 0
    Dim TerminalSetIDUE As Integer = 0
    Dim TerminalSetIDCPE As Integer = 0
    Dim OperateAtEnterprise As Boolean
    Dim TempServer As System.Data.DataTable

    Dim bCmToUeEnabled As Boolean = False

    Dim bCmInstalled As Boolean = False
    Dim bUeInstalled As Boolean = False
    Dim bCpeInstalled As Boolean = False
    Dim bDisplayTimezone As Boolean = False
    Dim BrickAndMortarLocationId As Long = 0
    Dim bDeployOffers As Boolean

    ' hidden variables
    Dim hExtLocationCode As String = ""
    Dim hLocationName As String = ""

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "store-edit.aspx"
    ' Open database connection
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    OperateAtEnterprise = ((MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso (MyCommon.Fetch_CPE_SystemOption(91) = "1")) OrElse (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso (MyCommon.Fetch_UE_SystemOption(91) = "1")))

    Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)

    If (Request.QueryString("LocationTypeID") <> "") Then
        LocationTypeID = MyCommon.Extract_Val(Request.QueryString("LocationTypeID"))
    Else
        LocationTypeID = MyCommon.Extract_Val(Request.Form("LocationTypeID"))
    End If

    Try
        ' fill in if it was a get method
        If Request.RequestType = "GET" Then
            LocationID = Request.QueryString("LocationID")
            If Logix.TrimAll(Request.QueryString("hExtLocationCode")) = "" Then
                ExtLocationCode = Logix.TrimAll(Request.QueryString("ExtLocationCode"))
            Else
                ExtLocationCode = Logix.TrimAll(Request.QueryString("hExtLocationCode"))
            End If
            If Logix.TrimAll(Request.QueryString("hLocationName")) = "" Then
                LocationName = Logix.TrimAll(Request.QueryString("LocationName"))
            Else
                LocationName = Logix.TrimAll(Request.QueryString("hLocationName"))
            End If
            Description = Logix.TrimAll(Request.QueryString("Description"))
            ContactName = Logix.TrimAll(Request.QueryString("ContactName"))
            PhoneNumber = Logix.TrimAll(Request.QueryString("PhoneNumber"))
            Address1 = Logix.TrimAll(Request.QueryString("Address1"))
            Address2 = Logix.TrimAll(Request.QueryString("Address2"))
            City = Logix.TrimAll(Request.QueryString("City"))
            State = Logix.TrimAll(Request.QueryString("State"))
            Zip = Logix.TrimAll(Request.QueryString("Zip"))
            If LocationTypeID = 2 Then
                CountryID = MyCommon.Fetch_SystemOption(65)
            Else
                CountryID = Request.QueryString("Country")
            End If
            EngineType = MyCommon.MakeInt(Request.QueryString("EngineID"), -1)
            BannerID = MyCommon.Extract_Val(Request.QueryString("banner"))
            If (Request.QueryString("TestingLocation") = "on") Then
                TestingLocation = True
            Else
                TestingLocation = False
            End If
            If (Request.QueryString("OldTestingLocation") = "1") Then
                OldTestingLocation = True
            Else
                OldTestingLocation = False
            End If
            If (Request.QueryString("TimeZone") <> "") Then
                TimeZone = Request.QueryString("TimeZone")
            End If
            If Request.QueryString("save") = "" Then
                bSave = False
            Else
                bSave = True
            End If
            If Request.QueryString("delete") = "" Then
                bDelete = False
            Else
                bDelete = True
            End If
            If Request.QueryString("mode") = "" Then
                bCreate = False
            Else
                bCreate = True
            End If
            bGenerateIPL = (Request.QueryString("generateIPL") <> "")
            'LocLanguageID = MyCommon.Extract_Val(Request.QueryString("loclanguageid"))
            CurrencyID = MyCommon.Extract_Val(Request.QueryString("currencyid"))
            UOMSetID = MyCommon.Extract_Val(Request.QueryString("uomsetid"))
            bDeployOffers = (Request.QueryString("deploy") <> "")
        Else
            LocationID = Request.Form("LocationID")
            If LocationID = 0 Then
                LocationID = MyCommon.Extract_Val(Request.QueryString("LocationID"))
            End If
            If Logix.TrimAll(Request.Form("hExtLocationCode")) = "" Then
                ExtLocationCode = Logix.TrimAll(Request.Form("ExtLocationCode"))
                BrickAndMortarLocationId = 0
            Else
                ExtLocationCode = Logix.TrimAll(Request.Form("hExtLocationCode"))
                MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where ExtLocationCode='" & ExtLocationCode.Substring(2) & "';"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    BrickAndMortarLocationId = MyCommon.NZ(rst.Rows(0).Item("LocationID"), 0)
                End If
            End If
            If Logix.TrimAll(Request.Form("hLocationName")) = "" Then
                LocationName = Logix.TrimAll(Request.Form("LocationName"))
            Else
                LocationName = Logix.TrimAll(Request.Form("hLocationName"))
            End If
            Description = Logix.TrimAll(Request.Form("Description"))
            ContactName = Logix.TrimAll(Request.Form("ContactName"))
            PhoneNumber = Logix.TrimAll(Request.Form("PhoneNumber"))
            Address1 = Logix.TrimAll(Request.Form("Address1"))
            Address2 = Logix.TrimAll(Request.Form("Address2"))
            City = Logix.TrimAll(Request.Form("City"))
            State = Logix.TrimAll(Request.Form("State"))
            Zip = Logix.TrimAll(Request.Form("Zip"))
            If LocationTypeID = 2 Then
                CountryID = MyCommon.Fetch_SystemOption(65)
            Else
                CountryID = Request.Form("Country")
            End If
            EngineType = MyCommon.MakeInt(Request.Form("EngineID"), -1)
            BannerID = MyCommon.Extract_Val(Request.Form("banner"))

            TerminalSetIDUE = MyCommon.MakeInt(Request.Form("terminalset9"))
            TerminalSetIDCPE = MyCommon.MakeInt(Request.Form("terminalset2"))

            TestingLocation = (Request.Form("TestingLocation") = "on")
            OldTestingLocation = (Request.Form("OldTestingLocation") = "1")

            If (Request.Form("TimeZone") <> "") Then
                TimeZone = Request.Form("TimeZone")
            End If
            If Request.Form("save") = "" Then
                bSave = False
            Else
                bSave = True
            End If
            If Request.Form("delete") = "" Then
                bDelete = False
            Else
                bDelete = True
            End If
            If Request.Form("mode") = "" Then
                bCreate = False
            Else
                bCreate = True
            End If
            bGenerateIPL = (Request.Form("generateIPL") <> "")
            bDeployOffers = (Request.Form("deploy") <> "")

            'LocLanguageID = MyCommon.Extract_Val(Request.Form("loclanguageid"))
            CurrencyID = MyCommon.Extract_Val(Request.Form("currencyid"))
            UOMSetID = MyCommon.Extract_Val(Request.Form("uomsetid"))
            If CurrencyID = 0 Then
                MyCommon.QueryStr = "select case isnumeric(OptionValue) when 1 then cast(OptionValue as int) else 1 end AS DefaultCurrencyID from UE_SystemOptions where OptionID=137;"
                rst = MyCommon.LRT_Select()
                If (rst.Rows.Count = 1) Then
                    For Each row In rst.Rows
                        CurrencyID = row.Item("DefaultCurrencyID")
                    Next
                End If
            End If
        End If

        BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
        bCmToUeEnabled = (MyCommon.Fetch_CM_SystemOption(135) = "1")

        If Request.QueryString("LocationTypeID") <> "" Then
            LocationTypeID = MyCommon.Extract_Val(Request.QueryString("LocationTypeID"))
        ElseIf Request.Form("LocationTypeID") <> "" Then
            LocationTypeID = MyCommon.Extract_Val(Request.Form("LocationTypeID"))
        Else
            MyCommon.QueryStr = "select LocationTypeID from Locations with (NoLock) where LocationID=" & LocationID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                LocationTypeID = MyCommon.NZ(rst.Rows(0).Item("LocationTypeID"), 1)
            End If
        End If
        If LocationTypeID = 2 Then
            Send_HeadBegin("term.server", , LocationID)
        Else
            Send_HeadBegin("term.store", , LocationID)
        End If
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 7)
        If LocationTypeID = 2 Then
            Send_Subtabs(Logix, 73, 4, , LocationID)
        Else
            Send_Subtabs(Logix, 72, 4, , LocationID)
        End If

        If (Logix.UserRoles.AccessStores = False) Then
            Send_Denied(1, "perm.admin-storesaccess")
            GoTo done
        End If
        If (BannersEnabled AndAlso LocationID > 0) Then
            ' check if the user is allowed to view this bannered location group
            MyCommon.QueryStr = "select BannerID from Locations LOC with (NoLock) " & _
                                "where LocationID = " & LocationID & " and (BannerID is Null or BannerID =0 " & _
                                "or BannerID in (select BannerID from AdminUserBanners where AdminUserID=" & AdminUserID & ") " & _
                                "      or EXISTS(select BAN.BannerID from Banners BAN with (NoLock) " & _
                                "                inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                                "                inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                "                where BE.EngineID = " & EngineType & " and AUB.AdminUserID=" & AdminUserID & " and BAN.AllBanners=1 and BAN.Deleted=0) " & _
                                ");"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count = 0) Then
                Send_Denied(1, "banners.access-denied-offer")
                Send_BodyEnd()
                GoTo done
            End If
        ElseIf (BannersEnabled AndAlso LocationID = 0) Then
            ' determine if any banners exist
            MyCommon.QueryStr = "select distinct BannerID from Banners with (NoLock) where Deleted=0 and AllBanners=0;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count = 0) Then
                Send_Denied(1, "banners.store-no-banners")
                Send_BodyEnd()
                GoTo done
            End If

            ' determine if the user has permission to any banners
            MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                "where BAN.Deleted=0 and BAN.AllBanners=0 and AdminUserID = " & AdminUserID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count = 0) Then
                Send_Denied(1, "banners.store-no-permissions")
                Send_BodyEnd()
                GoTo done
            End If
        End If

        If (Request.Form("new") <> "") Then
            If (Request.Form("LocationTypeID") <> "") Then
                Response.Redirect("store-edit.aspx?LocationTypeID=" & Request.Form("LocationTypeID"))
            Else
                Response.Redirect("store-edit.aspx")
            End If
        End If

        ' Lets see if they clicked save or delete
        If bSave Then
            ' check that the name and external location code are unique
            AllowSave = True

            ' Check if TimeZone selected for CPE Stores 
            If (EngineType = 2 AndAlso LocationTypeID = 1 AndAlso String.IsNullOrEmpty(TimeZone)) Then
                AllowSave = False
                infoMessage = Copient.PhraseLib.Lookup("store-edit.badtime", LanguageID)
            End If

            ' check that external location code is unique
            If (AllowSave) Then
                MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where ExtLocationCode='" & MyCommon.Parse_Quotes(ExtLocationCode) & "' and deleted=0 and LocationTypeID= " & LocationTypeID & " and LocationID <> " & LocationID & ";"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    AllowSave = False
                    If LocationTypeID = 2 Then
                        infoMessage = Copient.PhraseLib.Lookup("server-edit.codeused", LanguageID)
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("store-edit.codeused", LanguageID)
                    End If
                End If
                Dim BannerEngineType As Integer = 0
                If (BannersEnabled) Then
                    MyCommon.QueryStr = "select EngineID from BannerEngines where BannerID = " & BannerID
                    rst = MyCommon.LRT_Select

                    If rst.Rows.Count > 0 Then
                        BannerEngineType = rst.Rows(0)(0)
                    End If
                End If

                If ((EngineType = 9 Or (BannersEnabled AndAlso BannerEngineType = 9)) AndAlso LocationTypeID = 2) Then
                    Dim maxServers As Integer = Integer.Parse(MyCommon.Fetch_UE_SystemOption(174))
                    Dim totalServers As Integer
                    MyCommon.QueryStr = "select COUNT(*) As TotalServers from Locations where LocationTypeID = 2 and EngineID = 9 and Deleted = 0;"
                    rst = MyCommon.LRT_Select
                    totalServers = rst.Rows(0)(0)

                    If (totalServers >= maxServers And LocationID = 0) Then
                        AllowSave = False
                        infoMessage = Copient.PhraseLib.Lookup("maxallowedservers", LanguageID) & maxServers
                    End If
                End If
            End If
            If (AllowSave) Then
                ' check that location name is unique
                MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where LocationName='" & MyCommon.Parse_Quotes(LocationName) & "' and deleted=0 and LocationTypeID=" & LocationTypeID & " and LocationID <> " & LocationID & ";"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    AllowSave = False
                    If LocationTypeID = 2 Then
                        infoMessage = Copient.PhraseLib.Lookup("server-edit.nameused", LanguageID)
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("store-edit.nameused", LanguageID)
                    End If
                End If
            End If

            If (AllowSave) Then
                If ((EngineType = 9) AndAlso (LocationTypeID = 1) AndAlso (MyCommon.Fetch_SystemOption(311) = "1") AndAlso (BrickAndMortarLocationId = 0)) Then
                    AllowSave = False
                    infoMessage = Copient.PhraseLib.Lookup("store-edit.BrickAndMortarRequired", LanguageID)
                End If
            End If

            If (AllowSave) Then
                If BrickAndMortarLocationId > 0 Then
                    MyCommon.QueryStr = "select TimeZone from Locations with (NoLock) where Deleted=0 and LocationID=" & BrickAndMortarLocationId & ";"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        TimeZone = MyCommon.NZ(rst.Rows(0).Item("TimeZone"), "")
                    End If
                End If

                If (TimeZone = "" AndAlso LocationTypeID = 1 AndAlso Not (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE))) Then
                    If BrickAndMortarLocationId = 0 Then
                        AllowSave = False
                        infoMessage = Copient.PhraseLib.Lookup("store-edit.badtime", LanguageID)
                    End If
                End If
            End If

            If (AllowSave) Then
                ' overwrite the engine id if this store is a member of a banner (if necessary)
                If (BannersEnabled AndAlso BannerID > 0) Then
                    MyCommon.QueryStr = "select EngineID from BannerEngines BE with (NoLock) where BannerID=" & BannerID
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        EngineType = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
                    End If
                End If

                ' They want to save something -- do stored procedure for saving group
                If LocationTypeID = 2 Then
                    ' Servers don't have addresses, but the CountryID field can't be null,
                    ' so we simply set it to be whatever the system default country is.
                    CountryID = MyCommon.Fetch_SystemOption(65)
                End If

                If (LocationID = 0) Then
                    MyCommon.QueryStr = "dbo.pt_Locations_Insert"
                    MyCommon.Open_LRTsp()
                    ExtLocationCode = Logix.TrimAll(ExtLocationCode)
                    MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = ExtLocationCode
                    LocationName = Logix.TrimAll(LocationName)
                    MyCommon.LRTsp.Parameters.Add("@LocationName", SqlDbType.NVarChar, 100).Value = LocationName
                    MyCommon.LRTsp.Parameters.Add("@LocationTypeID", SqlDbType.Int).Value = LocationTypeID
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 4000).Value = Left(Description, 1000)
                    MyCommon.LRTsp.Parameters.Add("@Address1", SqlDbType.NVarChar, 200).Value = Address1
                    MyCommon.LRTsp.Parameters.Add("@Address2", SqlDbType.NVarChar, 200).Value = Address2
                    MyCommon.LRTsp.Parameters.Add("@City", SqlDbType.NVarChar, 100).Value = City
                    MyCommon.LRTsp.Parameters.Add("@State", SqlDbType.NVarChar, 50).Value = State
                    MyCommon.LRTsp.Parameters.Add("@Zip", SqlDbType.NVarChar, 20).Value = Zip
                    MyCommon.LRTsp.Parameters.Add("@CountryID", SqlDbType.Int).Value = CountryID
                    MyCommon.LRTsp.Parameters.Add("@ContactName", SqlDbType.NVarChar, 100).Value = ContactName
                    MyCommon.LRTsp.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar, 20).Value = PhoneNumber
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
                    MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                    If (LocationTypeID = 1) Then
                        MyCommon.LRTsp.Parameters.Add("@TimeZone", SqlDbType.NVarChar, 20).Value = TimeZone
                    End If
                    MyCommon.LRTsp.Parameters.Add("@CurrencyID", SqlDbType.Int).Value = CurrencyID
                    MyCommon.LRTsp.Parameters.Add("@UOMSetID", SqlDbType.Int).Value = UOMSetID
                    MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    If (LocationName = "") OrElse (ExtLocationCode = "") Then
                        If LocationTypeID = 2 Then
                            infoMessage = Copient.PhraseLib.Lookup("server-edit.noname", LanguageID)
                        Else
                            infoMessage = Copient.PhraseLib.Lookup("store-edit.noname", LanguageID)
                        End If
                    ElseIf (LocationID.ToString <> "-1") Then
                        MyCommon.LRTsp.ExecuteNonQuery()
                        LocationID = MyCommon.NZ(MyCommon.LRTsp.Parameters("@LocationID").Value, -1)
                        If (LocationID.ToString <> "-1") Then
                            MyCommon.Activity_Log(10, LocationID, AdminUserID, Copient.PhraseLib.Lookup("history.store-create", LanguageID))
                            If LocationTypeID = 2 Then
                                'We have just created a new server type location.  We need to populate the OfferLocUpdate table for this new location so that an IPL will return offers
                                If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then
                                    MyCommon.QueryStr = "dbo.pa_ECPE_Populate_Offers"
                                ElseIf MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) Then
                                    MyCommon.QueryStr = "dbo.pa_EUE_Populate_Offers"
                                Else
                                    ' Should not happen
                                    MyCommon.QueryStr = "dbo.pa_ECPE_Populate_Offers"
                                End If
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                                MyCommon.LRTsp.ExecuteNonQuery()
                            Else
                                If BrickAndMortarLocationId > 0 Then
                                    ' Update UE store with Brick and Mortar store
                                    MyCommon.QueryStr = "update Locations with (RowLock) set BrickAndMortarLocationId=" & BrickAndMortarLocationId & " where LocationID=" & LocationID & ";"
                                    MyCommon.LRT_Execute()

                                    ' Create a UE Location Group for this specifc store
                                    MyCommon.QueryStr = "dbo.pt_LocationGroups_Insert"
                                    MyCommon.Open_LRTsp()
                                    MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = ExtLocationCode
                                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = LocationName
                                    MyCommon.LRTsp.Parameters.Add("@ExtGroupId", SqlDbType.NVarChar, 20).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@ExtSeqNum", SqlDbType.NVarChar, 20).Value = ""
                                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
                                    If (BannersEnabled AndAlso BannerID > 0) Then
                                        MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                                    End If
                                    MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                    MyCommon.LRTsp.ExecuteNonQuery()
                                    Dim lGroupId As Long
                                    lGroupId = MyCommon.LRTsp.Parameters("@LocationGroupId").Value
                                    MyCommon.Close_LRTsp()
                                    If lGroupId > 0 Then
                                        MyCommon.Activity_Log(11, lGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-create", LanguageID))

                                        ' Add this specifc UE store to the Location Group
                                        MyCommon.QueryStr = "update LocationGroups with (RowLock) set BrickAndMortarLocationId=" & BrickAndMortarLocationId & " where LocationGroupID=" & lGroupId & ";"
                                        MyCommon.LRT_Execute()

                                        MyCommon.QueryStr = "dbo.pt_LocGroupItems_Insert"
                                        MyCommon.Open_LRTsp()
                                        MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt, 8).Value = lGroupId
                                        MyCommon.LRTsp.Parameters.Add("@LocationId", SqlDbType.BigInt, 8).Value = LocationID
                                        MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                                        MyCommon.LRTsp.ExecuteNonQuery()
                                        Dim lPkId As Long
                                        lPkId = MyCommon.LRTsp.Parameters("@PkId").Value
                                        MyCommon.Close_LRTsp()
                                        If lPkId > 0 Then
                                            MyCommon.Activity_Log(11, lGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-add", LanguageID) & ": " & LocationID)
                                        Else
                                            infoMessage = "Failed to add UE Location to new UE Location Group."
                                        End If
                                    Else
                                        infoMessage = "Failed to create new UE Location Group."
                                    End If
                                End If
                            End If
                        End If
                    End If
                    MyCommon.Close_LRTsp()
                Else
                    MyCommon.QueryStr = "dbo.pt_Locations_Update"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                    ExtLocationCode = Logix.TrimAll(ExtLocationCode)
                    MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = ExtLocationCode
                    LocationName = Logix.TrimAll(LocationName)
                    MyCommon.LRTsp.Parameters.Add("@LocationName", SqlDbType.NVarChar, 100).Value = LocationName
                    MyCommon.LRTsp.Parameters.Add("@LocationTypeID", SqlDbType.Int).Value = LocationTypeID
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 4000).Value = Left(Description, 1000)
                    MyCommon.LRTsp.Parameters.Add("@Address1", SqlDbType.NVarChar, 200).Value = Address1
                    MyCommon.LRTsp.Parameters.Add("@Address2", SqlDbType.NVarChar, 200).Value = Address2
                    MyCommon.LRTsp.Parameters.Add("@City", SqlDbType.NVarChar, 100).Value = City
                    MyCommon.LRTsp.Parameters.Add("@State", SqlDbType.NVarChar, 50).Value = State
                    MyCommon.LRTsp.Parameters.Add("@Zip", SqlDbType.NVarChar, 20).Value = Zip
                    MyCommon.LRTsp.Parameters.Add("@CountryID", SqlDbType.Int).Value = CountryID
                    MyCommon.LRTsp.Parameters.Add("@ContactName", SqlDbType.NVarChar, 100).Value = ContactName
                    MyCommon.LRTsp.Parameters.Add("@PhoneNumber", SqlDbType.NVarChar, 20).Value = PhoneNumber
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineType
                    MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                    If (LocationTypeID = 1) Then
                        MyCommon.LRTsp.Parameters.Add("@TimeZone", SqlDbType.NVarChar, 20).Value = TimeZone
                    End If
                    MyCommon.LRTsp.Parameters.Add("@CurrencyID", SqlDbType.Int).Value = CurrencyID
                    MyCommon.LRTsp.Parameters.Add("@UOMSetID", SqlDbType.Int).Value = UOMSetID

                    If (LocationName = "") OrElse (ExtLocationCode = "") Then
                        If LocationTypeID = 2 Then
                            infoMessage = Copient.PhraseLib.Lookup("server-edit.noname", LanguageID)
                        Else
                            infoMessage = Copient.PhraseLib.Lookup("store-edit.noname", LanguageID)
                        End If
                    ElseIf (LocationID.ToString <> "-1") Then
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Activity_Log(10, LocationID, AdminUserID, Copient.PhraseLib.Lookup("history.store-edit", LanguageID))
                        informationMessage = Copient.PhraseLib.Lookup("terminal-set.saved", LanguageID)
                        ' update the location languages
                        If LocationID > 0 AndAlso (EngineType = Copient.CommonInc.InstalledEngines.CPE OrElse EngineType = Copient.CommonInc.InstalledEngines.UE) Then
                            UpdateLocationLanguages(MyCommon, LocationID)
                        End If

                    End If
                    MyCommon.Close_LRTsp()
                End If

                If (LocationID = -1) And (LocationName <> "") And (ExtLocationCode <> "") Then
                    If LocationTypeID = 2 Then
                        infoMessage = Copient.PhraseLib.Lookup("server-edit.codeused", LanguageID)
                    Else
                        infoMessage = Copient.PhraseLib.Lookup("store-edit.codeused", LanguageID)
                    End If
                    LocationID = 0
                End If
                If (MyCommon.Fetch_SystemOption(88) = "1") AndAlso (TestingLocation <> OldTestingLocation) Then
                    ' force IPL
                    If (TestingLocation) Then
                        MyCommon.QueryStr = "update Locations with (RowLock) set TestingLocation=1, GenIPL=1 where LocationID=" & LocationID
                    Else
                        MyCommon.QueryStr = "update Locations with (RowLock) set TestingLocation=0, GenIPL=1 where LocationID=" & LocationID
                    End If
                Else
                    If (TestingLocation) Then
                        MyCommon.QueryStr = "update Locations with (RowLock) set TestingLocation=1 where LocationID=" & LocationID
                    Else
                        MyCommon.QueryStr = "update Locations with (RowLock) set TestingLocation=0 where LocationID=" & LocationID
                    End If
                End If
                MyCommon.LRT_Execute()

                ' update the local server settings
                If (EngineType = 2 AndAlso LocationID > 0) Then
                    ' update the site specific settings
                    MyCommon.QueryStr = "select Distinct OptionID from SiteSpecificOptions with (NoLock) order by OptionID;"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        For Each row In rst.Rows
                            tempstr = Request.Form("oid" & row.Item("OptionID"))
                            If (tempstr <> "") Then
                                MyCommon.QueryStr = "dbo.pt_LocationOptions_Update"
                                MyCommon.Open_LRTsp()
                                MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
                                MyCommon.LRTsp.Parameters.Add("@OptionID", SqlDbType.Int).Value = MyCommon.NZ(row.Item("OptionID"), -1)
                                MyCommon.LRTsp.Parameters.Add("@OptionValue", SqlDbType.NVarChar, 255).Value = tempstr
                                MyCommon.LRTsp.ExecuteNonQuery()
                                MyCommon.Close_LRTsp()
                            End If
                        Next
                    End If
                End If

                If ((EngineType = 2) Or (EngineType = 9)) Then
                    ImageFetchURL = Trim(Request.Form("imagefetchurl"))
                    If Not (Right(ImageFetchURL, 1) = "/") Then
                        ImageFetchURL = ImageFetchURL & "/"
                    End If
                    ImageFetchURL = MyCommon.Parse_Quotes(ImageFetchURL)

                    'incentive fetch url will be taken from system option 43 for UE
                    If EngineType <> 9 Then
                        IncentiveFetchURL = Trim(Request.Form("incentivefetchurl"))
                        If Not (Right(IncentiveFetchURL, 1) = "/") Then
                            IncentiveFetchURL = IncentiveFetchURL & "/"
                        End If
                        IncentiveFetchURL = MyCommon.Parse_Quotes(IncentiveFetchURL)
                    End If

                    PhoneHomeIPOverride = Trim(Request.Form("PhoneHomeIPOverride"))
                    PhoneHomeIPOverride = MyCommon.Parse_Quotes(PhoneHomeIPOverride)

                    OfflineFTPUser = Trim(Request.Form("FTPUser"))
                    OfflineFTPPass = Trim(Request.Form("FTPPass"))
                    OfflineFTPPath = Trim(Request.Form("FTPPath"))
                    OfflineFTPIP = Trim(Request.Form("FTPIP"))

                    MyCommon.QueryStr = "Update LocalServers with (RowLock) set ImageFetchURL='" & ImageFetchURL & "', IncentiveFetchURL='" & IncentiveFetchURL & "', PhoneHomeIPOverride='" & PhoneHomeIPOverride & "', " &
                                        "OfflineFTPUser='" & OfflineFTPUser & "', OfflineFTPIP='" & OfflineFTPIP & "', OfflineFTPPass='" & OfflineFTPPass & "', OfflineFTPPath='" & OfflineFTPPath & "' " &
                                        "where LocationID=" & LocationID & ";"
                    MyCommon.LRT_Execute()

                    If LocationTypeID = 1 Then
                        MyCommon.QueryStr = "dbo.pt_LocationTerminals_Update"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                        MyCommon.LRTsp.Parameters.Add("@TerminalSetID", SqlDbType.Int).Value = MyCommon.Extract_Val(GetCgiValue("terminalset" + EngineType.ToString()))
                        MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                    End If
                End If

                If infoMessage = "" And informationMessage = "" Then
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "store-edit.aspx?LocationID=" & LocationID)
                End If
            End If

        ElseIf bDelete Then
            'sQuery = "select distinct d.OfferId,d.Name from LocGroupItems a,LocationGroups b,OfferLocations c,Offers d with (NoLock)"
            'sQuery += " where d.deleted = 0 and d.OfferId = c.OfferId and c.Deleted = 0 and c.Excluded = 0 and c.LocationGroupId = b.LocationGroupId"
            'sQuery += " and b.LocationGroupId = a.LocationGroupId and a.Deleted = 0 and a.LocationID = " & LocationID
            If (EngineType = Copient.CommonInc.InstalledEngines.CPE OrElse EngineType = Copient.CommonInc.InstalledEngines.UE) Then
                sQuery = "SELECT DISTINCT I.IncentiveID, LGI.LocationID " &
                          "FROM dbo.CPE_Incentives AS I WITH (NoLock) INNER JOIN " &
                          "dbo.OfferLocations AS OL WITH (NoLock) ON I.IncentiveID = OL.OfferID AND I.Deleted = 0 AND OL.Deleted = 0 AND OL.Excluded = 0 AND " &
                          "I.IsTemplate = 0 INNER JOIN " &
                          "dbo.LocGroupItems AS LGI WITH (NoLock) ON LGI.LocationGroupID = OL.LocationGroupID AND LGI.Deleted = 0 INNER JOIN " &
                          "dbo.Locations AS L WITH (NoLock) ON LGI.LocationID = L.LocationID AND L.EngineID = " & EngineType & " " &
                          "Where L.LocationID = " & LocationID & " and getdate() between I.StartDate and dateadd(d,1,I.EndDate);"
            Else
                sQuery = "SELECT DISTINCT O.OfferID, LGI.LocationID " &
                          "FROM dbo.Offers AS O WITH (NoLock) INNER JOIN " &
                          "dbo.OfferLocations AS OL WITH (NoLock) ON O.OfferID = OL.OfferID AND O.Deleted = 0 AND OL.Deleted = 0 AND OL.Excluded = 0 INNER JOIN " &
                          "dbo.LocGroupItems AS LGI WITH (NoLock) ON LGI.LocationGroupID = OL.LocationGroupID AND LGI.Deleted = 0 " &
                          "Where LocationID = " & LocationID & " and getdate() between O.ProdStartDate and dateadd(d,1,O.ProdEndDate);"
            End If
            MyCommon.QueryStr = sQuery
            dtOffers = MyCommon.LRT_Select
            If (dtOffers.Rows.Count > 0) Then
                If LocationTypeID = 2 Then
                    infoMessage = Copient.PhraseLib.Lookup("server-edit.inuse", LanguageID)
                Else
                    infoMessage = Copient.PhraseLib.Lookup("store-edit.inuse", LanguageID)
                End If
            Else
                MyCommon.QueryStr = "dbo.pt_Locations_Delete"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Activity_Log(10, LocationID, AdminUserID, Copient.PhraseLib.Lookup("history.store-delete", LanguageID))
                MyCommon.Close_LRTsp()

                ' remove the location terminal set records for UE stores.
                If EngineType = Copient.CommonInc.InstalledEngines.UE Then
                    MyCommon.QueryStr = "delete from LocationTerminals with (RowLock) where LocationID=" & LocationID
                    MyCommon.LRT_Execute()
                End If

                If EngineType = Copient.CommonInc.InstalledEngines.UE Then
                    MyCommon.QueryStr = "update LocalServers set LocationID = NULL where LocationID =" & LocationID
                    MyCommon.LRT_Execute()
                End If

                ' remove the location languages records
                MyCommon.QueryStr = "update LocationLanguages with (RowLock) set Deleted=1, LastUpdate=GETDATE() where LocationID=" & LocationID
                MyCommon.LRT_Execute()

                If BrickAndMortarLocationId > 0 Then
                    Dim lGroupId As Long = 0
                    MyCommon.QueryStr = "select LocationGroupID from LocationGroups with (NoLock) where Deleted=0 and BrickAndMortarLocationId=" & BrickAndMortarLocationId & ";"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        lGroupId = MyCommon.NZ(rst.Rows(0).Item("LocationGroupID"), 0)
                    End If
                    If lGroupId > 0 Then
                        MyCommon.QueryStr = "dbo.pt_LocationGroups_Delete"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt).Value = lGroupId
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                        MyCommon.Activity_Log(11, lGroupId, AdminUserID, Copient.PhraseLib.Lookup("history.lgroup-delete", LanguageID))
                    End If
                End If

                'To Get how many Server type locations exists in system 
                MyCommon.QueryStr = "SELECT LocationID from dbo.Locations WHERE LocationTypeID = 2 AND Deleted = 0;"
                TempServer = MyCommon.LRT_Select

                Response.Status = "301 Moved Permanently"
                'If Operate at Enterprise is set to False and no Server exists in system , then server tab should not be visible else in all cases server tab will be available.
                If LocationTypeID = 2 AndAlso (OperateAtEnterprise OrElse TempServer.Rows.Count > 0) Then
                    Response.AddHeader("Location", "store-list.aspx?LocationTypeID=2")
                Else
                    Response.AddHeader("Location", "store-list.aspx?LocationTypeID=1")
                End If
                LocationID = 0
                ExtLocationCode = ""
                Description = ""
                LocationName = ""
                ContactName = ""
                PhoneNumber = ""
                Address1 = ""
                Address2 = ""
                City = ""
                State = ""
                Zip = ""
                CountryID = 1
                'LocLanguageID = 1
                UOMSetID = 0
                MyCommon.QueryStr = "select case isnumeric(OptionValue) when 1 then cast(OptionValue as int) else 1 end AS DefaultCurrencyID from UE_SystemOptions where OptionID=137;"
                rst = MyCommon.LRT_Select()
                If (rst.Rows.Count = 1) Then
                    For Each row In rst.Rows
                        CurrencyID = row.Item("DefaultCurrencyID")
                    Next
                Else
                    CurrencyID = 0
                End If
            End If
        ElseIf bGenerateIPL Then
            MyCommon.QueryStr = "update Locations with (RowLock) set GenIpl=1 where LocationID=" & LocationID
            MyCommon.LRT_Execute()
        ElseIf bDeployOffers Then
            MyCommon.QueryStr = "dbo.pa_CmToUeAgent_TranslateCmOffersForUeLocation"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
            MyCommon.LRTsp.ExecuteNonQuery()
            MyCommon.Close_LRTsp()
        End If

        LastUpdate = ""

        If Not bCreate Then
            ' no one clicked anything
            MyCommon.QueryStr = "select L.*, PE.Description as EngineName, PE.PhraseID as EnginePhraseID from Locations L with (NoLock) " & _
                                "left join PromoEngines PE with (NoLock) on PE.EngineID=L.EngineID " & _
                                "where Deleted=0 and LocationID=" & LocationID & ";"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
                For Each row In rst.Rows
                    If (ExtLocationCode = "") Then
                        If Not row.Item("ExtLocationCode").Equals(System.DBNull.Value) Then
                            ExtLocationCode = row.Item("ExtLocationCode")
                        End If
                    End If
                    If Not row.Item("TestingLocation").Equals(System.DBNull.Value) Then
                        TestingLocation = row.Item("TestingLocation")
                        OldTestingLocation = TestingLocation
                    End If
                    If (LocationName = "") Then
                        If Not row.Item("LocationName").Equals(System.DBNull.Value) Then
                            LocationName = row.Item("LocationName")
                        End If
                    End If
                    If (Description = "") Then
                        If Not row.Item("Description").Equals(System.DBNull.Value) Then
                            Description = Left(row.Item("Description"), 1000)
                        End If
                    End If
                    If (ContactName = "") Then
                        If Not row.Item("ContactName").Equals(System.DBNull.Value) Then
                            ContactName = row.Item("ContactName")
                        End If
                    End If
                    If (PhoneNumber = "") Then
                        If Not row.Item("PhoneNumber").Equals(System.DBNull.Value) Then
                            PhoneNumber = row.Item("PhoneNumber")
                        End If
                    End If
                    If (Address1 = "") Then
                        If Not row.Item("Address1").Equals(System.DBNull.Value) Then
                            Address1 = row.Item("Address1")
                        End If
                    End If
                    If (Address2 = "") Then
                        If Not row.Item("Address2").Equals(System.DBNull.Value) Then
                            Address2 = row.Item("Address2")
                        End If
                    End If
                    If (City = "") Then
                        If Not row.Item("City").Equals(System.DBNull.Value) Then
                            City = row.Item("City")
                        End If
                    End If
                    If (State = "") Then
                        If Not row.Item("State").Equals(System.DBNull.Value) Then
                            State = row.Item("State")
                        End If
                    End If
                    If (Zip = "") Then
                        If Not row.Item("Zip").Equals(System.DBNull.Value) Then
                            Zip = row.Item("Zip")
                        End If
                    End If
                    If (CountryID = 0) Then
                        If Not row.Item("CountryID").Equals(System.DBNull.Value) Then
                            CountryID = row.Item("CountryID")
                        End If
                    End If
                    If (EngineType = -1) Then
                        If Not row.Item("EngineID").Equals(System.DBNull.Value) Then
                            EngineType = row.Item("EngineID")
                        End If
                    End If
                    If (EngineName = "") Then
                        EngineName = Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("EngineName"), ""))
                    End If
                    If (LastUpdate = "") Then
                        If Not row.Item("LastUpdate").Equals(System.DBNull.Value) Then
                            LastUpdate = row.Item("LastUpdate")
                        End If
                    End If
                    If (TimeZone = "") Then
                        If Not row.Item("TimeZone").Equals(System.DBNull.Value) Then
                            TimeZone = row.Item("TimeZone")
                        End If
                    End If

                    If (CurrencyID = 0) Then
                        CurrencyID = MyCommon.NZ(row.Item("CurrencyID"), 0)
                    End If
                    If (UOMSetID = 0) Then
                        UOMSetID = MyCommon.NZ(row.Item("UOMSetID"), 0)
                    End If

                    If (BrickAndMortarLocationId = 0) Then
                        If Not row.Item("BrickAndMortarLocationID").Equals(System.DBNull.Value) Then
                            BrickAndMortarLocationId = row.Item("BrickAndMortarLocationID")
                            If BrickAndMortarLocationId > 0 Then
                                hExtLocationCode = ExtLocationCode
                                hLocationName = LocationName
                            End If
                        End If
                    End If

                    If (BannersEnabled) Then
                        BannerID = MyCommon.NZ(row.Item("BannerID"), 0)
                        MyCommon.QueryStr = "select Name from Banners with (NoLock) where BannerID=" & BannerID
                        rst2 = MyCommon.LRT_Select
                        If (rst2.Rows.Count > 0) Then
                            BannerName = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
                        End If
                        MyCommon.QueryStr = "select BAN.BannerID from Locations LOC with (NoLock) " & _
                                            "inner join BannerEngines BE with (NoLock) on BE.EngineId = LOC.EngineID and LOC.Deleted=0 " & _
                                            "inner join Banners BAN with (NoLock) on BAN.BannerID = BE.BannerID and BAN.Deleted=0 " & _
                                            "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                            "where LOC.LocationID = " & LocationID & " and AllBanners = 1 and AUB.AdminUserID=" & AdminUserID & ";"
                        rst = MyCommon.LRT_Select
                        HasAllBannersAccess = (rst.Rows.Count > 0)
                    End If
                Next
            ElseIf (Request.Form("new") <> "New") And (LocationID > 0) Then
                MyCommon.QueryStr = "select L.*, PE.Description as EngineName from Locations L with (NoLock) " & _
                                    "left join PromoEngines PE with (NoLock) on PE.EngineID=L.EngineID " & _
                                    "where LocationID=" & LocationID
                rst = MyCommon.LRT_Select()
                If (rst.Rows.Count > 0) Then
                    LocationName = MyCommon.NZ(rst.Rows(0).Item("LocationName"), "")
                End If
                Send("")
                Send("<div id=""intro"">")
                Sendb("  <h1 id=""title"">")
                If LocationTypeID = 2 Then
                    Sendb(Copient.PhraseLib.Lookup("term.server", LanguageID))
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID))
                End If
                Sendb(" " & ExtLocationCode & ": ")
                Send(MyCommon.TruncateString(LocationName, 40) & "</h1>")
                Send("</div>")
                Send("<div id=""main"">")
                Send("  <div id=""infobar"" class=""red-background"">")
                Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
                Send("  </div>")
                Send("</div>")
                GoTo done
            End If

            Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
            Dim conditionalQuery = String.Empty

            If (bEnableRestrictedAccessToUEOfferBuilder) Then
                conditionalQuery = GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "I")
            End If

            sQuery = "select distinct b.LocationGroupId,b.Name,b.AllLocations from LocGroupItems a with (NoLock),LocationGroups b with (NoLock)"
            sQuery += " where b.LocationGroupId = a.LocationGroupId and a.Deleted = 0 and a.LocationID = " & LocationID
            MyCommon.QueryStr = sQuery
            dtGroups = MyCommon.LRT_Select

            If EngineType = Copient.CommonInc.InstalledEngines.CPE Then
                sQuery = "select top 500 OI.EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name,buy.ExternalBuyerId as BuyerID " & _
                          "from CPE_IncentiveLocationsView ILV with (NoLock) " & _
                          "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = ILV.IncentiveID " & _
                          "inner join OfferIDs OI with (NoLock)  on I.IncentiveID = OI.OfferID " & _
                           "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                          "where I.Deleted=0 and I.IsTemplate=0 and LocationID = " & LocationID & _
                          " and getdate() between I.StartDate and dateadd(d,1,I.EndDate) " & _
                          "order by ILV.IncentiveID desc;"
            ElseIf EngineType = Copient.CommonInc.InstalledEngines.UE Then
                sQuery = "select top 500 OI.EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name,buy.ExternalBuyerId as BuyerID " & _
                          "from UE_IncentiveLocationsView ILV with (NoLock) " & _
                          "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = ILV.IncentiveID " & _
                          "inner join OfferIDs OI with (NoLock)  on I.IncentiveID = OI.OfferID " & _
                           "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                          "where I.Deleted=0 and I.IsTemplate=0 and LocationID = " & LocationID & _
                          " and getdate() between I.StartDate and dateadd(d,1,I.EndDate) "
                If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then sQuery &= conditionalQuery & " "
                sQuery &= "order by ILV.IncentiveID desc;"
            Else
                sQuery = "select top 500 O.EngineID, O.OfferID, O.Name,NULL as BuyerID " & _
                          "from CM_ST_OfferLocationsView OLV with (NoLock) " & _
                          "inner join CM_ST_Offers O with (NoLock) on O.OfferID = OLV.OfferID " & _
                          "where LocationID = " & LocationID & _
                          " and getdate() between O.ProdStartDate and dateadd(d,1,O.ProdEndDate) " & _
                          "order by OLV.OfferID desc;"
            End If
            MyCommon.QueryStr = sQuery
            dtOffers = MyCommon.LRT_Select
        End If
%>
<script type="text/javascript">
  function toggleDropdown() {
    if (document.getElementById("actionsmenu") != null) {
      bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
      if (bOpen) {
        document.getElementById("actionsmenu").style.visibility = 'visible';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▲';
              } else {
                document.getElementById("actionsmenu").style.visibility = 'hidden';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▼';
              }
            }
          }

          function handleTestingLoc(chkBox) {
            var elemOldTestingLocation = document.getElementById("OldTestingLocation");
            var genIPL = false;

            if (elemOldTestingLocation != null) {
              if (chkBox.checked) {
                if (elemOldTestingLocation.value != 1) {
                  genIPL = true;
                }
              }
              else {
                if (elemOldTestingLocation.value != 0) {
                  genIPL = true;
                }
              }
            }
            if (genIPL) {
              genIPL = confirm('<%Sendb(Copient.PhraseLib.Lookup("store-edit.ChangeRequiresIPL", LanguageID))%>');
          if (!genIPL) {
            chkBox.checked = !chkBox.checked;
          }
        }
      }

      function showLanguagesDialog(bShow) {
        var elem = document.getElementById('uploader');

        if (elem != null) {
          elem.style.display = (bShow) ? '' : 'none';
        }
      }

      function removeLocLang(langID, langName) {
        removeRow('trLocLang' + langID)
        addToAvailable(langID, langName);
      }

      function removeRow(rowName) {
        var trElem = document.getElementById(rowName);

        if (trElem != null) {
          trElem.parentNode.removeChild(trElem);
        }
      }

      function addToAvailable(langID, langName) {
        var elem = document.getElementById('availLangs');

        if (elem != null) {
          elem.options[elem.options.length] = new Option(langName, langID);
          sortAvailable();
        }
      }

      function sortAvailable() {
        var select = document.getElementById('availLangs');
        var x = [];
        var length = 0;

        if (select != null) {
          length = select.options.length;
          for (var i = length - 1; i >= 0; --i) {
            x.push(select.options[i]);
            select.removeChild(select.options[i]);
          }
          x.sort(function (o1, o2) {
            if (o1.text.toLowerCase() > o2.text.toLowerCase()) {
              return 1;
            } else if (o1.text.toLowerCase() < o2.text.toLowerCase()) {
              return -1;
            }
            return 0;
          });
          for (var i = 0; i < length; ++i) {
            select.appendChild(x[i]);
          }
        }
      }

      function addLanguages() {
        var elem = document.getElementById('availLangs');

        if (elem != null) {
          for (var i = 0; i < elem.options.length; i++) {
            if (elem.options[i].selected) {
              // add the language to the main page.
              addLanguageRow(elem.options[i].value, elem.options[i].text);
              removeRow('trNoLangs');
              // remove it from the available language box
              elem.options[i] = null;
              i--;
            }
          }
        }
        showLanguagesDialog(false);
      }

      function addLanguageRow(langID, langName) {
        var elemTable = document.getElementById('tableLangs');
        var elemTr = document.createElement('tr');
        var elemTd = null;
        var elemInput = null;

        if (elemTable != null) {
          elemTr.id = 'trLocLang' + langID;
          // first column (delete button and lang id hidden field)
          elemTd = document.createElement('td');
          elemInput = document.createElement('input');
          elemInput.type = 'hidden';
          elemInput.setAttribute("name", "loclanguageid");
          elemInput.value = langID;
          elemTd.appendChild(elemInput);
          elemInput = document.createElement('input');
          elemInput.type = 'button';
          elemInput.value = 'X';
          elemInput.setAttribute("class", "ex");
          elemInput.setAttribute("className", "ex");
          elemInput.onclick = function () { removeLocLang(langID, langName); }
          elemTd.appendChild(elemInput);
          elemTr.appendChild(elemTd);

          // second column (language name)
          elemTd = document.createElement('td');
          elemTd.appendChild(document.createTextNode(langName));
          elemTr.appendChild(elemTd);

          // third column (required checkbox)
          elemTd = document.createElement('td');
          elemInput = document.createElement('input');
          elemInput.type = 'checkbox';
          elemInput.setAttribute("name", "loclangrequired");
          elemInput.value = langID;
          elemTd.appendChild(elemInput);
          elemTr.appendChild(elemTd);

          elemTable.tBodies[0].appendChild(elemTr);
        }

      }
</script>
<form action="#" id="mainform" name="mainform" method="post">
  <%
    If OldTestingLocation Then
      Send("<input type=""hidden"" id=""OldTestingLocation"" name=""OldTestingLocation"" value=""1"" />")
    Else
      Send("<input type=""hidden"" id=""OldTestingLocation"" name=""OldTestingLocation"" value=""0"" />")
    End If

  %>
  <div id="intro">
    <h1 id="title">
      <%
        If LocationID = 0 Then
          If LocationTypeID = 2 Then
            Sendb(Copient.PhraseLib.Lookup("term.newserver", LanguageID))
          Else
            Sendb(Copient.PhraseLib.Lookup("term.newstore", LanguageID))
          End If
        Else
          If LocationTypeID = 2 Then
            Sendb(Copient.PhraseLib.Lookup("term.server", LanguageID) & " " & ExtLocationCode & ": ")
          Else
            Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID) & " " & ExtLocationCode & ": ")
          End If
          Sendb(MyCommon.TruncateString(LocationName, 40))
        End If
      %>
    </h1>
    <div id="controls">
      <%
        If (LocationID = 0) Then
          If (Logix.UserRoles.CreateStores) Then
            Send_Save()
          End If
        Else
          ShowActionButton = (Logix.UserRoles.CreateStores) OrElse (Logix.UserRoles.CRUDStoresAndTerminals) OrElse (Logix.UserRoles.DeleteStores)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.CRUDStoresAndTerminals) Then
              Send_Save()
            End If
            If (Logix.UserRoles.DeleteStores) Then
              Send_Delete()
            End If
            If (Logix.UserRoles.CreateStores) Then
              Send_New()
            End If
            If (Logix.UserRoles.CRUDStoresAndTerminals) AndAlso bCmToUeEnabled AndAlso BrickAndMortarLocationId > 0 Then
              Send("<input type=""submit"" accesskey=""s"" class=""regular"" id=""save"" name=""deploy"" value=""" & Copient.PhraseLib.Lookup("perm.offers-deploy", LanguageID) & """" & " />")
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(13, LocationID, AdminUserID)
            End If
          End If
        End If
        Send(" <input type=""hidden"" id=""LocationID"" name=""LocationID"" value=""" & LocationID & """ />")
        Send(" <input type=""hidden"" id=""LocationTypeID"" name=""LocationTypeID"" value=""" & LocationTypeID & """ />")
        Send(" <input type=""hidden"" id=""hExtLocationCode"" name=""hExtLocationCode"" value=""" & hExtLocationCode & """ />")
        Send(" <input type=""hidden"" id=""hLocationName"" name=""hLocationName"" value=""" & hLocationName & """ />")
      %>
    </div>
  </div>
  <div id="main">
    <%  If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      ElseIf (informationMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & informationMessage & "</div>")
      End If
    %>
    <div id="column1">
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <%
          If (LocationID > 0) Then
            If LocationTypeID <> 2 Then
              Send(Copient.PhraseLib.Lookup("term.storeid", LanguageID) & ": " & LocationID & "<br />")
            End If
            Send(Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ": " & EngineName & "<br />")
            Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineType & """ />")
            Send("<br class=""half"" />")
          End If
          
          If (LocationID = 0) Then
            If (BannersEnabled) Then
              MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                  "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                  "where BAN.Deleted=0 and BAN.AllBanners=0 and AdminUserID = " & AdminUserID & " order by BAN.Name;"
              rst = MyCommon.LRT_Select
              BannerCt = rst.Rows.Count
              If (BannerCt > 0) Then
                BannerID = MyCommon.Extract_Val(Request.Form("banner"))
                Send("<br class=""half"" />")
                Send("<label for=""banner"">" & Copient.PhraseLib.Lookup("term.banners", LanguageID) & ":</label><br />")
                Send("  <select class=""longest"" name=""banner"" id=""banner"">")
                For Each row In rst.Rows
                  Send("    <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """" & IIf(BannerID = MyCommon.NZ(row.Item("BannerID"), -1), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                Next
                Send("  </select><br />")
                Send("&nbsp;<br class=""half"" />")
              End If
              Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineType & """ />")
            Else
              MyCommon.QueryStr = "select EngineID, Description, DefaultEngine, PhraseID from PromoEngines with (NoLock) where Installed=1 and EngineID in (0, 2, 9);"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Send("<br class=""half"" />")
                Send("<label for=""EngineID"">" & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & ":</label><br />")
                                
                Dim strChange as String = ""
                If (bCmToUeEnabled) or (LocationTypeID = 1) Then
                   strChange = "onchange="""
                   If (bCmToUeEnabled) Then
                      strChange = strChange & "ActivateBrickAndMortar();"
                   End If
                   If (LocationTypeID = 1) Then
                      strChange = strChange & "ShoworHideTerminal(this);"
                   End If
                   strChange = strChange & """"
                End If
                                
                Send("<select class=""medium"" id=""EngineID"" name=""EngineID"" " & strChange & ">")
                For Each row In rst.Rows
                  If (EngineType = CInt(row.Item("EngineID")) OrElse (EngineType = -1 AndAlso row.Item("DefaultEngine") = 1)) Then
                    Send("<option selected=""selected"" value=" & row.Item("EngineID") & ">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("Description"), "")) & " </option>")
                    DefaultEngine = CInt(row.Item("EngineID"))
                  Else
                    Send("<option value=" & row.Item("EngineID") & ">" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("PhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("Description"), "")) & " </option>")
                  End If
                                
                Next
                           
                Send("</select><br />&nbsp;")
                Send("<br class=""half"" />")
                For Each row In rst.Rows
                  LoadTerminals(CInt(row.Item("EngineID")), LocationTypeID, LocationID, MyCommon, DefaultEngine, IIf(CInt(row.Item("EngineID")) = 2, TerminalSetIDCPE, TerminalSetIDUE))
                Next
              End If

              If bCmToUeEnabled And (LocationTypeID = 1) Then
                Send("<div class=""Brick"" id=""Brick"" style=""display:none; visibility: hidden;"" >")
                MyCommon.QueryStr = "select LocationId, ExtLocationCode, LocationName from Locations as LOC with (NoLock) where Deleted=0 and EngineID=0" & _
                                    " and not exists(select LocationID from Locations with (NoLock) where Deleted=0 and EngineID=9 and BrickAndMortarLocationID=LOC.LocationID);"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                  Send("<br class=""half"" />")
                  Send("<label for=""BrickCode"">" & Copient.PhraseLib.Lookup("term.BrickAndMortarStore", LanguageID) & ":</label><br />")
                  Send("<select class=""medium"" id=""BrickCode"" name=""BrickCode"" onchange=""SetExtCode();"" >")
                  Sendb("<option value=""0"" selected=""selected"" >" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
                  For Each row In rst.Rows
                    Send("<option value=""" & row.Item("ExtLocationCode") & "|" & row.Item("LocationName") & """>" & row.Item("ExtLocationCode") & " - " & row.Item("LocationName") & " </option>")
                  Next
                  Send("</select><br class=""half"" />")
                End If
                Send("</div>")
              End If
            End If
          Else
            Send("<br class=""half"" />")
            If (BannersEnabled AndAlso HasAllBannersAccess AndAlso (BannerID = 0 OrElse dtGroups.Rows.Count = 0)) Then
              MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                  "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                  "inner join BannerEngines BE with (NoLock) on BE.BannerID = AUB.BannerID " & _
                                  "where BAN.Deleted=0 and BAN.AllBanners=0 and AdminUserID = " & AdminUserID & " " & _
                                  "and BE.EngineID=" & EngineType & " order by BAN.Name;"
              rst = MyCommon.LRT_Select
              BannerCt = rst.Rows.Count
              If (BannerCt > 0) Then
                'BannerID = MyCommon.Extract_Val(Request.QueryString("banner"))
                Send("<br class=""half"" />")
                Send("<label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banners", LanguageID) & ":</label><br />")
                Send("  <select class=""longest"" name=""banner"" id=""banner"">")
                For Each row In rst.Rows
                  Send("    <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """" & IIf(BannerID = MyCommon.NZ(row.Item("BannerID"), -1), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                Next
                Send("  </select><br />")
                Send("&nbsp;<br class=""half"" />")
              End If
            ElseIf BannersEnabled Then
              Send(Copient.PhraseLib.Lookup("term.banner", LanguageID) & ": " & MyCommon.SplitNonSpacedString(BannerName, 25) & "<br />")
              Send("<input type=""hidden"" name=""banner"" id=""banner"" value=""" & BannerID & """ />")
            End If
                   
                  
            LoadTerminals(EngineType, LocationTypeID, LocationID, MyCommon, EngineType)
          End If
          Send("<br class=""half"" />")

          Send("<label for=""ExtLocationCode"">" & Copient.PhraseLib.Lookup("term.code", LanguageID) & ":</label><br />")
          Send("<input type=""text"" id=""ExtLocationCode"" name=""ExtLocationCode"" class=""longest"" maxlength=""20"" value=""" & ExtLocationCode & """ " & IIf(BrickAndMortarLocationId > 0, " disabled", "") & "/><br />")
          Send("<label for=""LocationName"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
          If (LocationName Is Nothing) Then
            LocationName = ""
          End If
          Send("<input type=""text"" id=""LocationName"" name=""LocationName"" class=""longest"" maxlength=""100"" value=""" & LocationName.Replace("""", "&quot;") & """ " & IIf(BrickAndMortarLocationId > 0, " disabled", "") & "/><br />")
          Send("<br class=""half"" />")
          Send("<label for=""Description"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</label><br />")
          Send("<textarea id=""Description"" name=""Description"" class=""longest"" rows=""3"" cols=""48"">" & Description & "</textarea><br />")
          If BrickAndMortarLocationId = 0 Then
            Send("<div id=""TestingDiv"">")
            Sendb("<input type=""checkbox"" id=""TestingLocation"" name=""TestingLocation""" & IIf(TestingLocation, " checked=""checked""", ""))
            If (MyCommon.Fetch_SystemOption(88) = "1") Then
              Sendb(" onclick=""handleTestingLoc(this);""")
            End If
            Send(" /><label for=""TestingLocation"">" & Copient.PhraseLib.Lookup("term.testinglocation", LanguageID) & "</label><br />")
            Send("</div>")
          End If
          If LastUpdate = Nothing Then
          Else
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
            longDate = LastUpdate
            longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
            Send(longDateString & "<br />")
          End If
          If (LocationID <> 0) AndAlso (Logix.UserRoles.AccessStoreHealth) Then
            Send("<br class=""half"" />")
            If EngineType = 9 Then
              If (MyCommon.Fetch_UE_SystemOption(91) = "1") Then
                If LocationTypeID = 2 Then
                  Sendb("<b><a href=""/logix/UE/UEServerHealthSummary.aspx"">")
                  Sendb(Copient.PhraseLib.Lookup("term.serverhealth", LanguageID))
                  Send("</a></b><br />")
                ElseIf (MyCommon.Fetch_UE_SystemOption(188) = "1") Then
                  Send("<b><a href=""/logix/UE/UEEngineHealthSummary.aspx?SEARCH=" & ExtLocationCode & "&FILTER=1 "">")
                  Sendb(Copient.PhraseLib.Lookup("term.enginehealth", LanguageID))
                  Send("</a></b><br />")
                End If
              Else
                Sendb("<b><a href=""/logix/store-detail.aspx?LocationID=" & LocationID & """>")
                If LocationTypeID = 1 Then
                  Sendb(Copient.PhraseLib.Lookup("term.storehealth", LanguageID))
                Else
                  Sendb(Copient.PhraseLib.Lookup("term.serverhealth", LanguageID))
                End If
                Send("</a></b><br />")
              End If
            Else
              If LocationTypeID = 2 Then
                Sendb("<b><a href=""/logix/store-health-cpe.aspx"">")
                Sendb(Copient.PhraseLib.Lookup("term.serverhealth", LanguageID))
              Else
                Sendb("<b><a href=""/logix/store-detail.aspx?LocationID=" & LocationID & """>")
                Sendb(Copient.PhraseLib.Lookup("term.storehealth", LanguageID))
              End If
              Send("</a></b><br />")
            End If
          End If
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="localserver" <% If ((EngineType <> 2) And (EngineType <> 9) And (DefaultEngine) <> 2 And (DefaultEngine <> 9)) Then Sendb(" style=""display:none;""")%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.localserver", LanguageID))%>
          </span>
        </h2>
        <%
          MyCommon.QueryStr = "select LocalServerID, ImageFetchURL, IncentiveFetchURL, PhoneHomeIPOverride, OfflineFTPUser, OfflineFTPPass, OfflineFTPPath, OfflineFTPIP " & _
                              "from LocalServers with (NoLock) where LocationID=" & LocationID & ";"
          rst = MyCommon.LRT_Select
          Dim LocalServerID As String = "&nbsp;"
          If rst.Rows.Count > 0 Then
            LocalServerID = MyCommon.NZ(rst.Rows(0).Item("LocalServerID"), "&nbsp;")
            ImageFetchURL = MyCommon.NZ(rst.Rows(0).Item("ImageFetchURL"), "")
            IncentiveFetchURL = MyCommon.NZ(rst.Rows(0).Item("IncentiveFetchURL"), "")
            PhoneHomeIPOverride = MyCommon.NZ(rst.Rows(0).Item("PhoneHomeIPOverride"), "")
            OfflineFTPUser = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPUser"), "")
            OfflineFTPPass = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPPass"), "")
            OfflineFTPPath = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPPath"), "")
            OfflineFTPIP = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPIP"), "")
            Send("<label for=""imagefetchurl"">" & Copient.PhraseLib.Lookup("term.imagefetchurl", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""imagefetchurl"" name=""imagefetchurl"" class=""longer"" maxlength=""255"" value=""" & ImageFetchURL & """ /><br />")
            
            'AMS-1597: Removing this field as incentive fetch url will be taken from system option 43 for UE
            If EngineType <> 9 Then
               Send("<label for=""incentivefetchurl"">" & Copient.PhraseLib.Lookup("term.incentivefetchurl", LanguageID) & ":</label><br />")
               Send("<input type=""text"" id=""incentivefetchurl"" name=""incentivefetchurl"" class=""longer"" maxlength=""255"" value=""" & IncentiveFetchURL & """ /><br />")
            End If

            Send("<label for=""PhoneHomeIPOverride"">" & Copient.PhraseLib.Lookup("term.phonehomeipoverride", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""PhoneHomeIPOverride"" name=""PhoneHomeIPOverride"" class=""longer"" maxlength=""255"" value=""" & PhoneHomeIPOverride & """ /><br />")
            'OfflineFTP
            If EngineType <> 9 Then
              Send("<label for=""FTPUser"">" & Copient.PhraseLib.Lookup("term.OfflineFTPUsername", LanguageID) & ":</label><br />")
              Send("<input type=""text"" id=""FTPUser"" name=""FTPUser"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPUser & """ /><br />")
              Send("<label for=""FTPPass"">" & Copient.PhraseLib.Lookup("term.OfflineFTPPassword", LanguageID) & ":</label><br />")
              Send("<input type=""password"" id=""FTPPass"" name=""FTPPass"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPPass & """ /><br />")
              Send("<label for=""FTPPath"">" & Copient.PhraseLib.Lookup("term.OfflineFTPPath", LanguageID) & ":</label><br />")
              Send("<input type=""text"" id=""FTPPath"" name=""FTPPath"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPPath & """ /><br />")
              Send("<label for=""FTPIP"">" & Copient.PhraseLib.Lookup("term.OfflineFTPIP", LanguageID) & ":</label><br />")
              Send("<input type=""text"" id=""FTPIP"" name=""FTPIP"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPIP & """ /><br />")
                        
              Send("<br class=""half"" />")
              Sendb(Copient.PhraseLib.Lookup("sanitycheck.status", LanguageID) & ": ")
              MyCommon.QueryStr = "select DBOK from SanityCheckStatus with (NoLock) where LocationID=" & LocationID
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                SanityCheckPassed = rst.Rows(0).Item("DBOK")
                If (SanityCheckPassed) Then
                  Send(Copient.PhraseLib.Lookup("term.passed", LanguageID))
                Else
                  Sendb(Copient.PhraseLib.Lookup("term.failed", LanguageID))
                  Sendb("<a href=""javascript:launchScReport(" & LocationID & ");"" alt=""" & Copient.PhraseLib.Lookup("store-detail.click-to-view", LanguageID) & _
                          """ title=""" & Copient.PhraseLib.Lookup("store-detail.click-to-view", LanguageID) & """ style=""margin-left:7px;"">")
                  Send("<img src=""../images/info.png"" border=""0"" style=""vertical-align: bottom;"" /></a>")
                End If
              Else
                Send(Copient.PhraseLib.Lookup("term.noresults", LanguageID))
              End If
            End If
          End If
        %>
        <hr class="hidden" />
      </div>
      <%If (LocationTypeID = 1 And BrickAndMortarLocationId = 0) Then%>
      <div class="box" id="address">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.address", LanguageID))%>
          </span>
        </h2>
        <table cellpadding="0" cellspacing="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.address", LanguageID))%>">
          <tr>
            <td colspan="3">
              <%
                Send("<label for=""address1"">" & Copient.PhraseLib.Lookup("term.address", LanguageID) & ":</label><br />")
                Send("<input type=""text"" id=""address1"" name=""address1"" class=""longer"" maxlength=""200"" value=""" & Address1 & """ /><br />")
                Sendb("<input type=""text"" id=""address2"" name=""address2"" class=""longer"" maxlength=""200"" value=""" & Address2 & """ /><br />")
              %>
            </td>
          </tr>
          <tr>
            <td style="width: 150px;">
              <%
                Send("<label for=""city"">" & Copient.PhraseLib.Lookup("term.city", LanguageID) & ":</label><br />")
                Sendb("<input type=""text"" id=""city"" name=""city"" maxlength=""100"" value=""" & City & """ /><br />")
              %>
            </td>
            <td>
              <%
                If (CountryID = 0) Then
                  MyCommon.QueryStr = "select SubPhraseID from Countries with (NoLock) where CountryID=" & MyCommon.Fetch_SystemOption(65)
                Else
                  MyCommon.QueryStr = "select SubPhraseID from Countries with (NoLock) where CountryID=" & CountryID
                End If
                rst = MyCommon.LRT_Select
                StateType = Copient.PhraseLib.Lookup(rst.Rows(0).Item(0), LanguageID)
                Send("<label for=""state"">" & StateType & ":</label><br />")
                Sendb("<input type=""text"" id=""state"" name=""state"" class=""short"" maxlength=""50"" value=""" & State & """ /><br />")
              %>
            </td>
            <td>
              <%
                Send("<label for=""zip"">" & Copient.PhraseLib.Lookup("term.postalcode", LanguageID) & ":</label><br />")
                Sendb("<input type=""text"" id=""zip"" name=""zip"" class=""short"" maxlength=""20"" value=""" & Zip & """ /><br />")
              %>
            </td>
          </tr>
          <tr>
            <td colspan="2">
              <label for="country">
                <% Sendb(Copient.PhraseLib.Lookup("term.country", LanguageID))%></label>:<br />
              <select id="country" name="country">
                <%
                  MyCommon.QueryStr = "select CountryID, PhraseID from Countries with (NoLock)"
                  rst = MyCommon.LRT_Select()
                  If CountryID = 0 Then CountryID = MyCommon.Fetch_SystemOption(65)
                  For Each row In rst.Rows
                    Sendb("    <option value=""" & row.Item("CountryID") & """")
                    If (row.Item("CountryID") = CountryID) Then
                      Sendb(" selected=""selected""")
                      CountryName = Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID)
                    End If
                    Send(">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                  Next
                %>
              </select>
            </td>
            <td style="text-align: right;" valign="bottom">
              <%
                If City <> "" AndAlso State <> "" Then
                  If Address1 <> "" Then GMapString = GMapString & Address1 & ", "
                  If Address2 <> "" Then GMapString = GMapString & Address2 & ", "
                  If City <> "" Then GMapString = GMapString & City & ", "
                  If State <> "" Then GMapString = GMapString & State
                  If Zip <> "" Then GMapString = GMapString & " " & Zip
                  If CountryID > 2 Then GMapString = GMapString & " " & CountryName
                  GMapString = GMapString & " (" & Copient.PhraseLib.Lookup("term.store", LanguageID) & " " & ExtLocationCode & ")"
                  GMapString = GMapString.Replace(" ", "+")
                  Send("<b><a href=""http://maps.google.com/maps?q=" & GMapString & """ target=""_blank"">" & Copient.PhraseLib.Lookup("term.map", LanguageID) & "</a></b>")
                End If
              %>
            </td>
          </tr>
        </table>
        <hr class="hidden" />
      </div>
      <%
          

          bCmInstalled = MyCommon.IsEngineInstalled(0)
          bCpeInstalled = MyCommon.IsEngineInstalled(2)
          bUeInstalled = MyCommon.IsEngineInstalled(9)

          If (bUeInstalled AndAlso bCmInstalled Or bCpeInstalled) Then
              bDisplayTimezone = True
          ElseIf (bUeInstalled AndAlso Not (bCmInstalled)) Or (bUeInstalled AndAlso Not (bCpeInstalled)) Then
              bDisplayTimezone = False
          ElseIf (bCmInstalled Or bCpeInstalled) Then
              bDisplayTimezone = True
          End If
          'If ((LocationTypeID = 1) AndAlso ((Not MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Or (bCmToUeEnabled))) Then
          If ((LocationTypeID = 1) AndAlso (bDisplayTimezone) Or (bCmToUeEnabled)) Then

         
          Send("<div class=""box"" id=""timezone"">")
          Send(" <h2><span>" & Copient.PhraseLib.Lookup("term.timezone", LanguageID) & "</span></h2>")
          Send("  <select id=""TimeZone"" name=""TimeZone"">")
          Send("   <option value="""">&nbsp;</option>")
          MyCommon.QueryStr = "select Code, UTCOffset, Description from TimeZones with (NoLock) order by UTCOffset;"
          rst = MyCommon.LRT_Select()
          If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
              Sendb("   <option value=""" & MyCommon.NZ(row.Item("Code"), "") & """" & If(MyCommon.NZ(row.Item("Code"), 0) = TimeZone, " selected=""selected""", "") & " alt=""" & MyCommon.NZ(row.Item("Description"), "") & """ title=""" & MyCommon.NZ(row.Item("Description"), "") & """>")
              Sendb(If(MyCommon.NZ(row.Item("UTCOffset"), 0) >= 0, "+", "") & MyCommon.NZ(row.Item("UTCOffset"), 0))
              Sendb(If(MyCommon.NZ(row.Item("Description"), "") <> "", " – " & MyCommon.NZ(row.Item("Description"), ""), ""))
              Send("</option>")
            Next
          End If
          Send("  </select>")
          Send("</div>")
        End If
      %>
      <div class="box" id="contact">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.contact", LanguageID))%>
          </span>
        </h2>
        <label for="contactname">
          <% Sendb(Copient.PhraseLib.Lookup("term.contactperson", LanguageID))%></label>:<br />
        <% Sendb("<input type=""text"" id=""contactname"" name=""contactname"" class=""medium"" maxlength=""100"" value=""" & ContactName & """ /><br />")%>
        <label for="phonenumber">
          <% Sendb(Copient.PhraseLib.Lookup("term.phone", LanguageID))%></label>:<br />
        <% Sendb("<input type=""text"" id=""phonenumber"" name=""phonenumber"" class=""medium"" maxlength=""20"" value=""" & PhoneNumber & """ /><br />")%>
        <hr class="hidden" />
      </div>
      <%  If (EngineType = 2) Then%>
      <div class="box" id="systemsetttings">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.systemsettings", LanguageID))%>
          </span>
        </h2>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID))%>">
          <%
            Dim rstLoc As DataTable
            MyCommon.QueryStr = "select OptionID, PhraseID from SiteSpecificOptions with (NoLock);"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              For Each row In rst.Rows
                OptionID = MyCommon.NZ(row.Item("OptionID"), 0)
                Send("")
                Send(" <tr>")
                Send("  <td><label for=""oid" & OptionID & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & ":</label>")
                MyCommon.QueryStr = "select OptionValue, PhraseID, DefaultVal " & _
                                    "from SiteSpecificOptionValues with (NoLock) " & _
                                    "where OptionID=" & OptionID & " order by OptionValue;"
                rst2 = MyCommon.LRT_Select
                If (rst2.Rows.Count > 0) Then
                  Send("   <select id=""oid" & OptionID & """ name=""oid" & OptionID & """>")
                  For Each row2 In rst2.Rows
                    MyCommon.QueryStr = "select OptionValue from LocationOptions with (NoLock) where OptionID=" & OptionID & _
                                        " and LocationId=" & LocationID & " and deleted = 0;"
                    rstLoc = MyCommon.LRT_Select
                    If (rstLoc.Rows.Count > 0) Then
                      LocationOptionValue = MyCommon.NZ(rstLoc.Rows(0).Item("OptionValue"), "")
                    Else
                      If (MyCommon.NZ(row2.Item("DefaultVal"), False)) Then
                        LocationOptionValue = MyCommon.NZ(row2.Item("OptionValue"), "")
                      Else
                        LocationOptionValue = ""
                      End If
                    End If
                    Sendb("    <option value=""" & MyCommon.NZ(row2.Item("OptionValue"), "") & """")
                    If MyCommon.NZ(row2.Item("OptionValue"), "") = LocationOptionValue Then Sendb(" selected=""selected""")
                    Send(">" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</option>")
                  Next
                  Send("   </select>")
                Else
                  Send("   <input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
                End If
                Send("  </td>")
                Send(" </tr>")
              Next
            End If
          %>
        </table>
      </div>
      <% End If%>
      <%End If%>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <%If (LocationTypeID = 1) Then%>
      <div class="box" id="localization" <% If (LocationID = 0) Then Sendb(" style=""display:none;""")%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.localization", LanguageID))%>
          </span>
        </h2>
        <%
          ' currency and UOM localization settings are only applicable to the UE Engine.
          If EngineType = Copient.CommonInc.InstalledEngines.UE Then
			    If CurrencyID = 0 Then
                            MyCommon.QueryStr = "select case isnumeric(OptionValue) when 1 then cast(OptionValue as int) else 1 end AS DefaultCurrencyID from UE_SystemOptions where OptionID=137;"
                            rst = MyCommon.LRT_Select()
                            If (rst.Rows.Count = 1) Then
                                For Each row In rst.Rows
                                    CurrencyID = row.Item("DefaultCurrencyID")
                                Next
                            End If
                        End If
            ' when enable multi-currency for UE option is on, selection is available; otherwise, the default currency type is always used.
            Send(Copient.PhraseLib.Lookup("term.currency", LanguageID) & ":<br />")
            If MyCommon.Extract_Val(MyCommon.Fetch_UE_SystemOption(136)) = 1 Then
              MyCommon.QueryStr = "select CurrencyID, NamePhraseTerm, AbbreviationPhraseTerm, Symbol " & _
                                  "from Currencies with (NoLock);"
              rst2 = MyCommon.LRT_Select
              If rst2.Rows.Count Then
                Send("        <select name=""currencyid"" class=""long"">")
                If CurrencyID = 0 Then
                  Send("          <option value=""0"">" & Copient.PhraseLib.Lookup("term.unspecified", LanguageID) & "</option>")
                End If
                For Each row2 In rst2.Rows
                  Sendb("          <option value=""" & row2.Item("CurrencyID") & """" & If(CurrencyID = row2.Item("CurrencyID"), " selected=""selected""", "") & ">")
                  Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("NamePhraseTerm"), ""), LanguageID))
                  Sendb(" (" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("AbbreviationPhraseTerm"), ""), LanguageID) & " " & MyCommon.NZ(row2.Item("Symbol"), "") & ")")
                  Send("</option>")
                Next
                Send("        </select>")
              End If
            Else
              ' if a CurrencyID isn't already saved for this location, then shows its default one.
              If CurrencyID = 0 Then CurrencyID = MyCommon.Extract_Val(MyCommon.Fetch_UE_SystemOption(137))
              MyCommon.QueryStr = "select CurrencyID, NamePhraseTerm, AbbreviationPhraseTerm, Symbol " & _
                                  "from Currencies with (NoLock) where CurrencyID=" & CurrencyID
              rst2 = MyCommon.LRT_Select
              If rst2.Rows.Count Then
                Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("NamePhraseTerm"), ""), LanguageID))
                Send(" (" & Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("AbbreviationPhraseTerm"), ""), LanguageID) & " " & MyCommon.NZ(rst2.Rows(0).Item("Symbol"), "") & ")")
                Send("        <input type=""hidden"" name=""currencyid"" value=""" & CurrencyID & """ />")
              End If
            End If
            Send("        <br /><br class=""half"" />")
          End If

          ' only show UOM set if the Enable multi unit of measure for UE is turned on.
          If MyCommon.Extract_Val(MyCommon.Fetch_UE_SystemOption(135)) = 1 Then
            Send(Copient.PhraseLib.Lookup("term.unitsofmeasureset", LanguageID) & ":<br />")

            MyCommon.QueryStr = "select UOMSetID, Name from UOMSets with (NoLock);"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count Then
              Send("        <select name=""uomsetid"" class=""long"">")
                Send("          <option value=""0"">" & Copient.PhraseLib.Lookup("term.unspecified", LanguageID) & "</option>")
              For Each row2 In rst2.Rows
                Sendb("          <option value=""" & row2.Item("UOMSetID") & """" & If(UOMSetID = row2.Item("UOMSetID"), " selected=""selected""", "") & ">")
                Sendb(MyCommon.NZ(row2.Item("Name"), ""))
                Send("</option>")
              Next
              Send("        </select>")
            End If
            Send("        <br /><br class=""half"" />")
          End If

          Send(Copient.PhraseLib.Lookup("term.languages", LanguageID) & ":<br />")

          Send("<table id=""tableLangs"">")
          Send("  <tr>")
          Send("    <th>" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & "</th>")
          Send("    <th>" & Copient.PhraseLib.Lookup("term.language", LanguageID) & "</th>")
          Send("    <th>" & Copient.PhraseLib.Lookup("term.required", LanguageID) & "</th>")
          Send("  </tr>")

          MyCommon.QueryStr = "select LANG.LanguageID, LANG.PhraseTerm, LL.Required " & _
                              "from Languages as LANG with (NoLock) " & _
                              "inner join LocationLanguages as LL with (NoLock) " & _
                              "  on LL.LanguageID = LANG.LanguageID " & _
                              "where LL.LocationID=" & LocationID & " and LL.Deleted=0;"
          rst2 = MyCommon.LRT_Select
          If rst2.Rows.Count Then
            For Each row2 In rst2.Rows
              LocLangID = row2.Item("LanguageID")
              LangName = Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseTerm"), ""), LanguageID)
              Send("  <tr id=""trLocLang" & LocLangID & """>")
              Send("    <td><input type=""hidden"" name=""loclanguageid"" value=""" & LocLangID & """ />")
              Send("       <input type=""button"" class=""ex"" value=""X"" onclick=""removeLocLang(" & LocLangID & ",'" & LangName & "');"" />")
              Send("    </td>")
              Send("    <td>" & LangName & "</td>")
              Send("    <td><input type=""checkbox"" name=""loclangrequired"" value=""" & LocLangID & """" & IIf(MyCommon.NZ(row2.Item("Required"), False), " checked=""checked""", "") & " /></td>")
              Send("  </tr>")
            Next
          Else
            Send("  <tr id=""trNoLangs""><td colspan=""3""><i>" & Copient.PhraseLib.Lookup("store-edit.NoSelectedLanguages", LanguageID) & "</i></td></tr>")
          End If
          Send("</table>")
          Send("<input type=""button"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""showLanguagesDialog(true);"" />")
          Send("<br />")

        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="groups" <% If (LocationID = 0) Then Sendb(" style=""display:none;""")%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedgroups", LanguageID))%>
          </span>
        </h2>
        <%
          If Not dtGroups Is Nothing AndAlso dtGroups.Rows.Count > 0 Then
            Send("    " & Copient.PhraseLib.Lookup("store-edit.groups", LanguageID))
            Send("    <div class=""boxscroll"">")
            For Each row In dtGroups.Rows
              If (row.Item("AllLocations").Equals(System.DBNull.Value) OrElse Not row.Item("AllLocations")) Then
                Send("     <a href=""lgroup-edit.aspx?LocationGroupId=" & row.Item("LocationGroupId") & """>" & row.Item("Name") & "</a><br />")
              Else
                Send("     " & row.Item("Name") & "<br />")
              End If
            Next
            Send("    </div>")
          Else
            Send("    <div class=""boxscroll"">")
            Sendb(Copient.PhraseLib.Lookup("store-edit.nogroups", LanguageID) & "<br />")
            Send("    </div>")
          End If
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="offers" <% If (LocationID = 0) Then Sendb(" style=""display:none;""")%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.recently", LanguageID) & " " & Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID).ToLower())%>
          </span>
        </h2>
        <%
          Dim assocName As String = ""
          If Not dtOffers Is Nothing AndAlso dtOffers.Rows.Count > 0 Then
            Send("    " & Copient.PhraseLib.Lookup("store-edit.offers", LanguageID))
            Send("    <div class=""boxscroll"">")
            For Each row In dtOffers.Rows
              If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
              Else
                assocName = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
              End If
              If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & assocName & "</a>")
              Else
                Sendb(assocName)
              End If
              Send("<br />")
            Next
            Send("    </div>")
          Else
            Send("    <div class=""boxscroll"">")
            Sendb(Copient.PhraseLib.Lookup("store-edit.nooffers", LanguageID))
            Send("    </div>")
          End If
        %>
        <hr class="hidden" />
      </div>
      <%End If%>
      <% skippast:%>
    </div>
    <br clear="all" />
  </div>
</form>
<div id="uploader" style="display: none;">
  <div id="uploadwrap">
    <div class="box" id="uploadbox">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.languages", LanguageID))%>
        </span>
      </h2>
      <%
        Sendb("<input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" ")
        Send("onclick=""showLanguagesDialog(false);"" />")
        Sendb(Copient.PhraseLib.Lookup("store-edit.SelectLanguage", LanguageID))
        Send("<br /><br />")

        MyCommon.QueryStr = "select LANG.LanguageID, case when PT.Phrase is not null then PT.Phrase else LANG.Name end as Name " & _
                            "from Languages as LANG " & _
                            "left join UIPhrases as UIP on UIP.Name = LANG.PhraseTerm " & _
                            "left join PhraseText as PT  on PT.PhraseID = UIP.PhraseID and PT.LanguageID=1 " & _
                            "where LANG.LanguageID not in (select LanguageID from LocationLanguages where Deleted=0 and LocationID=" & LocationID & ");"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          Send("<center>")
          Send("<select name=""availLangs"" id =""availLangs"" class=""long"" multiple=""multiple"" size=""8"">")
          For Each row In rst.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("LanguageID"), 1) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
          Next
          Send("</select>")
          Send("<br />")
          Send("<input type=""button"" name=""btnAddLang"" id=""btnAddLang"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""addLanguages();"" />")
          Send("<br />")
          Send("</center>")
        Else
          Send(Copient.PhraseLib.Lookup("store-edit.AllSelected", LanguageID))
        End If
      %>
      <hr class="hidden" />
    </div>
  </div>
  <%
    If Request.Browser.Type = "IE6" Then
      Send("<iframe src=""javascript:'';"" id=""uploadiframe-pg"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no""></iframe>")
    End If
  %>
</div>
<script type="text/javascript">
  if (window.captureEvents) {
    window.captureEvents(Event.CLICK);
    window.onclick = handlePageClick;
  } else {
    document.onclick = handlePageClick;
  }
  sortAvailable();
  ActivateBrickAndMortar();
</script>
<script runat="server">
  Sub UpdateLocationLanguages(ByRef Common As Copient.CommonInc, ByVal LocationID As Long)
    Dim LangIDs() As String
    Dim LangRequired() As String
    Dim IDList As String = ""
    Dim LangID As Integer
    Dim Required As Boolean
    Dim PKID As Integer

    LangIDs = Request.Form.GetValues("loclanguageid")
    LangRequired = Request.Form.GetValues("loclangrequired")

    If LangIDs IsNot Nothing Then
      IDList = String.Join(",", LangIDs)
      If IDList IsNot Nothing AndAlso IDList.Trim.Length > 0 Then
        ' mark as deleted all languages that are no longer selected for this location
        Common.QueryStr = "update LocationLanguages with (RowLock) set Deleted=1, LastUpdate=GETDATE() " & _
                          "where LocationID=" & LocationID & " and LanguageID not in(" & IDList & ");"
        Common.LRT_Execute()
      End If

      For Each LangStr As String In LangIDs
        ' check if this language is marked as required
        Required = False
        If LangRequired IsNot Nothing Then
          For Each s As String In LangRequired
            Required = s = LangStr
            If Required Then Exit For
          Next
        End If

        If Integer.TryParse(LangStr, LangID) AndAlso LangID > 0 Then
          Common.QueryStr = "dbo.pt_LocationLanguages_Update"
          Common.Open_LRTsp()
          Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
          Common.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int).Value = LangID
          Common.LRTsp.Parameters.Add("@Required", SqlDbType.Bit).Value = IIf(Required, 1, 0)
          Common.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
          Common.LRTsp.ExecuteNonQuery()
          PKID = Common.LRTsp.Parameters("@PKID").Value
          Common.Close_LRTsp()
        End If
      Next
    Else
      ' no languages are currently selected, so mark all existing records as deleted
      Common.QueryStr = "update LocationLanguages with (RowLock) set Deleted=1, LastUpdate=GETDATE() where LocationID=" & LocationID
      Common.LRT_Execute()
    End If
  End Sub
    
  Sub LoadTerminals(ByVal EngineType As Integer, ByVal LocationTypeID As Integer, ByVal LocationID As Integer, ByRef MyCommon As Copient.CommonInc, ByVal SelectedEngineID As Integer, Optional ByVal SelectedTerminalID As Integer = 0)
    If LocationTypeID = 1 And (EngineType = 9) Then
      Dim rst As DataTable
      Dim TerminalSetID As Integer
      Send("<div " + IIf(SelectedEngineID = EngineType, "", "style=""display:none""") + " id=""divTerminal" + EngineType.ToString() + """ > ")
      Send(Copient.PhraseLib.Lookup("term.terminal-set", LanguageID) & ":<br />")
      Send("<select name=""terminalset" + EngineType.ToString() + """ id=""terminalset" + EngineType.ToString() + """ class=""longer"">")
            
      If LocationID > 0 Then
        ' load the assigned location terminal set
        MyCommon.QueryStr = "select TerminalSetID from LocationTerminals with (NoLock) " & _
                            "where LocationID=" & LocationID
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          TerminalSetID = MyCommon.NZ(rst.Rows(0).Item("TerminalSetID"), 0)
        End If
      End If
            
      ' if no location terminal set is assigned to the location, look for a default terminal set
      If TerminalSetID = 0 Then
        If SelectedTerminalID > 0 Then
          TerminalSetID = SelectedTerminalID
        Else
          MyCommon.QueryStr = "select TerminalSetID from TerminalSets where TerminalSetTypeID=2 and PromoEngineID=" & EngineType
          rst = MyCommon.LRT_Select
          If rst.Rows.Count > 0 Then
            TerminalSetID = MyCommon.NZ(rst.Rows(0).Item("TerminalSetID"), 0)
          End If
        End If
      End If
            
      ' when no terminal set is assigned AND there is not a default specified, show that the terminal set is unassigned.
      If TerminalSetID <= 0 Then
        Send("  <option value=""-1"" selected=""selected"">[" & Copient.PhraseLib.Lookup("term.unassigned", LanguageID) & "]</option>")
      End If
            
      ' load up all the terminal sets for the locations engine.
      MyCommon.QueryStr = "select TerminalSetID, Name from TerminalSets as TS with (NoLock) " & _
                          "Where PromoEngineID=" & EngineType
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        For Each row In rst.Rows
          Send("  <option value=""" & MyCommon.NZ(row.Item("TerminalSetID"), 0) & """" & IIf(TerminalSetID = MyCommon.NZ(row.Item("TerminalSetID"), 0), " selected=""selected""", "") & _
                  ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
        Next
      End If
      Send("</select>")
      Send("<br />")
      Send("</div>")
    End If
  End Sub
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (LocationID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(13, LocationID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "ExtLocationCode")
MyCommon = Nothing
Logix = Nothing
%>