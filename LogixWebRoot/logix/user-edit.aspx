﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>

<%@ Import Namespace="CMS.AMS" %>

<%
  ' *****************************************************************************
  ' * FILENAME: user-edit.aspx 
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
  
  
%>
<script runat="server">


    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim MyCryptlib As New CMS.CryptLib
    Dim HashLib As New CMS.HashLib.CryptLib
    Dim dst As DataTable
    Dim dtAlertTypes As DataTable
    Dim rst As DataTable
    Dim HSrst As DataTable
    Dim HSrow As System.Data.DataRow
    Dim AHSrst As DataTable
    Dim pgName As String
    Dim u_UID As Long
    Dim u_Username As String
    Dim u_Firstname As String
    Dim u_Lastname As String
    Dim u_Fullname As String
    Dim u_Employeeid As String
    Dim u_JobTitle As String
    Dim u_Employer As String
    Dim u_Password As String
    Dim u_Email As String
    Dim u_Alertemail As String
    Dim u_LastLogin As String
    Dim u_Language As Integer
    Dim u_StartPage As Integer
    Dim u_Style As Integer
    Dim longDate As New DateTime
    Dim longDateString As String
    Dim EditIdentity As Boolean = True
    Dim EditRoles As Boolean = True
    Dim sqlBuf As New StringBuilder()
    Dim Restricted As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannersEnabled As Boolean = False
    Dim EngineList As String = ""
    Dim StartPage As String = ""
    Dim x As Integer = 0
    Dim PrefID As Integer = 0
    Dim PrefTypeID As Integer = 0
    Dim PrefName As String = ""
    Dim PrefValue As String = ""
    Dim File As HttpPostedFile
    Dim UserImagePath As String = ""
    Dim GraphicPath As String = ""
    Dim DEFAULT_GRAPHIC_PATH As String = MyCommon.Get_Install_Path
    Dim RegHSID As Integer = 0
    Dim AlertHSID As Integer = 0
    Dim Debug As Boolean = False
    Dim DefaultLanguageID As Integer
    Dim dtStoreUser as DataTable
    Dim row as DataRow
    Dim maxStoreUser as Integer
    Dim showStoreUser As Boolean = False
    Dim alreadyAddedLabels As Boolean = False
    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_EPM_Roles()

        Dim FormData As String
        Dim RawRequest As String
        Dim HostURI As String = ""
        Dim TargetAddress As String
        Dim dst As DataTable
        Dim ConnInc As New Copient.ConnectorInc()

        'The code for the contents of this box lives in customer.prefs.editbox.aspx.  

        RawRequest = Get_Raw_RequestData(Request.InputStream)
        If Debug Then
            Send("<!-- Raw data:")
            Send(RawRequest)
            Send("-->")
        End If

        MyCommon.QueryStr = "select isnull(HTTP_RootURI, '') as HTTP_RootURI from Integrations where IntegrationID=1;"
        dst = MyCommon.LRT_Select
        If dst.Rows.Count > 0 Then
            HostURI = dst.Rows(0).Item("HTTP_RootURI")
        End If
        dst = Nothing

        HostURI = Trim(HostURI)
        If HostURI = "" Then
            Send(Copient.PhraseLib.Lookup("user-edit.PrefManNotSet", LanguageID))
        Else
            If Not (Right(HostURI, 1) = "/") Then HostURI = HostURI & "/"
            If Debug Then Send("<!-- HostURI=" & HostURI & " -->")
            TargetAddress = HostURI & "UI/user.edit.rolesbox.aspx"
            If Debug Then Send("<! -- TargetAddress=" & TargetAddress & " -->")
            'Open_UI_Box
            Send("<div class=""box"" id=""epmrolesbox"">")
            Send("<h2>")
            Send("  <span>")
            Sendb(Copient.PhraseLib.Lookup("term.epm", LanguageID) & " " & Copient.PhraseLib.Lookup("term.roles", LanguageID))
            Send("  </span>")
            Send("</h2>")

            FormData = "AuthToken=" & HttpUtility.UrlEncode(Request.Cookies("AuthToken").Value) & "&ParentURI=user-edit.aspx&" & RawRequest
            Send(ConnInc.Retrieve_HttpResponse(TargetAddress, FormData))
            'Close_UI_Box()
            Send("</div><!-- epmrolesbox -->")
        End If
        ConnInc = Nothing

    End Sub


    '-------------------------------------------------------------------------------------------------------------  


    Sub Send_EPM_Alerts()

        Dim FormData As String
        Dim RawRequest As String
        Dim HostURI As String = ""
        Dim TargetAddress As String
        Dim dst As DataTable
        Dim ConnInc As New Copient.ConnectorInc()

        'The code for the contents of this box lives in customer.prefs.editbox.aspx.  

        RawRequest = Get_Raw_RequestData(Request.InputStream)
        If Debug Then
            Send("<!-- Raw data:")
            Send(RawRequest)
            Send("-->")
        End If

        MyCommon.QueryStr = "select isnull(HTTP_RootURI, '') as HTTP_RootURI from Integrations where IntegrationID=1;"
        dst = MyCommon.LRT_Select
        If dst.Rows.Count > 0 Then
            HostURI = dst.Rows(0).Item("HTTP_RootURI")
        End If
        dst = Nothing

        HostURI = Trim(HostURI)
        If HostURI = "" Then
            Send(Copient.PhraseLib.Lookup("user-edit.PrefManNotSet", LanguageID))
        Else
            If Not (Right(HostURI, 1) = "/") Then HostURI = HostURI & "/"
            If Debug Then Send("<!-- HostURI=" & HostURI & " -->")
            TargetAddress = HostURI & "UI/user.edit.alertsbox.aspx"
            If Debug Then Send("<! -- TargetAddress=" & TargetAddress & " -->")
            'Open_UI_Box
            Send("<div class=""box"" id=""epmalertsbox"">")
            Send("<h2>")
            Send("  <span>")
            Sendb(Copient.PhraseLib.Lookup("term.epm", LanguageID) & " " & Copient.PhraseLib.Lookup("term.alerts", LanguageID))
            Send("  </span>")
            Send("</h2>")

            FormData = "AuthToken=" & HttpUtility.UrlEncode(Request.Cookies("AuthToken").Value) & "&ParentURI=user-edit.aspx&" & RawRequest
            Send(ConnInc.Retrieve_HttpResponse(TargetAddress, FormData))
            'Close_UI_Box()
            Send("</div><!-- epmalertsbox -->")
        End If
        ConnInc = Nothing

    End Sub


    '---------------------------------------------------------------------  

    Sub Save_EPM_Alerts()

        Dim FormData As String
        Dim RawRequest As String
        Dim HostURI As String = ""
        Dim TargetAddress As String
        Dim dst As DataTable
        Dim ConnInc As New Copient.ConnectorInc()

        'The code for the contents of this procedure lives in EPM user.edit.alertsbox.aspx.  

        RawRequest = Get_Raw_RequestData(Request.InputStream)
        If Debug Then
            Send("<!-- Raw data:")
            Send(RawRequest)
            Send("-->")
        End If

        MyCommon.QueryStr = "select isnull(HTTP_RootURI, '') as HTTP_RootURI from Integrations where IntegrationID=1;"
        dst = MyCommon.LRT_Select
        If dst.Rows.Count > 0 Then
            HostURI = dst.Rows(0).Item("HTTP_RootURI")
        End If
        dst = Nothing

        HostURI = Trim(HostURI)
        If HostURI = "" Then
            Send(Copient.PhraseLib.Lookup("user-edit.PrefManNotSet", LanguageID))
        Else
            If Not (Right(HostURI, 1) = "/") Then HostURI = HostURI & "/"
            If Debug Then Send("<!-- HostURI=" & HostURI & " -->")
            TargetAddress = HostURI & "UI/user.edit.alertsbox.aspx"
            If Debug Then Send("<! -- TargetAddress=" & TargetAddress & " -->")
            FormData = "AuthToken=" & HttpUtility.UrlEncode(Request.Cookies("AuthToken").Value) & "&ParentURI=user-edit.aspx&" & RawRequest
            Send(ConnInc.Retrieve_HttpResponse(TargetAddress, FormData))
            'Save Display Page 
            '**************************************************************************************************************
            TargetAddress = HostURI & "UI/user.display.setting.aspx"
            If Debug Then Send("<! -- TargetAddress=" & TargetAddress & " -->")
            FormData = "AuthToken=" & HttpUtility.UrlEncode(Request.Cookies("AuthToken").Value) & "&ParentURI=user-edit.aspx&" & RawRequest
            Send(ConnInc.Retrieve_HttpResponse(TargetAddress, FormData))
            '**************************************************************************************************************
        End If


    End Sub

    '---------------------------------------------------------------------  
    Sub Send_EPM_Display()

        Dim FormData As String
        Dim RawRequest As String
        Dim RawURI As String = ""
        Dim HostURI As String = ""
        Dim TargetAddress As String
        Dim dst As DataTable
        Dim ConnInc As New Copient.ConnectorInc()

        'The code for the contents of this box lives in customer.prefs.editbox.aspx.  

        RawRequest = Get_Raw_RequestData(Request.InputStream)
        If Debug Then
            Send("<!-- Raw data:")
            Send(RawRequest)
            Send("-->")
        End If

        MyCommon.QueryStr = "select isnull(HTTP_RootURI, '') as HTTP_RootURI from Integrations where IntegrationID=1;"
        dst = MyCommon.LRT_Select
        If dst.Rows.Count > 0 Then
            HostURI = dst.Rows(0).Item("HTTP_RootURI")
        End If
        dst = Nothing

        HostURI = Trim(HostURI)
        If HostURI = "" Then
            Send(Copient.PhraseLib.Lookup("user-edit.PrefManNotSet", LanguageID))
        Else
            If Not (Right(HostURI, 1) = "/") Then HostURI = HostURI & "/"
            If Debug Then Send("<!-- HostURI=" & HostURI & " -->")
            TargetAddress = HostURI & "UI/user.display.setting.aspx"
            If Debug Then Send("<! -- TargetAddress=" & TargetAddress & " -->")
            'Open_UI_Box
            Send("<div class=""box"" id=""epmsavesbox"">")
            Send("<h2>")
            Send("  <span>")
            Sendb(Copient.PhraseLib.Lookup("term.epm", LanguageID) & " " & Copient.PhraseLib.Lookup("term.display", LanguageID))
            Send("  </span>")
            Send("</h2>")
            FormData = "AuthToken=" & HttpUtility.UrlEncode(Request.Cookies("AuthToken").Value) & "&ParentURI=user-edit.aspx&" & RawRequest
            Send(ConnInc.Retrieve_HttpResponse(TargetAddress, FormData))
            'Close_UI_Box() 
            Send("</div><!-- epmalertsbox -->")
        End If
        ConnInc = Nothing
    End Sub
    '----------------------------------------------------------------------------------------------------------------------

    Sub LoadHealthAlerts(ByRef AlertTypeID As Integer)
        MyCommon.QueryStr = "select distinct AT.AlertTypeID, AT.PhraseID from AlertTypes AT with (NoLock)  " & _
                 "inner join AlertTypeEngines ATE with (NoLock) on ATE.AlertTypeID = AT.AlertTypeID and AT.AlertTypeID="&AlertTypeID & _
                 "where EngineID in (select EngineID from PromoEngines with (NoLock) where Installed = 1) " & _
                 "order by AlertTypeId;"
        dst = MyCommon.LRT_Select
        MyCommon.QueryStr = "select RegHealthSeverityID As Reg, AlertHealthSeverityID As Alert from AlertReceivers with (NoLock) where AdminUserID=" & u_UID & " and AlertTypeID=" & AlertTypeID & ";"
        AHSrst = MyCommon.LRT_Select

        If dst.Rows.Count > 0 Then

            If Not alreadyAddedLabels Then
                Send("<br /><label for=""stdemailHS" & AlertTypeID.ToString() & """ ><b>" & Copient.PhraseLib.Lookup("term.reg", LanguageID) & "</b></label>")
                Sendb("<label for=""alertemailHS" & AlertTypeID.ToString() & """ ><span style=""margin-left:55px;""><b>" & Copient.PhraseLib.Lookup("term.alert", LanguageID) & "</b></span></label><br />")
            Else
                Sendb("<br /><br class=""half"" />")
            End If

            alreadyAddedLabels = True

            'Reg alert level
            Send("<select name=""reglevel" & AlertTypeID.ToString() & """ id=""stdemailHS" & AlertTypeID.ToString() & """" & IIf(EditIdentity, "", " disabled=""disabled""") & ">")
            MyCommon.QueryStr = "select HealthSeverityID, Description, PhraseID from LS_HealthSeverityTypes with (NoLock) "
            HSrst = MyCommon.LWH_Select
            For Each HSrow In HSrst.Rows
                Sendb("<option id=""HSID-" & HSrow.Item("HealthSeverityID") & "std""")
                If EditIdentity = False Then Sendb(" disabled=""disabled""")
                If (AHSrst.Rows.Count > 0) Then
                    If (HSrow.Item("HealthSeverityID") = AHSrst.Rows(0).Item("Reg")) Then
                        Sendb(" selected=""selected""")
                    End If
                End If
                Send(" name=""HSID-" & HSrow.Item("HealthSeverityID") & "std"" value=""" & HSrow.Item("HealthSeverityID") & """>" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(HSrow.Item("PhraseID"), 0)), LanguageID, HSrow.Item("Description").ToString) & "</option>")
            Next
            'Alert level
            Send("</select>")
            Send("<select name=""alertlevel" & AlertTypeID.ToString() & """ id=""alertemailHS" & AlertTypeID.ToString() & """" & IIf(EditIdentity, "", " disabled=""disabled""") & ">")
            MyCommon.QueryStr = "select HealthSeverityID, Description, PhraseID from LS_HealthSeverityTypes with (NoLock)"
            HSrst = MyCommon.LWH_Select
            For Each HSrow In HSrst.Rows
                Sendb("<option id=""HSID-" & HSrow.Item("HealthSeverityID") & "Alert""")
                If EditIdentity = False Then Sendb(" disabled=""disabled""")
                If (AHSrst.Rows.Count > 0) Then
                    If (HSrow.Item("HealthSeverityID") = AHSrst.Rows(0).Item("Alert")) Then
                        Sendb(" selected=""selected""")
                    End If
                End If
                Send(" name=""HSID-" & HSrow.Item("HealthSeverityID") & "Alert"" value=""" & HSrow.Item("HealthSeverityID") & """>" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(HSrow.Item("PhraseID"), 0)), LanguageID, HSrow.Item("Description").ToString) & "</option>")
            Next
            Send("</select>")
            Sendb("&nbsp;&nbsp;" & Copient.PhraseLib.Lookup(dst.Rows(0).Item("PhraseID"), LanguageID))
        End If
    End Sub
</script>
<%


    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "user-edit.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixWH()
    If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
        MyCommon.Open_PrefManRT()
    End If
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    CMS.AMS.CurrentRequest.Resolver.AppName = MyCommon.AppName
    infoMessage = ""
    u_UID = 0
    If Request.RequestType = "GET" Then
        u_UID = Server.HtmlEncode(Request.QueryString("UserID"))
    Else
        u_UID = Server.HtmlEncode(Request.Form("UserID"))
        If u_UID = 0 Then
            u_UID = Server.HtmlEncode(Request.QueryString("UserID"))
        End If
    End If


    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    ' lets check the logged in user and see if they are to be restricted to this page
    MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                        "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                        "where AU.AdminUserID=" & AdminUserID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        If (rst.Rows(0).Item("prestrict") = True) Then
            ' ok we got in here then we need to restrict the user from seeing any other pages
            Restricted = True
        End If
    End If

    infoMessage = Server.HtmlEncode(Request.QueryString("infoMessage"))

    If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(131), maxStoreUser) Then
        maxStoreUser = 0
    End If

    Send_HeadBegin("term.user", , u_UID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld, Restricted)
    Send_Scripts()
%>
<script type="text/javascript" language="javascript">
  var UserID = <%Sendb(u_UID)%>
  function LoadOfferBkgrd() {
    var elem = document.getElementById('uploader');

    if (elem != null) {
      elem.style.display = 'block';
    }
  }

  function LoadColorPalette() {
    var elem = document.getElementById("colorPalette");

    if (elem != null) {
      elem.style.display = "block";
    }
  }
//  function chooseFile() {
//      document.getElementById("browse").click();
//  }
//  function fileonclick() {
//      var filename = document.getElementById("browse").value;
//      document.getElementById("lblfileupload").innerText = filename.replace("C:\\fakepath\\", "");
//  }
  function populateColor(red, green, blue) {
    var elem = document.getElementById("pref3");
    var displayElem = document.getElementById("displayPref3");
    var clrElem = document.getElementById("colorPalette");
    var clr = RGB2HTML(red, green, blue);

    if (elem != null) {
      elem.value = clr;
    }

    if (displayElem != null) {
      displayElem.style.backgroundColor = "#" + clr;
    }

    if (clrElem != null) {
      clrElem.style.display = "none";
    }
  }

  function RGB2HTML(red, green, blue) {
    var hexRed = red.toString(16);
    var hexGreen = green.toString(16);
    var hexBlue = blue.toString(16);

    if (hexRed == "0") hexRed = "00";
    if (hexGreen == "0") hexGreen = "00";
    if (hexBlue == "0") hexBlue = "00";

    return hexRed + '' + hexGreen + '' + hexBlue;
  }

  function isValidPath() {
    var retVal = true;
    var frmElem = document.uploadform.browse
    var agt = navigator.userAgent.toLowerCase();
    var browser = '<% Sendb(Request.Browser.Browser) %>'

    if (browser == 'IE') {
      if (frmElem != null) {
        var filePath = frmElem.value

        if (filePath.length >= 2) {
          if (filePath.charAt(1) != ":") {
            alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidfilepath", LanguageID))%>');
            retVal = false;
          }
        } else {
          alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidfilepath", LanguageID))%>');
          retVal = false;
        }
      }
    }
    return retVal;
  }
   function StoreUserPopup()
   {
      var StoreUser = document.getElementById("storeuser").checked;
      var maxlocations = <%Sendb(maxStoreUser)%>
      if(StoreUser == true)
      {
        toggleDialog('foldercreate', true);

        $('#userLoc').empty();
        
        var table = document.getElementById("storelocations");
        var len = table.rows.length;
        for(var i=1;i<len; i++) //start at 1 to skip the header
        {
            var row = table.rows[i];
            var text = row.cells[0].innerText.substring(2); //get rid of the bullet
            var o = new Option(text, "userLoc-"+ row.cells[0].id.substring(6)); // Store-[LocationID] 
            // jquerify the DOM object 'o' so we can use the html method
            $(o).html(text);
            $("#userLoc").append(o);
        }
        
	if($('#userLoc > option').length >=maxlocations)
        {
          $('#storeselect').attr("disabled", true);
        }
	else
	{
	  $('#storeselect').attr("disabled", false);
        }
        
     }
     else
     {
        if(confirm('Are you sure you want to remove all locations for this store user?') == true)
        {
          $("#userLoc").empty();
          saveLocations();
        }
        else
          document.getElementById("storeuser").checked = true; 
     }
   }
   
 function locationSearch()
 {
    var xmlHttpReq = false;
    var self = this;
    var strURL ;
    var searchoption=1;
    
    // Mozilla/Safari
    if (window.XMLHttpRequest) {
      self.xmlHttpReq = new XMLHttpRequest();
    }
    // IE
    else { //if (window.ActiveXObject)
      self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    
    if($("#searchradio1").is(":checked") == true)
      searchoptions= 1;
    else if ($("#searchradio2").is(":checked") == true)
      searchoptions = 2;

    strURL = "OfferFeeds.aspx?mode=StoreUserLocationSearch&UserID=" + UserID + "&searchterms=" + $("#locationsearch").val() + "&searchoption=" +searchoptions
          
    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function() {
      if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
       // $("#storeLoc").empty();
         var resp =self.xmlHttpReq.responseText;
        var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;
        if(!isFireFox){
            document.getElementById("storeLocDiv").innerHTML = '<select id = "storeLoc" size="10" style="width:150px">' + resp + '</select>';
        }
        else
          document.getElementById("storeLoc").innerHTML = resp;
          removeUsed();
          if($("#storeLoc >options").length >=100)
            document.getElementById("maxstores").style.display = 'block';
          else
            document.getElementById("maxstores").style.display = 'none';
          
      }
    }
    self.xmlHttpReq.send();
 }

 function selectItem(source) 
 {
    var elemSource = document.getElementById(source);
    var elemDest = document.getElementById("userLoc");   
    var selOption = null;
    var selText ="", selVal = "";
    var selIndex = -1;
    var maxlocations = <%Sendb(maxStoreUser)%>
        
    if (elemSource != null && elemSource.options.selectedIndex == -1) {
      //alert('<% Sendb(Copient.PhraseLib.Lookup("CPE-rew-discounts.selectproducts", LanguageID)) %>');
      alert('Please select an item from the available stores.');
      elemSource.focus();
    } else {
      selIndex = elemSource.options.selectedIndex;
      selOption = elemSource.options[selIndex];
      selText = selOption.text;
      selVal = selOption.value.substring(9);
      
      var o = new Option(selText, "userLoc-"+ selVal);
      // jquerify the DOM object 'o' so we can use the html method
      $(o).html(selText);
      $("#userLoc").append(o);
       //removeUsed(false);
      locationSearch();
      if($('#userLoc > option').length >=maxlocations)
      {
        $('#storeselect').attr("disabled", true);
      }
    }
  }

  function deselectItem(source) 
  {
    var elemSource = document.getElementById(source);
    var selIndex =  -1;
    var selVal = null;
    var maxlocations = <%Sendb(maxStoreUser)%>

    if (elemSource != null && elemSource.options.selectedIndex == -1) {
      alert('Please select an item from the available stores.');
      elemSource.focus();
    } else {
      selIndex = elemSource.options.selectedIndex;
      selOption = elemSource.options[selIndex];
      selText = selOption.text;
      selVal = selOption.value.substring(8);

      var o = new Option(selText, "storeLoc-"+ selVal);
      // jquerify the DOM object 'o' so we can use the html method
      $(o).html(selText);
      $("#storeLoc").append(o);
      
      elemSource.options[selIndex] = null;
      locationSearch();
      if($('#userLoc > option').length < maxlocations)
      {
        $('#storeselect').attr("disabled", false);
      }
    }   
  }
  
  function removeUsed() 
  {
    // this function will remove items from the functionselect box that are used in 
    // selected and excluded boxes
  
    var funcSel = document.getElementById('storeLoc');
    var exSel = document.getElementById('userLoc');
    var elSel = document.getElementById('excluded');
    var i,j;
    if(exSel != null)
    {
      for (i = exSel.length - 1; i>=0; i--) {
        for(j=funcSel.length-1;j>=0; j--) {
          if(funcSel.options[j].text == exSel.options[i].text){
            funcSel.options[j] = null;
          }
        }
      }
    }
}

function saveLocations()
{
  var locationList ="";
  $("#userLoc > option").each(function() {
    locationList += "(" + UserID + "," + this.value.substring(8) + "),"; });

  var xmlHttpReq = false;
  var self = this;
  var strURL ;

  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  strURL = "OfferFeeds.aspx?mode=SaveStoreUserLocations&UserID=" + UserID + "&LocationList="+locationList;
           
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      //IE9 doesnt allow to edit html using the .html() function. 
      //the work around is to empty the element and append id.
      $("#storelocations > tbody").empty();
      $("#storelocations > tbody").append(self.xmlHttpReq.responseText);
      if(self.xmlHttpReq.responseText != "" ) 
      {
        document.getElementById("storelocations").style.display = 'block';
        document.getElementById("editstoreuser").style.display = 'block';
      }
      toggleDialog('foldercreate',false);
    }
  }
  self.xmlHttpReq.send();
}
  
function toggleDialog(elemName, shown) 
{
  var elem = document.getElementById(elemName);
  var fadeElem = document.getElementById('StoreUserFadeDiv');
    
  if (elem != null) {
    if(shown)
    {
      elem.style.display =  'block' ;
      locationSearch(); //populate the store locations
    }
    else{
      elem.style.display = 'none';
      document.getElementById("locationsearch").value = "";
      var defaultSeachOption = <%Sendb(MyCommon.Fetch_SystemOption(175))%>;
      if(defaultSeachOption == '1' )
          document.getElementById("searchradio1").checked= true;
      else if(defaultSeachOption == '2')
          document.getElementById("searchradio2").checked= true;
      if($("#storelocations tbody tr").length == 0 )
      {
        document.getElementById("storeuser").checked = false;
        document.getElementById("storelocations").style.display = 'none';
        document.getElementById("editstoreuser").style.display = 'none';
        
      }
    }
    
  }
  if (fadeElem != null) {
    fadeElem.style.display = (shown) ? 'block' : 'none';
  }
}
    
</script>
<style type="text/css">
#StoreUserFadeDiv {
  background-color: #e0e0e0;
  position: absolute;
  top: 0px;
  left: 0px;
  width:5000px;
  height:5000px;
  z-index: 99;
  display: none;
  opacity: .4;
  filter: alpha(opacity=40);
  }
</style>
<%  
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    If (Not Restricted) Then
        Send_Tabs(Logix, 8)
        Send_Subtabs(Logix, 8, 9)
    Else
        Send_Subtabs(Logix, 93, 3)
    End If

    Dim row As DataRow
    Try
        If Request.Files.Count >= 1 Then
            File = Request.Files.Get(0)
            If File.ContentType.IndexOf("image") > -1 AndAlso File.FileName.Trim <> "" Then
                ' save it to the graphics path first
                GraphicPath = MyCommon.Fetch_SystemOption(47)
                If (GraphicPath = "") Then
                    GraphicPath = DEFAULT_GRAPHIC_PATH
                End If
                If Not (Right(GraphicPath, 1) = "\") Then
                    GraphicPath = GraphicPath & "\"
                End If
                If System.IO.Directory.Exists(GraphicPath) Then
                    File.SaveAs(GraphicPath & File.FileName)
                End If

                ' save it to the web server the client is currently using
                UserImagePath = MyCommon.Get_Install_Path
                If (Right(UserImagePath, 1) <> "\") Then UserImagePath &= "\"
                UserImagePath &= "LogixWebRoot\images\users"
                If System.IO.Directory.Exists(UserImagePath) Then
                    'File.SaveAs(UserImagePath & "\" & File.FileName)

                    MyCommon.QueryStr = "Delete AdminUserPreferences with (RowLock) where AdminUserID=" & u_UID & " and PreferenceID=1"
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "Insert into AdminUserPreferences with (RowLock) (AdminUserID, PreferenceID, Value) values (" & u_UID & ", 1, '" & File.FileName & "');"
                    MyCommon.LRT_Execute()
                Else
                    infoMessage = Copient.PhraseLib.Detokenize("user-edit.InvalidImageDir", LanguageID, UserImagePath)
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("user-edit.InvalidImageFormat", LanguageID)
            End If
        End If
    Catch ex As Exception
        infoMessage = ex.ToString
    End Try

    ' any search terms inbound?
    If (Request.Form("Delete") <> "") Then
        If (MyCommon.Extract_Val(Request.Form("UserID")) <> "") Then
            If (AdminUserID = MyCommon.Extract_Val(Server.HtmlEncode(Request.Form("UserID"))) AndAlso Logix.IsAdministrator(AdminUserID, MyCommon)) Then  'See if the user is an Administrator and is trying to delete their own account
                infoMessage = Copient.PhraseLib.Lookup("user-edit.adminselfdelete", LanguageID)  'An Administrator may not delete their own account.
            Else
                If (Logix.UserRoles.DeleteAdminUsers = True) Then
                    MyCommon.QueryStr = "DELETE FROM AdminUsers with (RowLock) WHERE AdminUserID=" & MyCommon.Extract_Val(Server.HtmlEncode(Request.Form("UserID")))
                    u_UID = MyCommon.Extract_Val(Request.Form("UserID"))
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "delete from AlertReceivers with (RowLock) where AdminUserID=" & MyCommon.Extract_Val(Server.HtmlEncode(Request.Form("UserID")))
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "delete from AdminUserRoles with (RowLock) where AdminUserID=" & MyCommon.Extract_Val(Server.HtmlEncode(Request.Form("UserID")))
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "delete from AdminUserOffers with (RowLock) where AdminUserID=" & MyCommon.Extract_Val(Server.HtmlEncode(Request.Form("UserID")))
                    MyCommon.LRT_Execute()
                    MyCommon.Activity_Log(23, u_UID, AdminUserID, Copient.PhraseLib.Lookup("history.user-delete", LanguageID))
                    MyCommon.Activity_Log(23, AdminUserID, u_UID, Copient.PhraseLib.Lookup("history.user-deletedby", LanguageID))
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "user-list.aspx")
                    GoTo done
                Else
                    infoMessage = Copient.PhraseLib.Lookup("user-edit.nopermission", LanguageID)
                    GoTo done
                End If
            End If
        End If

    ElseIf (Request.Form("save") <> "") Then
        ' somebody clicked save
        infoMessage = ""
        Dim rst1 As DataTable
        Dim rst12 As DataTable
        CurrentRequest.Resolver.AppName = "user-edit.aspx"
        Dim m_adminUserDataService As IAdminUserData = CurrentRequest.Resolver.Resolve(Of IAdminUserData)()
        u_UID = MyCommon.Extract_Val(Server.HtmlEncode(Request.Form("UserID")))
        u_Username = Logix.TrimAll(Request.Form("username"))
        u_Employeeid = Logix.TrimAll(Request.Form("employeeid"))
        Dim u_Passwordvalid = m_adminUserDataService.ValidatePassword(Request.Form("password"), u_Username, LanguageID)
        Dim blnNewUser As Boolean = False
        ' need to check if its a new user and create one
        If (u_UID = 0) Then
            blnNewUser = True
            MyCommon.QueryStr = "Select UserName from AdminUsers with (NoLock) where UserName = @UserName and AdminUserID <> @AdminUserID"
            MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar, 50).Value = u_Username
            MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = u_UID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            MyCommon.QueryStr = "Select employeeid from AdminUsers with (NoLock) where employeeid = @employeeid and AdminUserID <> @AdminUserID and not(EmployeeID='')"
            MyCommon.DBParameters.Add("@employeeid", SqlDbType.NVarChar, 50).Value = u_Employeeid
            MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = u_UID
            rst1 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            MyCommon.QueryStr = "dbo.pt_AdminUsers_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@UserName ", SqlDbType.NVarChar, 50).Value = u_Username
            MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Direction = ParameterDirection.Output

            Dim IsEmployeeIdAlphaNumeric As Boolean
            If System.Text.RegularExpressions.Regex.IsMatch(u_Employeeid, "^[a-zA-Z0-9]+$") Then
                IsEmployeeIdAlphaNumeric = True
            Else
                IsEmployeeIdAlphaNumeric = False
            End If

            If (u_Username = "") Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.noname", LanguageID)
            ElseIf (Request.Form("firstname") = "" AndAlso Request.Form("lastname") = "") Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.noname", LanguageID)
            ElseIf (Request.Form("password") <> Request.Form("passwordagain")) Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.badpasswords", LanguageID)
            ElseIf (Request.Form("password") = "") Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.nopassword", LanguageID)
            ElseIf Not (u_Passwordvalid.ResultType = AMSResultType.Success) Then
                infoMessage = Copient.PhraseLib.Lookup("term.Passwordchange", LanguageID)
            ElseIf (rst.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.nameused", LanguageID)
            ElseIf (u_Employeeid <> "" AndAlso Not IsEmployeeIdAlphaNumeric) Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.employeealphanumeric", LanguageID)
            ElseIf (rst1.Rows.Count > 0) Then
                infoMessage = Copient.PhraseLib.Lookup("adminuser-employeeid-alreadyexists", LanguageID)
            Else
                MyCommon.LRTsp.ExecuteNonQuery()
                u_UID = MyCommon.LRTsp.Parameters("@AdminUserID").Value
                If (u_UID = -1) Then
                Else
                    ' Record history
                    MyCommon.Activity_Log(23, u_UID, AdminUserID, Copient.PhraseLib.Lookup("history.user-create", LanguageID))
                    MyCommon.Activity_Log(23, AdminUserID, u_UID, Copient.PhraseLib.Lookup("history.user-createdby", LanguageID))
                    ' Apply favored offers
                    MyCommon.QueryStr = "select OIDS.OfferID, OIDS.EngineID from OfferIDs as OIDS with (NoLock) " & _
                                        "left join Offers as CM with (NoLock) on CM.OfferID=OIDS.OfferID " & _
                                        "left join CPE_Incentives as CPE with (NoLock) on CPE.IncentiveID=OIDS.OfferID " & _
                                        "where (CM.Favorite = 1 or CPE.Favorite = 1) " & _
                                        "order by EngineID, OIDS.OfferID;"
                    rst = MyCommon.LRT_Select
                    x = 0
                    For Each row In rst.Rows
                        MyCommon.QueryStr = "insert into AdminUserOffers (AdminUserID, OfferID, Priority, FavoredBy, FavoredDate) " & _
                                            "values (" & u_UID & ", " & rst.Rows(x).Item("OfferID") & ", 1, " & AdminUserID & ", '" & Now() & "');"
                        MyCommon.LRT_Execute()
                        x = x + 1
                    Next
                End If
            End If
            MyCommon.Close_LRTsp()
            If infoMessage <> "" Then
                Response.Redirect("user-edit.aspx?new=New&infoMessage=" & infoMessage & _
                                  If(Request.Form("username") <> "", "&username=" & u_Username, "") & _
                                  If(Request.Form("firstname") <> "", "&firstname=" & Server.HtmlEncode(Request.Form("firstname")), "") & _
                                          If(Request.Form("lastname") <> "", "&lastname=" & Server.HtmlEncode(Request.Form("lastname")), "") & _
                                           If(Request.Form("employeeid") <> "", "&employeeid=" & Server.HtmlEncode(Request.Form("employeeid")), ""))

            End If
        End If
        If (u_UID <> -1) Then
            Dim IsEmployeeIdAlphaNumeric As Boolean
            If System.Text.RegularExpressions.Regex.IsMatch(Request.Form("employeeId"), "^[a-zA-Z0-9]+$") Then
                IsEmployeeIdAlphaNumeric = True
            Else
                IsEmployeeIdAlphaNumeric = False
            End If
            Dim IsSecondFactorEnable As Integer
            IsSecondFactorEnable = MyCommon.Fetch_SystemOption(298)
            If (u_Username = "") Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.noname", LanguageID)
            ElseIf (Request.Form("AlertEmail") <> "" AndAlso MyCommon.EmailAddressCheck(Request.Form("AlertEmail")) = False) Then
                infoMessage = Copient.PhraseLib.Lookup("emailValidation", LanguageID)
            ElseIf (IsSecondFactorEnable = 1 AndAlso Request.Form("email") = "" AndAlso blnNewUser = False) Then
                infoMessage = Copient.PhraseLib.Lookup("term.provideemail", LanguageID)
            ElseIf (Request.Form("email") <> "" AndAlso MyCommon.EmailAddressCheck(Request.Form("email")) = False) Then
                infoMessage = Copient.PhraseLib.Lookup("emailValidation", LanguageID)
            ElseIf (Request.Form("firstname") = "" AndAlso Request.Form("lastname") = "") Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.noname", LanguageID)
            ElseIf (Request.Form("employeeId") <> "" AndAlso Not IsEmployeeIdAlphaNumeric) Then
                infoMessage = Copient.PhraseLib.Lookup("user-edit.employeealphanumeric", LanguageID)
            ElseIf (Logix.UserRoles.UpdateOthersInfo AndAlso u_UID <> AdminUserID) OrElse (Logix.UserRoles.UpdateOwnInfo AndAlso u_UID = AdminUserID) OrElse (Logix.UserRoles.CreateAdminUsers = True AndAlso blnNewUser) Then
                If ((Request.Form("password") = Request.Form("passwordagain") AndAlso Request.Form("password") <> "") OrElse (Request.Form("password") = "" AndAlso Request.Form("passwordagain") = "")) Then
                    If Not (u_Passwordvalid.ResultType = AMSResultType.Success) And Request.Form("password") <> "" Then
                        infoMessage = u_Passwordvalid.MessageString
                    Else

                        MyCommon.QueryStr = "Select UserName from AdminUsers with (NoLock) where UserName = @UserName and AdminUserID <> @AdminUserID"
                        MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar, 50).Value = u_Username
                        MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = u_UID
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        MyCommon.QueryStr = "Select employeeid from AdminUsers with (NoLock) where employeeid = @employeeid and AdminUserID <> @AdminUserID and not(EmployeeID='')"
                        MyCommon.DBParameters.Add("@employeeid", SqlDbType.NVarChar, 50).Value = u_Employeeid
                        MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = u_UID
                        rst12 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        For Each row In rst.Rows
                            Send("<input type=""hidden"" id=""username_holder"" name=""username_holder"" value=""" & row.Item("UserName") & """ />")
                        Next
                        If (rst.Rows.Count = 0) Then
                            If (rst12.Rows.Count = 0) Then
                                Dim rst13 As DataTable
                                MyCommon.QueryStr = "Select UserName, Password from AdminUsers with (NoLock) where UserName = @UserName"
                                MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar, 50).Value = If(blnNewUser, u_Username, Logix.TrimAll(Server.HtmlEncode(Request.Form("username_hold"))))
                                rst13 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                Dim o_USalt As String = ""

                                If (blnNewUser) Then
                                    o_USalt = MyCommon.GetUserSalt(MyCommon, u_Username)
                                Else
                                    o_USalt = MyCommon.GetUserSalt(MyCommon, Logix.TrimAll(Server.HtmlEncode(Request.Form("username_hold"))))
                                End If

                                If (MyCommon.NZ(rst13.Rows(0).Item("Password"), "") <> HashLib.SQL_LoginHash(Server.HtmlEncode(Request.Form("password")), o_USalt)) Then
                                    sqlBuf.Append("Update AdminUsers with (RowLock) set ")
                                    If u_Username = "" Then
                                        sqlBuf.Append("UserName=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("username_hold")))) & "',")
                                    Else
                                        sqlBuf.Append("UserName=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(u_Username)) & "',")
                                    End If
                                    If MyCommon.Parse_Quotes(Logix.TrimAll(Request.Form("firstname"))) <> "" OrElse (MyCommon.Parse_Quotes(Logix.TrimAll(Request.Form("firstname"))) = "" AndAlso MyCommon.Parse_Quotes(Logix.TrimAll(Request.Form("lastname"))) <> "") Then
                                        sqlBuf.Append("FirstName=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("firstname")))) & "',")
                                    End If
                                    If MyCommon.Parse_Quotes(Logix.TrimAll(Request.Form("lastname"))) <> "" OrElse (MyCommon.Parse_Quotes(Logix.TrimAll(Request.Form("lastname"))) = "" AndAlso MyCommon.Parse_Quotes(Logix.TrimAll(Request.Form("firstname"))) <> "") Then
                                        sqlBuf.Append("LastName=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("lastname")))) & "',")
                                    End If
                                    If MyCommon.Parse_Quotes(Request.Form("employeeid")) <> "" Then
                                        sqlBuf.Append("EmployeeID=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("employeeid")))) & "',")
                                    Else
                                        sqlBuf.Append("EmployeeID=N'',")
                                    End If
                                    If MyCommon.Parse_Quotes(Request.Form("jobtitle")) <> "" Then
                                        sqlBuf.Append("JobTitle=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("jobtitle")))) & "',")
                                    Else
                                        sqlBuf.Append("JobTitle=N'',")
                                    End If
                                    If MyCommon.Parse_Quotes(Request.Form("employer")) <> "" Then
                                        sqlBuf.Append("Employer=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("employer")))) & "',")
                                    Else
                                        sqlBuf.Append("Employer=N'',")
                                    End If
                                    MyCommon.QueryStr = sqlBuf.ToString
                                    If (Request.Form("password") <> "") Then
                                        'Generate a Salt
                                        Dim USalt As String = HashLib.GenerateNewSalt()
                                        MyCommon.QueryStr = MyCommon.QueryStr & "password=N'" & HashLib.SQL_LoginHash(Server.HtmlEncode(Request.Form("password")), USalt) & "',"
                                        MyCommon.QueryStr = MyCommon.QueryStr & "PasswordChangedDate=N'" & DateAndTime.Now & "',"
                                        MyCommon.QueryStr = MyCommon.QueryStr & "USalt=N'" & USalt & "',"
                                    End If
                                    MyCommon.QueryStr = MyCommon.QueryStr & "Email=N'" & MyCryptlib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("email"))))) & "'," & _
                                                        "AlertEmail=N'" & MyCryptlib.SQL_StringEncrypt(MyCommon.Parse_Quotes(Logix.TrimAll(Server.HtmlEncode(Request.Form("AlertEmail"))))) & "'," & _
                                                        "LanguageID='" & Server.HtmlEncode(Request.Form("language")) & "'," & _
                                                        "StartPageID='" & MyCommon.Parse_Quotes(Server.HtmlEncode(Request.Form("startpage"))) & "'," & _
                                                        "StyleID='" & MyCommon.Parse_Quotes(Server.HtmlEncode(Request.Form("style"))) & "' where AdminUserID=" & u_UID
                                    If (MyCommon.Parse_Quotes(Request.Form("username")) = "" AndAlso MyCommon.Parse_Quotes(Request.Form("username_hold")) = "") Then
                                        infoMessage = Copient.PhraseLib.Lookup("user-edit.noname", LanguageID)
                                    Else
                                        MyCommon.LRT_Execute()

                                        If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then Save_EPM_Alerts()

                                        ' Let's delete all alerts for the user and put in new values to make it simple.
                                        MyCommon.QueryStr = "delete from AlertReceivers with (RowLock) where AdminUserID=" & u_UID
                                        MyCommon.LRT_Execute()



                                        ' get AlertTypes for installed engines
                                        MyCommon.QueryStr = "select distinct AT.AlertTypeID from AlertTypes AT with (NoLock)  " & _
                                              "inner join AlertTypeEngines ATE with (NoLock) on ATE.AlertTypeID = AT.AlertTypeID and AT.AlertTypeID not in (10000,10004,10005) " & _
                                              "where EngineID in (select EngineID from PromoEngines with (NoLock) where Installed = 1) " & _
                                              "order by AlertTypeId;"
                                        dtAlertTypes = MyCommon.LRT_Select

                                        Dim alert As Integer
                                        Dim amail As Integer
                                        Dim dr As DataRow
                                        For Each dr In dtAlertTypes.Rows
                                            alert = 0
                                            amail = 0
                                            If (Request.Form("alert-" & dr.Item(0)) = "on") Then alert = 1
                                            If (Request.Form("amail-" & dr.Item(0)) = "on") Then amail = 1
                                            If (amail Or alert) Then
                                                MyCommon.QueryStr = "insert into AlertReceivers with (RowLock) (AlertTypeID,AdminUserID,StdEmail,AlertEmail) " & _
                                                                    "values(" & dr.Item(0) & "," & u_UID & "," & alert & "," & amail & ")"
                                                MyCommon.LRT_Execute()
                                            End If
                                        Next

                                        MyCommon.QueryStr = "insert into AlertReceivers with (RowLock) (AlertTypeID,AdminUserID,StdEmail,AlertEmail,RegHealthSeverityID,AlertHealthSeverityID) " & _
                                                            "values(10000," & u_UID & "," & IIf(Val(Request.Form("reglevel10000")) > 0, 1, 0) & "," & IIf(Val(Request.Form("alertlevel10000")) > 0, 1, 0) & "," & Val(Request.Form("reglevel10000")) & "," & Val(Request.Form("alertlevel10000")) & ")"
                                        MyCommon.LRT_Execute()

                                        MyCommon.QueryStr = "insert into AlertReceivers with (RowLock) (AlertTypeID,AdminUserID,StdEmail,AlertEmail,RegHealthSeverityID,AlertHealthSeverityID) " & _
                                                            "values(10004," & u_UID & "," & IIf(Val(Request.Form("reglevel10004")) > 0, 1, 0) & "," & IIf(Val(Request.Form("alertlevel10004")) > 0, 1, 0) & "," & Val(Request.Form("reglevel10004")) & "," & Val(Request.Form("alertlevel10004")) & ")"
                                        MyCommon.LRT_Execute()

                                        MyCommon.QueryStr = "insert into AlertReceivers with (RowLock) (AlertTypeID,AdminUserID,StdEmail,AlertEmail,RegHealthSeverityID,AlertHealthSeverityID) " & _
                                                            "values(10005," & u_UID & "," & IIf(Val(Request.Form("reglevel10005")) > 0, 1, 0) & "," & IIf(Val(Request.Form("alertlevel10005")) > 0, 1, 0) & "," & Val(Request.Form("reglevel10005")) & "," & Val(Request.Form("alertlevel10005")) & ")"
                                        MyCommon.LRT_Execute()


                                        MyCommon.Activity_Log(23, u_UID, AdminUserID, Copient.PhraseLib.Lookup("history.user-edit", LanguageID))

                                        ' Write a style cookie
                                        If u_UID = AdminUserID Then
                                            MyCommon.QueryStr = "select StyleID from AdminUsers with (NoLock) where AdminUserID=" & AdminUserID
                                            dst = MyCommon.LRT_Select
                                            If dst.Rows.Count > 0 Then
                                                Write_StyleCookie(MyCommon.NZ(dst.Rows(0).Item("StyleID"), 1))
                                            End If
                                        End If

                                        ' Store the user's preferences
                                        Try
                                            ' image tiling for offer wallpaper
                                            MyCommon.QueryStr = "delete from AdminUserPreferences with (RowLock) where AdminUserID=" & u_UID & " and PreferenceID=2;"
                                            MyCommon.LRT_Execute()
                                            MyCommon.QueryStr = "insert into AdminUserPreferences with (RowLock) (AdminUserID, PreferenceID, Value) " & _
                                                                "  values (" & u_UID & ", 2," & IIf(Request.Form("pref2") <> "", "1", "0") & ")"
                                            MyCommon.LRT_Execute()

                                            ' background color
                                            MyCommon.QueryStr = "delete from AdminUserPreferences with (RowLock) where AdminUserID=" & u_UID & " and PreferenceID=3;"
                                            MyCommon.LRT_Execute()
                                            MyCommon.QueryStr = "insert into AdminUserPreferences with (RowLock) (AdminUserID, PreferenceID, Value) " & _
                                                                "  values (" & u_UID & ", 3, '" & IIf(Request.Form("pref3") <> "", "#" & Request.Form("pref3"), "#ffffff") & "')"
                                            MyCommon.LRT_Execute()

                                        Catch ex As Exception
                                            ' ignore - optional processing
                                            Send(ex.ToString)
                                            GoTo done
                                        End Try

                                        Response.Redirect("user-edit.aspx?UserID=" & u_UID)
                                    End If
                                Else
                                    infoMessage = Copient.PhraseLib.Lookup("term.Passwordexists", LanguageID)
                                End If
                            Else
                                infoMessage = Copient.PhraseLib.Lookup("adminuser-employeeid-alreadyexists", LanguageID)
                            End If
                        Else
                            infoMessage = Copient.PhraseLib.Lookup("user-edit.nameused", LanguageID) & " (""" & Server.HtmlEncode(Request.Form("username")) & """)"
                        End If
                    End If
                Else
                    infoMessage = Copient.PhraseLib.Lookup("user-edit.badpasswords", LanguageID)
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("user-edit.denied", LanguageID)
            End If
        End If

    ElseIf (Request.Form("role-add") <> "" AndAlso Request.Form("roles-avail") <> "") Then
        Dim UserCanEditRole As Boolean = IIf((AdminUserID = u_UID), Logix.UserRoles.EditOwnRoles, Logix.UserRoles.EditRoles)
        If UserCanEditRole Then
            Dim selections() As String
            Dim i As Integer
            selections = Server.HtmlEncode(Request.Form("roles-avail")).Split(",")
            Dim AddRole As DataTable
            Dim Role As DataTable
            For i = 0 To selections.GetUpperBound(0)
                MyCommon.QueryStr = "select RoleID from AdminRoles where RoleID=" & selections(i) & ";"
                Role = MyCommon.LRT_Select()
                If Role IsNot Nothing AndAlso Role.Rows.Count > 0 Then 'check if role exists in database
                    MyCommon.QueryStr = "select RoleID from AdminUserRoles with (NoLock) where AdminUserID=" & u_UID & " and RoleID=" & selections(i) & ";"
                    AddRole = MyCommon.LRT_Select()
                    If AddRole IsNot Nothing AndAlso AddRole.Rows.Count = 0 Then  'If the Role does not already exist for the user then it can be added. This is done to prevent duplicate records.
                        MyCommon.QueryStr = "INSERT into AdminUserRoles with (RowLock) (RoleID,AdminUserID) values(" & selections(i) & "," & u_UID & ")"
                        MyCommon.LRT_Execute()
                    End If
                Else
                    infoMessage = Copient.PhraseLib.Lookup("error.no-role", LanguageID)
                End If
            Next
        Else
            infoMessage = Copient.PhraseLib.Lookup("error.permission", LanguageID)
        End If
    ElseIf (Request.Form("role-rem") <> "" And Request.Form("roles-select") <> "") Then
        Dim UserCanEditRole As Boolean = IIf((AdminUserID = u_UID), Logix.UserRoles.EditOwnRoles, Logix.UserRoles.EditRoles)
        If UserCanEditRole Then
            Dim selections() As String
            Dim i As Integer
            selections = Server.HtmlEncode(Request.Form("roles-select")).Split(",")
            If ((Array.IndexOf(selections, "1") > -1) AndAlso AdminUserID = u_UID) Then   'If the Administrator role is included in the list of roles to be removed, do not remove the roles.
                infoMessage = Copient.PhraseLib.Lookup("error.removeadmin", LanguageID)
            Else
                For i = 0 To selections.GetUpperBound(0)
                    MyCommon.QueryStr = "DELETE from AdminUserRoles with (RowLock) where RoleID=" & selections(i) & " and AdminUserID=" & u_UID & ""
                    MyCommon.LRT_Execute()
                Next
            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("error.permission", LanguageID)
        End If
    ElseIf (Request.Form("banner-add") <> "" AndAlso Request.Form("banners-avail") <> "") Then
        Dim selections() As String
        selections = Server.HtmlEncode(Request.Form("banners-avail")).Split(",")
        Dim i As Integer
        Dim BannerAddRole As DataTable
        For i = 0 To selections.GetUpperBound(0)
            MyCommon.QueryStr = "Select BannerID from AdminUserBanners with (NoLock) where BannerID=" & selections(i) & " and AdminUserID=" & u_UID & ";"
            BannerAddRole = MyCommon.LRT_Select()
            If BannerAddRole.Rows.Count = 0 Then
                MyCommon.QueryStr = "INSERT into AdminUserBanners with (RowLock) (BannerID,AdminUserID) values(" & selections(i) & "," & u_UID & ")"
                MyCommon.LRT_Execute()
            End If
        Next
    ElseIf (Request.Form("banner-rem") <> "" And Request.Form("banners-select") <> "") Then
        Dim selections() As String
        Dim i As Integer
        selections = Server.HtmlEncode(Request.Form("banners-select")).Split(",")
        For i = 0 To selections.GetUpperBound(0)
            MyCommon.QueryStr = "DELETE from AdminUserBanners with (RowLock) where BannerID=" & selections(i) & " and AdminUserID=" & u_UID & ""
            MyCommon.LRT_Execute()
        Next
    ElseIf (Request.Form("epmrole-add") <> "" And Request.Form("epmroles-avail") <> "") Then
        Dim UserCanEditRole As Boolean = IIf((AdminUserID = u_UID), Logix.UserRoles.EditOwnRoles, Logix.UserRoles.EditRoles)
        If UserCanEditRole Then
            Dim selections() As String
            Dim i As Integer
            selections = Request.Form("epmroles-avail").Split(",")
            Dim AddRole As DataTable
            For i = 0 To selections.GetUpperBound(0)
                MyCommon.QueryStr = "select RoleID from AdminUserRoles with (NoLock) where AdminUserID=" & u_UID & " and RoleID=" & selections(i) & ";"
                AddRole = MyCommon.PMRT_Select()
                If AddRole.Rows.Count = 0 Then    'If the Role does not already exist for the user then it can be added. This is done to prevent duplicate records.
                    MyCommon.QueryStr = "INSERT into AdminUserRoles with (RowLock) (RoleID,AdminUserID) values(" & selections(i) & "," & u_UID & ")"
                    MyCommon.PMRT_Execute()
                End If
            Next
        Else
            infoMessage = Copient.PhraseLib.Lookup("error.permission", LanguageID)
        End If

    ElseIf (Request.Form("epmrole-rem") <> "" And Request.Form("epmroles-select") <> "") Then
        Dim UserCanEditRole As Boolean = IIf((AdminUserID = u_UID), Logix.UserRoles.EditOwnRoles, Logix.UserRoles.EditRoles)
        If UserCanEditRole Then
            Dim selections() As String
            Dim i As Integer
            selections = Request.Form("epmroles-select").Split(",")
            If ((Array.IndexOf(selections, "1") > -1) AndAlso AdminUserID = u_UID) Then   'If the Administrator role is included in the list of roles to be removed, do not remove the roles.
                infoMessage = Copient.PhraseLib.Lookup("error.removeadmin", LanguageID)
            Else
                For i = 0 To selections.GetUpperBound(0)
                    MyCommon.QueryStr = "DELETE from AdminUserRoles with (RowLock) where RoleID=" & selections(i) & " and AdminUserID=" & u_UID & ""
                    MyCommon.PMRT_Execute()
                Next
            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("error.permission", LanguageID)
        End If
    End If

    ' grab this user
    MyCommon.QueryStr = "SELECT U.AdminUserID,U.EmployeeID,U.FirstName,U.LastName,U.UserName,U.JobTitle,U.Employer,U.Password,U.Email,U.AlertEmail,U.AuthToken,U.LastAuth,U.LastLogin,U.LanguageID,U.StartPageID,U.StyleID FROM AdminUsers AS U WITH (NoLock) WHERE AdminUserID='" & u_UID & "';"
    dst = MyCommon.LRT_Select
    If (dst.Rows.Count > 0) Then
        u_UID = MyCommon.NZ(dst.Rows(0).Item("AdminUserID"), 0)
        u_Username = MyCommon.NZ(dst.Rows(0).Item("UserName"), "")
        u_Firstname = MyCommon.NZ(dst.Rows(0).Item("FirstName"), "")
        u_Lastname = MyCommon.NZ(dst.Rows(0).Item("LastName"), "")
        u_Employeeid = MyCommon.NZ(dst.Rows(0).Item("EmployeeID"), "")
        u_JobTitle = MyCommon.NZ(dst.Rows(0).Item("JobTitle"), "")
        u_Employer = MyCommon.NZ(dst.Rows(0).Item("Employer"), "")
        u_Password = MyCommon.NZ(dst.Rows(0).Item("Password"), "")
        u_Email = MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(dst.Rows(0).Item("Email"), ""))
        u_Alertemail = MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(dst.Rows(0).Item("AlertEmail"), ""))
        u_LastLogin = MyCommon.NZ(dst.Rows(0).Item("LastLogin"), "")
        u_Language = MyCommon.NZ(dst.Rows(0).Item("LanguageID"), 1)
        u_StartPage = MyCommon.NZ(dst.Rows(0).Item("StartPageID"), "")
        u_Style = MyCommon.NZ(dst.Rows(0).Item("StyleID"), 1)
        pgName = Regex.Replace(MyCommon.NZ(dst.Rows(0).Item("UserName"), ""), """", "")
    ElseIf (Request.QueryString("new") = "") AndAlso (u_UID > 0) Then
        Send("")
        Send("<div id=""intro"">")
        Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.user", LanguageID) & " #" & u_UID & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("  <div id=""infobar"" class=""red-background"">")
        Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        Send("  </div>")
        Send("</div>")
        GoTo done
    Else
        u_UID = "0"
        If (Request.QueryString("username") <> "") Then u_Username = Server.HtmlEncode(Request.QueryString("username")) Else u_Username = ""
        If (Request.QueryString("firstname") <> "") Then u_Firstname = Server.HtmlEncode(Request.QueryString("firstname")) Else u_Firstname = ""
        If (Request.QueryString("lastname") <> "") Then u_Lastname = Server.HtmlEncode(Request.QueryString("lastname")) Else u_Lastname = ""
        If (Request.QueryString("EmployeeID") <> "") Then u_Employeeid = Server.HtmlEncode(Request.QueryString("EmployeeID")) Else u_Employeeid = ""
        If (Request.QueryString("jobtitle") <> "") Then u_JobTitle = Server.HtmlEncode(Request.QueryString("jobtitle")) Else u_JobTitle = ""
        If (Request.QueryString("employer") <> "") Then u_Employer = Server.HtmlEncode(Request.QueryString("employer")) Else u_Employer = ""
        If (Request.QueryString("password") <> "") Then u_Password = Server.HtmlEncode(Request.QueryString("password")) Else u_Password = ""
        If (Request.QueryString("email") <> "") Then u_Email = Server.HtmlEncode(Request.QueryString("email")) Else u_Email = ""
        If (Request.QueryString("alertemail") <> "") Then u_Alertemail = Server.HtmlEncode(Request.QueryString("alertemail")) Else u_Alertemail = ""
        u_Language = MyCommon.Fetch_SystemOption(1)
        pgName = Copient.PhraseLib.Lookup("term.newuser", LanguageID)
    End If

    If (Logix.UserRoles.ViewOthersInfo = False AndAlso u_UID <> AdminUserID) Then
        Send_Denied(1, "perm.admin-users-seeothers")
        GoTo done
    End If
    If (Logix.UserRoles.UpdateOwnInfo = False) AndAlso (AdminUserID = u_UID) Then
        EditIdentity = False
    End If
    If (Logix.UserRoles.UpdateOthersInfo = False) AndAlso (AdminUserID <> u_UID) Then
        EditIdentity = False
    End If
    If (Logix.UserRoles.EditOwnRoles = False) AndAlso (AdminUserID = u_UID) Then
        EditRoles = False
    End If
%>
<form id="mainform" name="mainform" action="#" method="post">
<input type="hidden" id="username_hold" name="username_hold" value="<% Sendb(u_Username) %>" />
<input type="hidden" id="UserID" name="UserID" value="<%sendb(u_UID) %>" />
<div id="intro">
  <h1 id="title">
    <%If u_UID = 0 Then
        Sendb(Copient.PhraseLib.Lookup("term.newuser", LanguageID))
      Else
        u_Fullname = u_Firstname & " " & u_Lastname
        Sendb(Copient.PhraseLib.Lookup("term.user", LanguageID) & " #" & u_UID & ": " & MyCommon.TruncateString(u_Fullname, 40))
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If EditIdentity Then
        Send_Save()
      End If
      If (Logix.UserRoles.DeleteAdminUsers = True) Then
        If u_UID = 0 Then
        Else
          Send_Delete()
        End If
      End If
      If MyCommon.Fetch_SystemOption(75) Then
        If (u_UID > 0 And Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(19, u_UID, AdminUserID)
        End If
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    Send("<div id=""infobar"" class=""red-background""")
    If (infoMessage <> "") Then
      Send("style=""Display:Block"">")
      Send(infoMessage)
    Else
      Send("style=""Display:None"">")
    End If
    Send("</div>")
  %>
  <%--<% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>--%>
  <div id="column1x">
    <%
      MyCommon.QueryStr = "SELECT AdminUserID,EmployeeID,FirstName,LastName,UserName,JobTitle,Employer,Password,Email,AlertEmail,AuthToken,LastAuth,LastLogin,LanguageID,StartPageID FROM AdminUsers with (NoLock)"
      dst = MyCommon.LRT_Select
    %>
    <div class="box" id="identity">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.Identity", LanguageID))%>
        </span>
      </h2>
      <label for="username">
        <% Sendb(Copient.PhraseLib.Lookup("term.username", LanguageID))%>
        :</label><br />
      <input type="text" id="username" name="username" class="mediumlong" maxlength="50"
        value="<% Sendb(u_Username.Replace("""", "&quot;")) %>" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <label for="firstname">
        <% Sendb(Copient.PhraseLib.Lookup("term.firstname", LanguageID))%>
        :</label><br />
      <input type="text" id="firstname" name="firstname" class="mediumlong" maxlength="50"
        value="<% Sendb(u_Firstname.Replace("""", "&quot;")) %>" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <label for="lastname">
        <% Sendb(Copient.PhraseLib.Lookup("term.lastname", LanguageID))%>
        :</label><br />
      <input type="text" id="lastname" name="lastname" class="mediumlong" maxlength="50"
        value="<% Sendb(u_Lastname.Replace("""", "&quot;")) %>" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <label for="EmployeeID">
        <% Sendb(Copient.PhraseLib.Lookup("term.employeeid", LanguageID))%>
      </label>
      <br />
      <input type="text" id="employeeid" name="employeeid" class="mediumlong" maxlength="30"
        value="<%Sendb(u_Employeeid.Replace("""", "&quot;")) %>" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <% If (u_UID > 0) Then%>
      <label for="jobtitle">
        <% Sendb(Copient.PhraseLib.Lookup("term.title", LanguageID))%>
        :</label><br />
      <input type="text" id="jobtitle" name="jobtitle" class="mediumlong" maxlength="100"
        value="<% Sendb(u_JobTitle.Replace("""", "&quot;")) %>" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <label for="employer">
        <% Sendb(Copient.PhraseLib.Lookup("term.organization", LanguageID))%>
        :</label><br />
      <input type="text" id="employer" name="employer" class="mediumlong" maxlength="100"
        value="<% Sendb(u_Employer.Replace("""", "&quot;")) %>" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      
        
        <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%> 
        <% MyCommon.QueryStr = "select b.externalbuyerid from buyers as b inner join buyerroleusers as br on b.buyerid=br.buyerid where br.adminuserid=" & u_UID
        dst = MyCommon.LRT_Select
        Dim externalids As String = String.Empty
        If (dst.Rows.Count > 0) Then
                    
            %>
            <label for="BuyerID">
        <% Sendb(Copient.PhraseLib.Lookup("term.buyer", LanguageID))%>
      </label>
            <%        
          externalids += "<ul>"
          For Each row1 In dst.Rows
            externalids += "<li>"
            externalids += MyCommon.NZ(row1.Item("externalbuyerid"), "")
            externalids += "</li>"
          Next
        End If
        externalids += "</ul>"
        Send(externalids)%>
        <% End If %>

        <%
        If maxStoreUser > 0 Then ' CM_SystemOption 131
            MyCommon.QueryStr = "select loc.LocationName, loc.LocationID from Locations loc inner join storeusers su on su.LocationID = loc.LocationID where su.UserID = " & u_UID
            dtStoreUser = MyCommon.LRT_Select
            If dtStoreUser.Rows.Count > 0 Then
                showStoreUser = True
            End If
            Send("<input type=""checkbox"" name=""storeuser"" id=""storeuser""  " & IIf(showStoreUser, "checked =""checked""", "") & IIf(EditIdentity, "", " disabled=""disabled"" ") & " onclick=""StoreUserPopup();""/>")
            Send("<label for=""storeuser"">" & Copient.PhraseLib.Lookup("term.storeuser", LanguageID) & "</label>")
    
            Send("<table id =""storelocations"" " & IIf(showStoreUser, "", "style=""display:none""") & " >")
            Send("<thead>")
            Send("<tr>")
            Send("<th class=""th-rewardname"">" & Copient.PhraseLib.Lookup("term.storename", LanguageID) & "</th>")
            Send("</tr>")
            Send("</thead>")
            Send("<tbody =""storebody"">")

            For Each row In dtStoreUser.Rows
                Send("<tr>")
                Send("<td id=""Store-" & row.Item("LocationID") & """>&#8226; " & row.Item("LocationName") & "</td>")
                Send("</tr>")
            Next
   
            Send("</tbody>")
            Send("</table>")
            Send("</br><input  type=""button"" id = ""editstoreuser"" name=""editstoreuser"" class = ""regular"" " & IIf(showStoreUser, "", "style=""display:none""") & "value = ""Edit Stores"" " & IIf(EditIdentity, "", " disabled=""disabled"" ") & "onclick=""StoreUserPopup();""/>")
        End If
      %>
      <% End If%>
      <hr class="hidden" />
    </div>
    <div class="box" id="access">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.access", LanguageID))%>
        </span>
      </h2>
      <label for="password">
        <% Sendb(Copient.PhraseLib.Lookup("term.password", LanguageID))%>
        :</label><br />
      <input autocomplete="off" type="password" id="password" name="password" class="short"
        maxlength="20" value="" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <label for="passwordagain">
        <% Sendb(Copient.PhraseLib.Lookup("term.passwordagain", LanguageID))%>
        :</label><br />
      <input autocomplete="off" type="password" id="passwordagain" name="passwordagain"
        class="short" maxlength="20" value="" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <hr class="hidden" />
    </div>
    <div class="box" id="contact" <% if (u_uid = 0) then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.contact", LanguageID))%>
        </span>
      </h2>
      <label for="email">
        <% Sendb(Copient.PhraseLib.Lookup("term.email", LanguageID))%>
        :</label><br />
      <input type="text" id="email" name="email" class="mediumlong" maxlength="200" value="<% Sendb(u_Email) %>"
        <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <label for="alertemail">
        <% Sendb(Copient.PhraseLib.Lookup("term.alertemail", LanguageID))%>
        :</label><br />
      <input type="text" id="alertemail" name="alertemail" class="mediumlong" maxlength="200"
        value="<% Sendb(u_AlertEmail) %>" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %> /><br />
      <hr class="hidden" />
    </div>
    <div class="box" id="display" <% if (u_uid = 0) then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>
        </span>
      </h2>
      <label for="language">
        <% Sendb(Copient.PhraseLib.Lookup("term.language", LanguageID))%>
        :</label><br />
      <select id="language" name="language" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %>>
        <%  
          ' when multi-language system option is disabled and there is a specified default languageID, only show that LanguageID.
            If MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(124)) = 0 AndAlso Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID) Then
                If u_Language <> DefaultLanguageID Then
                    MyCommon.QueryStr = "select LanguageID,PhraseTerm from Languages with (NoLock) where LanguageID in (" & u_Language & "," & DefaultLanguageID & ")"
                Else
                     MyCommon.QueryStr = "select LanguageID, PhraseTerm from Languages with (NoLock) where InstalledForUI=1 and LanguageID = " & DefaultLanguageID
                End If
               
            Else
                MyCommon.QueryStr = "select LanguageID, PhraseTerm from Languages with (NoLock) where InstalledForUI=1;"
            End If
            
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            If (row.Item("LanguageID") = u_Language) Then
              Send("<option value=""" & row.Item("LanguageID") & """ selected=""selected"">" & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID) & "</option>")
            Else
              Send("<option value=""" & row.Item("LanguageID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID) & "</option>")
            End If
          Next
        %>
      </select>
      <br />
      <br class="half" />
      <label for="startpage">
        <% Sendb(Copient.PhraseLib.Lookup("term.startpage", LanguageID))%>
        :</label><br />
      <select id="startpage" name="startpage" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %>>
        <%  
          ' get all the installed engines
          MyCommon.QueryStr = "select EngineID from PromoEngines with (NoLock) where Installed=1;"
          rst = MyCommon.LRT_Select
          EngineList = "-99"
          For Each row In rst.Rows
            EngineList &= "," & MyCommon.NZ(row.Item("EngineID"), "0")
          Next
            
          MyCommon.QueryStr = "select StartPageID,PageName, DisplayName, PhraseID, OnlyEngineID, prestrict from AdminUserStartPages with (NoLock) " & _
                              "where OnlyEngineID is null or OnlyEngineID in (" & EngineList & ")"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            If (IsDBNull(row.Item("PhraseID"))) Then
              StartPage = MyCommon.NZ(row.Item("DisplayName"), "")
            Else
              StartPage = Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID, MyCommon.NZ(row.Item("DisplayName"), ""))
            End If
              
            If (row.Item("StartPageID") = u_StartPage) Then
              Send("<option value=""" & row.Item("StartPageID") & """ selected=""selected"">" & StartPage)
            Else
              Send("<option value=""" & row.Item("StartPageID") & """>" & StartPage)
            End If
            If (row.Item("prestrict") = True) Then
              Send(" (" & StrConv(Copient.PhraseLib.Lookup("term.restricted", LanguageID), VbStrConv.Lowercase) & ")")
            End If

            ' when this is a engine-specific page append the engine name
            If (MyCommon.NZ(row.Item("OnlyEngineID"), "-99") = 0) Then
              Send(" (CM)")
            ElseIf (MyCommon.NZ(row.Item("OnlyEngineID"), "-99") = 2) Then
              Send(" (CPE)")
            End If
              
            Send("</option>")
          Next
        %>
      </select>
      <br />
      <br class="half" />
      <label for="style">
        <% Sendb(Copient.PhraseLib.Lookup("term.style", LanguageID))%>:</label><br />
      <select id="style" name="style" <% if (EditIdentity = false) then sendb(" disabled=""disabled""") %>>
        <%
          MyCommon.QueryStr = "select StyleID, Name, PhraseID, DefaultStyle from UIStyles with (NoLock);"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            Sendb("<option value=""" & row.Item("StyleID") & """")
            If u_UID = 0 Then
              If (row.Item("DefaultStyle")) Then
                Sendb(" selected=""selected""")
              End If
            Else
              If (row.Item("StyleID") = u_Style) Then
                Sendb(" selected=""selected""")
              End If
            End If
            Sendb(">")
            If IsDBNull(row.Item("PhraseID")) Then
              Sendb(MyCommon.NZ(row.Item("Name"), "&nbsp;"))
            Else
              Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
            End If
            Send("</option>")
          Next
        %>
      </select>
      <br />
      <hr class="hidden" />
    </div>
    <%If u_UID > 0 Then
        If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then Send_EPM_Display()
      End If%>
  </div>
  <div id="gutter">
  </div>
  <div id="column2x" <% if(u_uid = 0)then sendb(" style=""visibility: hidden;""") %>>
    <% If (Logix.UserRoles.UpdateOwnInfo AndAlso (Logix.UserRoles.EditOwnRoles AndAlso u_UID = AdminUserID)) Or (Logix.UserRoles.UpdateOthersInfo) Then%>
    <div class="box" id="roles">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.ams", LanguageID) & " " & Copient.PhraseLib.Lookup("term.roles", LanguageID))%>
        </span>
      </h2>
      <div style="float: left; position: relative;">
        <label for="roles-select">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("user-edit.rolesselected", LanguageID))%>
          </b>
        </label>
        <br />
        <select class="narrowselector" multiple="multiple" id="roles-select" name="roles-select"
          <% if (EditRoles = false) then sendb(" disabled=""disabled""") %>>
          <%
            MyCommon.QueryStr = "select AUR.AdminUserID,AUR.RoleID,RoleName,PhraseID FROM AdminUserRoles AS AUR with (NoLock) LEFT JOIN AdminRoles AS AR with (NoLock) ON AUR.RoleID=AR.RoleID WHERE AdminUserID=" & u_UID & "ORDER BY RoleName"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              If (IsDBNull(row.Item("PhraseID"))) Then
                Send("<option value=""" & row.Item("RoleID") & """>" & row.Item("RoleName") & "</option>")
              Else
                If (row.Item("PhraseID") = 0) Then
                  Send("<option value=""" & row.Item("RoleID") & """>" & row.Item("RoleName") & "</option>")
                Else
                  Send("<option value=""" & row.Item("RoleID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                End If
              End If
            Next
          %>
        </select>
      </div>
      <div style="float: left; padding: 35px 2px 1px 2px; position: relative;">
        <input type="submit" class="arrowadd" id="role-add" name="role-add" value="&#171;"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" <% if (EditRoles = false) or (Logix.UserRoles.EditOwnRoles = False AndAlso u_UID = AdminUserID) then sendb(" disabled=""disabled""") %> /><br
            clear="all" />
        <br class="half" />
        <input type="submit" class="arrowrem" id="role-rem" name="role-rem" value="&#187;"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.remove", LanguageID)) %>" <% if (EditRoles = false) or (Logix.UserRoles.EditOwnRoles = False AndAlso u_UID = AdminUserID) then sendb(" disabled=""disabled""") %> />
      </div>
      <div style="float: left; position: relative;">
        <label for="roles-avail">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("user-edit.rolesavailable", LanguageID))%>
          </b>
        </label>
        <br />
        <select class="narrowselector" multiple="multiple" id="roles-avail" name="roles-avail"
          <% if (EditRoles = false) then sendb(" disabled=""disabled""") %>>
          <%
            MyCommon.QueryStr = "select R.RoleID,R.RoleName,R.PhraseID from AdminRoles as R with (NoLock) where R.RoleID not in(select RoleID from AdminUserRoles where AdminUserID=" & u_UID & ") ORDER BY RoleName"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              If (IsDBNull(row.Item("PhraseID"))) Then
                Send("<option value=""" & row.Item("RoleID") & """>" & row.Item("RoleName") & "</option>")
              Else
                Send("<option value=""" & row.Item("RoleID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
              End If
            Next
          %>
        </select>
        <br />
      </div>
      &nbsp;<br clear="left" />
      <hr class="hidden" />
    </div>
    <%If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then Send_EPM_Roles()%>
    <% End If%>
    <% If (BannersEnabled) Then%>
    <div class="box" id="banners">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.banners", LanguageID))%>
        </span>
      </h2>
      <div style="float: left; position: relative;">
        <label for="banners-select">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("user-edit.bannersselected", LanguageID))%>
            :</b></label>
        <br />
        <select class="narrowselector" multiple="multiple" id="banners-select" name="banners-select"
          <% if (true = false) then sendb(" disabled=""disabled""") %>>
          <%
            MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                "WHERE BAN.Deleted=0 and AdminUserID=" & u_UID & " ORDER BY BAN.Name"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              Send("<option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          %>
        </select>
      </div>
      <div style="float: left; padding: 35px 2px 1px 2px; position: relative;">
        <input type="submit" class="arrowadd" id="banner-add" name="banner-add" value="&#171;"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID)) %>" <% if (Logix.UserRoles.AddUsersToBanners = false) then sendb(" disabled=""disabled""") %> /><br
            clear="all" />
        <br class="half" />
        <input type="submit" class="arrowrem" id="banner-rem" name="banner-rem" value="&#187;"
          title="<% Sendb(Copient.PhraseLib.Lookup("term.remove", LanguageID)) %>" <% if (Logix.UserRoles.RemoveUsersFromBanners = false) then sendb(" disabled=""disabled""") %> />
      </div>
      <div style="float: left; position: relative;">
        <label for="banners-avail">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("user-edit.bannersavailable", LanguageID))%>
            :</b></label>
        <br />
        <select class="narrowselector" multiple="multiple" id="banners-avail" name="banners-avail">
          <%
            MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                "where BAN.Deleted =0 and BAN.BannerID Not IN " & _
                                "( select BAN.BannerID from Banners BAN with (NoLock) " & _
                                "  inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                "  where AUB.AdminUserID = " & u_UID & " and BAN.Deleted=0 ) " & _
                                "order by BAN.Name; "
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              Send("<option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          %>
        </select>
        <br />
      </div>
      <br clear="left" />
      <br class="zero" />
      <hr class="hidden" />
    </div>
    <% End If%>
    <div class="box" id="alerts">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.ams", LanguageID) & " " & Copient.PhraseLib.Lookup("term.alerts", LanguageID))%>
        </span>
      </h2>
      <%
        ' load all alert types assigned for the client's installed engines, except the special alert types that don't use checkboxes (ID >= 10000)
          MyCommon.QueryStr = "select distinct AT.AlertTypeID, AT.PhraseID from AlertTypes AT with (NoLock)  " & _
                              "inner join AlertTypeEngines ATE with (NoLock) on ATE.AlertTypeID = AT.AlertTypeID and AT.AlertTypeID Not In (10000,10004,10005) " & _
                              "where EngineID in (select EngineID from PromoEngines with (NoLock) where Installed = 1) " & _
                              "order by AlertTypeId;"
        dtAlertTypes = MyCommon.LRT_Select
        Dim dst2 As DataTable
        Dim u_stdEmail As Integer
        Dim u_alertEmailset As Integer
        Dim numAlerts As Integer = dtAlertTypes.Rows.Count
          
        If numAlerts > 0 Then
          Sendb(Copient.PhraseLib.Lookup("user-edit.alerts", LanguageID))
          Send("<br />")
          Send("<br class=""half"" />")
          Send("<b>" & Copient.PhraseLib.Lookup("term.reg", LanguageID) & " " & Copient.PhraseLib.Lookup("term.alert", LanguageID) & "</b>")
          Send("<br />")
        End If
          
        For Each row In dtAlertTypes.Rows
          u_alertEmailset = 0
          u_stdEmail = 0
          MyCommon.QueryStr = "Select StdEmail ,AlertEmail from AlertReceivers with (NoLock) where AlertTypeID=" & row.Item("AlertTypeID") & " and AdminUserID=" & u_UID
          dst2 = MyCommon.LRT_Select
            
          If (dst2.Rows.Count > 0) Then
            If (dst2.Rows(0).Item("AlertEmail")) Then u_alertEmailset = 1
            If (dst2.Rows(0).Item("stdEmail")) Then u_stdEmail = 1
            If (dst2.Rows(0).Item("AlertEmail") Or dst2.Rows(0).Item("StdEmail")) Then
              Sendb("<input class=""checkbox"" id=""alert-" & row.Item("AlertTypeID") & """ name=""alert-" & row.Item("AlertTypeID") & """ type=""checkbox""")
              If (u_stdEmail = 1) Then
                Sendb(" checked=""checked""")
              End If
              If (EditIdentity = False) Then
                Sendb(" disabled=""disabled""")
              End If
              Send(" />&nbsp;")
              Sendb("<input class=""checkbox"" id=""amail-" & row.Item("AlertTypeID") & """ name=""amail-" & row.Item("AlertTypeID") & """ type=""checkbox""")
              If (u_alertEmailset = 1) Then
                Sendb(" checked=""checked""")
              End If
              If (EditIdentity = False) Then
                Sendb(" disabled=""disabled""")
              End If
              Send(" />&nbsp;")
              Send("<label for=""amail-" & row.Item("AlertTypeID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</label><br />")
            Else
              Sendb("<input class=""checkbox"" id=""alert-" & row.Item("AlertTypeID") & """ name=""alert-" & row.Item("AlertTypeID") & """ type=""checkbox""")
              If (EditIdentity = False) Then
                Sendb(" disabled=""disabled""")
              End If
              Send(" />&nbsp;")
              Sendb("<input class=""checkbox"" id=""amail-" & row.Item("AlertTypeID") & """ name=""amail-" & row.Item("AlertTypeID") & """ type=""checkbox""")
              If (EditIdentity = False) Then
                Sendb(" disabled=""disabled""")
              End If
              Send(" />&nbsp;")
              Send("<label for=""amail-" & row.Item("AlertTypeID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</label><br />")
            End If
          Else
            Sendb("<input class=""checkbox"" id=""alert-" & row.Item("AlertTypeID") & """ name=""alert-" & row.Item("AlertTypeID") & """ type=""checkbox""")
            If (EditIdentity = False) Then
              Sendb(" disabled=""disabled""")
            End If
            Send(" />&nbsp;")
            Sendb("<input class=""checkbox"" id=""amail-" & row.Item("AlertTypeID") & """ name=""amail-" & row.Item("AlertTypeID") & """ type=""checkbox""")
            If (EditIdentity = False) Then
              Sendb(" disabled=""disabled""")
            End If
            Send(" />&nbsp;")
            Send("<label for=""alert-" & row.Item("AlertTypeID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</label><br />")
          End If
        Next
        ''Health Server Alerts
        'Select only the AlertTypeID and PhraseID of the health server alerts
          LoadHealthAlerts(10000)
          LoadHealthAlerts(10004)
          LoadHealthAlerts(10005)
        Send("<input type=""hidden"" id=""numAlerts"" name=""numAlerts"" value=""" & numAlerts & """ />")
      %>
      <hr class="hidden" />
    </div>
    <% If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then Send_EPM_Alerts()%>
    <div class="box" id="activity" <% if (u_uid = 0) then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.activity", LanguageID))%>
        </span>
      </h2>
      <%
        If u_LastLogin = Nothing Then
          Send(Copient.PhraseLib.Lookup("term.lastloggedin", LanguageID) & ": " & Copient.PhraseLib.Lookup("term.never", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.lastloggedin", LanguageID) & " ")
          longDate = u_LastLogin
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
        End If
        MyCommon.QueryStr = "select OfferID from AdminUserOffers with (NoLock) where AdminUserID=" & u_UID & ";"
        rst = MyCommon.LRT_Select
        Send("<span style=""display:none;"">" & Copient.PhraseLib.Detokenize("user-edit.FavoriteOffers", LanguageID, rst.Rows.Count) & "</span><br />")
        Send("<br class=""half"" />")
        If (Logix.UserRoles.ViewHistory = True) Then
          Send("<a class=""hidden"" href=""user-hist.aspx?UserID=" & u_UID & """>►</a>")
          Send("<a href=""javascript:openPopup('user-hist.aspx?UserID=" & u_UID & "')"">" & Copient.PhraseLib.Lookup("user-edit.recentactivity", LanguageID) & "</a><br />")
        End If
      %>
    </div>
    <div class="box" id="preferences" style="display: none;">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.preferences", LanguageID))%>
        </span>
      </h2>
      <%
        MyCommon.QueryStr = "select PREF.PreferenceID, PREF.PreferenceTypeID, PREF.Name, PREF.PhraseID, PREF.CallFunction, PREF.DefaultValue, AUP.Value " & _
                            "from Preferences PREF with (NoLock) " & _
                            "left join AdminUserPreferences AUP with (NoLock) on PREF.PreferenceID = AUP.PreferenceID and AdminUserID= " & u_UID & " " & _
                            "where PREF.Editable = 1;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.preferences", LanguageID) & """ style=""width:80%;"" >")
          For Each row In rst.Rows
            PrefID = MyCommon.NZ(row.Item("PreferenceID"), 0)
            PrefTypeID = MyCommon.NZ(row.Item("PreferenceTypeID"), 0)
            If Not IsDBNull(row.Item("PhraseID")) Then
              PrefName = Copient.PhraseLib.Lookup(Integer.Parse(row.Item("PhraseID")), LanguageID)
            Else
              PrefName = MyCommon.NZ(row.Item("Name"), "")
            End If
            If Not IsDBNull(row.Item("Value")) Then
              PrefValue = MyCommon.NZ(row.Item("Value"), "0")
            Else
              PrefValue = MyCommon.NZ(row.Item("DefaultValue"), "0")
            End If
              
            Select Case PrefTypeID
              Case 4 ' checkbox
                Send("<tr>")
                Send("<td><label for=""pref" & PrefID & """>" & PrefName & ":&nbsp;&nbsp;</label></td>")
                Sendb("<td><input type=""checkbox"" value=""1"" id=""pref" & PrefID & """ name=""pref" & PrefID & """ ")
                Sendb(" " & IIf(PrefValue = "1", "checked=""checked""", "") & " ")
                If Not IsDBNull(row.Item("CallFunction")) Then
                  Sendb(" onclick=""" & MyCommon.NZ(row.Item("CallFunction"), "") & "();"" ")
                End If
                Send(" /></td></tr>")
              Case 6 ' upload
                Send("<tr>")
                Send("<td>" & PrefName & "</td><td><span>" & PrefValue & "</span>&nbsp;&nbsp;")
                Sendb("<input type=""button"" value=""" & Copient.PhraseLib.Lookup("term.modify", LanguageID) & """ ")
                If Not IsDBNull(row.Item("CallFunction")) Then
                  Sendb(" onclick=""" & MyCommon.NZ(row.Item("CallFunction"), "") & "();"" ")
                End If
                Send(" /></td></tr>")
              Case 7 ' color
                Send("<tr>")
                Send("<td>" & PrefName & "</td><td><span id=""displayPref3"" style=""padding-left:60px;border:solid 1px black;background-color:" & PrefValue & ";"">&nbsp;&nbsp;</span>&nbsp;&nbsp;")
                Sendb("<input type=""button"" value=""" & Copient.PhraseLib.Lookup("term.modify", LanguageID) & """ ")
                If Not IsDBNull(row.Item("CallFunction")) Then
                  Sendb(" onclick=""" & MyCommon.NZ(row.Item("CallFunction"), "") & "();"" ")
                End If
                Send(" />")
                If (PrefValue.Length > 0 AndAlso Left(PrefValue, 1) = "#") Then PrefValue = PrefValue.Substring(1)
                Send("<input type=""hidden"" id=""pref3"" name=""pref3"" value=""" & PrefValue & """ />")
                Send("</td></tr>")
            End Select
          Next
          Send("</table>")
        End If
      %>
    </div>
  </div>
  <br clear="all" />
</div>
</form>
<div id="uploader" style="display: none;">
  <div id="uploadwrap">
    <div class="box" id="uploadbox">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.upload", LanguageID))%>
        </span>
      </h2>
      <form action="user-edit.aspx" id="uploadform" name="uploadform" onsubmit="return isValidPath();"
      method="post" enctype="multipart/form-data">
      <%
        Sendb("<input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" ")
        Send("onclick=""javascript:document.getElementById('uploader').style.display='none';"" />")
        Sendb(Copient.PhraseLib.Lookup("user-edit.SelectBackground", LanguageID))
          Send("<br /><br />")
      %>
      <br />
      <br class="half" />
      <%
        Send("     <input type=""hidden"" name=""UserID"" value=""" & u_UID & """ />")
        Send("     <input type=""hidden"" name=""StoreLocations"" value="""" />")
        Send("     <input type=""file"" id=""browse"" name=""browse"" value=""" & Copient.PhraseLib.Lookup("term.browse", LanguageID) & """ />")
        '  Send("<div id=""divfile"" style=""height:0px;overflow:hidden"">")
        'Send("<input type=""file"" id=""browse"" name=""fileInput"" onchange=""fileonclick()"" />")
        'Send("</div>")
        'Send("<button type=""button"" onclick=""chooseFile();"">"&Copient.PhraseLib.Lookup("term.browse", LanguageID) &"</button>")
        'Send("<label id=""lblfileupload"" name=""lblfileupload"">"&Copient.PhraseLib.Lookup("term.nofilesselected", LanguageID)&"</label>")
        Send("     <input type=""submit"" class=""regular"" id=""uploadfile"" name=""uploadfile"" value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """ />")
        Send("     <br />")
      %>
      </form>
      <hr class="hidden" />
    </div>
  </div>
  <%
    If Request.Browser.Type = "IE6" Then
      Send("<iframe src=""javascript:'';"" id=""uploadiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no""></iframe>")
    End If
  %>
</div>
<div id="colorPalette" style="display: none;">
  <div id="colorWrap">
    <div class="box" id="colorBox">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.colorselection", LanguageID))%>
        </span>
      </h2>
      <input type="button" class="ex" id="colorclose" name="colorclose" value="X" style="float: right;
        position: relative; top: -27px;" onclick="javascript:document.getElementById('colorPalette').style.display='none';" />
      <%
        Dim ColorValues(,) As Integer = {{255, 128, 128}, {255, 255, 128}, {128, 255, 128}, {0, 255, 128}, {128, 255, 255}, {0, 128, 255}, {255, 128, 192}, {255, 128, 255}, _
                                         {255, 0, 0}, {255, 255, 0}, {128, 255, 0}, {0, 255, 64}, {0, 255, 255}, {0, 128, 192}, {128, 128, 192}, {255, 0, 255}, _
                                         {128, 64, 64}, {255, 128, 64}, {0, 255, 0}, {0, 128, 128}, {0, 64, 128}, {128, 128, 255}, {128, 0, 64}, {255, 0, 128}, _
                                         {128, 0, 0}, {255, 128, 0}, {0, 128, 0}, {0, 128, 64}, {0, 0, 255}, {0, 0, 160}, {128, 0, 128}, {128, 0, 255}, _
                                         {64, 0, 0}, {128, 64, 0}, {0, 64, 0}, {0, 64, 64}, {0, 0, 128}, {0, 0, 64}, {64, 0, 64}, {64, 0, 128}, _
                                         {0, 0, 0}, {128, 128, 0}, {128, 128, 64}, {128, 128, 128}, {64, 128, 128}, {192, 192, 192}, {64, 0, 64}, {255, 255, 255}}
          
        For x = 0 To ColorValues.GetUpperBound(0)
          Sendb("<span style=""padding-left:30px;border:solid 1px #808080;cursor:hand;background-color:rgb(" & ColorValues(x, 0) & "," & ColorValues(x, 1) & "," & ColorValues(x, 2) & ");"" ")
          Send(" onclick=""populateColor(" & ColorValues(x, 0) & "," & ColorValues(x, 1) & "," & ColorValues(x, 2) & ");"">&nbsp;&nbsp;</span>")
          If ((x + 1) Mod 8) = 0 Then Send("<br /><br class=""half"" />")
        Next
        Send("<br />")
      %>
    </div>
  </div>
  
</div>
<div id="StoreUserFadeDiv"></div>
  
<div id="foldercreate" class="folderdialog" style="position:absolute; WIDTH: 466px; HEIGHT: 280px;">
  <div class="foldertitlebar">
    <span class="dialogtitle"><%Sendb(Copient.PhraseLib.Lookup("storeuser.selectlocations", LanguageID) & maxStoreUser & ")")%></span> <span class="dialogclose" onclick="toggleDialog('foldercreate', false);">X</span><!--!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!-->
  </div>
  <div class="dialogcontents">
    <table> 
      <tr>
        <td>
          <form>
            <input type="radio" id="searchradio1" name="searchradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> onclick ="locationSearch();" />
            <label for="searchradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
            <input type="radio" id="searchradio2" name="searchradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> onclick = "locationSearch();"/>
            <label for="searchradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
          </form>
          <input type = "text" class="large" id = "locationsearch" onkeyup="locationSearch();"/>
        </td>
      </tr>
    </table>
    <label id ="maxstores" style="display:none">Only displaying the top 100 stores, alphebetically</label>
    <div id ="StoreUserPane">
    <div id="storeLocDiv" style="float:left; width: 150px">
<select id = "storeLoc" size="10" style="width:150px">

</select>
</div>
<div style="float:left; width: 150px"">
<center>
</br>
</br>
</br>
<input  type="button" id="storeselect" class="regular select" value="Select &#9658" onclick="selectItem('storeLoc');"/>
</br>
</br>
<input  type="button" id="storedeselect" class="regular deselect" value="&#9668 Deselect" onclick="deselectItem('userLoc');"/>
</center>

</div> 
<div id="userLocDiv" style="float:left; width: 150px">
<select id="userLoc" size="10" style="width:150px;height: 172px;">
</select>
</div>
    </div>
    <br/>
    <input type="submit" align="right" class="regular" value = "Save" onclick="saveLocations();"/>
</div>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (u_UID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(19, u_UID, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd("mainform", "username")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixWH()
  MyCommon.Close_PrefManRT()
  Logix = Nothing
  MyCommon = Nothing
%>
