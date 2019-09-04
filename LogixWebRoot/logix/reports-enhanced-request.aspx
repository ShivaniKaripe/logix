<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: reports-enhanced.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2009.  All rights reserved by:
  ' *
  ' * NCR Corporation
  ' * 1435 Win Hentschel Boulevard
  ' * West Lafayette, IN  47906
  ' * voice: (888) 346-7199 fax: (765) 496-6489
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
  Dim Logix As New Copient.LogixInc
    Dim MyCryptlib As New Copient.CryptLib
  Dim AdminUserID As Long
  Dim ReportID As Integer = 0
  Dim dt As System.Data.DataTable = Nothing
  Dim row As System.Data.DataRow
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RECORD_LIMIT As Integer = GroupRecordLimit 'system option 126
    
  Dim ReportTypeID As Integer = 1
  Dim ReportName As String = ""
    
  Dim OldReportTypeID As Integer = -1
  Dim OldReportID As Integer = -1

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Response.Expires = 0
  MyCommon.AppName = "reports-enhanced.aspx"
  
  If (Request.QueryString("ReportID") <> "") Then
    ReportID = Request.QueryString("ReportID")
  End If
  
    If (Request.QueryString("ReportTypeID") <> "") Then
        ReportTypeID = Request.QueryString("ReportTypeID")
    End If
    
    If (Request.QueryString("ReportName") <> "") Then
        ReportName = Request.QueryString("ReportName")
    End If
    
    If (Request.QueryString("OldReportTypeID") <> "") Then
        OldReportTypeID = Request.QueryString("OldReportTypeID")
    End If
    
    If (Request.QueryString("OldReportID") <> "") Then
        OldReportID = Request.QueryString("OldReportID")
    End If
  Send_HeadBegin("term.reports", "term.enhanced")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script src="../javascript/jquery-1.10.2.min.js" type="text/javascript"></script>
<script type="text/javascript">
  $().ready(function () {
        toggleReportType("<%=ReportTypeID%>");
    $('#addOffers').click(function () {
        if (document.getElementById("reportTypeID").value != 4 && document.getElementById("reportTypeID").value != 2) {
                if (!$('#offerIDs-available option:selected').size() <= 0) {
               !$('#offerIDs-selected option').remove();
            }		
      }
      return !$('#offerIDs-available option:selected').remove().appendTo('#offerIDs-selected');
    });
    $('#removeOffers').click(function () {
      return !$('#offerIDs-selected option:selected').remove().appendTo('#offerIDs-available');
    });
    $('#addLocations').click(function () {
      return !$('#locationIDs-available option:selected').remove().appendTo('#locationIDs-selected');
    });
    $('#removeLocations').click(function () {
      return !$('#locationIDs-selected option:selected').remove().appendTo('#locationIDs-available');
    });
  
    $('#addCustomerGroups').click(function () {
        return !$('#CGs-available option:selected').remove().appendTo('#CGs-selected');
    });
    $('#removeCustomerGroups').click(function () {
        return !$('#CGs-selected option:selected').remove().appendTo('#CGs-available');
    });
  });
</script>
<script type="text/javascript">
    // Added the method to populate the offer, customer and location data for the selected report type.
    function OnChangeReportTypeID() {
        
        var reportID = "<%=ReportID %>";
        var bSendReportID = false;
        
        // OldreporttypeID and OldreportId are sent to populate the values of report filters for the generated report.
        var oldReportTypeId=document.getElementById("hidreporttypeid").value;
        var oldReportId=document.getElementById("hidreportid").value;

        var reportTypeID = document.getElementById("reportTypeID").value;
        if (reportTypeID == oldReportTypeId) bSendReportID = true;
        var reportName = document.getElementById("name").value;
        if ((reportID != undefined || reportID != "") && bSendReportID)
            location.href = "reports-enhanced-request.aspx?ReportID=" + oldReportId + "&ReportTypeID=" + reportTypeID + "&ReportName=" + reportName;
        else
            location.href = "reports-enhanced-request.aspx?ReportTypeID=" + reportTypeID + "&ReportName=" + reportName + "&OldReportTypeID=" + oldReportTypeId + "&OldReportID=" + oldReportId;;
    }
    function toggleReportType(reportTypeID) {
      document.getElementById("reportTypeID").value = reportTypeID;
    var typeElem = document.getElementById("reportTypeID");
    var cDiv = document.getElementById("customer");
    var oDiv = document.getElementById("offers");
    var eDiv = document.getElementById("employees");
    var dDiv = document.getElementById("discounts");
    var drDiv = document.getElementById("daterange");
    var lgDiv = document.getElementById("locationgroups");
    var allOffersElem = document.getElementById("allOffersTD");
    var offersSelectedElem = document.getElementById("offerIDs-selected");
    var CGDiv = document.getElementById("CustomerGroup");
    var allCGElem = document.getElementById("AllCustomerGroupsTD");
    var CGSelectedElem = document.getElementById("CGs-selected");
    var SearchByDiv = document.getElementById("SearchByEng");
    var engineDiv = document.getElementById("engines");
    var offerNameDiv = document.getElementById("SearchByOfferName");
    var offerStatusDiv = document.getElementById("offerstatus"); 
    var extInterfaceIDDiv = document.getElementById("ExtInterfaceID");
    var allOffersDivElem = document.getElementById("allOffersDiv");

    if (typeElem.value == 1) {
      cDiv.style.display = "block";
      oDiv.style.display = "none";
      eDiv.style.display = "none";
      dDiv.style.display = "none";
      drDiv.style.display = "block";
      lgDiv.style.display = "none";
      allOffersElem.style.display = "none";
      CGDiv.style.display = "none";
      allCGElem.style.display = "none";
      CGSelectedElem.style.display = "none";
      SearchByDiv.style.display = "none";
      engineDiv.style.display = "none";
      offerNameDiv.style.display = "none";
      offerStatusDiv.style.display = "none";
      extInterfaceIDDiv.style.display = "none";
      allOffersDivElem.style.display = "none";
    } else if (typeElem.value == 2) {
      cDiv.style.display = "none";
      oDiv.style.display = "block";
      eDiv.style.display = "none";
      dDiv.style.display = "none";
      drDiv.style.display = "none";
      lgDiv.style.display = "none";
      allOffersElem.style.display = "none";
      CGDiv.style.display = "none";
      allCGElem.style.display = "none";
      CGSelectedElem.style.display = "none";
      SearchByDiv.style.display = "none";
      offerStatusDiv.style.display = "block";
      extInterfaceIDDiv.style.display = "block";
      drDiv.style.display = "block";
      allOffersDiv.style.display = "block";
      engineDiv.style.display = "block";
      offerNameDiv.style.display = "block";
      allOffersDiv.style.display = "block";
    } else if (typeElem.value == 3) {
      cDiv.style.display = "none";
      oDiv.style.display = "block";
      eDiv.style.display = "block";
      dDiv.style.display = "block";
      drDiv.style.display = "block";
      lgDiv.style.display = "block";
      allOffersElem.style.display = "none";
      CGDiv.style.display = "none";
      allCGElem.style.display = "none";
      CGSelectedElem.style.display = "none";
      SearchByDiv.style.display = "none";
      engineDiv.style.display = "none";
      offerNameDiv.style.display = "none";
      offerStatusDiv.style.display = "none";
      extInterfaceIDDiv.style.display = "none";
      offerNameDiv.style.display = "none";
      allOffersDiv.style.display = "none";
    } else if (typeElem.value == 4) {
      cDiv.style.display = "none";
      oDiv.style.display = "block";
      eDiv.style.display = "block";
      dDiv.style.display = "block";
      drDiv.style.display = "block";
      lgDiv.style.display = "block";
      allOffersElem.style.display = "block";
      CGDiv.style.display = "none";
      allCGElem.style.display = "none";
      CGSelectedElem.style.display = "none";
      SearchByDiv.style.display = "none";
      engineDiv.style.display = "none";
      offerNameDiv.style.display = "none";
      offerStatusDiv.style.display = "none";
      extInterfaceIDDiv.style.display = "none";
      offerNameDiv.style.display = "none";
      allOffersDiv.style.display = "none";
    }
    else if (typeElem.value == 5) {
        cDiv.style.display = "none";
        oDiv.style.display = "none";
        eDiv.style.display = "none";
        dDiv.style.display = "none";
        drDiv.style.display = "none";
        lgDiv.style.display = "none";
        allOffersElem.style.display = "none";
        CGDiv.style.display = "block";
        allCGElem.style.display = "block";
        CGSelectedElem.style.display = "block";
        SearchByDiv.style.display = "block";
        engineDiv.style.display = "none";
        offerNameDiv.style.display = "none";
        offerStatusDiv.style.display = "none";
        extInterfaceIDDiv.style.display = "none";
        offerNameDiv.style.display = "none";
        allOffersDiv.style.display = "none";
    }
    //offersSelectedElem.options.length = 0;
  }

  function toggleOffers() {
    var allElem = document.getElementById("allOffers");
    var allElemOD = document.getElementById("allOffersOD");
    var listAvailableElem = document.getElementById("offerIDs-available");
    var listSelectedElem = document.getElementById("offerIDs-selected");
    var allTriggOffers = document.getElementById("triggerOffersOD");

    document.getElementById("triggerOffersOD").disabled = allElemOD.checked;
    document.getElementById("allOffersOD").disabled = allTriggOffers.checked;

    if (allElem.checked == true || (allElemOD.checked == true || allTriggOffers.checked == true)) {
      listAvailableElem.selectedIndex = -1;
      listSelectedElem.selectedIndex = -1;
      listAvailableElem.disabled = true;
      listSelectedElem.disabled = true;
      if (document.getElementById("functionradioOff1") != undefined || document.getElementById("functionradioOff1") != null)
          document.getElementById("functionradioOff1").disabled = true;
      if (document.getElementById("functionradioOff2") != undefined || document.getElementById("functionradioOff2") != null)
          document.getElementById("functionradioOff2").disabled = true;
      if (document.getElementById("functioninputOff") != undefined || document.getElementById("functioninputOff") != null)
          document.getElementById("functioninputOff").disabled = true;
    } else {
      listAvailableElem.disabled = false;
      listSelectedElem.disabled = false;
      if (document.getElementById("functionradioOff1") != undefined || document.getElementById("functionradioOff1") != null)
          document.getElementById("functionradioOff1").disabled = false;
      if (document.getElementById("functionradioOff2") != undefined || document.getElementById("functionradioOff2") != null)
          document.getElementById("functionradioOff2").disabled = false;
      if (document.getElementById("functioninputOff") != undefined || document.getElementById("functioninputOff") != null)
          document.getElementById("functioninputOff").disabled = false;
    }
  }
  function toggleCustomergroups() {
      var allElem = document.getElementById("allCustomerGroups");
      var listAvailableElem = document.getElementById("CGs-available");
      var listSelectedElem = document.getElementById("CGs-selected");

      if (allElem.checked == true) {
          listAvailableElem.selectedIndex = -1;
          listSelectedElem.selectedIndex = -1;
          listAvailableElem.disabled = true;
          listSelectedElem.disabled = true;
          if (document.getElementById("functionradio1") != undefined || document.getElementById("functionradio1") != null)
              document.getElementById("functionradio1").disabled = true;
          if (document.getElementById("functionradio2") != undefined || document.getElementById("functionradio2") != null)
              document.getElementById("functionradio2").disabled = true;
          if (document.getElementById("functioninputCG") != undefined || document.getElementById("functioninputCG") != null)
              document.getElementById("functioninputCG").disabled = true;
      } else {
          listAvailableElem.disabled = false;
          listSelectedElem.disabled = false;
          if (document.getElementById("functionradio1") != undefined || document.getElementById("functionradio1") != null)
              document.getElementById("functionradio1").disabled = false;
          if (document.getElementById("functionradio2") != undefined || document.getElementById("functionradio2") != null)
              document.getElementById("functionradio2").disabled = false;
          if (document.getElementById("functioninputCG") != undefined || document.getElementById("functioninputCG") != null)
              document.getElementById("functioninputCG").disabled = false;
      }
  }

  function toggleLocations() {
    var allElem = document.getElementById("allLocations");
    var listAvailableElem = document.getElementById("locationIDs-available");
    var listSelectedElem = document.getElementById("locationIDs-selected");

    if (allElem.checked == true) {
      listAvailableElem.selectedIndex = -1;
      listSelectedElem.selectedIndex = -1;
      listAvailableElem.disabled = true;
      listSelectedElem.disabled = true;
    } else {
      listAvailableElem.disabled = false;
      listSelectedElem.disabled = false;
    }
  }

  function toggleDiscount() {
    var dElem = document.getElementById("discountTypeID");
    var dMinRow = document.getElementById("discountMinimumRow");
    var dMaxRow = document.getElementById("discountMaximumRow");

    if (dElem.value == 0) {
      dMinRow.style.display = 'none';
      dMaxRow.style.display = 'none';
    } else {
      dMinRow.style.display = 'block';
      dMaxRow.style.display = 'block';
    }
  }

  function validateRequest() {
    var errorMessage = "";
    var reportType = document.getElementById("reportTypeID").value;
    var offerIDs = document.getElementById("offerIDs-selected").value;
    var CustomerGroupIDs = document.getElementById("CGs-selected").value;
    var startDate = document.getElementById("startDate").value;
    var endDate = document.getElementById("endDate").value;
    var discountMinimum = document.getElementById("discountMinimum").value;
    var discountMaximum = document.getElementById("discountMaximum").value;
    var save = document.getElementById("save");
    var status = document.getElementById("offerstatuses").value;
    var extInterfaceID = document.getElementById("ExtInterfaces").value;
    var engineID = document.getElementById("OfferEngineID").value;
    var allOffers = document.getElementById("allOffers");
    var allOffersOD = document.getElementById("allOffersOD");
    var triggerOffers = document.getElementById("triggerOffersOD");
	var allCustomerGroups = document.getElementById("allCustomerGroups");

    if (discountMinimum != "" && isNaN(discountMinimum)) {
      errorMessage = '<%Sendb(Copient.PhraseLib.Lookup("term.InvalidMinimumDiscount", LanguageID))%>';
    } else if (discountMaximum != "" && isNaN(discountMaximum)) {
      errorMessage = '<%Sendb(Copient.PhraseLib.Lookup("term.InvalidMaximumDiscount", LanguageID))%>';
    } else if (reportType == "2" && offerIDs == "" && allOffersOD.checked == false && triggerOffers.checked == false) {
      errorMessage = '<%Sendb(Copient.PhraseLib.Lookup("reports.PleaseSelectAnOffer", LanguageID))%>';
    } else if (reportType == "5" && CustomerGroupIDs == "" && allCustomerGroups.checked == false) {
      errorMessage = '<%Sendb(Copient.PhraseLib.Lookup("reports.PleaseASelectcustomergroup", LanguageID))%>';
    } else if (reportType == "3"  && offerIDs == "" && allOffers.checked == false) {
      errorMessage = '<%Sendb(Copient.PhraseLib.Lookup("reports.PleaseSelectAnOffer", LanguageID))%>';
    } else if (reportType == "4" && offerIDs == "" && allOffers.checked == false) {
      errorMessage = '<%Sendb(Copient.PhraseLib.Lookup("reports.PleaseSelectAnOffer", LanguageID))%>';
  }
    if (errorMessage != "") {
      alert(errorMessage);
    } else {
      save.value = "Save";
      document.mainform.submit();
    }
}
var isFireFox = (navigator.appName.indexOf('Mozilla') != -1) ? true : false;
var timer;
function xmlPostCG(strURL, mode) {
    clearTimeout(timer);
    timer = setTimeout("xmlhttpPostCG('" + strURL + "','" + mode + "')", 250);
}

function xmlhttpPostCG(strURL, mode) {
  
    var xmlHttpReq = false;
    var self = this;

    document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';

    // Mozilla/Safari
    if (window.XMLHttpRequest) {
        self.xmlHttpReq = new XMLHttpRequest();
    }
    // IE
    else if (window.ActiveXObject) {
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    var qryStr = getgroupqueryCG(mode);
    self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function () {
        if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
            updatepageCG(self.xmlHttpReq.responseText);
        }
    }

    self.xmlHttpReq.send(qryStr);
    //self.xmlHttpReq.send(getquerystring());
}

function getgroupqueryCG(mode) {
    var radioString;
    if (document.getElementById('functionradio2').checked) {
        radioString = 'functionradio2';
    }
    else {
        radioString = 'functionradio1';
    }
   
    return "Mode=" + mode + "&Search=" + document.getElementById('functioninputCG').value + "&SearchRadio=" + radioString;

}

function updatepageCG(str) {
    if (str.length > 0) {
        if (!isFireFox) {
            document.getElementById("CustomerGroupsDiv").innerHTML = '<select class="longer" id="CGs-available" name="CGs-available" size="12" multiple="multiple">' + str + '</select>';
        }
        else {
            document.getElementById("CGs-available").innerHTML = str;
        }
        document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
        if (document.getElementById("CGs-available").options.length > 0) {
            document.getElementById("CGs-available").options[0].selected = true;
        }
        var select = document.getElementById('CGs-selected').options;
        var available = document.getElementById('CGs-available').options;

        var opt = 0;
        for (opt = 0; opt < select.length; opt++) {
            if (select[opt] != null) {
                $("#CGs-available option[value=" + select[opt].value + " ]").remove();
            }
        }
    }
    else if (str.length == 0) {
        if (!isFireFox) {
            document.getElementById("CustomerGroupsDiv").innerHTML = '';
        }
        else {
            document.getElementById("CGs-available").innerHTML = '<select class="longer" id="CGs-available" name="CGs-available" size="12" multiple="multiple"></select>';
        }
        document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    }
}


function xmlPostOffer(strURL, mode) {
    clearTimeout(timer);
    timer = setTimeout("xmlhttpPostOff('" + strURL + "','" + mode + "')", 250);
}

function xmlhttpPostOff(strURL, mode) {

    var xmlHttpReq = false;
    var self = this;

    document.getElementById("searchLoadDivOff").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';

    // Mozilla/Safari
    if (window.XMLHttpRequest) {
        self.xmlHttpReq = new XMLHttpRequest();
    }
    // IE
    else if (window.ActiveXObject) {
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    var qryStr = getgroupqueryOff(mode);
    self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function () {
        if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
            updatepageOff(self.xmlHttpReq.responseText);
        }
    }

    self.xmlHttpReq.send(qryStr);
    //self.xmlHttpReq.send(getquerystring());
}

function getgroupqueryOff(mode) {
    var radioString;
    if (document.getElementById('functionradioOff2').checked) {
        radioString = 'functionradioOff2';
    }
    else {
        radioString = 'functionradioOff1';
    }

    return "Mode=" + mode + "&Search=" + document.getElementById('functioninputOff').value + "&SearchRadio=" + radioString;

}

function updatepageOff(str) {
    if (str.length > 0) {
        if (!isFireFox) {
            document.getElementById("OffersDiv").innerHTML = '<select class="longer" id="offerIDs-available" name="offerIDs-available" size="12" multiple="multiple">' + str + '</select>';
        }
        else {
            document.getElementById("offerIDs-available").innerHTML = str;
        }
        document.getElementById("searchLoadDivOff").innerHTML = '&nbsp;';
        if (document.getElementById("offerIDs-available").options.length > 0) {
            document.getElementById("offerIDs-available").options[0].selected = true;
        }
        var select = document.getElementById('offerIDs-selected').options;
        var available = document.getElementById('offerIDs-available').options;

        var opt = 0;
        for (opt = 0; opt < select.length; opt++) {
            if (select[opt] != null) {
                $("#offerIDs-available option[value=" + select[opt].value + " ]").remove();
            }
        }
    }
    else if (str.length == 0) {
        if (!isFireFox) {
            document.getElementById("OffersDiv").innerHTML = '';
        }
        else {
            document.getElementById("offerIDs-available").innerHTML = '<select class="longer" id="offerIDs-available" name="offerIDs-available" size="12" multiple="multiple"></select>';
        }
        document.getElementById("searchLoadDivOff").innerHTML = '&nbsp;';
    }
}
</script>
<%
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 8)
    Send_Subtabs(Logix, 8, 8)

    If (Logix.UserRoles.AccessReports = False) Then
        Send_Denied(1, "perm.admin-reports")
        GoTo done
    End If


    Dim Name As String = ""
    Dim StartDate As DateTime
    Dim StartDateString As String = ""
    Dim EndDate As DateTime
    Dim EndDateString As String = ""
    Dim CustomerID As String = ""
    Dim CustomerTypeID As Integer = 0
    Dim AllOffers As Boolean = False
    Dim allOffersVal As Integer = 0
    Dim OnlyTCOffers As Boolean = False
    Dim OfferIDs As String = ""
    Dim OfferStatus As String = "All"
    Dim ExtInterfaceID As integer = -1
    Dim AllLocations As Boolean = False
    Dim LocationGroupIDs As String = ""
    Dim EmployeeStatusID As Integer = 0
    Dim DiscountTypeID As Integer = 0
    Dim DiscountMinimum As Decimal = 0
    Dim DiscountMaximum As Decimal = 0
    Dim Deleted As Boolean = False
    Dim CustomerGroupIDs As String = ""
    Dim AllCustomerGroups As Boolean = False
    Dim EngineID As Integer = -1
    Dim OfferName As String = ""
    If (ReportID > 0) Then
        'Load the request
        MyCommon.QueryStr = "select * from Reports with (NoLock) where ReportID=" & ReportID & ";"
        Dim report As System.Data.DataTable = MyCommon.LRT_Select
        If report.Rows.Count > 0 Then
            ReportTypeID = MyCommon.NZ(report.Rows(0).Item("ReportTypeID"), 1)
            Name = MyCommon.NZ(report.Rows(0).Item("Name"), "")
            If Not IsDBNull(report.Rows(0).Item("StartDate")) Then
                StartDate = report.Rows(0).Item("StartDate")
                StartDateString = Logix.ToShortDateString(StartDate, MyCommon)
            End If
            If Not IsDBNull(report.Rows(0).Item("EndDate")) Then
                EndDate = report.Rows(0).Item("EndDate")
                EndDateString = Logix.ToShortDateString(EndDate, MyCommon)
            End If
            CustomerID = MyCommon.NZ(report.Rows(0).Item("CustomerID"), "")
            CustomerTypeID = MyCommon.NZ(report.Rows(0).Item("CustomerTypeID"), 0)
            AllOffers = MyCommon.NZ(report.Rows(0).Item("AllOffers"), False)
            OfferIDs = MyCommon.NZ(report.Rows(0).Item("OfferIDs"), "")
            OnlyTCOffers = MyCommon.NZ(report.Rows(0).Item("OnlyTCOffers"), False)
            AllCustomerGroups = MyCommon.NZ(report.Rows(0).Item("allCustomerGroups"), False)
            CustomerGroupIDs = MyCommon.NZ(report.Rows(0).Item("CustomerGroupIDs"), "")
            AllLocations = MyCommon.NZ(report.Rows(0).Item("AllLocations"), False)
            LocationGroupIDs = MyCommon.NZ(report.Rows(0).Item("LocationGroupIDs"), "")
            EngineId = MyCommon.NZ(report.Rows(0).Item("EngineId"), 0)
            EmployeeStatusID = MyCommon.NZ(report.Rows(0).Item("EmployeeStatusID"), 0)
            DiscountTypeID = MyCommon.NZ(report.Rows(0).Item("DiscountTypeID"), 0)
            DiscountMinimum = MyCommon.NZ(report.Rows(0).Item("DiscountMinimum"), 0)
            DiscountMaximum = MyCommon.NZ(report.Rows(0).Item("DiscountMaximum"), 0)
            Deleted = MyCommon.NZ(report.Rows(0).Item("Deleted"), False)
            OfferStatus = MyCommon.NZ(report.Rows(0).Item("OfferStatus"), "All")
            ExtInterfaceID = MyCommon.NZ(report.Rows(0).Item("ExtInterfaceID"), False)
            OfferName = MyCommon.NZ(report.Rows(0).Item("OfferName"), "")
            OldReportTypeID = ReportTypeID
            OldReportID = ReportID
        Else
            infoMessage = Copient.PhraseLib.Detokenize("reports.CouldNotFindReport", LanguageID, ReportID)
        End If
    End If

    If (Request.Form("save") <> "") Then

        If (Request.Form.Item("reportTypeID") <> "") Then ReportTypeID = Request.Form.Item("reportTypeID")
        If (Request.Form.Item("name") <> "") Then Name = Request.Form.Item("name")
        If (Request.Form.Item("startDate") <> "") Then
            If Not Date.TryParse(Request.Form.Item("startDate"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, StartDate) Then
                StartDateString = Request.Form.Item("startDate")
                infoMessage = Copient.PhraseLib.Lookup("term.InvalidStartDate", LanguageID)
            Else
                StartDateString = Logix.ToShortDateString(StartDate, MyCommon)
            End If
        End If

        If (Request.Form.Item("endDate") <> "") Then
            If Not Date.TryParse(Request.Form.Item("endDate"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, EndDate) Then
                EndDateString = Request.Form.Item("endDate")
                infoMessage = Copient.PhraseLib.Lookup("term.InvalidEndDate", LanguageID)
            Else
                EndDateString = Logix.ToShortDateString(EndDate, MyCommon)
            End If
        End If

        If (Request.Form.Item("customerID") <> "") Then
            CustomerID = Request.Form.Item("customerID")
            CustomerID = MyCommon.Pad_ExtCardID(MyCommon.Extract_Val(CustomerID), Copient.commonShared.CardTypes.CUSTOMER)
        End If
        If (Request.Form.Item("customerTypeID") <> "") Then CustomerTypeID = Request.Form.Item("customerTypeID")
        If (ReportTypeID = 2) Then
            If (Request.Form.Item("allOffersOD") <> "") Then
                allOffersVal = MyCommon.Extract_Val(Request.Form.Item("allOffersOD").ToString())
                If allOffersVal = 1 Then
                    AllOffers = True
                Else allOffersVal = 2
                    OnlyTCOffers = True
                End If
            End If
        Else
            If (Request.Form.Item("allOffers") <> "") Then AllOffers = True
        End If
        If (Request.Form.Item("offerIDs-selected") <> "") Then OfferIDs = Request.Form.Item("offerIDs-selected")
        If (Request.Form.Item("allLocations") <> "") Then AllLocations = True
        If (Request.Form.Item("locationIDs-selected") <> "") Then LocationGroupIDs = Request.Form.Item("locationIDs-selected")
        If (Request.Form.Item("employeeStatusID") <> "") Then EmployeeStatusID = MyCommon.Extract_Val(Request.Form.Item("employeeStatusID"))
        If (Request.Form.Item("discountTypeID") <> "") Then DiscountTypeID = MyCommon.Extract_Val(Request.Form.Item("discountTypeID"))
        If (Request.Form.Item("discountMinimum") <> "" And IsNumeric(Request.Form.Item("discountMinimum"))) Then DiscountMinimum = Request.Form.Item("discountMinimum")
        If (Request.Form.Item("discountMaximum") <> "" And IsNumeric(Request.Form.Item("discountMaximum"))) Then DiscountMaximum = Request.Form.Item("discountMaximum")
        If (Request.Form.Item("allCustomerGroups") <> "") Then AllCustomerGroups = True
        If (Request.Form.Item("CGs-selected") <> "") Then CustomerGroupIDs = Request.Form.Item("CGs-selected")
        If (Request.Form.Item("EngineType") <> "") Then EngineId = Request.Form.Item("EngineType")
        If ReportTypeID = 2 Then
            If (Request.Form.Item("OfferEngineID") <> "") Then EngineID = Request.Form.Item("OfferEngineID")
        End If
        If (Request.Form.Item("offerstatuses") <> "") Then OfferStatus = Request.Form.Item("offerstatuses")
        If (Request.Form.Item("ExtInterfaces") <> "") Then ExtInterfaceID = Request.Form.Item("ExtInterfaces")
        If (Request.Form.Item("searchByname") <> "") Then OfferName = Request.Form.Item("searchByname")

        If LocationGroupIDs = "" Then AllLocations = True
        If ReportTypeID <> 2 Then
            If OfferIDs = "" Then AllOffers = True
        End If

        If CustomerGroupIDs = "" Then AllCustomerGroups = True

        If infoMessage = "" Then
            'Save the request
            MyCommon.QueryStr = "dbo.pt_Reports_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ReportTypeID", SqlDbType.Int).Value = ReportTypeID
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = Name
            MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
            If (StartDateString <> "") Then MyCommon.LRTsp.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
            If (EndDateString <> "") Then MyCommon.LRTsp.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
            MyCommon.LRTsp.Parameters.Add("@CustomerID", SqlDbType.NVarChar, 100).Value = CustomerID
            MyCommon.LRTsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CustomerTypeID
            MyCommon.LRTsp.Parameters.Add("@AllOffers", SqlDbType.Bit).Value = AllOffers
            MyCommon.LRTsp.Parameters.Add("@OfferIDs", SqlDbType.NVarChar, OfferIDs.Length + 1).Value = OfferIDs
            MyCommon.LRTsp.Parameters.Add("@AllLocations", SqlDbType.Bit).Value = AllLocations
            MyCommon.LRTsp.Parameters.Add("@LocationGroupIDs", SqlDbType.NVarChar, LocationGroupIDs.Length + 1).Value = LocationGroupIDs
            MyCommon.LRTsp.Parameters.Add("@allCustomerGroups", SqlDbType.Bit).Value = AllCustomerGroups
            MyCommon.LRTsp.Parameters.Add("@CustomerGroupIDs", SqlDbType.NVarChar, CustomerGroupIDs.Length + 1).Value = CustomerGroupIDs
            MyCommon.LRTsp.Parameters.Add("@EmployeeStatusID", SqlDbType.Int).Value = EmployeeStatusID
            MyCommon.LRTsp.Parameters.Add("@DiscountTypeID", SqlDbType.Int).Value = DiscountTypeID
            If (DiscountMinimum.ToString <> "") Then MyCommon.LRTsp.Parameters.Add("@DiscountMinimum", SqlDbType.Decimal).Value = DiscountMinimum
            If (DiscountMaximum.ToString <> "") Then MyCommon.LRTsp.Parameters.Add("@DiscountMaximum", SqlDbType.Decimal).Value = DiscountMaximum
            MyCommon.LRTsp.Parameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = ExtInterfaceID
            MyCommon.LRTsp.Parameters.Add("@OfferStatus", SqlDbType.NVarChar, 100).Value = OfferStatus
            MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = IIf((ReportTypeID = 2 OrElse ReportTypeID = 5), EngineID, DBNull.Value)
            MyCommon.LRTsp.Parameters.Add("@OfferName", SqlDbType.NVarChar, 200).Value = IIf((ReportTypeID = 2 OrElse ReportTypeID = 5), OfferName, DBNull.Value)
            MyCommon.LRTsp.Parameters.Add("@OnlyTCOffers", SqlDbType.Bit).Value = OnlyTCOffers
            MyCommon.LRTsp.Parameters.Add("@ReportID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            ReportID = MyCommon.LRTsp.Parameters("@ReportID").Value
            MyCommon.Close_LRTsp()
            MyCommon.Activity_Log(29, ReportID, AdminUserID, Copient.PhraseLib.Lookup("history.report-request", LanguageID))
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "reports-enhanced-list.aspx")
        End If
    End If

    If Deleted Then
        infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    End If
%>
<form action="#" method="post" id="mainform" name="mainform">
<input type="hidden" id="save" name="save" value="" />
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.EnhancedReports", LanguageID))%>
  </h1>
  <div id="controls">
    <input type="button" class="regular" id="cancel" name="cancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID)) %>"
      onclick="location.href='reports-enhanced-list.aspx'" />
    <input type="button" class="regular" id="request" name="request" value="<% Sendb(Copient.PhraseLib.Lookup("term.request", LanguageID)) %>"
      onclick="javascript:validateRequest();" <% Sendb(IIf(Deleted, " style=""display:none;""", "")) %> />
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    If Deleted Then
      GoTo endmain
    End If
  %>
  <div id="column">
    <div class="box" id="general">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>
      </h2>
      <input type="hidden" id="hidreporttypeid" value="<%=OldReportTypeID %>" />
      <input type="hidden" id="hidreportid" value="<%=OldReportID %>"/>
      <select id="reportTypeID" name="reportTypeID" onchange="OnChangeReportTypeID();">
        <%
          If MyCommon.Fetch_SystemOption("318") = "1" Then
                MyCommon.QueryStr = "select ReportTypeID, Name, PhraseID from ReportTypes with (NoLock);"
          Else
                MyCommon.QueryStr = "select ReportTypeID, Name, PhraseID from ReportTypes with (NoLock) where ReportTypeID NOT IN (2,5)"
          End If
          dt = MyCommon.LRT_Select()
          If dt.Rows.Count > 0 Then
            For Each row In dt.Rows
              Sendb("<option value=""" & MyCommon.NZ(row.Item("ReportTypeID"), 0) & """" & IIf(ReportTypeID = MyCommon.NZ(row.Item("ReportTypeID"), 0), " selected=""selected""", "") & ">")
              If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
                Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
              Else
                Sendb(MyCommon.NZ(row.Item("Name"), "&nbsp;"))
              End If
              Send("</option>")
            Next
          End If
        %>
      </select>
      <%--
        <label for="sortfield"><% Sendb(Copient.PhraseLib.Lookup("term.sortedby", LanguageID))%></label>
        <select id="sortfield" name="sortfield">
          <option value="1">field</option>
          <option value="2">field</option>
          <option value="3">field</option>
        </select>
        <select id="sortorder" name="sortorder">
          <option value="1"><% Sendb(StrConv(Copient.PhraseLib.Lookup("term.ascending", LanguageID), VbStrConv.Lowercase))%></option>
          <option value="2"><% Sendb(StrConv(Copient.PhraseLib.Lookup("term.descending", LanguageID), VbStrConv.Lowercase))%></option>
        </select>
      --%>
      &nbsp;&nbsp;
      <label for="name">
        <% Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%> &nbsp; <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label>
      <input type="text" id="name" name="name" class="long" maxlength="100" value="<% Sendb(Name) %>" />
    </div>
  </div>
  <div id="column1">
    <div class="box" id="customer">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))%>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.customers", LanguageID))%>">
        <tr>
          <td>
            <input type="text" class="medium" id="customerID" name="customerID" value="<%Sendb(CustomerID)%>"
              maxlength="256" />
            <select id="customerTypeID" name="customerTypeID">
              <%
                  'Data will be populated for the selected report type.
                  If (ReportTypeID = 1) Then
                MyCommon.QueryStr = "select TypeID, Description, PhraseID from CustomerTypes with (NoLock) order by TypeID;"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count > 0 Then
                  For Each row In dt.Rows
                    Sendb("<option value=""" & MyCommon.NZ(row.Item("TypeID"), 0) & """" & IIf(CustomerTypeID = MyCommon.NZ(row.Item("TypeID"), 0), " selected=""selected""", "") & ">")
                    If (MyCommon.NZ(row.Item("PhraseID"), 0) > 0) Then
                      Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                    Else
                      Sendb(MyCommon.NZ(row.Item("Description"), "(" & MyCommon.NZ(row.Item("TypeID"), 0) & ")"))
                    End If
                    Send("</option>")
                  Next
                End If
                  End If
              %>
            </select>
          </td>
        </tr>
      </table>
    </div>
    <div class="box" id="offers">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>">
        <%
            Send("<tr>")
            Send("  <td colspan=""2"" id=""allOffersTD"">")
            Send("    <input type=""checkbox"" id=""allOffers"" name=""allOffers"" onclick=""javascript:toggleOffers();""" & IIf(AllOffers, IIf(ReportTypeID <> 2, " checked=""checked""", ""), "") & " /><label for=""allOffers"">" & Copient.PhraseLib.Lookup("term.alloffers", LanguageID) & "</label>")
            Send("  </td>")
            Send("</tr>")
            Send("<tr>")
            Send("  <td colspan=""2"" align=""left"">")
            Send("<div id = ""allOffersDiv"">")
            Send("    <input type=""checkbox"" id=""allOffersOD"" name=""allOffersOD"" value=""1"" onclick=""javascript:toggleOffers();""" & IIf(AllOffers, IIf(ReportTypeID = 2, " checked=""checked""", ""), "") & " /><label for=""allOffersOD"">" & Copient.PhraseLib.Lookup("term.alloffers", LanguageID) & "</label>")
            Send("   &nbsp;&nbsp; &nbsp;  &nbsp;<input type=""checkbox"" id=""triggerOffersOD"" name=""allOffersOD"" value=""2"" onclick=""javascript:toggleOffers();""" & IIf(OnlyTCOffers, " checked=""checked""", "") & " /><label for=""triggerOffersOD"">" & Copient.PhraseLib.Lookup("term.triggercodeoffers", LanguageID) & " </label>")
            Send("  </td>")
            Send("</div></tr>")
            Send("<tr>")
            Send("<td colspan=""2"">")
            Send("<input type=""radio"" id=""functionradioOff1"" name=""FunctionRadioOff"" checked=""checked"" /><label for=""functionradioOff1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
            Send("<input type=""radio"" id=""functionradioOff2"" name=""FunctionRadioOff"" /><label for=""functionradioOff2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
            Send("<input type=""text"" class=""medium"" id=""functioninputOff"" name=""functioninputOff"" maxlength=""100"" onkeyup=""javascript:xmlPostOffer('OfferFeeds.aspx','OffReportCondition');"" /><br />")
            Send("<div id=""searchLoadDivOff"" style=""display:block;"" >&nbsp;</div>")
            Send("  </td>")
            Send("</tr>")

            Send("<tr>")
            Send("  <td colspan=""2"">")
            Send("<div id=""OffersDiv"">")
            Send("    <select id=""offerIDs-available"" name=""offerIDs-available"" class=""full"" size=""10"" multiple=""multiple"">")
            'Data will be populated for the selected report type.
            If (ReportTypeID = 2 OrElse ReportTypeID=3 OrElse ReportTypeID=4 ) Then
                MyCommon.QueryStr = "select OfferID, Name from Offers with (NoLock) where Deleted=0 and IsTemplate=0 " & _
                                    IIf(OfferIDs <> "", "and OfferID not in (" & OfferIDs & ") ", "") & _
                                    " union " & _
                                    "select IncentiveID as OfferID, IncentiveName as Name from CPE_Incentives with (NoLock) where Deleted=0 and IsTemplate=0 " & _
                                    IIf(OfferIDs <> "", "and IncentiveID not in (" & OfferIDs & ") ", "") & _
                                    "order by Name;"
                dt = MyCommon.LRT_Select()
                If dt.Rows.Count > 0 Then
                    For Each row In dt.Rows
                        Send("      <option value=""" & MyCommon.NZ(row.Item("OfferID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
                    Next
                End If
            End If
            Send("    </select>")
            Send("  </td>")
            Send("</tr>")
            Send("<tr>")
            Send("<td colspan=""2"">")
            If (RECORD_LIMIT > 0) Then
                Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
            End If
            Send("  </td>")
            Send("</tr>")
            Send("<tr>")
            Send("  <td style=""text-align:center;"">")
            Send("    <button id=""addOffers"" name=""addOffers"" style=""color:#00aa00;font-size:10px;width:170px;"">" & Copient.PhraseLib.Lookup("term.add", LanguageID) & " ▼</button>")
            Send("  </td>")
            Send("  <td style=""text-align:center;"">")
            Send("    <button id=""removeOffers"" name=""removeOffers"" style=""color:#aa0000;font-size:10px;width:170px;"">" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & " ▲</button>")
            Send("  </td>")
            Send("</tr>")
            Send("<tr>")
            Send("  <td colspan=""2"">")
            Send("    <select id=""offerIDs-selected"" name=""offerIDs-selected"" class=""full"" size=""3"" multiple=""multiple"">")
            If OfferIDs <> "" Then
                'Data will be populated for the selected report type to avoid delaying in loading the data
                If (ReportTypeID = 2 OrElse ReportTypeID = 3 OrElse ReportTypeID = 4) Then
                    MyCommon.QueryStr = "select OfferID, Name from Offers with (NoLock) where Deleted=0" & _
                                        "and OfferID in (" & OfferIDs & ") " & _
                                        " union " & _
                                        "select IncentiveID as OfferID, IncentiveName as Name from CPE_Incentives with (NoLock) where Deleted=0 " & _
                                        "and IncentiveID in (" & OfferIDs & ") " & _
                                        "order by Name;"
                    dt = MyCommon.LRT_Select()
                    For Each row In dt.Rows
                        Send("      <option value=""" & MyCommon.NZ(row.Item("OfferID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
                    Next
                End If
            End If
            Send("    </select>")
            Send("  </td>")
            Send("</tr>")
        %>
      </table>
    </div>
      <div class="box" id="CustomerGroup">
      <h2>
        Customer Groups
      </h2>
      <table summary="CustomerGroup">
        <%
          Send("<tr>")
            Send("  <td colspan=""2"" id=""AllCustomerGroupsTD"">")
            Send("    <input type=""checkbox"" id=""allCustomerGroups"" name=""allCustomerGroups"" onclick=""javascript:toggleCustomergroups();""" & IIf(AllCustomerGroups, " checked=""checked""", "") & " /><label for=""AllCustomerGroups"">All Customer Groups</label>")
          Send("  </td>")
          Send("</tr>")
            Send("<tr>")
            Send("  <td colspan=""2"">")
            Send("<input type=""radio"" id=""functionradio1"" name=""functionradio"" checked=""checked"" /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
            Send("<input type=""radio"" id=""functionradio2"" name=""functionradio"" /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
            Send("<input type=""text"" class=""medium"" id=""functioninputCG"" name=""functioninputCG"" maxlength=""100"" onkeyup=""javascript:xmlPostCG('OfferFeeds.aspx','CGReportCondition');"" /><br />")
             Send("<div id=""searchLoadDiv"" style=""display:block;"" >&nbsp;</div>")
            Send("  </td>")
            Send("</tr>")
          Send("<tr>")
          Send("  <td colspan=""2"">")
          Send("<div id=""CustomerGroupsDiv"">")
            Send("    <select id=""CGs-available"" name=""CGs-available"" class=""full"" size=""10"">")
            Dim topString As String = ""
            If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "
            'Data will be populated for the selected report type to avoid delaying in loading the data
            If (ReportTypeID = 5) Then
            MyCommon.QueryStr = "Select " & topString & " CustomerGroupID, ExtGroupID, Name from CustomerGroups with (NoLock) " & _
            " where Deleted = 0 And CustomerGroupID <> 1 And CustomerGroupID <> 2" & _
             IIf(CustomerGroupIDs <> "", "and CustomerGroupID not in (" & CustomerGroupIDs & ") ", "") & _
            " And BannerID Is null And NewCardholders = 0 And AnyCAMCardholder = 0" & _
            " order by Name"
            dt = MyCommon.LRT_Select()
          If dt.Rows.Count > 0 Then
            For Each row In dt.Rows
                    Send("      <option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
            Next
                End If
            End If
          Send("    </select>")
           Send("  </td>")
            Send("</div>")
            Send("</tr>")
            Send("</tr>")
            Send("<tr>")
            Send("<td >")
            If (RECORD_LIMIT > 0) Then
                Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
            End If
          Send("  </td>")
          Send("</tr>")
          Send("<tr>")
          Send("  <td style=""text-align:center;"">")
          Send("    <button id=""addCustomerGroups"" name=""addCustomerGroups"" style=""color:#00aa00;font-size:10px;width:170px;"">Add ▼</button>")
          Send("  </td>")
          Send("  <td style=""text-align:center;"">")
          Send("    <button id=""removeCustomerGroups"" name=""removeCustomerGroups"" style=""color:#aa0000;font-size:10px;width:170px;"">Remove ▲</button>")
          Send("  </td>")
          Send("</tr>")
          Send("<tr>")
          Send("  <td colspan=""2"">")
          Send("    <select id=""CGs-selected"" name=""CGs-selected"" class=""full"" size=""3"" multiple=""multiple"">")
          If CustomerGroupIDs <> "" Then
            'Data will be populated for the selected report type to avoid delaying in loading the data
                If (ReportTypeID = 5) Then
                MyCommon.QueryStr = " select CustomerGroupID, ExtGroupID, Name from CustomerGroups with (NoLock) " & _
                " where Deleted = 0 And CustomerGroupID <> 1 And CustomerGroupID <> 2 And BannerID Is null And NewCardholders = 0 And AnyCAMCardholder = 0" & _
                 IIf(CustomerGroupIDs <> "", "and CustomerGroupID in (" & CustomerGroupIDs & ") ", "") & _
                " order by Name"
                 dt = MyCommon.LRT_Select()
            For Each row In dt.Rows
                    Send("      <option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
            Next
                End If
            End If
          Send("    </select>")
          Send("  </td>")
          Send("</tr>")
        %>
      </table>
    </div>
    <div class="box" id="employees">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.employees", LanguageID))%>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.employees", LanguageID))%>">
        <tr>
          <td>
            <select id="employeeStatusID" name="employeeStatusID">
              <option value="0" <%Sendb(IIf(EmployeeStatusID = 0, " selected=""selected""", ""))%>>
                <%Sendb(Copient.PhraseLib.Lookup("term.nolimitation", LanguageID))%></option>
              <option value="1" <%Sendb(IIf(EmployeeStatusID = 1, " selected=""selected""", ""))%>>
                <%Sendb(Copient.PhraseLib.Lookup("reports.EmployeesOnly", LanguageID))%></option>
              <option value="2" <%Sendb(IIf(EmployeeStatusID = 2, " selected=""selected""", ""))%>>
                <%Sendb(Copient.PhraseLib.Lookup("reports.NonEmployeesOnly", LanguageID))%></option>
            </select>
          </td>
        </tr>
      </table>
    </div>
    <div class="box" id="discounts">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.discounts", LanguageID))%>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.discounts", LanguageID))%>">
        <tr>
          <td colspan="2">
            <select id="discountTypeID" name="discountTypeID" onchange="javascsript:toggleDiscount();">
              <option value="0">
                <%Sendb(Copient.PhraseLib.Lookup("term.nolimitation", LanguageID))%></option>
              <%
                Dim CMInstalled As Boolean = MyCommon.IsEngineInstalled(0)
                Dim CPEInstalled As Boolean = MyCommon.IsEngineInstalled(2)
                MyCommon.QueryStr = "select DiscountTypeID, IsNull(CPE_AmountTypeID, 0), IsNull(CPE_DiscountTypeID, 0), IsNull(CM_AmountTypeID, 0), Description, PhraseID " & _
                                    "from ReportDiscountTypes with (NoLock) " & _
                                    "where DiscountTypeID>0"
                If (CMInstalled = True AndAlso CPEInstalled = False) Then
                  MyCommon.QueryStr &= " and CM_AmountTypeID>0"
                ElseIf (CMInstalled = False AndAlso CPEInstalled = True) Then
                  MyCommon.QueryStr &= " and CPE_AmountTypeID>0"
                End If
                MyCommon.QueryStr &= " order by DiscountTypeID;"
                dt = MyCommon.LRT_Select()
                For Each row In dt.Rows
                  Sendb("<option value=""" & MyCommon.NZ(row.Item("DiscountTypeID"), 0) & IIf(DiscountTypeID = MyCommon.NZ(row.Item("DiscountTypeID"), 0), " selected=""selected""", "") & """>")
                  If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
                    Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID))
                  Else
                    Sendb(MyCommon.NZ(row.Item("Description"), ""))
                  End If
                  Send("</option>")
                Next
              %>
            </select>
          </td>
        </tr>
        <tr id="discountMinimumRow" style="display: none;">
          <td style="width: 70px;">
            <label for="discountMinimum">
              <%Sendb(Copient.PhraseLib.Lookup("term.minimum", LanguageID))%>:</label>
          </td>
          <td>
            <input type="text" id="discountMinimum" name="discountMinimum" style="width: 60px;"
              value="<%Sendb(DiscountMinimum)%>" />
          </td>
        </tr>
        <tr id="discountMaximumRow" style="display: none;">
          <td style="width: 70px;">
            <label for="discountMaximum">
              <%Sendb(Copient.PhraseLib.Lookup("term.maximum", LanguageID))%>:</label>
          </td>
          <td>
            <input type="text" id="discountMaximum" name="discountMaximum" style="width: 60px;"
              value="<%Sendb(DiscountMaximum)%>" />
          </td>
        </tr>
      </table>
    </div>
  </div>
  <!-- End column1 -->
  <div id="gutter">
  </div>
  <div id="column2">
      <div class="box" id="engines">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.Engine", LanguageID))%>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.Engine", LanguageID))%>">
        <tr>
          <td>
         <%MyCommon.QueryStr = "select EngineID,Description,PhraseID,DefaultEngine from PromoEngines with (NoLock) where Installed=1 and (EngineID<3 or EngineID=9);"
              Dim rstOffer As DataTable = MyCommon.LRT_Select
             If rstOffer.Rows.Count > 0 Then%> 
           <label for="OfferEngineID"><%Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID) )%>:</label>
            <select id="OfferEngineID" name="OfferEngineID" >
           <%  For Each row In rstOffer.Rows
                   Sendb("  <option value=""" & row.Item("EngineID") & """" & IIf(row.Item("EngineID") = EngineID, " selected=""selected""", "") & ">")
                If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                Else
                    Sendb(row.Item("Description"))
                End If
                Send("</option>")
               Next
               %> 
            </select>
                 <%End If%>
          </td>
        </tr>
      </table>
    </div>
    <div class="box" id="SearchByOfferName">
      <h2>
        Search By
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.searchby", LanguageID))%>">
        <tr>
          <td>
      <label for="name">
        <% Sendb(Copient.PhraseLib.Lookup("term.offername", LanguageID))%>:</label>
      <input type="text" id="searchByname" name="searchByname" class="long" maxlength="100" value="<% Sendb(OfferName) %>" />
          </td>
        </tr>
      </table>
    </div>
    <div class="box" id="offerstatus">
      <h2>
        Offer Status
      </h2>
      <table summary="OfferStatus">
        <tr>
          <td>
             <select id="offerstatuses" name="offerstatuses" class="full">
             <%  
                 Dim statuses() As String = {"All", "Active", "Development", "Expired", "Testing", "Pending", "Scheduled"}
                 For Each status As String In statuses
                 Sendb("<option value=""" & status & """" & IIf(status = OfferStatus, " selected=""selected""", "") & ">" & status & "</option>")
                 Next        

                 %>
              </select>
          </td>
        </tr>
      </table>
    </div>
       <div class="box" id="ExtInterfaceID">
         <h2>
        External InterfaceID
      </h2>
      <table summary="ExtInterfaces">
          <tr>
             <td>
            <%
                Send("    <select id=""ExtInterfaces"" name=""ExtInterfaces"">")
                MyCommon.QueryStr = "select ExtInterfaceID,Name from ExtCRMInterfaces with (NoLock) where deleted = 0;"
                dt = MyCommon.LRT_Select()
                If dt.Rows.Count > 0 Then
                    Sendb("<option value=""-1"">" & Copient.PhraseLib.Lookup("term.select", LanguageID) & "</option>")
                    For Each row In dt.Rows
                        Sendb("<option value=""" & MyCommon.NZ(row.Item("ExtInterfaceID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ExtInterfaceID"), 0) = ExtInterfaceID, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
                    Next
                End If
                Send("    </select>")
             %>
          </td>
          </tr>
      </table>
        </div>
    <div class="box" id="daterange">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.daterange", LanguageID))%>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.daterange", LanguageID))%>">
        <tr>
          <td>
            <input type="text" class="short" id="startDate" name="startDate" maxlength="10" value="<%Sendb(StartDateString)%>" />
            –
            <input type="text" class="short" id="endDate" name="endDate" maxlength="10" value="<%Sendb(EndDateString)%>" />
            <% Sendb("(" & MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern & ")")%>
          </td>
        </tr>
      </table>
    </div>
    <div class="box" id="locationgroups">
      <h2>
        <% Sendb(Copient.PhraseLib.Lookup("term.stores", LanguageID))%>
      </h2>
      <table summary="<%Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID))%>">
        <%
          Send("<tr>")
          Send("  <td colspan=""2"" id=""allLocationsTD"">")
          Send("    <input type=""checkbox"" id=""allLocations"" name=""allLocations"" onclick=""javascript:toggleLocations();""" & IIf(AllLocations, " checked=""checked""", "") & " /><label for=""allLocations"">" & Copient.PhraseLib.Lookup("term.allstores", LanguageID) & "</label>")
          Send("  </td>")
          Send("</tr>")
          Send("<tr>")
          Send("  <td colspan=""2"">")
          Send("    <select id=""locationIDs-available"" name=""locationIDs-available"" class=""full"" size=""10"">")
            'Data will be populated for the selected report type to avoid delaying in loading the data
            If (ReportTypeID = 3 OrElse ReportTypeID = 4) Then
          MyCommon.QueryStr = "select LocationGroupID, Name from LocationGroups with (NoLock) " & _
                              "where Deleted=0 and LocationGroupID>1 " & _
                              IIf(LocationGroupIDs <> "", "and LocationGroupID not in (" & LocationGroupIDs & ") ", "") & _
                              "order by Name;"
          dt = MyCommon.LRT_Select()
          If dt.Rows.Count > 0 Then
            For Each row In dt.Rows
              Send("      <option value=""" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
            Next
          End If
            End If
          Send("    </select>")
          Send("  </td>")
          Send("</tr>")
          Send("<tr>")
          Send("  <td style=""text-align:center;"">")
          Send("    <button id=""addLocations"" name=""addLocations"" style=""color:#00aa00;font-size:10px;width:170px;"">Add ▼</button>")
          Send("  </td>")
          Send("  <td style=""text-align:center;"">")
          Send("    <button id=""removeLocations"" name=""removeLocations"" style=""color:#aa0000;font-size:10px;width:170px;"">Remove ▲</button>")
          Send("  </td>")
          Send("</tr>")
          Send("<tr>")
          Send("  <td colspan=""2"">")
          Send("    <select id=""locationIDs-selected"" name=""locationIDs-selected"" class=""full"" size=""3"" multiple=""multiple"">")
          If LocationGroupIDs <> "" Then
            'Data will be populated for the selected report type to avoid delaying in loading the data
                If (ReportTypeID = 3 OrElse ReportTypeID = 4) Then
            MyCommon.QueryStr = "select LocationGroupID, Name from LocationGroups with (NoLock) " & _
                                "where Deleted=0 and LocationGroupID>1 " & _
                                IIf(LocationGroupIDs <> "", "and LocationGroupID in (" & LocationGroupIDs & ") ", "") & _
                                "order by Name;"
            dt = MyCommon.LRT_Select()
                End If
            For Each row In dt.Rows
              Send("      <option value=""" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
            Next
          End If
          Send("    </select>")
          Send("  </td>")
          Send("</tr>")
        %>
      </table>
    </div>
    <div class="box" id="SearchByEng">
      <h2>
        Search By
      </h2>
      <select id="EngineType" name="EngineType" >
        <%
            MyCommon.QueryStr =  "select EngineID,Description  from promoengines WITH(NOLOCK) WHERE Installed=1 AND Description NOT IN ('Website','Excentus Fuel Integration')"
          dt = MyCommon.LRT_Select()
          If dt.Rows.Count > 0 Then
                For Each row In dt.Rows
                    Sendb("<option value=""" & MyCommon.NZ(row.Item("EngineID"), 0) & """" & IIf(ReportTypeID = MyCommon.NZ(row.Item("EngineID"), 0), " selected=""selected""", "") & ">")
                    Sendb(MyCommon.NZ(row.Item("Description"), ""))
                    Send("</option>")
                Next
            End If
        %>
      </select>
    </div>
  </div>
  <!-- End column2 -->
  <%endmain:%>
</div>
<!-- End main -->
</form>
<script type="text/javascript">
  toggleReportType("<%=ReportTypeID%>");
  toggleOffers();
  toggleLocations();
  toggleDiscount();
</script>
<%
done:
  Send_BodyEnd("mainform", "type")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>