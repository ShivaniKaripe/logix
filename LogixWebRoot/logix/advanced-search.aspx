<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
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
  Dim row As DataRow
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim TokenString As String = ""
  Dim XID As String = ""
  Dim ExternalbuyerId As String = ""
  Dim OfferID As String = ""
  Dim OfferName As String = ""
  Dim Desc As String = ""
  Dim ROID As String = ""
  Dim CreatedBy As String = ""
  Dim LastUpdatedBy As String = ""
  Dim Engine As String = ""
  Dim Banner As String = ""
  Dim Category As String = ""
  Dim Product As String = ""
  Dim Priority As String = ""
  Dim Created1 As String = ""
  Dim Created2 As String = ""
  Dim Starts1 As String = ""
  Dim Starts2 As String = ""
  Dim Ends1 As String = ""
  Dim Ends2 As String = ""
  Dim Mclu As String = ""
  Dim udfs(4) As String
  Dim udfOptions(4) As Integer
  Dim udfValues(4) As String
  Dim udfValue2(4) As String
  Dim ProductionID As String = ""
  
  Dim XIDSelected As Integer
  Dim OfferIDSelected As Integer
  'al-5111
  Dim BuyerIDSelected As Integer
  Dim NameSelected As Integer
  Dim DescSelected As Integer
  Dim ROIDSelected As Integer
  Dim CreatedBySelected As Integer
  Dim LastUpdatedBySelected As Integer
  Dim EngineSelected As Integer
  Dim BannerSelected As Integer
  Dim CategorySelected As Integer
  Dim ProductSelected As Integer
  Dim PrioritySelected As Integer
  Dim CreatedSelected As Integer
  Dim StartsSelected As Integer
  Dim EndsSelected As Integer
  Dim SourceSelected As Integer
  Dim FavoriteSelected As Integer
  Dim McluSelected As Integer
  Dim ProductionIDSelected As Integer
  Dim BannersEnabled As Boolean = False
  Dim CustomerInquiry As Boolean = False
  Dim bExternalOfferList As Boolean = False
  Dim ProductGroup As Boolean = False
  Dim EnginesInstalled(-1) As Integer
  Dim rst As DataTable
  Dim udfTokenCount As Integer = 0
  Dim AdditionalSpecialCharacters As String
  Dim myurl As String
  Dim target As String
  Dim title As String
  
  Dim bProductionSystem As Boolean = True
  Dim bTestSystem As Boolean = False
  Dim bArchiveSystem As Boolean = False
  Dim bCmInstalled As Boolean = False
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)  
    
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  If (Request.QueryString("CustomerInquiry") <> "") Then
    CustomerInquiry = True
  End If
  
  If (Request.QueryString("ExternalOfferList") <> "") Then
    bExternalOfferList = True
  ElseIf (Request.QueryString("PGroupList") <> "") Then
    ProductGroup = True
  End If
  Response.Expires = 0
  MyCommon.AppName = "advanced-search.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  EnginesInstalled = MyCommon.GetInstalledEngines()

  bCmInstalled = MyCommon.IsEngineInstalled(0)
  bTestSystem = (MyCommon.Fetch_CM_SystemOption(77) = "1")
  bArchiveSystem = (MyCommon.Fetch_CM_SystemOption(77) = "2")
  If bTestSystem Or bArchiveSystem Then
    bProductionSystem = False
  Else
    bProductionSystem = True
  End If

  
  AdditionalSpecialCharacters = MyCommon.Fetch_SystemOption(171)
  
  ' MCLU corresponds to Catalina Coupon (reward type 8)
  Dim isMCLUEnabled As Boolean = False
  MyCommon.QueryStr = "select Enabled from RewardTypes with (NoLock) where RewardTypeId=8 and Enabled=1;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    isMCLUEnabled = True
  End If

  If (ProductGroup) Then
    Send_HeadBegin("term.productgroups", "term.search")
  Else
    Send_HeadBegin("term.offers", "term.search")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts(New String() {"datePicker.js"})
%>
<script type="text/javascript" language="javascript">
  window.name = "AdvSearch"
  var datePickerDivID = "datepicker";

  if (window.captureEvents){
      window.captureEvents(Event.CLICK);
      window.onclick=handlePageClick;
  }
  else {
      document.onclick=handlePageClick;
  }

<% Send_Calendar_Overrides(MyCommon) %>

  function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el=(typeof event!=='undefined')? event.srcElement : e.target        
    
      if (el != null) {
          var pickerDiv = document.getElementById(datePickerDivID);
          if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
              if (el.id!="createdDate1picker" && el.id!="createdDate2picker"
                && el.id!="startdate1picker" && el.id!="startdate2picker"
                && el.id!="enddate1picker" && el.id!="enddate2picker") {
                  if (!isDatePickerControl(el.className)) {
                    pickerDiv.style.visibility = "hidden";
                    pickerDiv.style.display = "none";            
                    if (calFrame != null) {
                      calFrame.style.visibility = "hidden";
                      calFrame.style.display = "none";
                    }
                  }
              } else  {
                  pickerDiv.style.visibility = "visible";            
                  pickerDiv.style.display = "block";            
                  if (calFrame != null) {
                    calFrame.style.visibility = 'visible';
                    calFrame.style.display = "block";
                  }
              }
          }
      }
  }

  function isDatePickerControl(ctrlClass) {
    var retVal = false;
    
    if (ctrlClass != null && ctrlClass.length >= 2) {
      if (ctrlClass.substring(0,2) == "dp") {
        retVal = true;
      }
    }

    return retVal;
  }
  
  function handleDateToFrom(selIndex, tdName, elemName) {
    var elemTD = document.getElementById(tdName);
    var elem = document.getElementById(elemName);
    
    if (elemTD != null) {
      if (selIndex == 3) {
        elemTD.style.display = "";
      } else {
        if (elem !=null) elem.value = "";
        elemTD.style.display = "none";
      }
    }
  }
  
  function submitForm() {
    document.mainform.submit();
    if (window.opener != null && !window.opener.closed) {
      window.opener.focus();
      window.close();
    }
  }
    
  function updateDateControls() {
    var elemCreated = document.getElementById("createdOption");
    var elemTdCreated = document.getElementById("tdCreate");
    var elemStart = document.getElementById("startOption");
    var elemTdStart = document.getElementById("tdStart");
    var elemEnd = document.getElementById("endOption");
    var elemTdEnd = document.getElementById("tdEnd");
    
    if (elemCreated != null && elemCreated.value == "3") {
      if (elemTdCreated == null) 
        elemTdCreated.style.display = "block";
      else
        elemTdCreated.style.display = "";
    }
    if (elemStart != null && elemStart.value == "3") {
      if (elemTdStart == null) 
        elemTdStart.style.display = "block";
      else
        elemTdStart.style.display = "";
    }
    if (elemEnd != null && elemEnd.value == "3") {
      if (elemTdEnd == null) 
         elemTdEnd.style.display = "block";
      else
         elemTdEnd.style.display = "";
    }
  }
   function selectUDF(field, row, optionSelect, value1, value2)
  {
	var option = field.selectedIndex 
	var xmlHttpReq = false;
	var xmlHttpReq2 = false
	var self2 = this;
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
	strURL = "UDFOptions.aspx?mode=AdvSearchOption&udf=" + option +"&optionSelect=" + optionSelect + "&row=" + row;
	
    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
            updateUDFOption(row, self.xmlHttpReq.responseText);
        }
    }
     self.xmlHttpReq.send();
	 
	 
	     // Mozilla/Safari
    if (window.XMLHttpRequest) {
        self2.xmlHttpReq2 = new XMLHttpRequest();
    }
    // IE
    else if (window.ActiveXObject) {
        self2.xmlHttpReq2 = new ActiveXObject("Microsoft.XMLHTTP");
    }
	
	strURL =  "UDFOptions.aspx?mode=AdvSearch&udf=" + option + "&row=" + row + "&val=" + value1 + "&val2=" + value2; 
	self2.xmlHttpReq2.open('POST', strURL, true);
    self2.xmlHttpReq2.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self2.xmlHttpReq2.onreadystatechange = function() {
        if (self2.xmlHttpReq2 !=null && self2.xmlHttpReq2.readyState == 4) {
            updateUDF(row, self2.xmlHttpReq2.responseText);
        }
    }
    self2.xmlHttpReq2.send();
	
  }
   function updateUDF(row,  text)
   {
		// alert('reg' + text);
		var trEnd =text.indexOf("<tr>");
		if(trEnd >-1) {
			
			document.getElementById("udfdata-"+row).innerHTML  = text.substring(0,trEnd);
			document.getElementById("tdUdfEnd-"+row).innerHTML=text.substring(trEnd);
		 }
		 else
		 {
			 document.getElementById("udfdata-"+row).innerHTML  = text
		 }
		 handleDateToFrom(0, "trUdfEnd-"+row, "udfEnd-"+ row);
   }
   function updateUDFOption(row,  text)
   {
		 //alert('opt ' +text);
		 document.getElementById("udfmid-"+row).innerHTML  = text;
   }
	
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(3)
  
  If (Logix.UserRoles.AccessOffers = False AndAlso ProductGroup = False) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  ElseIf (Logix.UserRoles.AccessProductGroups = False AndAlso ProductGroup = True) Then
    Send_Denied(2, "perm.pgroup-access")
    GoTo done
  End If
  
  If (Request.QueryString("tokens") <> "") Then
    Dim TokenRows As String()
    Dim TokenCols As String()
    Dim i As Integer
    
    TokenString = Request.QueryString("tokens")
    TokenRows = TokenString.Split("|")
    For i = 0 To TokenRows.GetUpperBound(0)
      TokenCols = TokenRows(i).Split(",")
      If (TokenCols.Length >= 4) Then
        Select Case TokenCols(0).ToUpper
          Case "XID"
            XIDSelected = Integer.Parse(TokenCols(1))
            XID = TokenCols(2)
          Case "ID"
            OfferIDSelected = Integer.Parse(TokenCols(1))
            OfferID = TokenCols(2)
            'added for al-5111
          Case "ExternalbuyerId"
            BuyerIDSelected = Integer.Parse(TokenCols(1))
            ExternalbuyerId = TokenCols(2)
          Case "NAME"
            NameSelected = Integer.Parse(TokenCols(1))
            OfferName = TokenCols(2)
          Case "DESC"
            DescSelected = Integer.Parse(TokenCols(1))
            Desc = TokenCols(2)
          Case "ROID"
            ROIDSelected = Integer.Parse(TokenCols(1))
            ROID = TokenCols(2)
          Case "CREATEDBY"
            CreatedBySelected = Integer.Parse(TokenCols(1))
            CreatedBy = TokenCols(2)
          Case "LASTUPDATEDBY"
            LastUpdatedBySelected = Integer.Parse(TokenCols(1))
            LastUpdatedBy = TokenCols(2)
          Case "ENGINE"
            EngineSelected = Integer.Parse(TokenCols(1))
            Engine = TokenCols(2)
          Case "BAN.NAME"
            BannerSelected = Integer.Parse(TokenCols(1))
            Banner = TokenCols(2)
          Case "CATEGORY"
            CategorySelected = Integer.Parse(TokenCols(1))
            Category = TokenCols(2)
          Case "PRODUCT"
            ProductSelected = Integer.Parse(TokenCols(1))
            Product = TokenCols(2)
          Case "PRIORITY"
            PrioritySelected = Integer.Parse(TokenCols(1))
            Priority = TokenCols(2)  
          Case "CREATED"
            CreatedSelected = Integer.Parse(TokenCols(1))
            Created1 = TokenCols(2)
            Created2 = TokenCols(3)
          Case "STARTS"
            StartsSelected = Integer.Parse(TokenCols(1))
            Starts1 = TokenCols(2)
            Starts2 = TokenCols(3)
          Case "ENDS"
            EndsSelected = Integer.Parse(TokenCols(1))
            Ends1 = TokenCols(2)
            Ends2 = TokenCols(3)
          Case "SOURCE"
            SourceSelected = Integer.Parse(TokenCols(1))
          Case "STARRED"
            FavoriteSelected = Integer.Parse(TokenCols(1))
          Case "MCLU"
            McluSelected = Integer.Parse(TokenCols(1))
            Mclu = TokenCols(2)
          Case "PRODUCTIONID"
            ProductionIDSelected = Integer.Parse(TokenCols(1))
            ProductionID = TokenCols(2)
        End Select
		
		If  TokenCols(0).Length >=6 AndAlso TokenCols(0).ToUpper.Substring(0,6)  = "UDFROW" Then
          Dim udfrow As Integer = TokenCols(0).Substring(6)
			
          Dim udfpk As String = TokenCols(1).ToUpper.Substring(4)
			
			MyCommon.QueryStr = "Select * from UserDefinedFields where UDFPK = " & udfpk
			dt = MyCommon.LRT_Select
			udfs(udfTokenCount) = udfpk
			udfOptions(udfTokenCount) =  Integer.Parse(TokenCols(2))

			udfValues(udfTokenCount) = TokenCols(3)
			If dt.Rows(0).Item("DataType")  = 2 Then


				udfValue2(udfTokenCount) = TokenCols(4)
			End If
			udfTokenCount +=1 
		End If
      End If
    Next
  End If
  If (bExternalOfferList) Then
    myurl = "Enhanced-extoffer-list.aspx"
    target = "OfferListWin"
    title = "advanced-search.offer-title"
  ElseIf (ProductGroup) Then
    myurl = "pgroup-list.aspx"
    target = "PGroupListWin"
    title = "advanced-search.pg-title"
  Else
    myurl = "offer-list.aspx"
    target = "OfferListWin"
    title = "advanced-search.offer-title"
  End If
%>
<form action="<%Sendb(myurl)%><%Sendb(IIf(CustomerInquiry, "?CustomerInquiry=1", "")) %>"
id="mainform" name="mainform" method="post" target="<%Sendb(target)%>" onsubmit="return submitForm();">
<input type="hidden" name="mode" id="mode" value="advancedsearch" />
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup(title, LanguageID))%>
    </h1>
    <div id="controls">
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column">
      <div class="box" id="criteria">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.criteria", LanguageID))%>
          </span>
        </h2>
        <center>
        <table style="width: 70%;" summary="<% Sendb(Copient.PhraseLib.Lookup("term.advancedsearchcriteria", LanguageID))%>">
          <% If (Not ProductGroup) Then%>
            <tr>
              <td>
                <label for="xidOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="xidOption" name="xidOption" class="mediumshort">
                  <option value="1" <% If(XIDSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(XIDSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(XIDSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(XIDSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(XIDSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="xid" name="xid" type="text" value="<% Sendb(XID) %>" class="mediumshort" maxlength="20" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="idOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="idOption" name="idOption" class="mediumshort">
                  <option value="1" <% If(OfferIDSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(OfferIDSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(OfferIDSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(OfferIDSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(OfferIDSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
              <input id="idSearch" name="idSearch" type="text" value="<% Sendb(OfferID) %>" class="mediumshort"/>
            </td>
          </tr>
          <%--displays only when byerid system option and ue engine is enabled--%>
          <%  If (MyCommon.Fetch_UE_SystemOption(169) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
          <tr>
            <td>
              <label for="buyeridOption">
                <% Sendb(Copient.PhraseLib.Lookup("term.buyerid", LanguageID))%>
                :</label>
            </td>
            <td>
              <select id="buyeridOption" name="buyeridOption" class="mediumshort">
                <option value="1" <% If(BuyerIDSelected=1) Then Sendb(" selected=""selected""") %>>
                  <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                </option>
                <option value="2" <% If(BuyerIDSelected=2) Then Sendb(" selected=""selected""") %>>
                  <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                </option>
                <option value="3" <% If(BuyerIDSelected=3) Then Sendb(" selected=""selected""") %>>
                  <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                </option>
                <option value="4" <% If(BuyerIDSelected=4) Then Sendb(" selected=""selected""") %>>
                  <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                </option>
                <option value="5" <% If(BuyerIDSelected=5) Then Sendb(" selected=""selected""") %>>
                  <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                </option>
              </select>
            </td>
            <td>
              <input id="buyeridSearch" name="buyeridSearch" type="text" value="<% Sendb(ExternalbuyerId) %>"
                class="mediumshort" />
            </td>
          </tr>
          <%End If%>
          <% If bCmInstalled And Not bProductionSystem Then%>
          <tr>
            <td>
              <label for="productionIdOption">
                <% Sendb(Copient.PhraseLib.Lookup("term.productionid", LanguageID))%>:</label>
              </td>
              <td>
                <select id="Select1" name="productionIdOption" class="mediumshort">
                  <option value="1"<% If(ProductionIDSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2"<% If(ProductionIDSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3"<% If(ProductionIDSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4"<% If(ProductionIDSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5"<% If(ProductionIDSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
              <input id="productionIdSearch" name="productionIdSearch" type="text" value="<% Sendb(ProductionID) %>"
                class="mediumshort" />
            </td>
          </tr>
          <% End If%>
          <tr>
            <td>
              <label for="nameOption">
                <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label>
              </td>
              <td>
                <select id="nameOption" name="nameOption" class="mediumshort">
                  <option value="1" <% If(NameSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(NameSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(NameSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(NameSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(NameSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="offerName" name="offerName" type="text" value="<% If AdditionalSpecialCharacters <> "" Then Sendb(OfferName.Replace(chr(34),"&quot;")) Else Sendb(OfferName) %>"
                  class="mediumshort" maxlength="468" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="descOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="descOption" name="descOption" class="mediumshort">
                  <option value="1" <% If(DescSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(DescSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(DescSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(DescSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(DescSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
              <input id="desc" name="desc" type="text" value="<% If AdditionalSpecialCharacters <> "" Then Sendb(Desc.Replace(chr(34),"&quot;")) Else Sendb(Desc) %>"
                class="mediumshort" maxlength="1000" />
              </td>
            </tr>
            <% If MyCommon.IsEngineInstalled(2) OrElse MyCommon.IsEngineInstalled(6) Then%>
            <tr>
              <td>
                <label for="roidOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.roid", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="roidOption" name="roidOption" class="mediumshort">
                  <option value="1" <% If(ROIDSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(ROIDSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(ROIDSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(ROIDSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(ROIDSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="roid" name="roid" type="text" value="<% Sendb(ROID) %>" class="mediumshort"/>
              </td>
            </tr>
            <% End If %>
          <% End If%>
            <tr>
              <td>
                <label for="createdByOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.createdby", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="createdByOption" name="createdByOption" class="mediumshort">
                  <option value="1" <% If(CreatedBySelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(CreatedBySelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(CreatedBySelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(CreatedBySelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(CreatedBySelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="createdby" name="createdby" type="text" value="<% Sendb(CreatedBy) %>"
                  class="mediumshort" maxlength="50" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="lastupdatedByOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.lastupdatedby", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="lastupdatedByOption" name="lastupdatedByOption" class="mediumshort">
                  <option value="1" <% If(LastUpdatedBySelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(LastUpdatedBySelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(LastUpdatedBySelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(LastUpdatedBySelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(LastUpdatedBySelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="lastupdatedby" name="lastupdatedby" type="text" value="<% Sendb(LastUpdatedBy) %>"
                  class="mediumshort" maxlength="50" />
              </td>
            </tr>
          <% If (Not ProductGroup) Then%>
            <% If EnginesInstalled IsNot Nothing AndAlso EnginesInstalled.Length > 1  AndAlso Not bEnableRestrictedAccessToUEOfferBuilder Then%>
            <tr>
              <td>
                <label for="engineOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="engineOption" name="engineOption" class="mediumshort">
                  <option value="1" <% If(EngineSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(EngineSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(EngineSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(EngineSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(EngineSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="engine" name="engine" type="text" value="<% Sendb(Engine) %>" class="mediumshort" maxlength="255" />
              </td>
            </tr>
            <% 
            Else
                 Dim dt1 As DataTable
                 MyCommon.QueryStr="Select Description From PromoEngines where Installed=1 and OfferBuilderSupported=1 "
                If((myurl = "offer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers) OrElse ((myurl = "Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.CreateUEOffers)  OrElse  (myurl = "Enhanced-extoffer-list.aspx" AndAlso Logix.UserRoles.AccessTranslatedUEOffers)) ) Then
                    'Do nothing
                Else
                    'Add condition to remove UE Engine as user has no permission to view UE offers
                     MyCommon.QueryStr &= " and EngineID <> 9"
                End If    
                   dt1 = MyCommon.LRT_Select
                    
            %>
                <tr>
              <td>
                <label for="engineOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID))%>
                  :</label>
              </td>
              <td>
                 <select id="engineOption" name="engineOption" class="mediumshort">
                  <option value="2" <% If(EngineSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                </select>
                </td>
               <td>
                <select id="engine" name="engine" class="mediumshort">
                <%
                 If dt1.Rows.Count > 0 Then
                    For Each row In dt1.Rows
                %>
                 <option value="<% Sendb(row.Item("Description"))%>" <% If(Engine=row.Item("Description")) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(row.Item("Description"))%>
                  </option>
                   <% 
                   Next 
                 End If 
                   %>
                </select>
                </td>
            </tr>
            <% End If %>
          <% If (BannersEnabled) Then%>
            <tr>
              <td>
                <label for="bannerOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.banner", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="bannerOption" name="bannerOption" class="mediumshort">
                  <option value="1" <% If(EngineSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(EngineSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(EngineSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(EngineSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(EngineSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="banner" name="banner" type="text" value="<% Sendb(Banner) %>" class="mediumshort" />
              </td>
            </tr>
            <% End If %>
            <tr>
              <td>
                <label for="categoryOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="categoryOption" name="categoryOption" class="mediumshort">
                  <option value="1" <% If(CategorySelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(CategorySelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(CategorySelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(CategorySelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(CategorySelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="category" name="category" type="text" value="<% Sendb(Category) %>" class="mediumshort" maxlength="255" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="productOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.product", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="productOption" name="productOption" class="mediumshort">
                  <option value="1" <% If(ProductSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(ProductSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(ProductSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(ProductSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(ProductSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="product" name="product" type="text" value="<% Sendb(Product) %>" class="mediumshort" maxlength="100" />
              </td>
            </tr>
            <% If bCmInstalled AndAlso Not (MyCommon.IsEngineInstalled(2) OrElse MyCommon.IsEngineInstalled(9)) Then%>
            <tr>
              <td>
                <label for="priorityOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.cm-offer-priority", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="priorityOption" name="priorityOption" class="mediumshort">
                  <option value="1" <% If(PrioritySelected=1) Then Sendb(" selected=""selected""") %>>
                    =</option>
                  <option value="2" <% If(PrioritySelected=2) Then Sendb(" selected=""selected""") %>>
                    &lt;=</option>
                  <option value="3" <% If(PrioritySelected=3) Then Sendb(" selected=""selected""") %>>
                    &gt;=</option>
                </select>
              </td>
              <td>
                <input id="priority" name="priority" type="text" value="<% Sendb(Priority) %>" class="mediumshort" />
              </td>
            </tr>
            <% End If %>
            <%  If bCmInstalled AndAlso isMCLUEnabled = True Then%>
            <tr>
              <td>
                <label for="mcluOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.mclu", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="mcluOption" name="mcluOption" class="mediumshort">
                  <option value="1" <% If(McluSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))%>
                  </option>
                  <option value="2" <% If(McluSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))%>
                  </option>
                  <option value="3" <% If(McluSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))%>
                  </option>
                  <option value="4" <% If(McluSelected=4) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))%>
                  </option>
                  <option value="5" <% If(McluSelected=5) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="mclu" name="mclu" type="text" value="<% Sendb(Mclu) %>" class="mediumshort" />
              </td>
            </tr>
            <% End If %>
            <tr>
              <td>
                <label for="createdOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="createdOption" name="createdOption" class="mediumshort" onchange="handleDateToFrom(this.selectedIndex, 'tdCreate', 'createdDate2');">
                  <option value="0" <% If(CreatedSelected=0) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.on", LanguageID))%>
                  </option>
                  <option value="1" <% If(CreatedSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.before", LanguageID))%>
                  </option>
                  <option value="2" <% If(CreatedSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.after", LanguageID))%>
                  </option>
                  <option value="3" <% If(CreatedSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.between", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="createdDate1" name="createdDate1" type="text" value="<% Sendb(Created1) %>"
                  class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="createdDate1picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                  title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('createdDate1',event);" />
              </td>
            </tr>
            <tr id="tdCreate" style="display: none;">
              <td>
              </td>
              <td>
              </td>
              <td>
                <input id="createdDate2" name="createdDate2" type="text" value="<% Sendb(Created2) %>"
                  class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="createdDate2picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                  title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('createdDate2',event);" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="startOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.starts", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="startOption" name="startOption" class="mediumshort" onchange="handleDateToFrom(this.selectedIndex, 'tdStart', 'startDate2');">
                  <option value="0" <% If(StartsSelected=0) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.on", LanguageID))%>
                  </option>
                  <option value="1" <% If(StartsSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.before", LanguageID))%>
                  </option>
                  <option value="2" <% If(StartsSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.after", LanguageID))%>
                  </option>
                  <option value="3" <% If(StartsSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.between", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="startDate1" name="startDate1" type="text" value="<% Sendb(Starts1) %>"
                  class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="startdate1picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                  title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('startDate1',event);" />
              </td>
            </tr>
            <tr id="tdStart" style="display: none;">
              <td>
              </td>
              <td>
              </td>
              <td>
                <input id="startDate2" name="startDate2" type="text" value="<% Sendb(Starts2) %>"
                  class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="startdate2picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                  title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('startDate2',event);" />
              </td>
            </tr>
            <tr>
              <td>
                <label for="endOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.ends", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="endOption" name="endOption" class="mediumshort" onchange="handleDateToFrom(this.selectedIndex, 'tdEnd', 'endDate2');">
                  <option value="0" <% If(EndsSelected=0) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.on", LanguageID))%>
                  </option>
                  <option value="1" <% If(EndsSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.before", LanguageID))%>
                  </option>
                  <option value="2" <% If(EndsSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.after", LanguageID))%>
                  </option>
                  <option value="3" <% If(EndsSelected=3) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.between", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
                <input id="endDate1" name="endDate1" type="text" value="<% Sendb(Ends1) %>" class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="enddate1picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                  title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('endDate1',event);" />
              </td>
            </tr>
            <tr id="tdEnd" style="display: none;">
              <td>
              </td>
              <td>
              </td>
              <td>
                <input id="endDate2" name="endDate2" type="text" value="<% Sendb(Ends2) %>" class="mediumshort" />
                <img src="../images/calendar.png" class="calendar" id="enddate2picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>"
                  title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('endDate2',event);" />
              </td>
            </tr>
            <% Send("<tr" & IIf(bExternalOfferList = True, "", " style=""display:none;""") & ">") %>
              <td>
                <label for="sourceOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.source", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="sourceOption" name="sourceOption" class="mediumshort">
                  <option value="0" <% If(SourceSelected=0) Then Sendb(" selected=""selected""") %>>
                    &nbsp;</option>
                  <%
                If ((MyCommon.IsEngineInstalled(0) Or MyCommon.IsEngineInstalled(9)) AndAlso MyCommon.Fetch_SystemOption(167) = "1") Then
                          Send("<option value=""-1""" & IIf(SourceSelected = "-1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.allsources", LanguageID) & "</option>")
                      End If
                    MyCommon.QueryStr = "select ExtInterfaceID, ExtCode, Name, PhraseID from ExtCRMInterfaces with (NoLock) " & _
                                        "where Active=1 and Deleted=0 and ExtInterfaceTypeID < 2 and ExtInterfaceID>0 order by Name;"
                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count > 0 Then
                      For Each row In dt.Rows
                        Sendb("<option value=""" & MyCommon.NZ(row.Item("ExtInterfaceID"), 0) & """" & IIf(SourceSelected = MyCommon.NZ(row.Item("ExtInterfaceID"), 0), " selected=""selected""", "") & ">")
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
              </td>
              <td>
              </td>
            </tr>
            <tr>
              <td>
                <label for="favoriteOption">
                  <% Sendb(Copient.PhraseLib.Lookup("term.favorite", LanguageID))%>
                  :</label>
              </td>
              <td>
                <select id="favoriteOption" name="favoriteOption" class="mediumshort">
                  <option value="0" <% If(FavoriteSelected=0) Then Sendb(" selected=""selected""") %>>
                      <% Sendb(Copient.PhraseLib.Lookup("term.selectoption", LanguageID))%>
                     &nbsp; </option>
                  <option value="6" <% If(FavoriteSelected=1) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.yes", LanguageID))%>
                  </option>
                  <option value="7" <% If(FavoriteSelected=2) Then Sendb(" selected=""selected""") %>>
                    <% Sendb(Copient.PhraseLib.Lookup("term.no", LanguageID))%>
                  </option>
                </select>
              </td>
              <td>
              </td>
            </tr>
          <% End If%>
          </table>
			<% 
				If MyCommon.Fetch_SystemOption(156) = "1" Then
					MyCommon.QueryStr = "select * from UserDefinedFields where AdvancedSearch = 1"
					dt = MyCommon.LRT_Select
					Dim rowCount as Integer =-1
					Dim row2 as DataRow
					Dim DataType as Integer
					Dim dst2 as DataTable
					Dim selectCounter as Integer = 0
					Dim UDFStringValue as String
					' Dim udfcount as Integer =0
					If dt.Rows.Count <=  5 Then
						If dt.Rows.Count > 0 Then
							rowCount = dt.Rows.Count-1
						End If
					Else
						rowCount = 4
					End If	
                    Send("<table style=""width: 75%;"" summary=""UDF Advanced Search Criteria"">")                    
					For udfcount as Integer = 0 to rowCount
						row = dt.Rows(udfcount)
						Dim optionSelect as Integer = udfOptions(udfcount)
						Dim selectCount2 as Integer = 0
						
						Send("<tr id=""udfrow-" & udfcount & """>")
						Send("<td>")

						Send("<select class=""medium"" id=""UDFDataType-" & udfcount & """ name=""UDFDataType-" & udfcount & """ onchange=""selectUDF(this," & udfcount  & ")"" >") 
			            'add each udf with advanced search to the drop down
			            For Each row2 In dt.Rows
			                If row2.Item("UDFPK") = udfs(udfcount) Then
			                    Sendb("<option value=""" & row2.Item("UDFPK") & """ selected=""selected"">" & row2.Item("Description") & "</option>")
			                Else
			                    If selectCount2 = selectCounter Then
			                        Sendb("<option value=""" & row2.Item("UDFPK") & """ selected=""selected"">" & row2.Item("Description") & "</option>")
			                    Else
			                        Sendb("<option value=""" & row2.Item("UDFPK") & """ >" & row2.Item("Description") & "</option>")
			                    End If
			                End If
			                selectCount2 = selectCount2 + 1
			            Next
						Send("</select>")
						Send("</td>")
						Send("<td id = ""udfmid-" & udfcount & """>")
						If udfs(udfCount) <> 0 Then
							MyCommon.QueryStr = "select DataType from UserDefinedFields where UDFPK = " & udfs(udfCount)
							dst2 = MyCommon.LRT_Select
							 DataType = dst2.Rows(0).Item("DataType")
						Else
							DataType = dt.Rows(udfcount).Item("DataType")
						End If
						Select Case DataType
			                Case 0, 1, 4, 5, 6 ' string , integer,listbox,nueric range,likert
			                    Send("<select id=""udfOption-" & udfcount & """ name=""udfOption-" & udfcount & """ class=""mediumshort"">")
			                    Send("<option value=""1""" & IIf(optionSelect = 1, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))

			                    Send("</option>")
			                    Send("<option value=""2""" & IIf(optionSelect = 2, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))

			                    Send("</option>")
			                    Send("<option value=""3""" & IIf(optionSelect = 3, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))

			                    Send("</option>")
			                    Send("<option value=""4""" & IIf(optionSelect = 4, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))

			                    Send("</option>")
			                    Send("<option value=""5""" & IIf(optionSelect = 5, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))

			                    Send("</option>")
			                    Send("</select>")
			                Case 2 ' date
			                    Send("<select id=""udfOption-" & udfcount & """ name=""udfOption-" & udfcount & """ class=""mediumshort"" onchange=""handleDateToFrom(this.selectedIndex, 'trUdfEnd-" & udfcount & "', 'udfEnd-" & udfcount & "');"">")
			                    Send("<option value=""0""" & IIf(optionSelect = 0, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.on", LanguageID))

			                    Send("</option>")
			                    Send("<option value=""1""" & IIf(optionSelect = 1, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.before", LanguageID))

			                    Send("</option>")
			                    Send("<option value=""2""" & IIf(optionSelect = 2, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.after", LanguageID))

			                    Send("</option>")
			                    Send("<option value=""3""" & IIf(optionSelect = 3, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.between", LanguageID))

			                    Send("</option>")
			                    Send("</select>")
			                Case 3 'boolean
			                    Send("<select id=""udfOption-" & udfcount & """ name=""udfOption-" & udfcount & """ class=""mediumshort"">")
			                    Send("<option value=""0""" & IIf(optionSelect = 0, " selected=""selected""", "") & ">")
			                    Send("&nbsp; </option>")
			                    Send("<option value=""6""" & IIf(optionSelect = 6, " selected=""selected""", "") & ">")
			                    Send(Copient.PhraseLib.Lookup("term.yes", LanguageID))
			                    Send("</option>")
			                    Send("<option value=""7""" & IIf(optionSelect = 7, " selected=""selected""", "") & ">")
			                    Sendb(Copient.PhraseLib.Lookup("term.no", LanguageID))
			                    Send("</option>")
			                    Send("</select>")
			            End Select
						Send("</td>")
						Send("<td id = ""udfdata-" & udfcount & """>")
						
						Select Case DataType
			                Case 0, 4, 6 ' string,list box, likert
			                    UDFStringValue = udfValues(udfcount)
			                    If AdditionalSpecialCharacters <> "" Then
			                        If UDFStringValue <> "" Then UDFStringValue = UDFStringValue.Replace(Chr(34), "&quot;")
			                    End If
			                    Send("<input id=""udf-" & udfcount & """ name=""udf-" & udfcount & """ type=""text"" value=""" & UDFStringValue & """ class=""medium"" />")
			                Case 1, 5 ' integer, numeric range
			                    Send("<input id=""udf-" & udfcount & """ name=""udf-" & udfcount & """ type=""text"" value=""" & udfValues(udfcount) & """ class=""medium"" />")
			                Case 2 'date
			                    Send("<input id=""udf-" & udfcount & """ name=""udf-" & udfcount & """ type=""text"" value=""" & udfValues(udfcount) & """  class=""mediumshort"" />")
			                    Send("<img src=""../images/calendar.png"" class=""calendar"" id=""enddate2picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """" & _
			                     " title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('udf-" & udfcount & "',event);"" />")
			            End Select
						Send("</td>")
						Send("</tr>")
						Send("<tr id=""trUdfEnd-" & udfcount & """" & IIF(udfValue2(udfCount) ="","style =""display: none;""","") & " >") 
						Send("<td/>")
						Send("<td/>")
						Send("<td id =""tdUdfEnd-"& udfCount & """>")
						Send("<input id=""udfEnd-" & udfcount &""" name=""udfEnd-" & udfcount & """ type=""text"" value=""" & udfValue2(udfCount) & """ class=""mediumshort"" />") 
						Send("<img src=""../images/calendar.png"" class=""calendar"" id=""enddate2picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) &"""" & _
									" title=" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('udfEnd-" & udfcount & "',event);"" />")
						Send("</td></tr>")
						selectCounter = selectCounter + 1
					Next
                    Send("</table>")
			    End If
			%>
          <br />
          <input type="submit" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
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
<script type="text/javascript">
<% Send_Date_Picker_Terms() %>
  updateDateControls();
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform")
  Logix = Nothing
  MyCommon = Nothing
%>
