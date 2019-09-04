<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" EnableEventValidation="false" %>

<%@ Import Namespace="Newtonsoft.Json" %>
<%@ Import Namespace="Copient.CommonIncConfigurable" %>
<%@ Import Namespace="System.Web.Script.Serialization" %>
<%@ Import Namespace="System.ServiceModel.Web" %>
<%@ Import Namespace="System.Web.Services" %>
<%@ Register Src="UserControls/ProductAttributeFilter.ascx" TagName="ProductAttributeFilter"
  TagPrefix="uc1" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="System.Web.UI.WebControls" %>
<%@ Import Namespace="CMS" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="System.IO" %>
<%
  ' *****************************************************************************
  ' * FILENAME: pgroup-edit.aspx
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
<script language="JavaScript" type="text/javascript">
    <% If (Logix.UserRoles.EditProductGroups = True And EditProductRegardlessOfBuyer  And ProductGroupID > 0) %>
    window.onunload= function(){
        if(document.mainform.GroupName.value != document.mainform.GroupName.defaultValue) {
            saveChanges = confirm('<%Sendb(Copient.PhraseLib.Lookup("sv-edit.ChangesMade", LanguageID)) %>');
            if (saveChanges) {
                if (document.mainform.elements['Save'] == null) {
                    saveElem = document.createElement("input");
                    saveElem.type = 'hidden';
                    saveElem.id = 'Save';
                    saveElem.name = 'Save';
                    saveElem.value = 'save';
                    document.mainform.appendChild(saveElem);
                }
                handleAutoFormSubmit();
            }
        }
    };
    <% End If %>

 window.onload = function () {
 try{
       if((document.getElementById("producthierarchy")!=null) && document.getElementById("producthierarchy").innerHTML!="")
       {
               document.getElementById("hdnProductHierarchyExist").value=true;
       }
       else
       {
              document.getElementById("hdnProductHierarchyExist").value=false;

       }
       ShoworHideDivs();
       }catch(ex){
       //do nothing
       }
    }

    // callback function for save changes on unload during navigate away
    function handleAutoFormSubmit() {
        window.onunload = null;
        document.mainform.action="pgroup-edit.aspx"
        document.mainform.submit();
    }

    function disableSaveCheck() {
        if (typeof spin === 'function' && btnClicked != 'Download' ) {
      spin('attributepgbuilder');//PAB
    }
    if(document.mainform.GroupName.value != document.mainform.GroupName.defaultValue) {
      document.getElementById("hdnIsPGNameUpdated").value = true;
    }    
        window.onunload = null;
        return true;
    }

    function SubmitForm() {
        document.mainform.submit();
    }

    function launchSearch() {
        openPopup('phierarchy-search.aspx?ProductGroupID=<% Sendb(ProductGroupID) %>' );
    }

    function isValidID() {
        var retVal = true;
        var elemNumericOnly = document.getElementById("NumericOnly");
        var elemID = document.getElementById("productid");
        var elemProductType = document.getElementById("producttype");
        var productType = elemProductType.options[elemProductType.selectedIndex].value;
        var elemMFCenabled = document.getElementById("hdnMFCenabled");
        if((elemID != null) && (elemID.value.length == 0)) {
            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.invalid", LanguageID) + " " + Copient.PhraseLib.Lookup("term.productid", LanguageID)) %>');
        }
        if ((elemNumericOnly != null) && (elemNumericOnly.value != "")) {
           if ((elemID != null) && (isNaN(elemID.value))) {
              retVal = false;
              alert('<%Sendb(Copient.PhraseLib.Lookup("product.mustbenumeric", LanguageID)) %>');
           }
        }
         // Manuf Family code enhancement
        if (elemMFCenabled.value && productType == 4 && (elemID.value.length > 8 || elemID.value.length < 5)) {
            retVal = false;
            alert('UPC must be between 5 to 8 characters for Manufacturer Family code');
        }
        return retVal;
    }

    function isValidPath() {
        var retVal = true;
        var frmElem = document.uploadform.browse
        var agt = navigator.userAgent.toLowerCase();
        var browser = '<% Sendb(Request.Browser.Browser) %>'

        if (browser == 'IE') {
          if (frmElem != null) {
             var filePath = frmElem.value

              if (filePath.length >=2) {
                  if (filePath.charAt(1)!=":") {
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

    function launchHierarchy(isLinking) {
        var popW = 700;
        var popH = 570;
        var url = 'phierarchytree.aspx?ProductGroupID=<% Sendb(ProductGroupID) %>&OfferID=<%Sendb(OfferID)%>&EngineID=<%Sendb(EngineID)%>&BuyerID=<%Sendb(BuyerID)%>';

        if (isLinking) {
          url += '&Linking=1'
        }

        hierWindow = window.open(url,"hierTree", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
        hierWindow.focus();
    }

    function launchNodes() {
        var popW = 700;
        var popH = 570;
        <% If (ProductGroupID > 0) Then %>
        var url = 'pgroup-edit-nodes.aspx?ProductGroupID=<% Sendb(ProductGroupID) %>&Name=' + escape('<%Sendb(GName.replace("'", "\'"))%>');
        <% Else %>
        var url = '';
        <% End If %>

        nodeWindow = window.open(url,"Nodes", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
        nodeWindow.focus();
    }

    function adjustItemsSelectBox() {
      var elemItems = document.getElementById("PKID");
      var elemDiv = document.getElementById("itemsDiv");
      var BORDER_WIDTH = 2;

      if (elemItems != null && elemDiv != null) {
        if (elemItems.clientWidth < elemDiv.clientWidth) {
          elemItems.style.width = (elemDiv.clientWidth - BORDER_WIDTH) + 'px';
        }
      }
    }

    function toggleDropdown1() {
        if (document.getElementById("actionsmenu1") != null) {
            bOpen = (document.getElementById("actionsmenu1").style.visibility != 'visible')
            if (bOpen) {
                document.getElementById("actionsmenu1").style.visibility = 'visible';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
            } else {
                document.getElementById("actionsmenu1").style.visibility = 'hidden';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
            }
        }
    }

    function submitShowAll() {
      var elem = document.getElementById("showall");

      if (elem != null) {
          elem.value = "true";
      }
      document.mainform.submit();
    }

	var isProductgroupModifiedfailed = null
	function PopUpModifyGroup()
	{
	    isProductgroupModifiedfailed = false;
		$('#modifypg').css("display","block");
		ResetModifyGroupDiv();
	}

	function IsValidRegularExpression()
	{
		var re = new RegExp("[^_A-Za-z0-9-/\r/\n,]");
		var bAllowSpacesTab  = '<% Sendb(MyCommon.Fetch_SystemOption(207))%>';
		var bAllowHyphen  = '<% Sendb(MyCommon.Fetch_SystemOption(208))%>';
		if(bAllowSpacesTab == 1) {
            re = new RegExp("[^_A-Za-z0-9 /\r/\n/\t,]");
		} 
		if(bAllowHyphen == 1) {
            re = new RegExp("[^-_A-Za-z0-9/\r/\n,]");
		}
		if(bAllowSpacesTab == 1 && bAllowHyphen == 1) {
            re = new RegExp("[^-_A-Za-z0-9 /\r/\n/\t,]");
        }		

		if (document.getElementById("pasteproducts").value.match(re)) {
			return false;

		}
		else {
		return true;
		}
	}

	function ModifyGroup()
	{
	    var prods=document.getElementById("pasteproducts").value;
		var confirmMsg = "" ;

		if(document.getElementsByName("modifyoperation")[0].checked==true){
			confirmMsg = '<% Sendb(Copient.PhraseLib.Lookup("confirm.productsreplace", LanguageID)) %>';
			}
		else if(document.getElementsByName("modifyoperation")[1].checked==true){
			confirmMsg = '<% Sendb(Copient.PhraseLib.Lookup("confirm.productsadd", LanguageID)) %>';
			}
		else if (document.getElementsByName("modifyoperation")[2].checked==true){
			confirmMsg= '<% Sendb(Copient.PhraseLib.Lookup("confirm.productsdelete", LanguageID)) %>';

		}
		if (IsValidRegularExpression())
		{
		if(prods!=null && prods!="" ){
			if(confirm(confirmMsg)){
				    var bCreateProducts  = '<% Sendb(MyCommon.Fetch_SystemOption(150))%>';
					if(document.getElementsByName("modifyoperation")[0].checked == true || document.getElementsByName("modifyoperation")[1].checked  == true)
					{
						if(bCreateProducts == 0)
						{
							alert('<%Sendb(Copient.PhraseLib.Lookup("pgroup-edit.productnotexist", LanguageID)) %>');
							return false;
						}
					}
					document.getElementById("modifyprodgroup").disabled = true;
					document.getElementById("modifypgclose").disabled=true;
					xmlhttpPost('XMLFeeds.aspx', 'ModifyProductsProductGroups');
				}

			else{return false}
			}
			else{
					alert('<% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.invalidproducts", LanguageID))%>');
					return false;
				}
	}
		else{
					alert('<% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.invaliddata", LanguageID))%>');
					return false;
			}
	}

    function ResetModifyGroupDiv()
    {
	    document.getElementById("modifypginfobar").innerText = "";

        if(document.getElementById("pasteproducts")!=undefined && document.getElementById("pasteproducts")!=null)
		    document.getElementById("pasteproducts").value=""
        if(document.getElementById("modifyproducttype")!=undefined && document.getElementById("modifyproducttype")!=null)
            document.getElementById("modifyproducttype").selectedIndex = 0;
        if(document.getElementsByName("modifyoperation")!=undefined && document.getElementsByName("modifyoperation")!=null)
            document.getElementsByName("modifyoperation")[0].checked=true;
    }

	function xmlhttpPost(strURL, mode) {

        var xmlHttpReq = false;
        var self = this;

        handleWait(true);

        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
        // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }

        self.xmlHttpReq.open('POST', strURL, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.send(getQueryString(mode));
        self.xmlHttpReq.onreadystatechange = function() {
            if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
                updatePage(self.xmlHttpReq.responseText);
            }
        }

    }

	 function handleWait(bShow) {
        var elem = document.getElementById("disabledBkgrd");

        if (elem != null) {
            elem.style.display = (bShow) ? 'block' : 'none';
        }
    }

	function getQueryString(mode) {

        var products = document.getElementById("pasteproducts").value;
        var opertaionType = get_radio_value();
        var productType = document.getElementById("modifyproducttype").value;
        var productGroupID=document.getElementById("ModifyProductGroupID").value;
        var bAllowHyphen  = '<% Sendb(MyCommon.Fetch_SystemOption(208))%>';
        if(bAllowHyphen == 0) {
            products = (products.toString().trim().replace(/\s/g, ', ')).replace(/-/g, '');
        } else {
            products = products.replace(/\r?\n/g, ', ');
        }
        return "Mode=" + mode + "&ProductGroupID="+ productGroupID + "&Products="+ products + "&OpertaionType="+ opertaionType + "&ProductType="+ productType;

    }
	function get_radio_value() {
            var inputs = document.getElementsByName("modifyoperation");
            for (var i = 0; i < inputs.length; i++) {
              if (inputs[i].checked) {
                return inputs[i].value;
              }
            }
          }
	 function updatePage(responseMsg) {

        var allUserMsg ='<% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.invalidproductcodes", LanguageID)) %>';

        if (responseMsg == 'OK') {
           	$('#modifypg').css("display","none");
			location.replace('pgroup-edit.aspx?ProductGroupID=<% Sendb(ProductGroupID) %>');
        }
		else {
			if (responseMsg.indexOf("~|") !=-1) {
				responseMsg=responseMsg.replace("~|","#");
				var responseArray =new Array();
				responseArray=responseMsg.split('#');
				if(responseArray[0]=="Fail")
				{
					isProductgroupModifiedfailed=true;
					document.getElementById("modifypginfobar").style.padding = "2px 2px 3px 4px";
					document.getElementById("modifypginfobar").innerHTML = allUserMsg;
					document.getElementById("modifyprodgroup").disabled = false;
					document.getElementById("modifypgclose").disabled=false;
					document.getElementById("pasteproducts").value=responseArray[1];
				}
				else if(responseArray[0]=="Invalid")
				{
					isProductgroupModifiedfailed=true;
					document.getElementById("modifypginfobar").style.padding = "2px 2px 3px 4px";
					document.getElementById("modifypginfobar").innerHTML = responseArray[1];
					document.getElementById("modifyprodgroup").disabled = false;
					document.getElementById("modifypgclose").disabled=false;
				}
			}
			else
			{
				ResetModifyGroupDiv();
				document.getElementById('modifypg').style.display='none';
	    		handleWait(false);
				location.reload();
			}

        }
        handleWait(false);

    }

        function closeModifyGroup()
        {
          ResetModifyGroupDiv();
          document.getElementById('modifypg').style.display='none';

          handleWait(false);
          if(isProductgroupModifiedfailed)
             location.replace('pgroup-edit.aspx?ProductGroupID=<% Sendb(ProductGroupID) %>');
        }

        function CancelReturnKey() {

            if (window.event.keyCode == 13){
                document.getElementById("save").click();
                return false;
                }
        }
   function getIDList()
   { 
    	if(typeof Groupgrid !== "undefined")
            UpdateProductChanges();//For groupgrid update the product and level exclude details
       if(typeof idList !== "undefined"){
        var nodelist = document.getElementById('NodeListID');
        nodelist.value=idList;
       }
   }




</script>
<form id="mainform" name="mainform" runat="server" action="#" onkeypress="return CancelReturnKey()"
method="post" onsubmit="return disableSaveCheck();">
<%
    If CreatedFromOffer Then
        Send("<input type=""hidden"" id=""OfferID"" name=""OfferID"" value=""" & OfferID & """ />")
        Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineID & """ />")
        Send("<input type=""hidden"" id=""slct"" name=""slct"" value=""" & GetCgiValue("slct") & """ />")
        Send("<input type=""hidden"" id=""ex"" name=""ex"" value=""" & GetCgiValue("ex") & """ />")
        Send("<input type=""hidden"" id=""condChanged"" name=""condChanged"" value=""" & GetCgiValue("condChanged") & """ />")
    End If
    Send("<input type=""hidden"" id=""hdnProductHierarchyExist"" name=""hdnProductHierarchyExist"" value=""" & ProductHierarchyExist.ToString.ToLower & """ />")
    Send("<input type=""hidden"" name=""showall"" id=""showall"" value=""" & ShowAllItems.ToString.ToLower & """ />")
    Send("<input id=""NodeListID"" type=""hidden"" name=""NodeListID"" value=""N""/>")
  If MyCommon.Fetch_SystemOption(iProductIdNumericOnly) = "1" Then
    Send("<input type=""hidden"" id=""NumericOnly"" name=""NumericOnly"" value=""true"" />")
  End If
  If bMFCenabled Then
    Send("<input type=""hidden"" id=""hdnMFCenabled"" name=""hdnMFCenabled"" value=""true"" />")
  End If

  If (IsSpecialGroup) Then
    Send("<div id=""intro"" style=""background:url('../images/notbg.png');"">")
  Else
    Send("<div id=""intro"">")
  End If
  If ProductGroupID = 0 Then
    If bStaticPG Then
      GNameTitle = Copient.PhraseLib.Lookup("term.newstaticPG", LanguageID)
    Else
      GNameTitle = Copient.PhraseLib.Lookup("term.newproductgroup", LanguageID)
    End If
    GFullName = GNameTitle
  Else
    MyCommon.QueryStr = "SELECT Name,buyerid, ProductGroupID, isnull(IsStatic,0) as IsStatic FROM ProductGroups with (NoLock) WHERE ProductGroupId = " & ProductGroupID & ";"
    rst2 = MyCommon.LRT_Select
    If (rst2.Rows.Count > 0) Then
      GNameTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
      GFullName = GNameTitle
      If (Len(GNameTitle) > 30) Then
        GNameTitle = Left(GNameTitle, 27) & "..."
      End If
      If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(rst2.Rows(0).Item("Buyerid"))) Then
        Dim buyerid As Int32 = rst2.Rows(0).Item("Buyerid")
        GNameTitle = Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " #" & ProductGroupID & ": " & "Buyer " & ExternalBuyerId & " - " & GNameTitle
      Else
        If bStaticEnabled Then
          bStaticPG = rst2.Rows(0).Item("IsStatic")
          If bStaticPG Then
            GNameTitle = Copient.PhraseLib.Lookup("term.staticPG", LanguageID) & " #" & ProductGroupID & ": " & GNameTitle
          Else
            GNameTitle = Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " #" & ProductGroupID & ": " & GNameTitle
                    End If
                Else
                    GNameTitle = Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " #" & ProductGroupID & ": " & GNameTitle
                End If
      End If
     End If

  End If
%>
<h1 id="title" title="<% Sendb(GFullName)%>">
  <% Sendb(GNameTitle)%>
</h1>
<div id="controls">
  <%
        
        
      Dim isPGEditable As Boolean = CheckIsPGEditable(MyCommon, Logix, bStoreUser, ProductGroupID,associatedOfferDT)
        
    If (ProductGroupID = 0) Then
      If ((Logix.UserRoles.CreateProductGroups AndAlso Not IsSpecialGroup) OrElse (bStaticPG AndAlso bCreateStaticPG)) Then
        Send_Save("onclick=""getIDList()""")
      End If
    Else
            If(bEnableRestrictedAccessToUEOfferBuilder) Then
                Dim tempdt As DataTable
                MyCommon.QueryStr = " Select ProductGroupID  from ProductGroups with (NoLock) where deleted=0 and ProductGroupID="& ProductGroupID &" and isnull(TranslatedFromOfferID,0) > 0 "
                tempdt = MyCommon.LRT_Select
                If tempdt.Rows.Count > 0 Then
                   IsTranslatedPG =True
                End If
            End If
            
    If bStaticPG Then
      ShowActionButton = bCreateStaticPG OrElse bEditStaticPG
    Else
      ShowActionButton = (Logix.UserRoles.CreateProductGroups) OrElse (Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) OrElse (Logix.UserRoles.DeleteProductGroups) OrElse (bDownloadPG)
      ' hide the action button if this user does not have permission to edit the system-wide exclusion group
      If (IsSpecialGroup AndAlso Not CanEditSpecialGroup) Then
        ShowActionButton = False
      End If
    End If

      If (ShowActionButton) Then
        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown1();"" />")
        Send("<div class=""actionsmenu"" id=""actionsmenu1"">")
          If bStaticPG Then
            If (bEditStaticPG) Then
              Send_Save("onclick=""getIDList()""")
            End If
            If bEditStaticPG Then
              Send_CopyGroup()
              Send_Upload()
              Send_Download()
            End If
            If bCreateStaticPG Then
              Sendb("<input type=""submit"" accesskey=""n"" class=""regular"" id=""newstatic"" name=""newstatic"" value=""" & Copient.PhraseLib.Lookup("term.new", LanguageID) & " " & Copient.PhraseLib.Lookup("term.static", LanguageID) & """" & " />")
            End If
            If bEditStaticPG Then
              Send_ReDeploy()
            End If
            If bEditStaticPG Then
              Send_ModifyGroup()
            End If
          Else
            If ((Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) AndAlso Not IsSpecialGroup AndAlso Not IsTranslatedPG) Then
              Send_Save("onclick=""getIDList()""")
            End If
            If (Logix.UserRoles.DeleteProductGroups AndAlso Not IsSpecialGroup AndAlso Not IsTranslatedPG AndAlso isPGEditable) Then
              Send_Delete()
            End If
            If (Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) AndAlso (ProductGroupID > 0) AndAlso (ProductGroupTypeID = 1 OrElse (ProductGroupTypeID = 2 AndAlso AttributePGEnabled = False)) AndAlso Not IsTranslatedPG AndAlso isPGEditable Then
              Send_Upload()
            End If
            If (((Logix.UserRoles.EditProductGroups OrElse bDownloadPG) And EditProductRegardlessOfBuyer ) AndAlso Not CreatedFromOffer AndAlso Not bStaticPG) Then
              Send_Download()
            End If
            If ((((Logix.UserRoles.EditProductGroups AndAlso Not bStaticPG) Or (bEditStaticPG AndAlso bStaticPG)) And EditProductRegardlessOfBuyer) AndAlso Not IsSpecialGroup AndAlso Not IsTranslatedPG) Then
              Send_CopyGroup()
            End If
            If (Logix.UserRoles.CreateProductGroups And Not CreatedFromOffer AndAlso Not IsTranslatedPG AndAlso isPGEditable) Then
              Send_New()
            End If
            If ((Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) And Not CreatedFromOffer AndAlso Not IsTranslatedPG AndAlso isPGEditable) Then
              Send_ReDeploy()
            End If
            If CreatedFromOffer Then
              Send_Close()
            End If
            If (Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) AndAlso (ProductGroupID > 0) AndAlso (ProductGroupTypeID = 1 OrElse (ProductGroupTypeID = 2 AndAlso AttributePGEnabled = False)) AndAlso Not IsTranslatedPG AndAlso isPGEditable Then
             Send_ModifyGroup()
           End If
          End If
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:169px;""></iframe>")
        End If
        Send("</div>")
      End If
            If MyCommon.Fetch_SystemOption(75) And Not CreatedFromOffer AndAlso Not IsTranslatedPG AndAlso isPGEditable Then
        If (Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(7, ProductGroupID, AdminUserID)
        End If
      End If
    End If
    Send("<input type=""hidden"" id=""ProductGroupID1"" name=""ProductGroupID"" value=""" & ProductGroupID & """ />")
  %>
</div>
</div>
<%
  If Request.Browser.Type = "IE6" Then
    IE6ScrollFix = " onscroll=""javascript:document.getElementById('uploader').style.display='none';document.getElementById('modify').style.display='none';document.getElementById('actionsmenu1').style.visibility='hidden';"""
  End If
%>
<div id="main" <% Sendb(IE6ScrollFix) %>>
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% If (statusMessage <> "") Then Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")%>
  <div class="column1">
    <div class="box" id="identity">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <label for="GroupName">
        <%Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
      <%
        If (GName Is Nothing) Then
          GName = ""
        End If
        Send("<input type=""text"" class=""" & If(CreatedFromOffer, "long", "longest") & """ id=""GroupName"" name=""GroupName"" maxlength=""100"" value=""" & GName.Replace("""", "&quot;") & """ /><br />")
        Send("<input type=""hidden"" id=""hdnIsPGNameUpdated"" name=""hdnIsPGNameUpdated"" value=""false"" />")
        Send("<input type=""hidden"" id=""hdnProductGroupTypeID"" name=""hdnProductGroupTypeID"" value=""" & ProductGroupTypeID & """ />")
        Send("<input type=""hidden"" id=""hdnBuyerID"" name=""hdnBuyerID"" value=""" & BuyerID & """ />")
        Sendb("<br class=""half"" />")
        'If UE Engine is installed, then onlu show Product Group
        If (MyCommon.IsEngineInstalled(9)) Then
          If (ProductGroupID <> 0) Then
            Sendb(Copient.PhraseLib.Lookup("term.buyer", LanguageID) & ": " & ExternalBuyerId & "<br />")
          Else
            '*******************************************************************'
            Dim tempbuyerId As Int32 = -2
            ' find all Buyers Created
            MyCommon.QueryStr = "select B.BuyerId,ExternalBuyerId from Buyers B inner join buyerroleusers BU on B.BuyerId = BU.BuyerId where BU.AdminUserID=" & AdminUserID & ";"

            rst2 = MyCommon.LRT_Select
            'if We have any buyers in Logix, then show Buyers section
            If rst2.Rows.Count > 0 Then

              'preselect the BuyerID, which is used before
              If Not Request.Cookies("DefaultBuyerForProd") Is Nothing Then
                tempbuyerId = Convert.ToInt32(Request.Cookies("DefaultBuyerForProd").Value)
              End If
              Send("<tr id=""buyers"">")
              Send("  <td valign=""top"">")
              Send("    <label for=""buyerID"">Buyer ID:</label>")
              Send("  </td>")
              Send("  <td valign=""top"">")
              Sendb("    <select id=""buyerID"" name=""buyerID""  class=""medium"" >")
              Send("      <option value=""-1"">— Select a Buyer —</option>")
              For Each temprow In rst2.Rows
                'Send("      <option value=""" & MyCommon.NZ(temprow.Item("BuyerId"), -1) & """>" & MyCommon.NZ(temprow.Item("ExternalBuyerId"), "") & "</option>")
                Send("      <option value=""" & MyCommon.NZ(temprow.Item("BuyerId"), -1) & """")
                If (tempbuyerId = MyCommon.NZ(temprow.Item("BuyerId"), -1)) Then
                  Send(" selected=""selected""")
                End If
                Send(">" & MyCommon.NZ(temprow.Item("ExternalBuyerId"), "") & "</option>")
              Next
              Send("    </select>")
              Send("  </td>")
              Send("</tr>")
            End If
            '*******************************************************************'
          End If
        End If
        If XID <> "" AndAlso XID <> "0" Then
          Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & XID & "<br />")
        End If
        If CreatedDate = Nothing Then
        Else
          Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
          longDate = CType(CreatedDate, Date)
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
          Send("<br />")
        End If
        If LastUpdate = Nothing Then
        Else
          Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
          longDate = CType(LastUpdate, Date)
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
          Send("<br />")
        End If
        If ProductGroupID <> 0 Then
          Send("<br class=""half"" />")
          Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
          If (ProductGroupTypeID = 2 AndAlso AttributeUpdatePending = True) Then
            Sendb("<span class=""red"">" & Copient.PhraseLib.Lookup("pgroup-edit.productsawaiting", LanguageID) & " </span>")
            Send("<small><a href=""pgroup-edit.aspx?ProductGroupID=" & ProductGroupID & "&OfferID=" & OfferID & "&EngineID=" & EngineID & """>" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "</a></small>")
          ElseIf (GroupSize = 1) Then
            Response.Write(GroupSize & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.product", LanguageID), VbStrConv.Lowercase))
          Else
            Response.Write(GroupSize & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.products", LanguageID), VbStrConv.Lowercase))
          End If
		  Send("<br class=""half"" />")
		  Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
		  If (ProductsWithoutDesc = 1) Then
            Response.Write(ProductsWithoutDesc & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.product", LanguageID) & " ", VbStrConv.Lowercase))
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.withoutdesc", LanguageID), VbStrConv.Lowercase))			
		  Else
            Response.Write(ProductsWithoutDesc & " ")
            Sendb(StrConv(Copient.PhraseLib.Lookup("term.products", LanguageID) & " ", VbStrConv.Lowercase))
			Sendb(StrConv(Copient.PhraseLib.Lookup("term.withoutdesc", LanguageID), VbStrConv.Lowercase))
		  End If
        End If
        MyCommon.QueryStr = "select ProductGroupID from ProdInsertQueue with (NoLock) where ProductGroupID=" & ProductGroupID
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          Send("<br />")
          Send("<span class=""red"">" & rst2.Rows.Count & " " & Copient.PhraseLib.Lookup("pgroup-edit.awaiting", LanguageID) & "</span>")
          Send("<small><a href=""pgroup-edit.aspx?ProductGroupID=" & ProductGroupID & "&OfferID=" & OfferID & "&EngineID=" & EngineID & """>" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "</a></small>")
        End If
        ' check if the hierarchy is resyncing right now.
        MyCommon.QueryStr = "select top 1 HRQ.StatusFlag from HierarchyResyncQueue as HRQ with (NoLock) " & _
                            "inner join ProdGroupItems as PGI with (NoLock) on PGI.ExtHierarchyID = HRQ.ExtHierarchyID " & _
                            "where PGI.Deleted=0 and PGI.ProductGroupID=" & ProductGroupID & " and PGI.ExtHierarchyID is not null;"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          Send("<br />")
          Send("<span class=""red"">" & Copient.PhraseLib.Lookup("pgroup-edit.resyncing", LanguageID) & "</span>")
          Send("<small><a href=""pgroup-edit.aspx?ProductGroupID=" & ProductGroupID & "&OfferID=" & OfferID & "&EngineID=" & EngineID & """>" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "</a></small>")
        End If

      %>
      <br />
      <hr class="hidden" />
    </div>
    <% If (ProductGroupID = 0 AndAlso AttributePGEnabled) Then%>
    <div class="box" id="attributetype">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.productgrouptype", LanguageID))%>
        </span>
      </h2>
      <asp:radiobuttonlist id="RadioButtonList1" runat="server">
      </asp:radiobuttonlist>
      <br />
    </div>
    <% End If%>
  </div>
  <div style="height: 25px; float: left; width: 10px;">
  </div>
  <% If (ProductGroupID > 0) AndAlso (ProductGroupTypeID = 1 OrElse (ProductGroupTypeID = 2 AndAlso AttributePGEnabled = False)) Then%>
  <div class="column2">
    <div class="box" id="hierarchylinking" <% if(productgroupid=0)then sendb(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.hierarchylinking", LanguageID))%>
        </span>
      </h2>
      <br class="half" />
      <input type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.LinkToHierarchy", LanguageID))%>"
                onclick="launchHierarchy(true);" <% sendb(IIf(ProductGroupID > 0 AndAlso (ProductGroupTypeID = 2 AndAlso AttributePGEnabled = False) , "disabled=''", "")) %>  <% sendb(IIf(Not isPGEditable or bstaticPG, "disabled='disabled'", "")) %> />
      <% If (ProductGroupID > 0 AndAlso (ProductGroupTypeID = 2 AndAlso AttributePGEnabled = False)) Then%>
      <br />
      <br class="half" />
      <span class="red">
        <% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.hierarchylinkingdisabled", LanguageID))%></span>
      <% End If%>
      <hr />
      <%
        ' Check if ProductHierarchyLink agent is currently processing the nodes for this group and only display node count if it is not
        MyCommon.QueryStr = "SELECT T1.ProductGroupID " & _
                            "FROM ProductAddAllFromNodeQueue T1 WITH (NoLock) " & _
                            "FULL OUTER JOIN ProductAjustLinksQueue T2 " & _
                            "ON T1.ProductGroupID = T2.ProductGroupID " & _
                            "WHERE T1.ProductGroupID = " & ProductGroupID & " OR T2.ProductGroupID = " & ProductGroupID & ";"

        Dim linkAgentProcessing As Boolean = MyCommon.LRT_Select.Rows.Count > 0

        MyCommon.QueryStr = "select ProductTypeID, PhraseID from ProductTypes with (NoLock)"
        rstProdTypes = MyCommon.LRT_Select

        ' find the number of linking items
        MyCommon.QueryStr = "select COUNT(*) as LinkSize from ProdGroupItems with (NoLock) " & _
                            "where ProductGroupID=" & ProductGroupID & " and Deleted=0 " & _
                            "  and ISNULL(ExtHierarchyID, '') <> '' and ISNULL(ExtNodeID, '') <> '';"
        rst = MyCommon.LRT_Select
        LinkSize = rst.Rows(0).Item("LinkSize")

        ' write linked hierarchy nodes
     
              If linkAgentProcessing Then
                    Sendb(Copient.PhraseLib.Lookup("term.linkagentprocessing", LanguageID))
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.linked", LanguageID))
                    If LinkSize > 0 Then
                        Send(" (" & LinkSize & " " & Copient.PhraseLib.Lookup("term.items", LanguageID).ToLower & ")")
                    End If
                    Send(":")
                End If
		

          Dim divContentCount As Integer = 0
          Dim MaxLimit As Integer = 5000
          MyCommon.QueryStr = "select top 5000 PGH.PKID, PGH.ExtHierarchyID, PGH.ExtNodeID, PGH.ExtHierarchyID + ' - ' + PH.Name + ' : ' +  " & _
                            "  PGH.ExtNodeID + ' ' +  PHN.Name as Label from ProdGroupHierarchies  as PGH with (NoLock) " & _
                            "inner join ProdHierarchies as PH with (NoLock) on PH.ExternalID = PGH.ExtHierarchyID " & _
                            "inner join PHNodes as PHN with (NoLock) on PHN.ExternalID = PGH.ExtNodeID and PHN.HierarchyID = PH.HierarchyID " & _
                            "where ProductGroupID=" & ProductGroupID & " order by Label;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          Send("<br />")
          Send("<div id=""linkitemsDiv"" onscroll=""handlePageClick(this);"" class=""boxscroll""" & If(CreatedFromOffer, " style=""width:310px;white-space: nowrap;""", " style=""white-space: nowrap;""") & """>")
          For Each row As DataRow In rst.Rows
                  If divContentCount = MaxLimit Then
                      Exit For
                  End If
            Send("<b><i>" & MyCommon.NZ(row.Item("Label"), "") & "</i></b><br />")
                  MyCommon.QueryStr = "select top 5000 PKID, PROD.ExtProductID, PROD.ProductTypeID, PROD.Description  from ProdGroupItems as PGI with (NoLock) " & _
                                "inner join Products as PROD with (NoLock) on PROD.ProductID = PGI.ProductID " & _
                                "where PGI.ProductGroupID=" & ProductGroupID & " and PGI.Deleted=0 " & _
                                "  and IsNull(PGI.ExtHierarchyID, '')= '" & MyCommon.NZ(row.Item("ExtHierarchyID"), "") & "' " & _
                                "  and ISNULL(PGI.ExtNodeID, '') = '" & MyCommon.NZ(row.Item("ExtNodeID"), "") & "' " & _
                                "order by PROD.ExtProductID;"
            rst2 = MyCommon.LRT_Select
            For Each row2 As DataRow In rst2.Rows
                      If divContentCount = MaxLimit Then
                          Exit For
                      End If
              descriptionItem = MyCommon.NZ(row2.Item("ExtProductID"), " ") & " " & MyCommon.NZ(row2.Item("Description"), " ") & "-" & Copient.PhraseLib.Lookup(rstProdTypes.Rows(MyCommon.NZ(row2.Item("ProductTypeID"), 1) - 1).Item("PhraseID"), LanguageID)
              Send(StrClone("&nbsp;", 2) & descriptionItem & "<br />")
                      divContentCount += 1
            Next
          Next
          Send("</div>")

        Else
          Send(Copient.PhraseLib.Lookup("term.none", LanguageID))
        End If
        Send("<br />")
          If LinkSize <> 0 AndAlso Not linkAgentProcessing Then
              Send(Copient.PhraseLib.Lookup("term.showing", LanguageID) & " " & divContentCount & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & LinkSize & " " & Copient.PhraseLib.Lookup("term.items", LanguageID).ToLower())
              Send("<br />")
              Send("<br />")
          End If
        ' write excluded hierarchy nodes
        LinkSize = 0
        LabelSize = 0
        Dim top10 As Integer =0
        HTMLBuf = New StringBuilder()
        HTMLBuf.AppendLine(Copient.PhraseLib.Lookup("term.excluded", LanguageID) & " (###):")
        MyCommon.QueryStr = "select  PGHE.PKID, PGHE.ExtHierarchyID, PGHE.LevelID as ExtNodeID, PGHE.ExtHierarchyID + ' - ' + PH.Name + ' : ' +   " & _
                            "  ' Node ' + PGHE.LevelID + ' ' + PHN.Name as Label from ProdGroupHierarchyExclusions  as PGHE with (NoLock) " & _
                            "inner join ProdHierarchies as PH with (NoLock) on PH.ExternalID = PGHE.ExtHierarchyID " & _
                            "inner join PHNodes as PHN with (NoLock) on PHN.ExternalID = PGHE.LevelID and PHN.HierarchyID = PH.HierarchyID " & _
                            "where ProductGroupID=" & ProductGroupID & " and PGHE.HierarchyLevel=1 " & _
                            "order by ExtHierarchyID, ExtNodeID;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          HasExcludedNodes = True
          HTMLBuf.AppendLine("<br />")
          HTMLBuf.AppendLine("<div id=""excludeditemsDiv"" onscroll=""handlePageClick(this);"" class=""boxscroll""" & If(CreatedFromOffer, " style=""width:310px;""", " style=""white-space: nowrap;""") & """>")
          For Each row As DataRow In rst.Rows                   
            If top10 <= 10 then
                  HTMLBuf.AppendLine("<b><i>" & MyCommon.NZ(row.Item("Label"), "") & "</i></b><br />") 
            End If                  
            LabelSize += 1
            MyCommon.QueryStr = "dbo.pa_Products_FindAllFromNode"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ExtHierarchyID", SqlDbType.NVarChar, 100).Value = MyCommon.NZ(row.Item("ExtHierarchyID"), "")
            MyCommon.LRTsp.Parameters.Add("@ExtNodeID", SqlDbType.NVarChar, 100).Value = MyCommon.NZ(row.Item("ExtNodeID"), "")
            rst2 = MyCommon.LRTsp_select
            LinkSize += rst2.Rows.Count                 
            i = 0
            If top10 <= 10 then   
            For Each row2 As DataRow In rst2.Rows
              i += 1
              descriptionItem = MyCommon.NZ(row2.Item("ExtProductID"), " ") & " " & MyCommon.NZ(row2.Item("Description"), " ") & "-" & Copient.PhraseLib.Lookup(rstProdTypes.Rows(MyCommon.NZ(row2.Item("ProductTypeID"), 1) - 1).Item("PhraseID"), LanguageID)
              HTMLBuf.AppendLine(StrClone("&nbsp;", 2) & descriptionItem & "<br />")
             
              ' show only the first 100 products
              If i >= 100 Then
                HTMLBuf.AppendLine(StrClone("&nbsp;", 2) & Copient.PhraseLib.Lookup("term.more", LanguageID) & "...<br />")
                LabelSize += 1
                Exit For
              End If
            Next
                      End If
            MyCommon.Close_LRTsp()
            top10 += 1 
          Next
        End If

        ' add the individual items
        MyCommon.QueryStr = "select PGHE.PKID, PROD.ExtProductID, PROD.ProductTypeID, PROD.Description, " & _
                            "  PGHE.ExtHierarchyID + ' - ' + PH.Name + ' : ' +   PGHE.LevelID + ' ' + PROD.Description as Label " & _
                            "from ProdGroupHierarchyExclusions  as PGHE with (NoLock) " & _
                            "inner join ProdHierarchies as PH with (NoLock) on PH.ExternalID = PGHE.ExtHierarchyID " & _
                            "inner join Products as PROD with (NoLock) on PROD.ProductID= PGHE.LevelID " & _
                            "where ProductGroupID=" & ProductGroupID & " and PGHE.HierarchyLevel=2 order by Label;"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          HasExcludedItems = True
          If Not HasExcludedNodes Then
            HTMLBuf.AppendLine("<br />")
            HTMLBuf.AppendLine("<div id=""excludeditemsDiv"" onscroll=""handlePageClick(this);"" class=""boxscroll""" & If(CreatedFromOffer, " style=""width:310px;""", " style=""white-space: nowrap;""") & """>")
          End If
          HTMLBuf.AppendLine("<b><i>" & Copient.PhraseLib.Lookup("pgroup-edit.ExcludedItems", LanguageID) & "</i></b><br />")
          LabelSize += 1
          i = 0
          For Each row As DataRow In rst.Rows
            i += 1
            descriptionItem = MyCommon.NZ(row.Item("ExtProductID"), " ") & " " & MyCommon.NZ(row.Item("Description"), " ") & "-" & Copient.PhraseLib.Lookup(rstProdTypes.Rows(MyCommon.NZ(row.Item("ProductTypeID"), 1) - 1).Item("PhraseID"), LanguageID)
            HTMLBuf.AppendLine(StrClone("&nbsp;", 2) & descriptionItem & "<br />")
            LinkSize += 1
            If i >= 100 Then
              HTMLBuf.AppendLine(StrClone("&nbsp;", 2) & Copient.PhraseLib.Lookup("term.more", LanguageID) & "...<br />")
              LabelSize += 1
              Exit For
            End If
          Next
        End If

        If HasExcludedNodes Or HasExcludedItems Then
          HTMLBuf.AppendLine("</div>")
        End If

        If LinkSize > 0 Then
          Send(HTMLBuf.ToString.Replace("(###)", "(" & LinkSize & " " & Copient.PhraseLib.Lookup("term.items", LanguageID) & ")"))
        Else
          Send(Copient.PhraseLib.Lookup("term.excluded", LanguageID) & " : " & Copient.PhraseLib.Lookup("term.none", LanguageID))
        End If

        'If Not HasExcludedNodes AndAlso Not HasExcludedItems Then
        '  Send(Copient.PhraseLib.Lookup("term.excluded", LanguageID) & " : " & Copient.PhraseLib.Lookup("term.none", LanguageID))
        'End If
      %>
      <br class="half" />
      <hr class="hidden" />
    </div>
  </div>
  <% End If%>
  <% If (ProductGroupID > 0) AndAlso (ProductGroupTypeID = 1 OrElse (ProductGroupTypeID = 2 AndAlso AttributePGEnabled = False)) Then%>
  <div class="columnfull" <% if(productgroupid=0)then send(" style=""visibility: hidden;""") %>>
    <div class="box" style="overflow: auto" id="addproducts">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("pgroup-edit.addremove", LanguageID))%>
        </span>
      </h2>
      <br class="half" />
      <div style="float: left; width: 360px;">
        <span style="position: relative">
          <%
            If (ShowAllItems OrElse GroupSize <= 100) AndAlso rstItems IsNot Nothing Then
              Sendb(Copient.PhraseLib.Lookup("pgroup-edit.all-items-note", LanguageID) & " (" & rstItems.Rows.Count & " ")
              If (rstItems.Rows.Count = 1) Then
                Sendb(Copient.PhraseLib.Lookup("term.product", LanguageID).ToString.ToLower & ")<br />")
              Else
                Sendb(Copient.PhraseLib.Lookup("term.products", LanguageID).ToString.ToLower & ")<br />")
              End If
            Else
              Sendb(Copient.PhraseLib.Lookup("pgroup-edit.listnote", LanguageID) & "<br />")
            End If
          %>
        </span>
        <select name="PKID" id="PKID" size="12" multiple="multiple" onscroll="handlePageClick(this);"
          class="longer" <%Sendb(If(CreatedFromOffer, " style=""width:310px;overflow-x: scroll;""", "style=""width:348px;overflow-x: scroll;""")) %>>
          <%
            descriptionItem = String.Empty
            If (GroupSize > 0) Then
              For Each row As DataRow In rstItems.Rows
                descriptionItem = MyCommon.NZ(row.Item("ExtProductID"), " ") & " " & MyCommon.NZ(row.Item("Description"), " ") & "-"
                If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
                  descriptionItem &= Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
                Else
                  If MyCommon.NZ(row.Item("ProductType"), "") <> "" Then
                    descriptionItem &= row.Item("ProductType")
                  Else
                    descriptionItem &= Copient.PhraseLib.Lookup("term.unknown", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.type", LanguageID), VbStrConv.Lowercase) & " " & MyCommon.NZ(row.Item("ProductTypeID"), 0)
                  End If
                End If
                Send("     <option style=""width: 300%;"" value=""" & row.Item("PKID") & """>" & descriptionItem & "</option>")
              Next
            End If
          %>
        </select>
        <br />
        <%
        If bStaticPG Then
          If (bEditStaticPG) Then
            Send("    <br class=""half"" /><input type=""submit"" class=""large"" id=""remove"" name=""remove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width:150px;"" value=""" & Copient.PhraseLib.Lookup("term.removefromlist", LanguageID) & """  />")
          Else
            Send("    <br class=""half"" /><input type=""submit"" class=""large"" id=""remove"" name=""remove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width:150px;"" value=""" & Copient.PhraseLib.Lookup("term.removefromlist", LanguageID) & """ disabled=""disabled""  />")
          End If
        Else
          If (Not IsSpecialGroup OrElse (IsSpecialGroup And CanEditSpecialGroup)) AndAlso Not IsTranslatedPG AndAlso isPGEditable Then
            If (Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) Then
              Send("    <br class=""half"" /><input type=""submit"" class=""large"" id=""remove"" name=""remove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width:150px;"" value=""" & Copient.PhraseLib.Lookup("term.removefromlist", LanguageID) & """ />")
            End If
          End If
        End If
         If (Not ShowAllItems AndAlso GroupSize > 100 AndAlso Not IsTranslatedPG) Then
            Send("<input class=""regular"" id=""btnShowAll"" name=""btnShowAll"" type=""button"" value=""" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & """ onclick=""submitShowAll();"" />")
          End If%>
        <br />
      </div>
      <div style="float: left;">
        <% If (Not IsSpecialGroup OrElse (IsSpecialGroup And CanEditSpecialGroup)) Then%>
        <div style="padding-top: 15px;">
          <input type="button" id="hierarchy" name="hierarchy" value="<%Sendb(Copient.PhraseLib.Lookup("pgroup-edit.modifyusinghierarchy", LanguageID) & "...")%>"
                        onclick="launchHierarchy(false);" <% sendb(IIf(IsTranslatedPG or bStaticPG, "disabled='disabled'", "")) %>  <% sendb(IIf(Not isPGEditable , "disabled='disabled'", "")) %>/>
          <% If (ShowViewSelected) Then%>
          <input type="button" id="btnNodes" name="btnNodes" value="<%Sendb(Copient.PhraseLib.Lookup("pgroup-edit.viewselected", LanguageID) & "...")%>"
                        onclick="launchNodes();" <% sendb(IIf(IsTranslatedPG , "disabled='disabled'", "")) %> <% sendb(IIf(Not isPGEditable , "disabled='disabled'", "")) %>/>
          <% End If%>
        </div>
        <br class="clear" />
        <hr />
        <% End If%>
        <% If (Not IsSpecialGroup OrElse (IsSpecialGroup And CanEditSpecialGroup)) Then%>
        <table cellpadding="1" cellspacing="1">
          <tr>
            <td>
              <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>:
            </td>
            <td>
              <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>:
            </td>
          </tr>
          <tr>
                        <% 
                            MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
                            rst = MyCommon.LRT_Select
                            If rst IsNot Nothing Then
                                IDLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
                            End If
                            'Integer.TryParse(MyCommon.Fetch_SystemOption(52), IDLength)%>
            <td>
              <input type="text" id="productid" maxlength="120" name="ExtProductID" <% Sendb(If(IDLength > 0, " maxlength=""" & IDLength & """", ""))%>
                style="<%Sendb(If(CreatedFromOffer, "width:115px;", "width:137px;")) %>" value="" />
            </td>
            <td>
              <select id="producttype" name="producttype">
                <%
                
                  'BZ2079: UE-feature-removal #: Remove unsupported product types for UE (Mix/Match Code, Manufacturer Family code, Pool Code)
                  '        To restore previous functionality: remove the all code in the If statement checking engines except the query without a where clause.
                  If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.Catalina) _
                     OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) Then
                    MyCommon.QueryStr = "select ProductTypeID,PhraseID from ProductTypes with (NoLock)"
                  Else
                    'Allow Manufacturer Family Codes when system option enabled
                    If (bMFCenabled = 0) Then
                    MyCommon.QueryStr = "select ProductTypeID,PhraseID from ProductTypes with (NoLock) where ProductTypeID not in (3,4,5);"
                    Else
                      MyCommon.QueryStr = "select ProductTypeID,PhraseID from ProductTypes with (NoLock) where ProductTypeID not in (3,5);"
                    End If
                  End If
                  rst2 = MyCommon.LRT_Select
                  For Each row2 As DataRow In rst2.Rows
                    ProductTypeID = row2.Item("ProductTypeID")
                    If (Not CpeEngineOnly Or (CpeEngineOnly And (ProductTypeID <= 2 or (ProductTypeID = 4 And bMFCenabled)))) Then
                      Send("     <option value=""" & row2.Item("ProductTypeID") & """>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</option>")
                    End If
                  Next
                %>
              </select>
            </td>
          </tr>
        </table>
        <div>
          <label for="productdesc">
            <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
          <input type="text" id="productdesc" style="<%Sendb(IIf(CreatedFromOffer, "width:310px;", "width:347px;")) %>"
            name="productdesc" maxlength="200" value="" /><br />
          <br class="half" />
        </div>
        <%
          Dim bEdit As Boolean = bStaticPG Or (Not bStaticPG And Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer)
          Dim bDisabled As Boolean = (bStaticPG AndAlso Not bEditStaticPG) Or (Not bStaticPG AndAlso IsTranslatedPG)
        %>
        <% If bEdit Then%>
        <div style="float: left;">
          <input type="submit" class="large" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID)) %>"
                        onclick="return isValidID();" <% sendb(IIf(bDisabled , "disabled='disabled'", "")) %> <% sendb(IIf(Not isPGEditable , "disabled='disabled'", "")) %>/></div>
        <div style="float: right; margin-right: 20px">
          <input type="submit" class="large" id="mremove" name="mremove" onclick="if(confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.remove", LanguageID)) %>')){}else{return false}"
                        style="width: 150px;" value="<% Sendb(Copient.PhraseLib.Lookup("term.removemanually", LanguageID)) %>" <% sendb(IIf(bDisabled , "disabled='disabled'", "")) %> <% sendb(IIf(Not isPGEditable , "disabled='disabled'", "")) %>/></div>
        <br />
        <% End If%>
        <% End If%>
      </div>
      <hr class="hidden" />
    </div>
  </div>
  <% End If%>
   <% If Not IsPostBack then                 
            GetHierarchyHTML()
            End If %>
  <% If (ProductGroupID > 0) AndAlso (ProductGroupTypeID = 2) AndAlso (AttributePGEnabled) Then%>
  <div class="columnfull"> 
   <div class="box" id="attributepgbuilder" style="float:left;height:auto;width:98%;min-height:250px;">
         <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.attributepgbuilder", LanguageID))%>
        </span>
      </h2>
      <br class="half" />
      <uc1:ProductAttributeFilter ID="ucProductAttributeFilter" runat="server" AppName="pgroup-edit.aspx" />
      <hr class="hidden" />
        
 <%--   </div>  --%>  
    
        <%--<div class="box" id="producthierarchy1" style="width: auto; height: 550px;">
    
           <h2>
            <span>
              <% Sendb(Copient.PhraseLib.Lookup("term.includedproducts", LanguageID)) %>
            </span>
          </h2>--%>

       <%-- <div id="divHierarchyContent" runat="server" style="float:left;position:relative;width: 100%; height: 93%;">
        </div>--%>
        <%--<div style="float:left;position:relative;" id="throbber">
            <span id="pcount" style="display:none">
            <label id="warning" style="display:none">Warning:</label>
            Contains &nbsp;
            <label id="lblProductsCount" style="" >
            </label>         
             <img id="Img1" src="../../images/loader.gif" style="display:none"  alt="hello" height="10px" width="40px" />
            &nbsp;
            Products
            </span>
            </div>--%>
    </div>   
  </div>        
  
  <% End If%>
  <div class="column1">
    <% If (MyCommon.IsEngineInstalled(2) And Not CreatedFromOffer) Then%>
    <div class="box" id="CPEvalidation" <% if(productgroupid=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <%
            Dim dtEngine As DataTable
            Dim sEngine As String = ""
            MyCommon.QueryStr = "select Description, PhraseID from PromoEngines where EngineID=2;"
            dtEngine = MyCommon.LRT_Select()
            If dtEngine.Rows.Count > 0 Then
              sEngine = " (" & IIf(MyCommon.NZ(dtEngine.Rows(0).Item("PhraseID"), 0) > 0, Copient.PhraseLib.Lookup(dtEngine.Rows(0).Item("PhraseID"), LanguageID), Trim(MyCommon.NZ(dtEngine.Rows(0).Item("Description"), ""))) & ")"
            End If
            Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID) & sEngine)
          %>
        </span>
      </h2>
      <%
        Dim dtValid As DataTable
        Dim rowOK(), rowWatches(), rowWarnings() As DataRow
        Dim objTemp As Object
        Dim GraceHours As Integer
        Dim GraceCount As Double
        objTemp = MyCommon.Fetch_CPE_SystemOption(41)
        If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
          GraceHours = 4
        End If
        objTemp = MyCommon.Fetch_CPE_SystemOption(42)
        If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
          GraceCount = 0.1D
        End If
        MyCommon.QueryStr = "dbo.pa_ValidationReport_ProdGroup"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.Int).Value = ProductGroupID
        MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
        MyCommon.LRTsp.Parameters.Add("@GraceCount", SqlDbType.Decimal, 2).Value = GraceCount
        dtValid = MyCommon.LRTsp_select()

        rowOK = dtValid.Select("Status=0", "LocationName")
        rowWatches = dtValid.Select("Status=1", "LocationName")
        rowWarnings = dtValid.Select("Status=2", "LocationName")

        Send("<a id=""validLinkCPE"" href=""javascript:openPopup('validation-report.aspx?type=pg&id=" & ProductGroupID & "&level=0');"">")
        Send(Copient.PhraseLib.Lookup("term.valid", LanguageID) & " " & Copient.PhraseLib.Lookup("term.locations", LanguageID) & " (" & rowOK.Length & ")</a><br />")
        Send("<a id=""watchLinkCPE"" href=""javascript:openPopup('validation-report.aspx?type=pg&id=" & ProductGroupID & "&level=1');"">")
        Send(Copient.PhraseLib.Lookup("term.watch", LanguageID) & " " & Copient.PhraseLib.Lookup("term.locations", LanguageID) & " (" & rowWatches.Length & ")</a><br />")
        Send("<a id=""warningLinkCPE"" href=""javascript:openPopup('validation-report.aspx?type=pg&id=" & ProductGroupID & "&level=2');"">")
        Send(Copient.PhraseLib.Lookup("term.warning", LanguageID) & " " & Copient.PhraseLib.Lookup("term.locations", LanguageID) & " (" & rowWarnings.Length & ")</a><br />")
      %>
      <hr class="hidden" />
    </div>
    <% End If%>
    <% If ((MyCommon.IsEngineInstalled(0) Or MyCommon.IsEngineInstalled(1)) And Not CreatedFromOffer) Then%>
    <div class="box" id="CMvalidation" <% if(productgroupid=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <%
            Dim dtEngine As DataTable
            Dim sEngine As String = ""
            Dim drEngine As DataRow
            MyCommon.QueryStr = "select Description, PhraseID from PromoEngines where EngineID in (0,1) and Installed='True' order by EngineID;"
            dtEngine = MyCommon.LRT_Select()
            If dtEngine.Rows.Count > 0 Then
              For Each drEngine In dtEngine.Rows
                sEngine += " (" & IIf(MyCommon.NZ(drEngine.Item("PhraseID"), 0) > 0, Copient.PhraseLib.Lookup(drEngine.Item("PhraseID"), LanguageID), Trim(MyCommon.NZ(drEngine.Item("Description"), ""))) & ")"
              Next
            End If
            Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID) & sEngine)
          %>
        </span>
      </h2>
      <%
        Dim dtValid As DataTable
        Dim rowOK(), rowWaiting(), rowWatches(), rowWarnings() As DataRow
        Dim objTemp As Object
        Dim GraceHours As Integer
        Dim GraceHoursWarn As Integer
        Dim iGroupLocations As Integer

        objTemp = MyCommon.Fetch_CM_SystemOption(10)
        If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
          GraceHours = 4
        End If

        objTemp = MyCommon.Fetch_CM_SystemOption(11)
        If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
          GraceHoursWarn = 24
        End If

        MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_ProdGroup"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.Int).Value = ProductGroupID
        MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
        MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn
        dtValid = MyCommon.LRTsp_select()
        iGroupLocations = dtValid.Rows.Count

        rowOK = dtValid.Select("Status=0", "LocationName")
        rowWaiting = dtValid.Select("Status=1", "LocationName")
        rowWatches = dtValid.Select("Status=2", "LocationName")
        rowWarnings = dtValid.Select("Status=3", "LocationName")

        Send("<a id=""validLinkCM"" href=""javascript:openPopup('CM-validation-report.aspx?type=pg&id=" & ProductGroupID & "&level=0&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iGroupLocations & ")</a><br />")
        Send("<a id=""waitingLinkCM"" href=""javascript:openPopup('CM-validation-report.aspx?type=pg&id=" & ProductGroupID & "&level=1&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.waitlocations", LanguageID) & " (" & rowWaiting.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iGroupLocations & ")</a><br />")
        Send("<a id=""watchLinkCM"" href=""javascript:openPopup('CM-validation-report.aspx?type=pg&id=" & ProductGroupID & "&level=2&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iGroupLocations & ")</a><br />")
        Send("<a id=""warningLinkCM"" href=""javascript:openPopup('CM-validation-report.aspx?type=pg&id=" & ProductGroupID & "&level=3&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iGroupLocations & ")</a><br />")
      %>
      <hr class="hidden" />
    </div>
    <% End If%>
    <div class="box" id="lastuploadattempt" <% if(productgroupid=0)then send(" style=""visibility: hidden;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("pgroup-edit-lastuploaded", LanguageID))%>
        </span>
      </h2>
      <br class="half" />
      <%
        ' last update date
        Sendb(Copient.PhraseLib.Lookup("term.lastupload", LanguageID) & ": ")
        If (LastUpload Is Nothing) OrElse (LastUpload = "1/1/1900") Then
          If (ProductGroupID <> 0) Then
            Sendb(Copient.PhraseLib.Lookup("term.neveruploaded", LanguageID))
            Send("<br />")
          End If
        Else
          longDate = MyCommon.NZ(LastUpload, "1/1/1900")
          longDateString = Logix.ToLongDateTimeString(longDate, MyCommon)
          Send(longDateString)
          Send("<br />")
        End If

        ' last update message
        If LastUploadMsg IsNot Nothing AndAlso LastUploadMsg.Trim <> "" Then
          Sendb(Copient.PhraseLib.Lookup("term.statusmessage", LanguageID) & ": ")
          If LastUploadMsg.ToLower.IndexOf("upload processing completed") > -1 OrElse LastUploadMsg.ToLower.IndexOf("upload processing of item data completed") > -1 Then
            Sendb(Copient.PhraseLib.Lookup("term.successful", LanguageID))
          Else
            Sendb("<br /><p style=""margin:2px 10px;font-family:courier;font-size:11px;"">" & LastUploadMsg.Trim & "</p>")
          End If
        End If

      %>
      <hr class="hidden" />
    </div>
  </div>
  <div style="height: 25px; float: left; width: 10px;">
  </div>
  <div class="column2">
    <% If Not CreatedFromOffer Then%>
    <div class="box" id="offers" <% If(ProductGroupID=0 OrElse IsSpecialGroup)Then Send(" style=""display:none;""") %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <%
			
          If (associatedOfferDT IsNot Nothing) Then
            rowCount = associatedOfferDT.Rows.Count
            Dim assocName As String = ""
            If rowCount > 0 Then
                For Each row As DataRow In associatedOfferDT.Rows
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

                If (MyCommon.NZ(row.Item("ProdEndDate"), Today) < Today) Then
                  Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                End If
                Send("<br />")
              Next
            Else
              Send("     " & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>
    <% End If%>
    <%
      If (ProductGroupID > 0 And Not CreatedFromOffer) Then
        MyCommon.QueryStr = "select CMOADeployStatus,CMOADeployRpt,CMOARptDate,CMOADeploySuccessDate from ProductGroups with (nolock) where ProductGroupID=" & ProductGroupID & " and Deleted=0"
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("CMOARptDate").ToString, "") = "" AndAlso MyCommon.NZ(row.Item("CMOADeployRpt").ToString, "") = "") Then
              GoTo nodeployment
            End If
          Next
          'Send("<div class=""column"">")
          Sendb("<div class=""box"" id=""deployment"">")
          Send("  <h2>")
          Send("    <span>" & Copient.PhraseLib.Lookup("term.deployment", LanguageID) & "</span>")
          Send("  </h2>")
          For Each row In rst.Rows
          Next
          Send("    <h3>" & Copient.PhraseLib.Lookup("term.lastattempted", LanguageID) & ":</h3>")
          deployDate = MyCommon.NZ(row.Item("CMOARptDate"), "")
          If deployDate = "" Then
            Send(Copient.PhraseLib.Lookup("term.never", LanguageID) & "<br />")
          Else
            Send(deployDate & "<br />")
          End If
          Send("    <br class=""half"" />")
          Send("    <h3>" & Copient.PhraseLib.Lookup("term.lastsuccessful", LanguageID) & ":</h3>")
          deployDate = MyCommon.NZ(row.Item("CMOADeploySuccessDate"), "")
          If deployDate = "" Then
            Send(Copient.PhraseLib.Lookup("term.never", LanguageID) & "<br />")
          Else
            Send(deployDate & "<br />")
          End If
          Send("    <br class=""half"" />")
          Send("    <h3>" & Copient.PhraseLib.Lookup("term.status", LanguageID) & ":</h3>")
          Send(MyCommon.NZ(row.Item("CMOADeployRpt"), "") & "<br />")
          Send("<hr class=""hidden"" />")
          Send("</div>")
          'Send("</div>")
nodeployment:
        End If
      End If
    %>
  </div>
</div>

</form>
<%If (ProductGroupID > 0) AndAlso (ProductGroupTypeID = 1 OrElse (ProductGroupTypeID = 2 AndAlso AttributePGEnabled = False)) Then%>
<div id="uploader" style="display: none;">
  <div id="uploadwrap">
    <div class="box" id="uploadbox">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.upload", LanguageID))%>
        </span>
      </h2>
      <form action="pgroup-edit.aspx<%Sendb(If(CreatedFromOffer, "?OfferID=" & OfferID & "&EngineID=" & EngineID & "&slct=" & GetCgiValue("slct") & "&ex=" & GetCgiValue("ex"), "")) %>"
      id="uploadform" name="uploadform" onsubmit="return isValidPath();" method="post"
      enctype="multipart/form-data">
      <%
        Sendb("<input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" ")
        Send("onclick=""javascript:document.getElementById('uploader').style.display='none';"" />")
        Sendb(Copient.PhraseLib.Lookup("pgroup-edit.upload", LanguageID))
        Send("<br /><br />")
        Sendb("<input type=""radio"" name=""operation"" id=""operation1"" value=""0"" checked=""checked"" />")
        Send("<label for=""operation1"">" & Copient.PhraseLib.Lookup("term.FullReplace", LanguageID) & "</label>&nbsp;&nbsp;")
        Sendb("<input type=""radio"" name=""operation"" id=""operation2"" value=""1""  />")
        Send("<label for=""operation2"">" & Copient.PhraseLib.Lookup("term.AddToGroup", LanguageID) & "</label>&nbsp;&nbsp;")
        Sendb("<input type=""radio"" name=""operation"" id=""operation3"" value=""2""  />")
        Send("<label for=""operation3"">" & Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID) & "</label>")
        Send("<br />")
      %>
      <br />
      <br class="half" />
      <%
        If (Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) Then
          Send("     <input type=""hidden"" id=""ProductGroupID2"" name=""ProductGroupID"" value=""" & ProductGroupID & """ />")
          Send("     <input type=""file"" id=""browse"" name=""browse"" value=""" & Copient.PhraseLib.Lookup("term.browse", LanguageID) & """ />")
        '         Send("<div id=""divfile"" style=""height:0px;overflow:hidden"">")
        'Send("<input type=""file"" id=""browse"" name=""fileInput"" onchange=""fileonclick()"" />")
        'Send("</div>")
        'Send("<button type=""button"" onclick=""chooseFile();"">"&Copient.PhraseLib.Lookup("term.browse", LanguageID) &"</button>")
        'Send("<label id=""lblfileupload"" name=""lblfileupload"">"&Copient.PhraseLib.Lookup("term.nofilesselected", LanguageID)&"</label>")
          Send("     <input type=""submit"" class=""regular"" id=""uploadfile"" name=""uploadfile"" value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """ />")
          Send("     <br />")
        End If
      %>
      </form>
      <hr class="hidden" />
    </div>
  </div>
  <%
    If Request.Browser.Type = "IE6" Then
      Send("<iframe src=""javascript:'';"" id=""uploadiframe-pg"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no""></iframe>")
    End If
  %>
</div>
<div id="modifypg" style="display: none;">
  <div id="modifypgwrap" style="float: center;">
    <div class="box" id="modifypgbox">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.modifygroup", LanguageID))%>
        </span>
      </h2>
      <%
        Sendb("<input type=""button"" class=""ex"" id=""modifypgclose"" name=""modifypgclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" ")
        Send("onclick=""javascript:closeModifyGroup();"" />")
        Sendb("<span id=""modifypginfobar"" class=""red-background"" style=""color:#ffffff;font-weight: bold;""> </span>")
        Sendb("<br />")
        Sendb("<br />")
        Sendb("<label for=""operation2"">Product Type : </label>")
        Send("<select id=""modifyproducttype"" name=""modifyproducttype"">")

        'BZ2079: UE-feature-removal #: Remove unsupported product types for UE (Mix/Match Code, Manufacturer Family code, Pool Code)
        '        To restore previous functionality: remove the all code in the If statement checking engines except the query without a where clause.
        If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then
          MyCommon.QueryStr = "select ProductTypeID,PhraseID from ProductTypes with (NoLock)"
        Else
          MyCommon.QueryStr = "select ProductTypeID,PhraseID from ProductTypes with (NoLock) where ProductTypeID not in (3,4,5);"
        End If
        rst2 = MyCommon.LRT_Select
        For Each row2 As DataRow In rst2.Rows
          ProductTypeID = row2.Item("ProductTypeID")
          If (Not CpeEngineOnly Or (CpeEngineOnly And ProductTypeID <= 2)) Then
            Send("<option value=""" & row2.Item("ProductTypeID") & """>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</option>")
          End If
        Next

        Sendb("</select>")
        Sendb("<br /><br />")
        Sendb("<label for=""operation3"">Paste product list : </label> ")
        Send("<br />")
        Sendb("<textarea name=""pasteproducts"" id=""pasteproducts"" style=""width: 290px; height: 150px"">")
        Sendb("</textarea>")
        Send("<br />")
        Send("<br />")

        Sendb("<input type=""radio"" name=""modifyoperation"" value=""0"" checked=""checked"" />")
        Send("<label for=""operation4"">" & Copient.PhraseLib.Lookup("term.FullReplace", LanguageID) & "</label>&nbsp;&nbsp;")
        Sendb("<input type=""radio"" name=""modifyoperation""  value=""1""  />")
        Send("<label for=""operation5"">" & Copient.PhraseLib.Lookup("term.AddToGroup", LanguageID) & "</label>&nbsp;&nbsp;")
        Sendb("<input type=""radio"" name=""modifyoperation"" value=""2""  />")
        Send("<label for=""operation6"">" & Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID) & "</label>")
        Send("<br />")
      %>
      <br />
      <br class="half" />
      <%
        If (Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer) Then
          Send("     <input type=""hidden"" id=""ModifyProductGroupID"" name=""ModifyProductGroupID"" value=""" & ProductGroupID & """ />")
          Send("     <input type=""button"" class=""regular"" id=""modifyprodgroup"" name=""modifyprodgroup"" onclick=""javascript:ModifyGroup();"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
          Send("     <br />")
        End If
      %>
      <hr class="hidden" />
    </div>
  </div>
  <%
    If Request.Browser.Type = "IE6" Then
      Send("<iframe src=""javascript:'';"" id=""modifyframe-pg"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no""></iframe>")
    End If
  %>
 
</div>
<%End If%>
<script type="text/javascript">
  //adjustItemsSelectBox()
</script>
<script type="text/javascript">
  if (window.captureEvents) {
    window.captureEvents(Event.CLICK);
    window.onclick = handlePageClick;
  }
  else {
    document.onclick = handlePageClick;
  }
</script>
<script runat="server">
  Dim CopientFileName As String
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""

  Dim AdminUserID As Long
    Dim ProductGroupID As Long = -1
  Dim GName As String
  Dim CreatedDate As String
  Dim LastUpdate As String
  Dim LastUpload As String = Nothing
  Dim LastUploadMsg As String = ""
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
     Dim Hierarchy As New Copient.HierarchyResync(MyCommon, "UI", "HierarchyFeeds.txt")
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim rstItems As DataTable = Nothing
  Dim rstProdTypes As DataTable = Nothing
  Dim upc As String
  Dim GroupSize As Integer
  Dim outputStatus As Integer
  Dim DefaultIDType As Integer
  Dim File As HttpPostedFile
  Dim InstallPath As String
  Dim rowCount As Integer
  Dim ProdAvailableCount As Integer
  Dim squery As String
  Dim dtProdAvailable As DataTable
  Dim ProductList As String
  Dim Products() As String
  Dim bAdd As Boolean
  Dim bAddAll As Boolean
  Dim bRemove As Boolean
  Dim typeST As DataTable
  Dim iType As Integer
  Dim deployDate As String
  Dim longDate As New DateTime
  Dim longDateString As String
  Dim statusMessage As String = ""
  Dim ExtProductID As String = ""
  Dim IDLength As Integer = 0
  Dim GNameTitle As String = ""
  Dim GFullName As String = ""
  Dim XID As String = ""
  Dim ShowActionButton As Boolean = False
  Dim ListBoxSize, LinkSize, LabelSize As Integer
    Dim ShowAllItems As Boolean
    Dim ProductHierarchyExist As Boolean=false
  Dim OfferCtr As Integer = 0
  Dim IE6ScrollFix As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim IsSpecialGroup As Boolean = False
  Dim CanEditSpecialGroup As Boolean = False
  Dim OfferID As Integer = 0
  Dim EngineID As Integer = -1
  Dim CreatedFromOffer As Boolean = False
  Dim UploadOperation As Integer = 0
  Dim prodDT As DataTable
  Dim Description As String = ""
  Dim HasExcludedNodes, HasExcludedItems As Boolean
  Dim bMFCenabled As Boolean = MyCommon.NZ(MyCommon.Fetch_SystemOption(297),0)
  Dim bDownloadPG As Boolean = IIF(MyCommon.NZ(MyCommon.Fetch_SystemOption(302), 0) = "1", 1, 0)
  ' Hierarchy stuff
  Dim ParentNodeIdList As String
  Dim SelectedNodeId As String
  Dim NodeName As String
  Dim bUp As Boolean
  Dim bDown As Boolean
  Dim NodeIds() As String
  Dim ParentId As String
  Dim dtParents As DataTable = Nothing
  Dim dtParents1 As DataTable = Nothing
  Dim dtChildren As DataTable = Nothing
  Dim i As Integer
  Dim CpeEngineOnly As Boolean = False
  Dim ProductTypeID As Integer = 0
  Dim ItemPKID As Integer = -1
  Dim ShowViewSelected As Boolean = False
  Dim iProductIdNumericOnly As Integer = 97
  Dim NewProductGroupID As Long = 0
  Dim HTMLBuf As New StringBuilder()
  Dim descriptionItem As String = String.Empty
  Dim AttributePGEnabled As Boolean = False
  Dim ProductGroupTypeID As Byte = 1
  Dim m_ProductGroupService As IProductGroupService
  Dim BuyerID As Int32 = -1
  Dim ExternalBuyerId As String = ""
  Dim ProductsWithoutDesc As Integer
  Dim NodeID As String=""
  Dim LinkedItems As String=""
  Dim PABStage As Int16 = 1
  Dim bStoreUser As Boolean = False
  Dim sValidLocIDs As String = ""
  Dim sValidSU As String = ""
  Dim wherestr As String = "" 
  Dim sJoin As String = "" 
  Dim iLen As Integer = 0
  Dim associatedOfferDT As DataTable =Nothing
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
  Dim IsTranslatedPG As Boolean = False
  Dim bStaticPG As Boolean = False
  Dim bStaticEnabled As Boolean = False
  Dim bCreateStaticPG As Boolean = False
  Dim bEditStaticPG As Boolean = False
    
    Property EditProductRegardlessOfBuyer As Boolean
        Get
            If MyCommon.IsEngineInstalled(9) Then
                Return (Logix.UserRoles.EditProductgroupsRegardlessBuyer Or MyCommon.IsProductCreatedWithUserAssociated(AdminUserID, ProductGroupID))
            Else
                Return True
            End If
        End Get
        Set(value As Boolean)

        End Set
    End Property
    Dim AttributeUpdatePending As Boolean = False
   
    Private Function GetHierarchyHTML() As String
        Dim writer As New StringWriter
        'Dim locateHierarchyURL As String = ""
        'If Not String.IsNullOrWhiteSpace(Request.QueryString("LocateHierarchyURL")) Then
        '    locateHierarchyURL = HttpUtility.UrlDecode(Request.QueryString("LocateHierarchyURL"))
        'End If
        If (Not IsPostBack Or GetCgiValue("hdnProductHierarchyExist") = "true") Then
            Dim bttbl As DataTable = New DataTable()
          
            If ProductGroupID > 0 Then
                MyCommon.QueryStr = "select Buyerid from productgroups where ProductGroupID=" & ProductGroupID
                bttbl = MyCommon.LRT_Select()
                If bttbl.Rows.Count > 0 Then
                    If bttbl.Rows.Count > 0 Then
                        BuyerID = IIf(IsDBNull(bttbl.Rows(0).Item("BuyerID")), 0, bttbl.Rows(0).Item("BuyerID"))
                    End If
                End If
            End If
            
            ucProductAttributeFilter.BuyerID = BuyerID
            ucProductAttributeFilter.HierarchyTreeURL = "/logix/phierarchytree.aspx?ProductGroupID=" & ProductGroupID & "&PAB=1&PopupFlag=0" '+ locateHierarchyURL
            ucProductAttributeFilter.ProductGroupID = ProductGroupID
        End If
        
      
       
     
        Return writer.ToString
    End Function
    <WebMethod()>
    Public Shared Function LoadLevelsGV(PageIndex As Int16, PageSize As Int32, _sortBy As String, _sortOrder As String, ProductGroupID As Int32, strExcludedProductIDs As String, strExcludedLevels As String) As String
        CurrentRequest.Resolver.AppName = "pgroup-edit.aspx"
        Dim m_Product As IProductService = CurrentRequest.Resolver.Resolve(Of IProductService)()
        Dim amsResult As AMSResult(Of DataSet)
        Dim AttributeTypeToDisplay As List(Of AttributeType)
        Dim TotalPages As Int32
        Dim IsFilterUpdated As Boolean
        Dim TotalProductCount As Long, TotalLevelCount As Long
        Dim AttributeValues As DataTable = New DataTable()
        Dim AllChildNodes As DataTable = New DataTable()
        Dim ds As DataSet = New DataSet("Table")
        Dim LevelDt As DataTable
        Dim CountDt As DataTable = New DataTable("TotalPages")
        CountDt.Columns.Add("TotalPages", Type.GetType("System.String"))
        Try
            AttributeValues = HttpContext.Current.Session("AttributeValues")
            AttributeTypeToDisplay = HttpContext.Current.Session("AttributeTypeToDisplay")
            AllChildNodes = HttpContext.Current.Session("AllChildNodes")
            IsFilterUpdated = HttpContext.Current.Session("IsLevelFilterUpdated")
            _sortBy = IIf(_sortBy = "", Nothing, _sortBy)
            _sortOrder = IIf(_sortOrder = "", Nothing, _sortOrder)
            Dim ExcludeLeveldt As DataTable = New DataTable()
            Dim dc As DataColumn = New DataColumn("DisplayLevel", GetType(String))
            ExcludeLeveldt.Columns.Add(dc)
            Dim exclList As List(Of String)
            If (IsFilterUpdated) Then
                exclList = strExcludedLevels.Split(",").ToList()
                exclList.RemoveAll(Function(item) String.IsNullOrEmpty(item))
                exclList.ForEach(Function(r) ExcludeLeveldt.Rows.Add(r))
            End If
            Dim sortOrder = "ASC", sortBy = "DisplayLevel"
            If _sortBy = sortBy Then
                sortOrder = _sortOrder
            End If
            amsResult = m_Product.GetLevelGroups(PageIndex, PageSize, sortBy, sortOrder, AttributeValues, AttributeTypeToDisplay, ProductGroupID,
                                                 strExcludedProductIDs, IsFilterUpdated, AllChildNodes, ExcludeLeveldt, TotalLevelCount, TotalProductCount)
            LevelDt = amsResult.Result.Tables(0).Copy()
            LevelDt.TableName = "Level"
            If (TotalLevelCount > 0) Then
                TotalPages = (TotalLevelCount / PageSize) + IIf(((TotalLevelCount Mod PageSize) > 0), 1, 0)
            End If
            CountDt.Rows.Add(TotalPages)
            ds.Tables.AddRange({LevelDt, CountDt})
        Catch ex As ApplicationException
        End Try
        Return JsonConvert.SerializeObject(ds)
    End Function

    <WebMethod()>
    Public Shared Function LoadProductsByLevel(Level As String, PageIndex As Int16, PageSize As Int32, _sortBy As String, _sortOrder As String, ProductGroupID As Int32, strExcludedProductIDs As String) As String
        CurrentRequest.Resolver.AppName = "pgroup-edit.aspx"
        Dim m_Product As IProductService = CurrentRequest.Resolver.Resolve(Of IProductService)()
        Dim amsResult As AMSResult(Of DataTable)
        Dim AttributeTypeToDisplay As List(Of AttributeType)
        Dim TotalPages As Int32
        Dim TotalCount As Int32
        Dim IsFilterUpdated As Boolean
        Dim AttributeValues As DataTable = New DataTable()
        Dim AllChildNodes As DataTable = New DataTable()
        Dim ds As DataSet = New DataSet("Table")
        Dim LevelDt As DataTable
        Dim CountDt As DataTable = New DataTable("TotalPages")
        CountDt.Columns.Add("TotalPages", Type.GetType("System.String"))
        Try
            AttributeValues = HttpContext.Current.Session("AttributeValues")
            AttributeTypeToDisplay = HttpContext.Current.Session("AttributeTypeToDisplay")
            AllChildNodes = HttpContext.Current.Session("AllChildNodes")
            IsFilterUpdated = HttpContext.Current.Session("IsLevelFilterUpdated")
            If (Not IsFilterUpdated) Then
                strExcludedProductIDs = ""
            End If
            amsResult = m_Product.GetProductsByLevel(PageIndex, PageSize, _sortBy, _sortOrder, AttributeValues, AttributeTypeToDisplay, ProductGroupID, strExcludedProductIDs, IsFilterUpdated, AllChildNodes, Level, TotalCount)
            LevelDt = amsResult.Result.Copy()
            LevelDt.TableName = "Product"
            If (TotalCount > 0) Then
                TotalPages = (TotalCount / PageSize) + IIf(((TotalCount Mod PageSize) > 0), 1, 0)
            End If
            CountDt.Rows.Add(TotalPages)
            ds.Tables.AddRange({LevelDt, CountDt})
        Catch ex As ApplicationException
        End Try
        Return JsonConvert.SerializeObject(ds, New JsonSerializerSettings() With {.ContractResolver = New Serialization.DefaultContractResolver()})
    End Function
  
    <WebMethod()>
    Public Shared Function GetProductCount(FetchProductCountInNodesFlag As String, AVPairs As String) As String
        CurrentRequest.Resolver.AppName = "pgroup-edit.aspx"
        Dim attService As IAttributeService = CurrentRequest.Resolver.Resolve(Of IAttributeService)()
        Dim amsResult As AMSResult(Of Int32) = Nothing
        
        Dim nodeIdsDT As DataTable = HttpContext.Current.Session("AllChildNodes")
        If HttpContext.Current.Session("AllChildNodes") IsNot Nothing Then
            nodeIdsDT = CType(HttpContext.Current.Session("AllChildNodes"), DataTable)
        End If
        
        If JSONHelper.ToObject(Of String)(FetchProductCountInNodesFlag) = "1" Then
            amsResult = attService.GetCountOfProductsInNode(nodeIdsDT)
        Else
            amsResult = attService.GetLatestCountOfProducts(nodeIdsDT, PrepareDTFromList(JSONHelper.ToObject(Of List(Of Dictionary(Of String, Object)))(AVPairs)))
        End If
        Return JSONHelper.ToJSON(amsResult.Result)
    End Function
    Private Shared Function PrepareDTFromList(avPairs As List(Of Dictionary(Of String, Object))) As DataTable
        Dim dt As DataTable = New DataTable()
        For Each dict As Dictionary(Of String, Object) In avPairs
            If dt.Columns.Count = 0 Then
                For Each col As String In dict.Keys
                    dt.Columns.Add(col)
                Next
            End If
            Dim dr As DataRow = dt.NewRow()
            For Each item As KeyValuePair(Of String, Object) In dict
                dr(item.Key) = item.Value
            Next
            dt.Rows.Add(dr)
        Next
        Return dt
    End Function

    'Function to be called asynchronously by javascript for fetching paged attribute values
    <WebMethod>
    <WebInvoke(Method:="POST")>
    Public Shared Function GetAttributes(term As String, attributetype As Int16, pageindex As Int32, excludeattr As String, keyValue As String) As List(Of Attributes)
        CurrentRequest.Resolver.AppName = "pgroup-edit.aspx"
        Dim attributeService As IAttributeService = CurrentRequest.Resolver.Resolve(Of IAttributeService)()
        Dim nodeIdsDT As DataTable = HttpContext.Current.Session("AllChildNodes")
        If HttpContext.Current.Session("AllChildNodes") IsNot Nothing
            nodeIdsDT = CType(HttpContext.Current.Session("AllChildNodes"), DataTable)
        End If
        Dim amsResultAttributes As AMSResult(Of List(Of Attributes)) = attributeService.GetAttributesInChunks(term, attributetype, pageindex, nodeIdsDT, excludeattr, keyValue)
        If amsResultAttributes.ResultType <> AMSResultType.Success AndAlso amsResultAttributes.MessageString <> String.Empty Then
            Dim activityFields As ActivityLogFields = New ActivityLogFields
            Dim myCommon As New Copient.CommonInc
            myCommon.Open_LogixRT
            activityFields.LinkID = attributetype
            activityFields.Description = amsResultAttributes.MessageString
            myCommon.Activity_Log3(activityFields)
            myCommon.Close_LogixRT
        End If
        Return amsResultAttributes.Result
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
       
        If Not String.IsNullOrWhiteSpace(Request.QueryString("PABStage")) Then
            PABStage = Convert.ToInt16(Request.QueryString("PABStage"))
        End If
        If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
            Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
        End If
        CopientFileName = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))

        Response.Expires = 0
        MyCommon.AppName = "pgroup-edit.aspx"
        CurrentRequest.Resolver.AppName = MyCommon.AppName
        m_ProductGroupService = CurrentRequest.Resolver.Resolve(Of IProductGroupService)()
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)
        NodeID = GetCgiValue("NodeListID")
        'Store User
        If (MyCommon.Fetch_CM_SystemOption(131) = "1") Then
            MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
            rst = MyCommon.LRT_Select
            iLen = rst.Rows.Count
            If iLen > 0 Then
                bStoreUser = True
                sValidSU = AdminUserID
                For i = 0 To (iLen - 1)
                    If i = 0 Then
                        sValidLocIDs = rst.Rows(0).Item("LocationID")
                    Else
                        sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
                    End If
                Next
      
                MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
                rst = MyCommon.LRT_Select
                iLen = rst.Rows.Count
                If iLen > 0 Then
                    For i = 0 To (iLen - 1)
                        sValidSU &= "," & rst.Rows(i).Item("UserID")
                    Next
                End If
            End If
        End If
    
        DefaultIDType = MyCommon.Fetch_SystemOption(30)
        CanEditSpecialGroup = Logix.UserRoles.EditSystemConfiguration

        bStaticEnabled = MyCommon.Fetch_SystemOption(280)
        If bStaticEnabled Then
          bCreateStaticPG = Logix.UserRoles.CreateStaticProductGroups
          bEditStaticPG = Logix.UserRoles.EditStaticProductGroups
        End If
           
        If Request.RequestType = "GET" Then
            ' fill in if it was a get method
            ProductGroupID = MyCommon.Extract_Val(GetCgiValue("ProductGroupID"))
            GName = Logix.TrimAll(GetCgiValue("GroupName"))
            ParentNodeIdList = GetCgiValue("ParentNodeIdList")
            SelectedNodeId = GetCgiValue("SelectedNodeId")
            NodeName = GetCgiValue("NodeName")
            UploadOperation = MyCommon.Extract_Val(GetCgiValue("Operation"))
            ' check in case it was a POST instead of get
        Else
            ParentNodeIdList = Request.Form("ParentNodeIdList")
            SelectedNodeId = Request.Form("SelectedNodeId")
            NodeName = Request.Form("NodeName")
            GName = Logix.TrimAll(Request.Form("GroupName"))
            BuyerID = Request.Form("buyerID")
            'If buyer option doesn't exist in PG creation page set it to -1
            BuyerID = If(BuyerID = 0, -1, BuyerID)
            UploadOperation = MyCommon.Extract_Val(Request.Form("Operation"))
            ProductGroupID = Request.Form("ProductGroupID")
            If ProductGroupID = 0 Then
                ProductGroupID = MyCommon.Extract_Val(GetCgiValue("ProductGroupID"))
            End If
        End If
        If (ProductGroupID > 0) Then
            'Dim rst As DataTable = Hierarchy.GetNodesLinkedToProductGroupID(ProductGroupID)
            'LinkedItems = ""
            'If (rst.Rows.Count > 0) Then
            '    For i As Integer = 0 To rst.Rows.Count - 1
            '        LinkedItems = LinkedItems & rst.Rows(i)("ExtNodeID").ToString()
            '        LinkedItems = LinkedItems & If((i < rst.Rows.Count - 1), ",", String.Empty)
            '    Next
            'End If
            BuyerID = MyCommon.Extract_Val(GetCgiValue("hdnBuyerID")).ToString().ConvertToInt32()
            ProductGroupTypeID = MyCommon.Extract_Val(GetCgiValue("hdnProductGroupTypeID")).ToString().ConvertToByte()
            If ProductGroupTypeID = 0 Then
                ProductGroupTypeID = 1
            End If
        End If
      

        AttributePGEnabled = (MyCommon.Fetch_UE_SystemOption(157) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE))
        If (AttributePGEnabled AndAlso ProductGroupID = 0) Then
            If Not Page.IsPostBack Then
                Dim ProductGroupTypes As AMSResult(Of List(Of ProductGroupTypes)) = m_ProductGroupService.GetProductGroupTypes()
                If (ProductGroupTypes.ResultType <> AMSResultType.Success) Then
                    infoMessage = ProductGroupTypes.MessageString
                Else
                    RadioButtonList1.DataSource = ProductGroupTypes.Result
                    RadioButtonList1.DataTextField = "PhraseID"
                    RadioButtonList1.DataValueField = "ProductGroupTypeID"
                    RadioButtonList1.DataBind()
                    If (RadioButtonList1.Items.Count > 0) Then
                        Dim DefaultProductGroupID As String = MyCommon.Fetch_UE_SystemOption(156)
                        Dim radioBtn As ListItem = RadioButtonList1.Items.FindByValue(DefaultProductGroupID)
                        If radioBtn Is Nothing Then
                            RadioButtonList1.Items(0).Selected = True
                        Else
                            radioBtn.Selected = True
                        End If
                    End If
                End If
            Else
                ProductGroupTypeID = RadioButtonList1.SelectedItem.Value.ConvertToByte()
            End If

        End If
        'Status message
        If MyCommon.IsEngineInstalled(2) Then
            MyCommon.QueryStr = "select distinct PG.ProductGroupID " & _
                                "from ProductGroups as PG with (NoLock) INNER JOIN ProductGroupLocUpdate PGLU WITH (NoLock) ON PGLU.ProductGroupID=PG.ProductGroupID " & _
                                "WHERE PG.Deleted = 0 AND PGLU.EngineID=2 AND (PG.CPEStatusFlag=2 OR PGLU.StatusFlag=2) " & _
                                "and PG.ProductGroupID=" & ProductGroupID & ";"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
                statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
            End If
        End If

        If MyCommon.IsEngineInstalled(9) Then
            MyCommon.QueryStr = "select distinct PG.ProductGroupID " & _
                                "from ProductGroups as PG with (NoLock) INNER JOIN UE_ProductGroupLocUpdate PGLU WITH (NoLock) ON PGLU.ProductGroupID=PG.ProductGroupID " & _
                                "WHERE PG.Deleted = 0 AND PGLU.EngineID=9 AND (PG.UEStatusFlag=2 OR PGLU.StatusFlag=2) " & _
                                "and PG.ProductGroupID=" & ProductGroupID & ";"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
                statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
            End If
        End If

        OfferID = MyCommon.Extract_Val(GetCgiValue("OfferID"))
        EngineID = If(GetCgiValue("EngineID") = "", -1, MyCommon.Extract_Val(GetCgiValue("EngineID")))
        CreatedFromOffer = OfferID > 0 And EngineID > 0
        If CreatedFromOffer And ProductGroupID = 0 Then
            GName = Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.group", LanguageID), VbStrConv.Lowercase)
            MyCommon.QueryStr = "select count(*) as GroupCount from ProductGroups where Name like @Name + '%'"
            MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If rst.Rows(0).Item("GroupCount") > 0 Then
                GName = GName & " (" & rst.Rows(0).Item("GroupCount") & ")"
            End If
        End If

        ' Send("DEBUG: SelectedNodeId=" & SelectedNodeId)
        InstallPath = MyCommon.Get_Install_Path(Request.PhysicalPath)

        If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then
            ' The user wants to redeploy, so do a quick check on offerRewards to make sure there are at least some rewards
            MyCommon.QueryStr = "select CMOAStatusFlag from productgroups with (NoLock) where ProductGroupID=" & ProductGroupID
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
                If rst.Rows(0).Item(0) = "2" Then
                    ' only show the awaiting deployment if there are Locations this group will be sent to.
                    MyCommon.QueryStr = "select distinct LocationID from ProductGroupLocUpdate with (NoLock) where EngineID=1 and ProductGroupID = " & ProductGroupID
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
                    End If
                ElseIf rst.Rows(0).Item(0) = "-1" Then
                    infoMessage = Copient.PhraseLib.Lookup("status.warning", LanguageID)
                End If
            End If
        End If

        If GetCgiValue("up") = "" Then
            bUp = False
        Else
            bUp = True
        End If
        If GetCgiValue("down") = "" Then
            bDown = False
        Else
            bDown = True
        End If
        If GetCgiValue("product-add") = "" Then
            bAdd = False
        Else
            bAdd = True
        End If
        If GetCgiValue("product-add-all") = "" Then
            bAddAll = False
        Else
            bAddAll = True
        End If
        If GetCgiValue("product-rem") = "" Then
            bRemove = False
        Else
            bRemove = True
        End If

        If (GetCgiValue("new") <> "") Then
            Response.Redirect("pgroup-edit.aspx")
        End If
        If (GetCgiValue("newstatic") <> "") Then
          bStaticPG = True
          If ProductGroupID > 0 Then
            Response.Redirect("pgroup-edit.aspx?newstatic=1")
          End If
        End If

        MyCommon.QueryStr = "select ProductGroupID from ProductGroups where ProductGroupID=@PGID AND ProductGroupTypeID = 2 AND IsAttributeUpdated = 1"
        MyCommon.DBParameters.Add("@PGID", SqlDbType.BigInt).Value = ProductGroupID
        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        AttributeUpdatePending = (rst.Rows.Count > 0)

        If (GetCgiValue("download") <> "") Then
            If (GetCgiValue("hdnProductHierarchyExist") = "true") Then
                GetHierarchyHTML()
            End If
            If (ProductGroupTypeID = 2 AndAlso AttributeUpdatePending = True) Then
                infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.changespending", LanguageID)
            End If
            If String.IsNullOrWhiteSpace(infoMessage) Then
                ' they want to download the group get it from the database and stream it to the client
                MyCommon.QueryStr = "select GM.ProductID,ExtProductID, ProductTypeID, CID.Description from Products as CID with (NoLock) Inner Join ProdGroupItems as GM with (NoLock) on CID.ProductID=GM.ProductID where GM.ProductGroupID = " & ProductGroupID & " And GM.Deleted = 0"
                rst = MyCommon.LRT_Select()
                If (rst.Rows.Count > 0) Then
                    Response.Clear()
                    Response.AddHeader("Content-Disposition", "attachment; filename=PG" & ProductGroupID & ".txt")
                    Response.ContentType = "application/octet-stream"
                    Response.ContentEncoding = Encoding.GetEncoding(1252)
                    Response.ClearContent()
                    For Each row In rst.Rows
                        Sendb(MyCommon.NZ(row.Item("ExtProductID"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
                        Sendb(",")
                        Sendb(MyCommon.NZ(row.Item("ProductTypeID"), 1))
                        Sendb(",")
                        Send(MyCommon.NZ(row.Item("Description"), ""))
                    Next
                    Response.End()
                    GoTo done
                Else
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.noelements", LanguageID)
                End If
            End If
        End If

        If GetCgiValue("LargeFile") = "true" Then
            infoMessage = Copient.PhraseLib.Lookup("error.UploadTooLarge", LanguageID)
        End If

        If Request.Files.Count >= 1 Then
            File = Request.Files.Get(0)
            If File.ContentType <> "text/plain" Then
                infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badfile", LanguageID)
            Else
                Dim UploadFileName As String
                Dim TimeStampStr As String
                Dim UploadedText As String = Copient.PhraseLib.Lookup("history.pgroup-upload", LanguageID)

                TimeStampStr = MyCommon.Leading_Zero_Fill(Day(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Hour(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Minute(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Second(Date.Now), 2)
                UploadFileName = "P" & ProductGroupID & "-" & TimeStampStr & ".dat"
                File.SaveAs(MyCommon.Fetch_SystemOption(29) & "\" & UploadFileName)
                System.IO.File.ReadAllText(MyCommon.Fetch_SystemOption(29) & "\" & UploadFileName)

                ' add entry into table for agent to pick it up
                ' dbo.pt_GMInsertQueue_Insert @FileName nvarchar(255), @ProductGroupID bigint
                MyCommon.QueryStr = "dbo.pt_ProdInsertQueue_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = UploadFileName
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@OperationType", SqlDbType.Int).Value = UploadOperation
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()

                Select Case UploadOperation
                    Case 1
                        UploadedText &= " (" & Copient.PhraseLib.Lookup("term.addedtogroup", LanguageID) & ")"
                    Case 2
                        UploadedText &= " (" & Copient.PhraseLib.Lookup("term.removefromgroup", LanguageID) & ")"
                    Case Else
                        UploadedText &= " (" & Copient.PhraseLib.Lookup("term.fullreplacement", LanguageID) & ")"
                End Select
                MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, UploadedText)
                'MyCommon.QueryStr = "Insert into GMInsertQueue (FileName,UploadTime,ProductGroupID,StatusFlag) values ('" & File.FileName & "', getdate()," & ProductGroupID & ",0);"
                If Not CreatedFromOffer Then
                    Response.Redirect("pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
                Else
                    Select Case EngineID
                        Case 7
                            Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID & _
                                               "&OfferID=" & OfferID & "&EngineID=" & EngineID & "&slct=" & GetCgiValue("slct") & _
                                               "&ex=" & GetCgiValue("ex") & "&condChanged=" & GetCgiValue("condChanged"))
                        Case Else
                            Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
                    End Select
                End If
            End If
        End If

        ' lets see if they clicked save or delete
        Dim isInvalid As Boolean
        ' lets see if they clicked save or delete
	    If (GetCgiValue("save") <> "") OrElse (CreatedFromOffer AndAlso ProductGroupID = 0) Then
 	 	GName = MyCommon.NZ(Logix.TrimAll(GName), "")
        Dim regex As Regex = New Regex("['\""\\]")
 	 	Dim match As Match = regex.Match(GName)
 	 	isInvalid = match.Success
 	 	If match.Success Then
 	 	    infoMessage = Copient.PhraseLib.Lookup("error.invalid", LanguageID)
 	 	End If
 	 	
 	 	If (ProductGroupID = 0) AndAlso (isInvalid <> True) Then
                MyCommon.QueryStr = "dbo.pt_ProductGroups_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
                MyCommon.LRTsp.Parameters.Add("@AnyProduct", SqlDbType.Bit).Value = 0
                MyCommon.LRTsp.Parameters.Add("@BuyerId", SqlDbType.Int).Value = If(BuyerID = -1, DBNull.Value, BuyerID)
                MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output

                MyCommon.LRTsp.Parameters.Add("@ProductGroupTypeID", SqlDbType.TinyInt).Value = ProductGroupTypeID
                If (GName = "") Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.noname", LanguageID)
                Else
                    MyCommon.QueryStr = "SELECT ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = @Name AND Deleted=0"
                    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If (rst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.nameused", LanguageID)
                    Else
                        MyCommon.LRTsp.ExecuteNonQuery()
                        ProductGroupID = MyCommon.LRTsp.Parameters("@ProductGroupID").Value

                        If bStaticPG Then
                          MyCommon.QueryStr = "update productgroups with (RowLock) set IsStatic=1 where ProductGroupID=" & ProductGroupID & ";"
                          MyCommon.LRT_Execute()
                        End If
                        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-create", LanguageID), BuyerID)

                        'Save Offer Folder selection,
                        Dim cookie As HttpCookie
                        If Request.Cookies("DefaultBuyerForProd") Is Nothing Then
                            cookie = New HttpCookie("DefaultBuyerForProd")
                        Else
                            cookie = HttpContext.Current.Request.Cookies("DefaultBuyerForProd")
                        End If

                        cookie.Value = BuyerID
                        cookie.Expires = DateTime.MaxValue
                        Response.Cookies.Add(cookie)

                    End If
                End If
                MyCommon.Close_LRTsp()
            Else
                MyCommon.QueryStr = "dbo.pt_ProductGroups_Update"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = GName
                GName = MyCommon.Parse_Quotes(Logix.TrimAll(GName))
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                If (GName = "") Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.noname", LanguageID)
                Else
                    MyCommon.QueryStr = "SELECT Name, ProductGroupID FROM ProductGroups with (NoLock) WHERE Name = '" & GName & "' AND Deleted=0 AND ProductGroupID <> " & ProductGroupID
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.nameused", LanguageID)
                    Else
                        MyCommon.LRTsp.ExecuteNonQuery()
                        If GetCgiValue("hdnIsPGNameUpdated") = "true" Then MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-rename", LanguageID))
                        If (ProductGroupTypeID = 2 AndAlso AttributePGEnabled) Then
                            If (NodeID = "") Then
                                NodeID = "N0"
                            End If
                            ucProductAttributeFilter.SelectedNodeIDs = NodeID
                            ucProductAttributeFilter.ProductGroupID = ProductGroupID
                            ucProductAttributeFilter.SaveData = True
                            MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, "updated product group")
                        End If
                        SendNotificationsOfItemChange(ProductGroupID, 2)
                    End If
                End If
                MyCommon.Close_LRTsp()
        
            If bStaticEnabled Then
              MyCommon.QueryStr = "SELECT isnull(IsStatic,0) as IsStatic FROM ProductGroups with (NoLock) WHERE ProductGroupId = " & ProductGroupID & ";"
              rst2 = MyCommon.LRT_Select
              If (rst2.Rows.Count > 0) Then
                bStaticPG = rst2.Rows(0).Item("IsStatic")
                If bStaticPG Then
                  Dim SetFlags As String = ""
                  If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then SetFlags = " CMOAStatusFlag=2"
                  If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
                    If (SetFlags <> "") Then SetFlags = SetFlags & ","
                    SetFlags = SetFlags & " CPEStatusFlag=2"
                  End If
                  If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                    If (SetFlags <> "") Then SetFlags = SetFlags & ","
                    SetFlags = SetFlags & " UEStatusFlag=2"
                  End If
                  MyCommon.QueryStr = "update productgroups with (RowLock) set TCRMAStatusFlag=2," & SetFlags & ", updatelevel=updatelevel+1, LastUpdatedByAdminID=" & AdminUserID & " where ProductGroupID=" & ProductGroupID
                  MyCommon.LRT_Execute()
                  statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
                  MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-redeploy", LanguageID))
                End If
              End If
            End If
        End If
            If infoMessage = "" Then

                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
            End If
        ElseIf (GetCgiValue("delete") <> "") Then
            MyCommon.QueryStr = "select 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate from offerconditions as OC with (NoLock) " & _
                                " left join Offers as O with (NoLock) on O.Offerid=OC.offerID " & _
                                " where (linkid=" & ProductGroupID & " or excludedid=" & ProductGroupID & ") and  ConditionTypeID=2  and " & _
                                " OC.deleted=0 and O.deleted=0 and O.IsTemplate=0 " & _
                                " union " & _
                                " select 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate from offerrewards as OFR with (NoLock) " & _
                                " left join Offers as O with (NoLock) on O.Offerid=OFR.offerID " & _
                                " where (ProductGroupID=" & ProductGroupID & " or ExcludedProdGroupID=" & ProductGroupID & ")   and " & _
                                " OFR.deleted=0 and O.deleted=0 and O.IsTemplate=0 " & _
                                " UNION " & _
                                "SELECT DISTINCT 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate " & _
                                "  FROM CPE_IncentiveProductGroups IPG with (NoLock)" & _
                                " left  join ProductConditionProductGroups P1G on P1G.IncentiveProductGroupID= IPG.IncentiveProductGroupID" & _
                                "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID" & _
                                "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                "  INNER JOIN ProductGroups PG with (NoLock) on IPG.ProductGroupID = PG.ProductGroupID " & _
                                "  WHERE IPG.Deleted = 0 and RO.Deleted = 0 and I.Deleted=0 and I.IsTemplate=0 and PG.Deleted = 0 " & _
                                 "  AND (IPG.ProductGroupID = " & ProductGroupID & " or  P1G.ProductGroupID = " & ProductGroupID & ")" & _
                                "UNION " & _
                                "  SELECT DISTINCT 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate " & _
                                "  FROM CPE_Deliverables D with (NoLock) " & _
                                "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                                "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                "  INNER JOIN CPE_Discounts DISC with (NoLock) on DISC.DiscountID = D.OutputID " & _
                                " left  join DiscountProductGroups PDG on PDG.DiscountID= DISC.DiscountID " & _
                                "  WHERE  D.Deleted=0 and D.DeliverableTypeId =2 and DISC.Deleted=0  and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 " & _
                                "    AND (DISC.DiscountedProductGroupID=" & ProductGroupID & " or PDG.ProductGroupID=" & ProductGroupID & " or DISC.ExcludedProductGroupID=" & ProductGroupID & ") " & _
                                "ORDER BY Name;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count = 0 Then
                ' they want to delete a group
                If (ProductGroupID <= 1) Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.nodelete", LanguageID)
                Else
                    ' check that there are no deployed offers that use this product group
                    MyCommon.QueryStr = "dbo.pa_AssociatedOffers_ST"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@LinkType", SqlDbType.Int).Value = 2
                    MyCommon.LRTsp.Parameters.Add("@LinkID", SqlDbType.Int).Value = ProductGroupID
                    rst2 = MyCommon.LRTsp_select
                    MyCommon.Close_LRTsp()

                    If (rst2.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("term.inusedeployment", LanguageID) & " : ("
                        For OfferCtr = 0 To rst2.Rows.Count - 1
                            infoMessage &= MyCommon.NZ(rst2.Rows(OfferCtr).Item("IncentiveID"), "")
                        Next
                        infoMessage &= ")"
                    Else
                        MyCommon.QueryStr = "dbo.pt_ProductGroups_Delete"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                        MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-delete", LanguageID))

                        ' remove the flags from the system-wide exclusions if this is the product group designated as either option
                        MyCommon.QueryStr = "update CPE_SystemOptions set OptionValue='' where OptionID in (67,69) and OptionValue='" & ProductGroupID & "';"
                        MyCommon.LRT_Execute()
                        SendNotificationsOfItemChange(ProductGroupID, 2)
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "pgroup-list.aspx")
                        ProductGroupID = 0
                        GName = ""
                    End If
                End If
            Else : infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.inuse", LanguageID)
            End If

        ElseIf (GetCgiValue("redeploy") <> "") Then
            Dim SetFlags As String = ""
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then SetFlags = " CMOAStatusFlag=2"
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
                If (SetFlags <> "") Then SetFlags = SetFlags & ","
                SetFlags = SetFlags & " CPEStatusFlag=2"
            End If
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                If (SetFlags <> "") Then SetFlags = SetFlags & ","
                SetFlags = SetFlags & " UEStatusFlag=2"
            End If
            MyCommon.QueryStr = "update productgroups with (RowLock) set TCRMAStatusFlag=2," & SetFlags & ", updatelevel=updatelevel+1, LastUpdatedByAdminID=" & AdminUserID & " where ProductGroupID=" & ProductGroupID
            MyCommon.LRT_Execute()
            statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
            MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-redeploy", LanguageID))

            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)

        ElseIf (GetCgiValue("add") <> "") Then
            '4722 Below
            Dim upcValidationMessage As String = String.Empty
            If Not ValidUPC(MyCommon, Trim(GetCgiValue("ExtProductID")), upcValidationMessage) Then
                infoMessage = upcValidationMessage
            Else
                '4722 Above
                Dim bGoodItemCode As Boolean = True
                Dim bProductExist As Boolean = True
                Dim bCreateProducts As Boolean = MyCommon.Fetch_SystemOption(150)
                Dim bAddProduct As Boolean = True
                Dim bMFCode As Boolean = IIF(bMFCenabled AndAlso Int(GetCgiValue("producttype")) = 4, True, False)
                ' desired product add to group
                ' dbo.pt_ProdGroupItems_Insert  @ExtProductID nvarchar(20), @ProductGroupID bigint, @ProductTypeID int, @Status int OUTPU
                'Send("Inserting product type : " & GetCgiValue("producttype"))
                If (Int(GetCgiValue("producttype")) = 1) Then
                    MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
                    rst = MyCommon.LRT_Select
                    If rst IsNot Nothing Then
                        IDLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
                    End If
                ElseIf (Int(GetCgiValue("producttype")) = 2) Then
                    Integer.TryParse(MyCommon.Fetch_SystemOption(54), IDLength)
                ElseIf (bMFCode) Then
                  'Manuf Family Code
                  IDLength = 8
                Else
                    IDLength = 0
                End If
                If (IDLength > 0) Then
                  'Manuf Family Code begin
                  If (bMFCode) Then
					If (Int(GetCgiValue("producttype")) = 2) Then
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120)).PadRight(IDLength, "0")
					Else
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26)).PadRight(IDLength, "0")
					End If
                  Else
					If (Int(GetCgiValue("producttype")) = 2) Then
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120)).PadLeft(IDLength, "0")
					Else
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26)).PadLeft(IDLength, "0")
					End If
                  End If
                Else
					If (Int(GetCgiValue("producttype")) = 2) Then
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120))
					Else
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26))
					End If
                End If
				
                'Don't change the product description if it is saved as blank
                MyCommon.QueryStr = "Select Description from Products where ExtProductID='" & ExtProductID & "' and ProductTypeID=" & Int(GetCgiValue("producttype")) & ";"
                prodDT = MyCommon.LRT_Select()
                If prodDT.Rows.Count > 0 Then
                    Description = MyCommon.NZ(prodDT.Rows(0).Item("Description"), "")
                Else
                    bProductExist = False
                End If
                If GetCgiValue("productdesc") <> "" Then
                    Description = GetCgiValue("productdesc")
                End If

				If bProductExist = False AndAlso bCreateProducts = False Then
					infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.productnotexist", LanguageID)
					bAddProduct = False
				ElseIf (CleanUPC(GetCgiValue("ExtProductID")) = False) Then
					infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
					bGoodItemCode = False
                End If

                Const Product_NotChanged_Status As Integer = 0
                Const Not_Changed_Status As Integer = 0
                Const Product_Add_Status As Integer = 1
                Const Add_Status As Integer = 1
                Const Product_Update_Status As Integer = 2
                Const Update_Status As Integer = 2
                Dim productOutputStatus As Integer = 0

                bGoodItemCode = True
                'If (MyCommon.Extract_Val(GetCgiValue("ExtProductID")) < 1) Or (Int(MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) <> MyCommon.Extract_Val(GetCgiValue("ExtProductID"))) Then
                '    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
                '    bGoodItemCode = False
                If (MyCommon.Fetch_CM_SystemOption(82) = "1" AndAlso MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) = True) Then
                    Dim sItemCode As String = GetCgiValue("ExtProductID").ToString
                    Dim productType As Integer = Int(GetCgiValue("producttype"))
                    If (productType = 1) Then
                        If (CheckItemCode(sItemCode, infoMessage) = False) Then
                            bGoodItemCode = False
                        End If
                    End If
                    'ElseIf (CleanUPC(GetCgiValue("ExtProductID")) = False) Then
                    '    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
                    '    bGoodItemCode = False
                End If

                If bGoodItemCode = True AndAlso bAddProduct = True Then
                    MyCommon.QueryStr = "dbo.pa_ProdGroupItems_ManualInsert"
                    MyCommon.Open_LRTsp()
                    If (Int(GetCgiValue("producttype")) = 2) Then
                        MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
                    Else
                        MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 19).Value = ExtProductID
                    End If
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                    MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = Int(GetCgiValue("producttype"))
                    MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 200).Value = Description
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.Parameters.Add("@ProductStatus", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                    productOutputStatus = MyCommon.LRTsp.Parameters("@ProductStatus").Value
                End If
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = "Select PhraseID,Name from ProductTypes where ProductTypeID=" & Int(GetCgiValue("producttype")) & ";"
                Dim productTypeTable As DataTable = MyCommon.LRT_Select()
                Dim typePhrase As Integer = 0
                If (productTypeTable.Rows.Count > 0) Then
                    typePhrase = MyCommon.NZ(productTypeTable.Rows(0).Item("PhraseID"), 0)
                End If
                If (productOutputStatus > Product_NotChanged_Status) Then
                    If (productOutputStatus = Product_Add_Status) Then
                        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & Copient.PhraseLib.Lookup("term.product", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & ExtProductID & _
                                              If(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
                    ElseIf (productOutputStatus = Product_Update_Status) Then
                        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("term.updated", LanguageID) & " " & Copient.PhraseLib.Lookup("term.product", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & ExtProductID & _
                                              If(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
                    End If
                End If
                If (outputStatus > Not_Changed_Status) Then
                    If (outputStatus = Add_Status) Then
                        MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-add", LanguageID) & " " & ExtProductID & IIf(typePhrase > 0, "(" & Copient.PhraseLib.Lookup(typePhrase, LanguageID) & ")", ""))
                    ElseIf (outputStatus = Update_Status) Then
                        'Product was updated to be a manual product entry from a linked product.
                        'MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, "Updated product ID " & ExtProductID)
                    End If
                End If

                If (outputStatus <> 0 OrElse productOutputStatus <> 0) Then
                    MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where ProductGroupID=" & ProductGroupID
                    MyCommon.LRT_Execute()
                    SendNotificationsOfItemChange(ProductGroupID, 2)
                End If
                If infoMessage = "" Then
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
                End If
            End If  '4722
        ElseIf (GetCgiValue("remove") <> "" AndAlso GetCgiValue("PKID") <> "") Then
            MyCommon.Open_LogixRT()
            For i = 0 To Request.Form.GetValues("PKID").GetUpperBound(0)
                MyCommon.QueryStr = "select P.ExtProductID from Products as P with (NoLock) Inner Join ProdGroupItems as PGI " & _
                                    "with (NoLock) on P.ProductID=PGI.ProductID where PGI.PKID=" & Request.Form.GetValues("PKID")(i)
                rst = MyCommon.LRT_Select()
                If rst.Rows.Count > 0 Then
                    upc = rst.Rows(0).Item("ExtProductID")
                End If
                rst = Nothing
                MyCommon.QueryStr = "dbo.pt_ProdGroupItems_Delete_ByID"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = Request.Form.GetValues("PKID")(i)
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
                SendNotificationsOfItemChange(ProductGroupID, 2)
                MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & upc)
            Next
            MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where ProductGroupID=" & ProductGroupID
            MyCommon.LRT_Execute()
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
        ElseIf (GetCgiValue("mremove") <> "") Then
            ' desired product remove from group  dbo.pt_GroupMembership_Delete_ByID  @MembershipID bigint
            ' dbo.pt_ProdGroupItems_Delete  @ExtProductID nvarchar(20), @ProductGroupID bigint, @ProductTypeID int, @Status int OUTPUT
            If (GetCgiValue("ExtProductID") <> "") Then
                If (Int(GetCgiValue("producttype")) = 1) Then
                    MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
                    rst = MyCommon.LRT_Select
                    If rst IsNot Nothing Then
                        IDLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
                    End If
                ElseIf (Int(GetCgiValue("producttype")) = 2) Then
                    Integer.TryParse(MyCommon.Fetch_SystemOption(54), IDLength)
                Else
                    IDLength = 0
                End If
                If (IDLength > 0) Then
					If (Int(GetCgiValue("producttype")) = 2) Then
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120)).PadLeft(IDLength, "0")
					Else
					    ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26)).PadLeft(IDLength, "0")
					End If
                Else
					If (Int(GetCgiValue("producttype")) = 2) Then
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 120))
					Else
						ExtProductID = MyCommon.Parse_Quotes(Left(Trim(GetCgiValue("ExtProductID")), 26))
					End If
                End If

                ' check if the product is linked to the group, if so then exclude the product from the group.
                MyCommon.QueryStr = "select PROD.ProductID, PGI.ExtHierarchyID from Products as PROD with (NoLock) " & _
                                    "inner join ProdGroupItems as PGI with (NoLock) on PGI.ProductID = PROD.ProductID " & _
                                    "where PGI.Deleted=0 and PGI.ProductGroupID=" & ProductGroupID & " and IsNull(PGI.ExtHierarchyID, '') <> ''" & _
                                    "   and PROD.ExtProductID='" & MyCommon.Parse_Quotes(ExtProductID) & "' and PROD.ProductTypeID=" & Int(GetCgiValue("producttype"))
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    If rst.Rows(0).Item("ProductID") > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("ExtHierarchyID"), "") <> "" Then
                        MyCommon.QueryStr = "insert into ProdGroupHierarchyExclusions (ExtHierarchyID,ProductGroupID, HierarchyLevel, LevelID) " & _
                                            "      values ('" & MyCommon.Parse_Quotes(MyCommon.NZ(rst.Rows(0).Item("ExtHierarchyID"), "")) & "', " & ProductGroupID & ", 2, '" & rst.Rows(0).Item("ProductID") & "')"
                        MyCommon.LRT_Execute()
                    End If
                End If

                MyCommon.Open_LogixRT()
                MyCommon.QueryStr = "dbo.[pt_ProdGroupItems_DeleteItem]"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
                MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = Int(GetCgiValue("producttype"))
                MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                MyCommon.Close_LRTsp()
                MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & ExtProductID)
                If (outputStatus <> 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.notmember", LanguageID)
                Else
                    MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where ProductGroupID=" & ProductGroupID
                    MyCommon.LRT_Execute()
                    SendNotificationsOfItemChange(ProductGroupID, 2)
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.badid", LanguageID)
            End If
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
        ElseIf (GetCgiValue("close") <> "") Then
            Response.Status = "301 Moved Permanently"
        ElseIf (GetCgiValue("copygroup") <> "") Then

            'AL-4362 Check that the generated name + "Copy of: " is not longer than 200
            If GetCgiValue("GroupName").Length > 190 Then
                infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.nametoolong", LanguageID)
            Else

                Try
                    MyCommon.QueryStr = "dbo.pa_ProductGroup_Copy"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
                    MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                    MyCommon.LRTsp.Parameters.Add("@NewProductGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    NewProductGroupID = MyCommon.LRTsp.Parameters("@NewProductGroupID").Value
                    MyCommon.Close_LRTsp()
                Catch
                    NewProductGroupID = -1
                End Try

                If NewProductGroupID = -1 Then
                    infoMessage = Copient.PhraseLib.Lookup("pgroup-edit.invalidcopy", LanguageID)
                Else
                    MyCommon.Activity_Log(5, NewProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-copy", LanguageID) & " " & ProductGroupID)
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & NewProductGroupID)
                End If
            End If
        End If

        ' START - Hierarchy ******************************************
        If bDown Then
            If ParentNodeIdList = "" Then
                ParentNodeIdList = SelectedNodeId
            Else
                ParentNodeIdList += "," & SelectedNodeId
            End If
            SelectedNodeId = ""
            NodeName = ""
            '  Send("DEBUG: PArtentNodeIDList=" & ParentNodeIdList & "<br />")
        End If

        If bUp Then
            If ParentNodeIdList <> "" Then
                Dim n As Integer
                n = ParentNodeIdList.LastIndexOf(",")
                If n > 0 Then
                    SelectedNodeId = ParentNodeIdList.Substring(n + 1)
                    ParentNodeIdList = ParentNodeIdList.Substring(0, n)
                Else
                    SelectedNodeId = ParentNodeIdList
                    ParentNodeIdList = ""
                End If
            End If
            NodeName = ""
        End If

        If ParentNodeIdList = "" Then
            ' No hierarchy is selected, so no parents
            ' existing Hierarchies makeup children
            ParentId = ""
            'squery = "select HierarchyId as Id, [Name] as Name from ProdHierarchies with (NoLock)"
            squery = "select HierarchyId as Id,Name = " & _
                     "   case  " & _
                     "       when ExternalID is NULL then Name " & _
                     "       when ExternalID = '' then Name " & _
                     "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                     "       else ExternalID " & _
                     "   end " & _
                     "from ProdHierarchies with (NoLock);"
        Else
            NodeIds = ParentNodeIdList.Split(",")
            'MyCommon.QueryStr = "select HierarchyId as Id, [Name] as Name from ProdHierarchies with (NoLock) where HierarchyId = " & NodeIds(0)
            MyCommon.QueryStr = "select HierarchyId as Id,Name = " & _
                                "   case  " & _
                                "       when ExternalID is NULL then Name " & _
                                "       when ExternalID = '' then Name " & _
                                "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                                "       else ExternalID " & _
                                "   end " & _
                                "from ProdHierarchies with (NoLock) where HierarchyID=" & NodeIds(0)
            dtParents = MyCommon.LRT_Select()
            If NodeIds.Length = 1 Then
                ' No node is selected, selected hierarchy is parent
                ' existing root nodes for this hierarchy makeup children
                ParentId = "0"
                'squery = "select NodeId as Id, Name as Name from PHNodes with (NoLock) where ParentId = 0 and HierarchyId =  " & NodeIds(0)
                squery = "select NodeId as Id, Name = " & _
                         "   case  " & _
                         "       when ExternalID is NULL then Name " & _
                         "       when ExternalID = '' then Name " & _
                         "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                         "       else ExternalID " & _
                         "   end " & _
                         "from PHNodes with (NoLock) where ParentId = 0 and HierarchyId = " & NodeIds(0)
            Else
                For i = 1 To NodeIds.Length - 1
                    'MyCommon.QueryStr = "select NodeId as Id, Name as Name from PHNodes with (NoLock) where NodeId = " & NodeIds(i)
                    MyCommon.QueryStr = "select NodeId as Id, Name= " & _
                                        "   case  " & _
                                        "       when ExternalID is NULL then Name " & _
                                        "       when ExternalID = '' then Name " & _
                                        "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                                        "       else ExternalID " & _
                                        "   end " & _
                                        "from PHNodes with (NoLock) where NodeId = " & NodeIds(i)
                    dtParents1 = MyCommon.LRT_Select()
                    dtParents.Merge(dtParents1)
                Next
                ' parents consist of hierarchy and listed nodes
                ' children made-up of nodes with parentId = last parent
                ParentId = dtParents.Rows(dtParents.Rows.Count - 1).Item("id").ToString
                'squery = "select NodeId as Id, Name as Name from PHNodes with (NoLock) where ParentId = " & ParentId
                squery = "select NodeId as Id,Name= " & _
                         "   case  " & _
                         "       when ExternalID is NULL then Name " & _
                         "       when ExternalID = '' then Name " & _
                         "       when ExternalID not like '%' + Name + '%' then ExternalID + '-' + Name " & _
                         "       else ExternalID " & _
                         "   end " & _
                         "from PHNodes with (NoLock) where ParentId = " & ParentId
            End If
        End If

        MyCommon.QueryStr = squery
        dtChildren = MyCommon.LRT_Select()

        If SelectedNodeId = "" AndAlso dtChildren.Rows.Count > 0 Then
            SelectedNodeId = dtChildren.Rows(0).Item("id")
        End If

        ' END - Hierarchy ******************************************

        If bAdd Then
            ' In this case the LocationList contains LocationId's from the LHContainer table
            ProductList = GetCgiValue("level-avail")
            If ProductList <> "" Then
                Products = ProductList.Split(",")
                For i = 0 To Products.Length - 1
                    ' Send("Product: " & Products(i) & "<br />")
                    ' determine the type of item for the selected product
                    MyCommon.QueryStr = "select ProductTypeID from Products with (NoLock) where ExtProductId = '" & Products(i) & "'"
                    typeST = MyCommon.LRT_Select
                    'Send("Get name type: " & MyCommon.QueryStr)
                    iType = MyCommon.NZ(typeST.Rows(0).Item("ProductTypeID"), 1)
                    MyCommon.QueryStr = "dbo.pt_ProdGroupItems_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = Products(i)
                    MyCommon.LRTsp.Parameters.Add("@ProductGroupId", SqlDbType.BigInt, 8).Value = ProductGroupID.ToString
                    MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = iType
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                    MyCommon.Close_LRTsp()
                    ' Send("<br /><br />Adding: " & outputStatus & "<br /><br />")
                Next
            End If
        ElseIf bAddAll Then
            If SelectedNodeId <> "" Then
                Dim HierarchyId As String
                Dim CurrentNode As String
                Dim dtProducts As DataTable
                If ParentNodeIdList = "" Then
                    ' Hierarchy level
                    HierarchyId = SelectedNodeId
                    CurrentNode = "0"
                Else
                    ' Node level
                    HierarchyId = NodeIds(0)
                    CurrentNode = SelectedNodeId
                End If
                squery = "select distinct ProductId as Id from GetBranchProducts(" & CurrentNode & "," & HierarchyId & ") where ProductId not in"
                squery += " (select ProductId from ProdGroupItems with (NoLock) where Deleted = 0 and ProductGroupId = " & ProductGroupID & ")"
                MyCommon.QueryStr = squery
                dtProducts = MyCommon.LRT_Select()
                If dtProducts.Rows.Count > 0 Then
                    For Each dr As DataRow In dtProducts.Rows
                        ' determine the type of item for the selected product
                        MyCommon.QueryStr = "select ProductTypeID,ExtProductId from Products with (NoLock) where ProductId = '" & dr.Item("Id") & "'"
                        typeST = MyCommon.LRT_Select
                        'Send("Get name type: " & MyCommon.QueryStr)
                        iType = MyCommon.NZ(typeST.Rows(0).Item("ProductTypeID"), 1)
                        MyCommon.QueryStr = "dbo.pt_ProdGroupItems_Insert"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = typeST.Rows(0).Item("ExtProductID")
                        MyCommon.LRTsp.Parameters.Add("@ProductGroupId", SqlDbType.BigInt, 8).Value = ProductGroupID.ToString
                        MyCommon.LRTsp.Parameters.Add("@ProductTypeID", SqlDbType.Int).Value = iType
                        MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        outputStatus = MyCommon.LRTsp.Parameters("@Status").Value
                        MyCommon.Close_LRTsp()
                        'MyCommon.QueryStr = "dbo.pt_LocGroupItems_Insert"
                        'MyCommon.Open_LRTsp()
                        'MyCommon.LRTsp.Parameters.Add("@LocationGroupId", SqlDbType.BigInt, 8).Value = ProductGroupID.ToString
                        'MyCommon.LRTsp.Parameters.Add("@LocationId", SqlDbType.BigInt, 8).Value = dr.Item("Id")
                        'MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        'MyCommon.LRTsp.ExecuteNonQuery()
                        'MyCommon.Close_LRTsp()
                    Next
                End If
            End If
        ElseIf bRemove Then
            ' In this case the LocationList contains Primary Key values (PkId) from the LocationGroupItems table
            ProductList = GetCgiValue("level-group")
            If ProductList <> "" Then
                Products = ProductList.Split(",")
                For i = 0 To Products.Length - 1
                    MyCommon.QueryStr = "dbo.pt_LocGroupItems_Delete"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@PkId", SqlDbType.BigInt, 8).Value = Products(i)
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                Next
            End If
        End If

        If Not dtParents Is Nothing AndAlso dtParents.Rows.Count > 0 AndAlso dtChildren.Rows.Count > 0 Then
            squery = "select a.ProductId as Id,a.ExtProductID as Code, a.Description, PT.ProductTypeID, Name, PT.PhraseID, b.PKID from Products a with (NoLock), PHContainer b with (NoLock), ProductTypes as PT"
            squery += " with (NoLock) where a.ProductId = b.ProductId and PT.ProductTypeID=a.ProductTypeID and b.NodeId = " & SelectedNodeId
            If ProductGroupID > 0 Then
                squery += " and a.ProductId not in ("
                squery += "select ProductId from ProdGroupItems with (NoLock) where Deleted = 0 and ProductGroupId = " & ProductGroupID & ")"
            End If
            MyCommon.QueryStr = squery
            dtProdAvailable = MyCommon.LRT_Select()
            ProdAvailableCount = dtProdAvailable.Rows.Count
        Else
            ProdAvailableCount = 0
        End If

        ' If ProductGroupID > 0 Then
        '     squery = "select b.PkId as Id, a.ExtProductID as Code from Products a, ProdGroupItems b"
        '     squery += " with (NoLock) where a.ProductId = b.ProductId and b.deleted = 0 and b.ProductGroupId = " & ProductGroupID
        '     MyCommon.QueryStr = squery
        '     dtProdAssigned = MyCommon.LRT_Select()
        '     ProdAssignedCount = dtProdAssigned.Rows.Count
        ' End If
        ' Send("DEBUG: squery" & squery)

        If (GetCgiValue("mode") <> "Create") Then
            MyCommon.QueryStr = "select Name,ExtGroupID,CreatedDate,LastUpdate,LastLoaded,LastLoadMsg,PointsNotApplyGroup,NonDiscountableGroup,ProductGroupTypeID, B.BuyerID ,B.ExternalBuyerId " & _
                " from ProductGroups P with (NoLock) left join Buyers B with (NoLock)  on B.BuyerId=p.BuyerId where ProductGroupID = @ProductGroupID and deleted = 0"
            MyCommon.DBParameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (rst.Rows.Count > 0) Then
                For Each row In rst.Rows
                    GName = row.Item("Name")
                    CreatedDate = row.Item("CreatedDate")
                    LastUpdate = row.Item("LastUpdate")
                    LastUpload = MyCommon.NZ(row.Item("LastLoaded"), "1/1/1900")
                    LastUploadMsg = MyCommon.NZ(row.Item("LastLoadMsg"), "")
                    XID = MyCommon.NZ(row.Item("ExtGroupID"), "")
                    IsSpecialGroup = (MyCommon.NZ(row.Item("PointsNotApplyGroup"), False) OrElse MyCommon.NZ(row.Item("NonDiscountableGroup"), False))
                    ProductGroupTypeID = MyCommon.NZ(row.Item("ProductGroupTypeID"), 1)
                    BuyerID = MyCommon.NZ(row.Item("BuyerID"), -1)
                    ExternalBuyerId = MyCommon.NZ(row.Item("ExternalBuyerId"), "")
                Next

                If (ProductGroupTypeID = 2) Then
                    ucProductAttributeFilter.IsPGAttributeType = (AttributePGEnabled)
                    ucProductAttributeFilter.ProductGroupID = ProductGroupID
                    ucProductAttributeFilter.BuyerID = BuyerID
                    ucProductAttributeFilter.LanguageID = LanguageID
                    ucProductAttributeFilter.IsEditPermitted = Logix.UserRoles.EditProductGroups And EditProductRegardlessOfBuyer
                End If
                MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & ProductGroupID & " And Deleted = 0"
                rst = MyCommon.LRT_Select()
                For Each row In rst.Rows
                    GroupSize = row.Item("GCount")
                Next
                ShowAllItems = (GetCgiValue("showall") = "true")
                MyCommon.QueryStr = "select count(*) as PCount from ProdGroupItems PGI with (NoLock) inner join products PRD on PGI.productid = PRD.productid " & _
                              "where PGI.ProductGroupID = " & ProductGroupID & " And PGI.Deleted = 0 And (PRD.Description IS NULL OR  PRD.Description = '')"
                rst = MyCommon.LRT_Select()
                If rst.Rows.Count > 0 Then
                    ProductsWithoutDesc = rst.Rows(0).Item("PCount")
                End If

                Dim bBlankDescProd As Boolean = (MyCommon.Fetch_SystemOption(206) = "1")
                Dim sBlankDescOrderByStr As String = " order by case when Description is null or Description = '' then nullif(Description, '') else CAST(ExtProductID as bigint) end, CAST(ExtProductID as bigint), ExtProductID DESC;"
                MyCommon.QueryStr = "select" & If(ShowAllItems, "", " top 100") & " GM.ProductID, PKID, CID.ProductTypeID, ExtProductID, Description, PT.Name as ProductType, PT.PhraseID " & _
                                      "from Products as CID with (NoLock) " & _
                                      "inner join ProdGroupItems as GM with (NoLock) on CID.ProductID=GM.ProductID " & _
                                      "left join ProductTypes as PT with (NoLock) on PT.ProductTypeID=CID.ProductTypeID " & _
                                      "where GM.ProductGroupID=" & ProductGroupID & " and GM.Deleted=0 and IsNull(GM.ExtHierarchyID, '')='' and IsNull(GM.ExtNodeID, '')='' " & _
                              "and IsNull(GM.ExtNodeID, '')='' " & If(bBlankDescProd, sBlankDescOrderByStr, " order by ExtProductID;")
                rstItems = MyCommon.LRT_Select()
                ListBoxSize = rstItems.Rows.Count
            ElseIf (GetCgiValue("new") = "") And (ProductGroupID > 0) Then
                ' check if this is a deleted product group
                MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID=" & ProductGroupID & " and deleted =1"
                rst = MyCommon.LRT_Select()
                If (rst.Rows.Count > 0) Then
                    GName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
                Else
                    GName = ""
                End If

                Send_HeadBegin("term.productgroup", , ProductGroupID)
                Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
                Send_Metas()
                Send_Links(Handheld)
                Send_Scripts()
                Send_HeadEnd()
                If CreatedFromOffer Then
                    Send_BodyBegin(3)
                Else
                    Send_BodyBegin(1)
                    Send_Bar(Handheld)
                    Send_Help(CopientFileName)
                    Send_Logos()
                    Send_Tabs(Logix, 4)
                    Send_Subtabs(Logix, 41, 4, , ProductGroupID)
                End If
                Send("<div id=""intro"">")
                Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " #" & GetCgiValue("ProductGroupID") & " - " & GName & "</h1>")
                Send("</div>")
                Send("<div id=""main"">")
                Send("  <div id=""infobar"" class=""red-background"">")
                Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
                Send("  </div>")
                Send("</div>")
                Send_BodyEnd()
                Response.End()
                GoTo done
            End If
        End If

        If (GetCgiValue("ItemPK") <> "") Then
            Integer.TryParse(GetCgiValue("ItemPK"), ItemPKID)
        End If

        CpeEngineOnly = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) And _
                        Not (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) And _
                        Not (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.Catalina))
        Send_HeadBegin("term.productgroup", , ProductGroupID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        If CreatedFromOffer Then
            Send_BodyBegin(3)
        Else
            Send_BodyBegin(1)
            Send_Bar(Handheld)
            Send_Help(CopientFileName)
            Send_Logos()
            Send_Tabs(Logix, 4)
            Send_Subtabs(Logix, 41, 4, , ProductGroupID)
        End If

        ' determine whether to show the View selected hierarchy nodes button
        MyCommon.QueryStr = "select count(NodeID) as NodeCount from ProductGroupNodes with (NoLock) where ProductGroupID=" & ProductGroupID
        rst = MyCommon.LRT_Select
        ShowViewSelected = (MyCommon.NZ(rst.Rows(0).Item("NodeCount"), 0) > 0)

        ' Retreiving offers asscoiated with the product group
        '-------------------------------------------------------------------------------------------------------------------------------
        Dim conditionalQuery As String = String.Empty
  
        If (bEnableRestrictedAccessToUEOfferBuilder) Then
            conditionalQuery = GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "I")
        End If
                
        If bStoreUser Then
            sJoin = "Full Outer Join OfferLocUpdate olu with (NoLock) on O.OfferID=olu.OfferID "
            wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) "
        End If
        If (ProductGroupID <> 0) Then
            MyCommon.QueryStr = "select 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate,NULL as BuyerID from offerconditions as OC with (NoLock) " & _
                                "  left join Offers as O with (NoLock) on O.Offerid=OC.offerID " & _
                                " " & sJoin & " " & _
                                "  where (linkid=" & ProductGroupID & " or excludedid=" & ProductGroupID & ") and  ConditionTypeID=2  and " & _
                                "  OC.deleted=0 and O.deleted=0 and O.IsTemplate=0 " & wherestr & _
                                " UNION " & _
                                " select 1 as EngineID, O.OfferID,O.Name,O.ProdEndDate,NULL as BuyerID from offerrewards as OFR with (NoLock) " & _
                                "  left join Offers as O with (NoLock) on O.Offerid=OFR.offerID " & _
                                " " & sJoin & " " & _
                                "  where (ProductGroupID=" & ProductGroupID & " or ExcludedProdGroupID=" & ProductGroupID & ")   and " & _
                                "  OFR.deleted=0 and O.deleted=0 and O.IsTemplate=0 " & wherestr & _
                                " UNION " & _
                                " select distinct 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID " & _
                                "  FROM CPE_IncentiveProductGroups IPG with (NoLock)" & _
                                " left JOIN ProductConditionProductGroups P1G on P1G.IncentiveProductGroupID= IPG.IncentiveProductGroupID" & _
                                "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID" & _
                                "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                "  INNER JOIN ProductGroups PG with (NoLock) on IPG.ProductGroupID = PG.ProductGroupID " & _
                                 "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                "  WHERE  IPG.Deleted = 0 and RO.Deleted = 0 and I.Deleted=0 and I.IsTemplate=0 and PG.Deleted = 0 " & _
                                "  AND (IPG.ProductGroupID = " & ProductGroupID & " or  P1G.ProductGroupID = " & ProductGroupID & ")"
            If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
            MyCommon.QueryStr &= " UNION " & _
                                " select distinct 2 as EngineID, I.IncentiveID as OfferID, I.IncentiveName as Name, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID " & _
                                "  FROM CPE_Deliverables D with (NoLock) " & _
                                "  INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                                "  INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                                "  INNER JOIN CPE_Discounts DISC with (NoLock) on DISC.DiscountID = D.OutputID " & _
                                " left JOIN DiscountProductGroups PDG on PDG.DiscountID= DISC.DiscountID " & _
                                 "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                "  WHERE  D.Deleted=0 and D.DeliverableTypeId =2 and DISC.Deleted=0  and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 " & _
                                "    AND (DISC.DiscountedProductGroupID=" & ProductGroupID & " or PDG.ProductGroupID = " & ProductGroupID & " or DISC.ExcludedProductGroupID=" & ProductGroupID & ") "
            If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
            MyCommon.QueryStr &= "ORDER BY Name;"
            associatedOfferDT = MyCommon.LRT_Select
        End If
        '-----------------------------------------------------------------------------------------------------------------------------

done:
    
       
    End Sub

  Protected Sub RadioButtonList1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButtonList1.DataBound
    Dim PhraseID As Integer
    For i As Integer = 0 To sender.Items.Count - 1
      If (Not String.IsNullOrWhiteSpace(sender.Items(i).Text) AndAlso Int32.TryParse(sender.Items(i).Text, PhraseID)) Then
        sender.Items(i).Text = Copient.PhraseLib.Lookup(PhraseID, LanguageID)
      End If
    Next
  End Sub

  Function TestFileUpload(ByVal TestString As String)
    Dim IsMatch As Boolean
    Dim RegString As String
    'RegString = "^[^,]*[^:a,].*$"
    RegString = "^[:a]*,"
    Dim rx As New Regex(RegString, RegexOptions.Multiline)

    If (rx.Matches(TestString).Count > 0) Then
      IsMatch = True
    Else
      IsMatch = False
    End If

    'Return IsMatch
    Return rx.Matches(TestString).Count
  End Function

  Function StrClone(ByVal Source As String, ByVal Count As Integer) As String
    Return Replace(Space(Count), " ", Source)
  End Function
    
  Function CheckIsPGEditable(ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc, ByVal bStoreUser As Boolean, ByVal ProductGroupID As Long ,ByVal dt As DataTable) As Boolean
      Dim isEditable As Boolean = True
      Dim subQuery As New StringBuilder
      Dim Query As New StringBuilder

      Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
      
       If (bEnableAdditionalLockoutRestrictionsOnOffers AndAlso Not Logix.UserRoles.EditOfferPastLockoutPeriod AndAlso (Not dt Is Nothing AndAlso dt.Rows.Count > 0)) Then
          
          'retrieving all offerid's from datatable associated with product group
          Dim offerIds = String.Join(",", (From u In dt.AsEnumerable() Select u.Field(Of Int64)("OfferID")).Distinct().ToArray())
          
              
              Query.Append(" WITH query_CTE AS ( ")
              Query.Append(" SELECT  bo.OfferID,FI.FolderID,f.FolderName,f.StartDate as FolderStartDate,bt.Lockoutdays,DateADD(day,bt.Lockoutdays,GETDATE()) as Lockoutdate FROM BannerOffers BO WITH(NOLOCK) ")
              Query.Append(" INNER JOIN folderItems FI WITH(NOLOCK) ON FI.LinkID=BO.OfferID  ")
              Query.Append(" INNER JOIN FolderThemes FT WITH(NOLOCK) ON FT.FolderID=FI.FolderID ")
              Query.Append(" INNER JOIN Folders F WITH(NOLOCK) ON F.FolderID=FT.FolderID ")
              Query.Append(" INNER JOIN BannerThemes BT WITH(NOLOCK) ON BT.BannerID=BO.BannerID AND FT.ThemeID=BT.ThemeID ")
              Query.Append(" WHERE bo.OfferID IN( ").Append(offerIds).Append(" ) ")
              Query.Append(" ) ")
              Query.Append(" SELECT offerID, FolderStartDate, Lockoutdate,Lockoutdays ")
              Query.Append(" FROM query_CTE where FolderStartDate <= Lockoutdate ")

              MyCommon.QueryStr = Query.ToString()
              rst = MyCommon.LRT_Select
              If (Not rst Is Nothing AndAlso rst.Rows.Count > 0) Then
                  'One or more offers associated with this product group is in lockout period and so the PG is uneditable
                  isEditable = False
              End If
          End If
      
      Return isEditable
  End Function
    
</script>
<%
 'it should be in the page load but cannot be done as of now because its leading to nullpointer exception for some of the reference variables.
    If Logix.UserRoles.AccessProductGroups = False Then
        Send_Denied("1", "perm.product-access")
        GoTo done
    ElseIf (MyCommon.IsEngineInstalled(9) AndAlso ProductGroupID > 0) Then
        MyCommon.QueryStr = "select * from buyerroleusers where adminuserid=" & AdminUserID
        If MyCommon.LRT_Select().Rows.Count > 0 Then
            MyCommon.QueryStr = "dbo.pa_IsProductGroupAccessibleToBuyer"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@ProductGroupID", SqlDbType.BigInt).Value = ProductGroupID
            MyCommon.LRTsp.Parameters.Add("@checkPGlistBasedOnBuyerID", SqlDbType.Bit).Value = (Logix.UserRoles.ViewProductgroupRegardlessBuyer = False)
            MyCommon.LRTsp.Parameters.Add("@adminuserid", SqlDbType.Int).Value = AdminUserID
            Dim IsPGAccessible As Boolean = (MyCommon.LRTsp_select().Rows.Count > 0)
            MyCommon.Close_LRTsp()
            If (IsPGAccessible = False) Then
                Send_Denied("1", "perm.viewProductgroupsRegardlessBuyer")
                GoTo done
            End If
        End If
    End If
 %>
<%
  
  If MyCommon.Fetch_SystemOption(75) Then
    If (ProductGroupID > 0 And Logix.UserRoles.AccessNotes And Not CreatedFromOffer) Then
      Send_Notes(7, ProductGroupID, AdminUserID)
    End If
  End If
  Send_BodyEnd("mainform", "GroupName")
  Send("<div id=""disabledBkgrd"" style=""position:absolute;top:0px;left:0px;right:0px;width:100%;height:100%;background-color:Gray;display:none;z-index:99;filter:alpha(opacity=50);-moz-opacity:.50;opacity:.50;""></div>")
done:
  MyCommon.Close_LogixRT()
  Logix = Nothing
    MyCommon = Nothing
%>