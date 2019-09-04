<%@ Control Language="C#" AutoEventWireup="true" CodeFile="UDFJavaScript.ascx.cs"
  Inherits="logix_UserControls_UDFJavaScript" %>
<% // version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%
  //' *****************************************************************************
  //  ' * FILENAME: UDFJavaScript.ascx 
  //  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  //  ' * Copyright © 2002 - 2014.  All rights reserved by:
  //  ' *
  //  ' * NCR Corporation
  //  ' * 2651 Satellite Blvd
  //  ' * Duluth, GA 30096     
  //  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  //  ' *
  //  ' * PROJECT : NCR Advanced Marketing Solution
  //  ' *
  //  ' * MODULE  : Logix
  //  ' *
  //  ' * PURPOSE : 
  //  ' *
  //  ' * NOTES   : 
  //  ' *
  //  ' * Version : 7.3.1.138972 
  //  ' *
  //  ' *****************************************************************************
%>
<script type="text/javascript" src="/javascript/jquery.min.js"></script>
<script type="text/javascript" src="/javascript/jquery-ui-1.10.3/jquery-ui-min.js"></script>
<link rel="stylesheet" href="/javascript/jquery-ui-1.10.3/style/jquery-ui.min.css" />
<script type="text/javascript">

    function controlAddSelection()
    {
        var x = document.getElementById("UDFDataType");
        if(x.length > 0)
        {
            document.getElementById("UDFAddButton").disabled=false;
            document.getElementById("UDFDataType").disabled=false;
        }
        else
        {
            document.getElementById("UDFAddButton").disabled=true;
            document.getElementById("UDFDataType").disabled=true;
        }
    }
    
    /*this function associates the udf with the offer, but does NOT save the value*/
    /*begin UDF javascript functions*/
    function addUDF(offerID) {
        var x = document.getElementById("UDFDataType").selectedIndex;
        if (x==-1)
        {
            setTimeout(controlAddSelection,100);
            return;
        }
        var y = document.getElementById("UDFDataType").options;
        var udfval;
        var xmlHttpReq = false;

        var self = this;
        var strURL;
        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
        // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }
        strURL = "/logix/UDFOptions.aspx?mode=OfferAdd&udf=" + y[x].index + "&OfferID=" + offerID;
        didUdfChange=true;

        var table = document.getElementById("udftable");
        for (var i = 0, row; row = table.rows[i]; i++) {

            var val = row.id;
            if (val.length > 0) {                
                val = val.substring(2);
                saveValue(document.getElementById(val));
            }
        }

        self.xmlHttpReq.open('POST', strURL, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.onreadystatechange = function () {
            if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
                $("#udftable > tbody ").append(self.xmlHttpReq.responseText);
                y[x] = null;
            }
        }
        self.xmlHttpReq.send();
        
        setTimeout(controlAddSelection,100);         
    }
    function saveValue(ele) {	
        if (typeof (ele) == 'undefined' || ele == null) {
            return;
        }
        if (ele.value)
            ele.setAttribute("value", ele.value);
        if (ele.checked)
            ele.setAttribute("checked", ele.checked);
    }

    function onUDFListBoxChange(elem)
    {
        didUdfChange=true;
    }
/*
   this function deletes row from udftable and resets the user defined fields drop down
   also makes call to server to stage the userdefinedfeildsvalues row for deletion
*/
    function deleteUDF(udfpk, offerID) {
        
        if (confirm(<% Response.Write("\"" +Copient.PhraseLib.Lookup("confirm.delete", LanguageID) + "\""); %>)) {

            var index = document.getElementById("TRudfVal-" + udfpk).rowIndex;
            document.getElementById("udftable").deleteRow(index);
            var xmlHttpReq = false;

            var self = this;
            var strURL;

            didUdfChange = true;
            // Mozilla/Safari
            if (window.XMLHttpRequest) {
                self.xmlHttpReq = new XMLHttpRequest();
            }
            // IE
            else if (window.ActiveXObject) {
                self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
            }

            var table = document.getElementById("udftable");
            for (var i = 0, row; row = table.rows[i]; i++) {

                var val = row.id;
                if (val.length > 0) {
                    val = val.substring(2);
                    saveValue(document.getElementById(val));
                }
            }

            strURL = "/logix/UDFOptions.aspx?mode=OfferDel&udf=" + udfpk + "&OfferID=" + offerID
            self.xmlHttpReq.open('POST', strURL, true);
            self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
            self.xmlHttpReq.onreadystatechange = function () {
                if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
                    $("#udfSelect").html(self.xmlHttpReq.responseText);
                }
            }
            self.xmlHttpReq.send();

        }
        
        setTimeout(controlAddSelection,100);         
    }


    function assigntextboxcontrol(cntrl, elemName, shown) {

        var selectedUDFString = cntrl.name;
        var selectedUDFPK = selectedUDFString.split("-");
        document.mainform.SelectedUDF.value = selectedUDFPK[1];
        var elem = document.getElementById(elemName);
        var fadeElem = document.getElementById('UDFfadeDiv');
        var newValue = '';

        if (elem != null) {        
            elem.style.display = (shown) ? 'block' : 'none';
        }
        if (fadeElem != null) {        
            fadeElem.style.display = (shown) ? 'block' : 'none';
        }
        if (shown) {        
          xmlhttpPost_UDFString('/logix/OfferFeeds.aspx', 'Mode=UDFStringValue&UDFPK=' + document.mainform.SelectedUDF.value + '&OfferID=' + document.mainform.form_OfferID.value, 'UDFStringValue');
          var imageId = 'Image_' + selectedUDFPK[1];
          var imageElem = document.getElementById(imageId);
          if (imageElem != null) {
            var src = document.getElementById("udfVal-" + selectedUDFPK[1]).value;
            var src1 = '/logix/show-image.aspx?caller=udf&src=' + src
            imageElem.src = src1;
            imageElem.onclick = function() {
            showFullSizedImage(src1)
          };
        }
      }
    }

    function xmlhttpPost_UDFString(strURL, qryStr, action) {
        var xmlHttpReq = false;
        var self = this;
        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
        // IE
        else if (window.ActiveXObject) {        
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");            
        }
        self.xmlHttpReq.open('POST', strURL + '?' + qryStr, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.setRequestHeader("Content-length", qryStr.length);
        self.xmlHttpReq.setRequestHeader("Connection", "close");
        self.xmlHttpReq.onreadystatechange = function () {        
            if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
                 document.body.style.cursor='default';
                if (action == 'UDFStringValue') {
                    document.getElementById("txtOfferUDFstringValue").value = self.xmlHttpReq.responseText;
                    var opt = document.getElementById("txtOfferUDFstringValue").value;      
                    opt = opt.replace(/(\r\n|\n|\r)/gm, "");
                    document.getElementById("txtOfferUDFstringValue").value = opt;
                }
            }
        }
        document.body.style.cursor='wait';
        self.xmlHttpReq.send(qryStr);
        return false;
    }


     /*
   **  Returns the caret (cursor) position of the specified text field.
   **  Return value range is 0-oField.length.
   reference : http://www.webdeveloper.com/forum/showthread.php?74982-How-to-set-get-caret-position-of-a-textfield-in-IE
   */
   function doGetCaretPosition (oField) {

     // Initialize
     var iCaretPos = 0;

     // IE Support
     if (document.selection) { 

       // Set focus on the element
       oField.focus ();
  
       // To get cursor position, get empty selection range
       var oSel = document.selection.createRange ();
  
       // Move selection start to 0 position
       oSel.moveStart ('character', -oField.value.length);
  
       // The caret position is selection length
       iCaretPos = oSel.text.length;
     }

     // Firefox support
     else if (oField.selectionStart || oField.selectionStart == '0')
       iCaretPos = oField.selectionStart;

     // Return results
     return (iCaretPos);
   }

    function validateNumericRange(control,evt)
    {
        if(control.value == "")//if it hasn't been set yet, we'll allow it
        {
            return true;
        }
        var label = document.getElementById(control.id+"_label");
        var ranges = label.innerHTML;
        var rangeList = ranges.split(',');
        for(var i = 0; i< rangeList.length;i++)
        {
            var parts = rangeList[i].split(':');
            if(parts.length==2)
            {
                parts[0] = parts[0].replace("{","");
                parts[1] = parts[1].replace("}","");
                if(parseInt(control.value) >= parseInt(parts[0]) && parseInt(control.value) <= parseInt(parts[1]))
                {
                    return true;
                }

            }
            else{
                if(parseInt(control.value) == parseInt(parts[0]))
                {
                    return true;
                }
            }
        }

        alert("integer entered is not in a valid range");
        control.focus();
        return false;
    }

    function isNumber(control,evt) 
    {        
        var cp = doGetCaretPosition(control);
        var evtobj=window.event? event : e //distinguish between IE's explicit event object (window.event) and Firefox's implicit.
        var unicode=evtobj.charCode? evtobj.charCode : evtobj.keyCode
        var actualkey=String.fromCharCode(unicode)
        if(actualkey == "0" ||
            actualkey == "1" || 
            actualkey == "2" ||
            actualkey == "3" || 
            actualkey == "4" ||
            actualkey == "5" || 
            actualkey == "6" ||
            actualkey == "7" || 
            actualkey == "8" ||
            actualkey == "9" || 
            actualkey == "-" )//||
           // actualkey == ".")
        {
            if(actualkey == "-" && cp > 0)
            {//only allow negative as the first character
                return false;
            }
            else if(actualkey == ".")
            {
                var currentString = control.value;
                if(currentString.indexOf(".") != -1) //only allow one decimal
                {
                    return false;
                }
                if(cp <= currentString.indexOf("-"))//make sure decimal is not before negative
                {
                    return false;
                }
            }
            else
            {
                return true;
            }
        }
        else
        {
            return false;
        }
    }

function addUDFTextmessagetoOffer(elemName) {
    var changedValue = document.getElementById("txtOfferUDFstringValue").value;

    changedValue = encodeURIComponent(changedValue);
	//changedValue = replaceAll("&","%26",changedValue);
	//changedValue = replaceAll("+","%2B",changedValue);
	//changedValue = replaceAll("#","%23",changedValue);
	
    if (document.getElementById("txtOfferUDFstringValue").value.length > 1000){
	alert('Maximum 1000 characters can only be entered!');
	} else 
	{
    xmlhttpPost_UDFString('/logix/OfferFeeds.aspx', 'Mode=UDFStringValue&UDFPK=' + document.mainform.SelectedUDF.value + '&OfferID=' + document.mainform.form_OfferID.value + '&UDFValue=' + changedValue,'UDFStringValue');
	  document.getElementById("udfVal-" + document.mainform.SelectedUDF.value).value = document.getElementById("txtOfferUDFstringValue").value;
	  document.getElementById("txtOfferUDFstringValue").value = "";
		var imageId = 'Image_' + document.mainform.SelectedUDF.value;
		var imageElem = document.getElementById(imageId);
		if (imageElem != null) {
			var src = document.getElementById("udfVal-" + document.mainform.SelectedUDF.value).value;
			var src1 = '/logix/show-image.aspx?caller=udf&src=' + src;
			imageElem.src = src1;
			imageElem.onclick = function() {
				showFullSizedImage(src1)
			};
		}
	//  // hide the receipt messages popup
	  toggleDialog('foldercreate',false);
    }
} 

    /*End UDF javascript functions*/
</script>

<script type="text/javascript">

  function showFullSizedImage(imagesrc)
  {
    var elemImg = document.getElementById('fullSizedImage');
    var elemWin = document.getElementById('imagepopup');
    didViewImage = true;

    if (elemImg != null) {
      elemImg.src = imagesrc;
      elemImg.onload = function ()
      {
        resizeImage(elemImg, 600, 600)
      };
    }
    if (elemWin != null) {
      elemWin.style.display = '';
    }
  }

  function resizeImage(obj, maxW, maxH)
  {
    var iw = parseInt(obj.naturalWidth);
    var ih = parseInt(obj.naturalHeight);
    var aspect = (iw * 1.0) / (ih * 1.0);
    if (iw > maxW) {
      iw = maxW;
      ih = iw / aspect;
    }
    if (ih > maxH) {
      ih = maxH;
      iw = ih * aspect
    }
    obj.setAttribute("width", iw);
    obj.setAttribute("height", ih);
  }

  function closeImage()
  {
    didViewImage = true;
    var elemWin = document.getElementById('imagepopup');
    elemWin.style.display = 'none';
  }
</script>
