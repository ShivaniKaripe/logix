/// Customer website JavaScript functions ///
// version:7.3.1.138972.Official Build (SUSDAY10202)

var dtCh= "/";
var minYear=1900;
var maxYear=2100;



/// Popup window launcher ///
function openPopup(url) {
	popW = 700;
	popH = 570;
	siteWindow = window.open(url,"Popup", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
	siteWindow.focus();
}

function openMiniPopup(url) {
	popW = 250;
	popH = 440;
	siteWindow = window.open(url,"MiniPopup", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
	siteWindow.focus();
}

function openNamedPopup(url, name) {
	popW = 400;
	popH = 200;
	siteWindow = window.open(url,name, "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
	siteWindow.focus();
}


/// Allows a link in a popup to target the opener ///
function targetOpener(mylink, closeme, closeonly) {
	if (! (window.focus && window.opener))return true;
		window.opener.focus();
	if (! closeonly)window.opener.location.href=mylink.href;
	if (closeme)window.close();
	return false;
}



/// Focuser, which sets the cursor to a particular form input ///
var formInUse = false;
function setFocus(form, input)
{
 if(!formInUse) {
  document.form.input.focus();
 }
}



/// Browser and platfom detection ///
function detectBrowser() {
	document.write("<b>Browser:</b> " + navigator.appName + " " + navigator.appVersion + "<br />");
	document.write("<b>Platform:</b> " + navigator.platform + "<br />");
}



/// Utility functions ///
function calcLeft(popW) {
	return Math.round(( screen.width - popW ) / 2);
}
function calcTop(popH) {
	return Math.round((( screen.height - popH ) / 2) - 10);
}



/// Confirm exit ///
function confirmExit() {
	input_box=confirm("Do you want to end your Logix session?");
	if (input_box == true) {
		top.location.href = ('/cgi-bin/yellowbox/logix/adminlogoff.exe');
		top.close();
	}
	else {
	}
}



// DHTML date validation script.
// Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
// Declaring valid date character, minimum year and maximum year

function isInteger(s){
	var i;
	for (i = 0; i < s.length; i++){   
	// Check that current character is number.
		var c = s.charAt(i);
		if (((c < "0") || (c > "9"))) return false;
		}
	// All characters are numbers.
	return true;
}

function isSignedInteger(s) {
	var i;
	var retVal = true;
	
  if (s.length > 0) {
    var c = s.charAt(0);
    if (c == "-") {
      if (s.length > 1) {
        retVal = isInteger(s.substring(1));
      } else {
        retVal = false;
      } 
    } else {
      retVal = isInteger(s);  
    }
  }
  
	return retVal;
}

function stripCharsInBag(s, bag){
	var i;
	var returnString = "";
	// Search through string's characters one by one.
	// If character is not in bag, append to returnString.
	for (i = 0; i < s.length; i++){   
		var c = s.charAt(i);
		if (bag.indexOf(c) == -1) returnString += c;
		}
	return returnString;
}

function daysInFebruary (year){
// February has 29 days in any year evenly divisible by four,
// EXCEPT for centurial years which are not also divisible by 400.
	//return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
	return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
}

function DaysArray(n) {
	for (var i = 1; i <= n; i++) {
		this[i] = 31;
		if (i==4 || i==6 || i==9 || i==11) {this[i] = 30;}
		if (i==2) {this[i] = 29;}
	} 
	return this;
}

function isDate(dtStr){
	var daysInMonth = DaysArray(12);
	var pos1=dtStr.indexOf(dtCh);
	var pos2=dtStr.indexOf(dtCh,pos1+1);
	var strMonth=dtStr.substring(0,pos1);
	var strDay=dtStr.substring(pos1+1,pos2);
	var strYear=dtStr.substring(pos2+1);
	strYr=strYear;
	if (strDay.charAt(0)=="0" && strDay.length>1) strDay=strDay.substring(1);
	if (strMonth.charAt(0)=="0" && strMonth.length>1) strMonth=strMonth.substring(1);
	for (var i = 1; i <= 3; i++) {
		if (strYr.charAt(0)=="0" && strYr.length>1) strYr=strYr.substring(1);
	}
	month=parseInt(strMonth);
	day=parseInt(strDay);
	year=parseInt(strYr);
	if (pos1==-1 || pos2==-1){
		alert("Please enter a date in the format mm/dd/yyyy");
		return false;
	}
	if (strMonth.length<1 || month<1 || month>12){
		alert("Please enter a valid month");
		return false;
	}
	if (strDay.length<1 || day<1 || day>31 || (month==2 && day>daysInFebruary(year)) || day > daysInMonth[month]){
		alert("Please enter a valid day");
		return false;
	}
	if (strYear.length != 4 || year==0 || year<minYear || year>maxYear){
		alert("Please enter a valid 4 digit year between "+minYear+" and "+maxYear);
		return false;
	}
	if (dtStr.indexOf(dtCh,pos2+1)!=-1 || isInteger(stripCharsInBag(dtStr, dtCh))==false){
		alert("Please enter a valid date");
		return false;
	}
	return true;
}

function ValidateOfferForm(){
	var dt3=document.mainform.form_ProdStartDate;
	var dt4=document.mainform.form_ProdEndDate;


    if(document.getElementById('form_Name').value == ""){
        alert("Please enter a name for this offer.")
        document.getElementById('form_Name').focus();
		return false;
    }

    if (isDate(dt3.value)==false){
		dt3.focus();
		return false;
	}
	else if (isDate(dt4.value)==false){
		dt4.focus();
		return false;
	}
	return true;
}

function ValidateOfferConTimeForm(){
	var dt1=document.mainform.start_hour;
	var dt2=document.mainform.end_hour;
	var dt3=document.mainform.start_minute;
	var dt4=document.mainform.end_minute;

	if(dt1.value < 0 || dt1.value > 23){
		alert("Please enter a valid Start Hour between 0 and 23");
		dt1.focus();
		return false;
	}
	else if(dt2.value < 0 || dt2.value > 23){
		alert("Please enter a valid End Hour between 0 and 23");
		dt2.focus();
		return false;
	}
	else if(dt3.value < 0 || dt3.value > 59){
		alert("Please enter a valid Start Minute between 0 and 59");
		dt3.focus();
		return false;
	}
	else if(dt4.value < 0 || dt4.value > 59){
		alert("Please enter a valid End Minute between 0 and 59");
		dt4.focus();
		return false;
	}
	return true;
}


function ValidateCPEOfferForm(){
	var dt1=document.form1.testingstart;
	var dt2=document.form1.testingend;
	var dt3=document.form1.eligibilitystart;
	var dt4=document.form1.eligibilityend;
    var dt5=document.form1.productionstart;
	var dt6=document.form1.productionend;


	if (isDate(dt1.value)==false){
		dt1.focus();
		return false;
	}
	else if (isDate(dt2.value)==false){
		dt2.focus();
		return false;
	}
	else if (isDate(dt3.value)==false){
		dt3.focus();
		return false;
	}
	else if (isDate(dt4.value)==false){
		dt4.focus();
		return false;
	}
	else if (isDate(dt5.value)==false){
		dt5.focus();
		return false;
	}
	else if (isDate(dt6.value)==false){
		dt6.focus();
		return false;
	}
	else if(document.getElementById('form_name').value == ""){
        alert("Please enter a name for this offer.")
        document.getElementById('form_name').focus();
		return false;
    }

	return true;
}



/// Resizer, used to collapse and expand box-class divs ///
function resizeDiv(divName, imgName, altTitle) {
    var divElem = document.getElementById(divName);
    var imgElem = document.getElementById(imgName);

    if (divElem != null) {
        divElem.style.display = (divElem.style.display == 'none') ? '' : 'none';
        imgElem.src = (divElem.style.display == 'none') ? '../images/arrowdown-off.png' : '../images/arrowup-off.png';
        imgElem.alt = (divElem.style.display == 'none') ? 'Show ' + altTitle : 'Hide ' + altTitle;
	imgElem.title = (divElem.style.display == 'none') ? 'Show ' + altTitle : 'Hide ' + altTitle;
    }
}

function handleResizeHover(bOver, divName, imgName) {
    var divElem = document.getElementById(divName);
    var imgElem = document.getElementById(imgName);
    var imgOn = "", imgOff = "";
    
    if (divElem != null && imgElem != null) {
        imgOn = (divElem.style.display == 'none') ? '../images/arrowdown-on.png' : '../images/arrowup-on.png';
        imgOff = (divElem.style.display == 'none') ? '../images/arrowdown-off.png' : '../images/arrowup-off.png';
        imgElem.src = (bOver) ? imgOn : imgOff;
    }            
}

function createCookie(name,value,days) {
	if (days) {
		var date = new Date();
		date.setTime(date.getTime()+(days*24*60*60*1000));
		var expires = "; expires="+date.toGMTString();
	}
	else var expires = "";
	document.cookie = name+"="+value+expires+"; path=/";
}

function readCookie(name) {
	var nameEQ = name + "=";
	var ca = document.cookie.split(';');
	for(var i=0;i < ca.length;i++) {
		var c = ca[i];
		while (c.charAt(0)==' ') c = c.substring(1,c.length);
		if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
	}
	return null;
}

function eraseCookie(name) {
	createCookie(name,"",-1);
}

function updateBoxesCookie(divElems, divVals) {
        var value = null, newValue = null, elem = null;
        var divClosed = false, valPresent = false;
        
	if (divElems != null && divVals != null) {
        value = readCookie("BoxesCollapsed");
        if (value == null) {
            value = 0;
        } else {
            value = parseInt(value);
        }
        newValue = value;
    
    	// loop through the pages div elements
        for (var i=0; i < divElems.length; i++) {
            elem = document.getElementById(divElems[i]);
            if (elem != null) {
                divClosed = (elem.style.display=='none');
                valPresent = (value & divVals[i]) > 0;
                if (divClosed && !valPresent) {
                    newValue += divVals[i];
                } else if(!divClosed && valPresent) {
                    newValue -= divVals[i];
                }
            }
        }
        createCookie("BoxesCollapsed", newValue, 9999);   
	}
}

function updatePageBoxes(divElems, divVals, divImages, value) {
	var valPresent = false;
	var elem = null, imgElem = null;

	if (divElems != null && divVals != null && value != null) {
    	// loop through the pages div elements
        for (var i=0; i < divElems.length; i++) {
            elem = document.getElementById(divElems[i]);
            if (elem != null) {
                valPresent = (value & divVals[i]) > 0;
                if (valPresent) {
                    elem.style.display='none';
                    imgElem = document.getElementById(divImages[i]);
                    if (imgElem != null) {
                        imgElem.src = '../images/arrowdown-off.png';
                    }
                }
            }
        }
    }
}

function handleNavAway(frm) {
    var saveChanges = false;
    
    if (IsFormChanged(frm)) {
        saveChanges = confirm("Changes were made. Do you wish to save?");
        if (saveChanges) {
           if (frm.elements['Save'] == null) {
                saveElem = document.createElement("input");
                saveElem.type = 'hidden';
                saveElem.id = 'Save';
                saveElem.name = 'Save';
                saveElem.value = 'save';
                frm.appendChild(saveElem);
            }
            handleAutoFormSubmit();
        }
    }
}

function IsFormChanged(frm) {
    var result = false;
    var output = '';

    for (var i=0, j=frm.elements.length; i<j; i++) {
        myType = frm.elements[i].type;
        if (myType == 'checkbox' || myType == 'radio') {
            if (frm.elements[i].checked != frm.elements[i].defaultChecked) {
                result = true;
                break;
            }
        }
        if (myType == 'hidden' || myType == 'password' || myType == 'text' || myType == 'textarea') {
            if (frm.elements[i].value != frm.elements[i].defaultValue) {
                result = true;
                break;
            }
        }
        if (myType == 'select-one' || myType == 'select-multiple') {
            for (var k=0, l=frm.elements[i].options.length; k<l; k++) {
                if (frm.elements[i].options[k].selected != frm.elements[i].options[k].defaultSelected) {
                    result = true;
                    break;
                }
            }
        }
    }

    return result;
}

function handleFormElements(frm, bDisabled) {
    if (frm != null) {
        for (var i=0, j=frm.elements.length; i<j; i++) {
            if (frm.elements[i] != null) { 
               frm.elements[i].disabled = bDisabled;
            }
        }    
    }    
}

function replaceAll(sString, sReplaceThis, sWithThis) { 
    if (sReplaceThis != "" && sReplaceThis != sWithThis && sString!= null && sString!= "" ) { 
        var counter = 0; 
        var start = 0; 
        var before = ""; 
        var after = ""; 

        while (counter<sString.length) { 
            start = sString.indexOf(sReplaceThis, counter); 
            if (start == -1) { 
                break; 
            } else { 
                before = sString.substr(0, start); 
                after = sString.substr(start + sReplaceThis.length, sString.length); 
                sString = before + sWithThis + after; 
                counter = before.length + sWithThis.length; 
            } 
        } 
    } 
    return sString; 
} 

function cleanRegExpString(str) {
    str = replaceAll(str, "\\", "\\\\");
    str = replaceAll(str, "(", "\\(");
    str = replaceAll(str, ")", "\\)");
    str = replaceAll(str, "[", "\\[");
    str = replaceAll(str, "]", "\\]");
    str = replaceAll(str, "{", "\\{");
    str = replaceAll(str, "}", "\\}");
    str = replaceAll(str, ".", "\\.");
    str = replaceAll(str, "$", "\\$");
    str = replaceAll(str, "*", "\\*");
    str = replaceAll(str, "+", "\\+");
    str = replaceAll(str, "|", "\\|");
    str = replaceAll(str, "_", "\_");
    
    return str;
}
