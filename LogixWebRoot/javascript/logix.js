/// Logix 5 JavaScript functions ///
// version:7.3.1.138972.Official Build (SUSDAY10202)

var dtCh = "/";
var minYear = 1900;
var maxYear = 2100;

// text that needs overridden by the calling page with the text translated into the user's language.
var termSelectPrinter = 'Please select a printer before preview';
var termBrowser = 'Browser';
var termPlatform = 'Platform';
var termEndLogix = 'Do you want to end your Logix session?';
var termDateFormat = 'Please enter a date in the format mm/dd/yyyy';
var termValidMonth = 'Please enter a valid month';
var termValidDay = 'Please enter a valid day';
var termValidYear = 'Please enter a valid 4 digit year between {0} and {1}';
var termValidDate = 'Please enter a valid date';
var termEnterName = 'Please enter a name for this offer.';
var termValidStartHour = 'Please enter a valid Start Hour between 0 and 23';
var termValidEndHour = 'Please enter a valid End Hour between 0 and 23';
var termStartMinute = 'Please enter a valid Start Minute between 0 and 59';
var termEndMinute = 'Please enter a valid End Minute between 0 and 59';
var termPromptForSave = 'Changes were made. Do you wish to save?';
var termSave = 'Save';
var termMarkupTagWarning = 'Markup tags are not permitted as they represent a potential security risk to the server when submitted.';

/// Popup window launcher ///
function openPopup(url) {
    popW = 700;
    popH = 570;
    siteWindow = window.open(url, "Popup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    siteWindow.focus();
}

/// Popup window launcher ///
function openpatternprevPopup(url) {
    popW = 570;
    popH = 470;
    siteWindow = window.open(url, "Popup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    siteWindow.focus();
}

/// Popup window launcher ///
function openWidePopup(url) {
    popW = 780;
    popH = 570;
    siteWindow = window.open(url, "Popup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    siteWindow.focus();
}

/// Popup window launcher ///
function openExtraWidePopup(url) {
    popW = 955;
    popH = 570;
    siteWindow = window.open(url, "extrawidePopup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    siteWindow.focus();
}
function openPreviewPopup(url) {

    if (document.getElementById('printerselect').value == "999") {
        alert(termSelectPrinter);
    }
    else {

        popW = 700;
        popH = 522;
        siteWindow = window.open(url, "MiniPopup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
        siteWindow.focus();
    }
}

function openMiniPopup(url) {
    popW = 480;
    popH = 522;
    siteWindow = window.open(url, "MiniPopup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    siteWindow.focus();
}

function openReports(url) {
    popW = 750;
    popH = 540;
    siteWindow = window.open(url, "Popup", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    siteWindow.focus();
}



/// Allows a link in a popup to target the opener ///
function targetOpener(mylink, closeme, closeonly) {
    if (!(window.focus && window.opener)) return true;
    window.opener.focus();
    if (!closeonly) window.opener.location.href = mylink.href;
    if (closeme) window.close();
    return false;
}



/// Focuser, which sets the cursor to a particular form input ///
var formInUse = false;
function setFocus(form, input) {
    if (!formInUse) {
        document.form.input.focus();
    }
}



/// Browser and platfom detection ///
function detectBrowser() {
    document.write("<b>" + termBrowser + ":</b> " + navigator.appName + " " + navigator.appVersion + "<br />");
    document.write("<b>" + termPlatform + ":</b> " + navigator.platform + "<br />");
}



/// Function to allow only numeric values to be input ///
function NumberCheck(evt, src, allowDecimal) {
    var nkeycode = (window.event) ? window.event.keyCode : evt.which;
    var exceptionKeycodes = new Array(8, 9, 13, 16, 35, 36, 37, 38, 39, 40, 46);

    if ((nkeycode >= 48 && nkeycode <= 57) || (nkeycode >= 96 && nkeycode <= 105)) {
        return true;
    } else {
        for (var i = 0; i < exceptionKeycodes.length; i++) {
            if (nkeycode == exceptionKeycodes[i]) {
                return true;
            }
        }
        if (allowDecimal == true && (nkeycode == 110 || nkeycode == 190)) {
            if (src != null && src.value.indexOf(".") < 0) {
                return true;
            }
        }
        return false;
    }
}
function IsNumeric(value, allowdecimal, allownegative) {
    if (value == '')
        return true;
    var reg = '^';
    if (allownegative == true)
        reg = reg + '-{0,1}'
    if (allowdecimal == true)
        reg = reg + '\\d*\\.{0,1}';
    reg = reg + '\\d+$';
    var regExp = new RegExp(reg);
    return regExp.test(value)

}
function ForceNumericInput(This, AllowDot, AllowMinus) {

    var e = event

    if (arguments.length == 1) {
        var s = This.value;
        // if "-" exists then it better be the 1st character
        var i = s.lastIndexOf("-");
        if (i == -1)
            return;
        if (i != 0)
            This.value = s.substring(0, i) + s.substring(i + 1);
        return;
    }

    var code = e.keyCode;

    switch (code) {
        case 8:     // backspace
        case 9: //Tab
        case 13: //eNTER
        case 16: //Shift
        case 35: //End
        case 36: //Home
        case 37:    // left arrow
        case 38:
        case 39:    // right arrow
        case 40:
        case 46:    // delete
            return true;

    }
    if (e.ctrlKey == true && code == 86) {
        return true;
    }
    if (e.shiftKey) {
        return false;
    }

    if (code == 189)     // minus sign
    {
        if (AllowMinus == false) {
            return false;

        }
        if (This.value.indexOf("-") == 0) {
            return false;

        }

        // wait until the element has been updated to see if the minus is in the right spot
        var s = "ForceNumericInput(document.getElementById('" + This.id + "'))";
        setTimeout(s, 250);
        return;
    }
    if (AllowDot && (code == 110 || code == 190)) {
        if (This.value.indexOf(".") >= 0) {
            // don't allow more than one dot
            return false;

        }
        return true;

    }


    //    // allow character of between 0 and 9
    if ((code >= 48 && code <= 57) || (code >= 96 && code <= 105)) {
        return true;

    }
    return false;
}

/// Help button rollover ///
if (document.images) {
    helpbuttonup = new Image();
    helpbuttonup.src = "/images/help.png";
    helpbuttondown = new Image();
    helpbuttondown.src = "/images/help-on.png";
}
function buttondown(buttonname) {
    if (document.images) {
        document[buttonname].src = eval(buttonname + "down.src");
    }
}
function buttonup(buttonname) {
    if (document.images) {
        document[buttonname].src = eval(buttonname + "up.src");
    }
}
String.prototype.PadLeft = function (pad_length, pad_string) {
    var output = this
    while (output.length < pad_length) {
        output = pad_string + output;
    }
    return output;
}
String.prototype.PadRight = function (pad_length, pad_string) {
    var output = this;
    while (output.length < pad_length) {
        output = output + pad_string;
    }
    return output;
}


/// Utility functions ///
function calcLeft(popW) {
    return Math.round((screen.width - popW) / 2);
}
function calcTop(popH) {
    return Math.round(((screen.height - popH) / 2) - 10);
}



/// Confirm exit ///
function confirmExit() {
    input_box = confirm(termEndLogix);
    if (input_box == true) {
        top.location.href = ('/cgi-bin/yellowbox/logix/adminlogoff.exe');
        top.close();
    } else {
    }
}



/// Close the actions menu if user clicks elsewhere ///
function handlePageClick(e) {
    try {
        var el = (typeof event !== 'undefined') ? event.srcElement : e.target
        if (el != null && el.id != 'actions') {
            if (document.getElementById("actionsmenu") != null) {
                var bOpen = (document.getElementById("actionsmenu").style.visibility == 'visible');
                if (bOpen) {
                    toggleDropdown();
                }
            }
        }
    } catch (err) {
        //do nothing
    }
}
function handlePageClickNew(e, divid, buttonid) {
    var el = (typeof event !== 'undefined') ? event.srcElement : e.target
    if (el != null && el.id != buttonid) {
        if (document.getElementById(divid) != null) {
            var bOpen = (document.getElementById(divid).style.display == 'block');
            if (bOpen) {
                toggleDropdown();
            }
        }
    }
}


/// Select all checkboxes on a page ///
function checkAll(form) {
    with (document.form) {
        var d;
        d = document.getElementsByTagName("input");
        for (i = 0; i < d.length; i++) {
            if (d[i].type == "checkbox") {
                d[i].checked = true;
            }
        }
    }
}



/// Toggling notes display ///
function toggleNotes() {
    var divNotes = document.getElementById('notes');

    if (divNotes != null) {
        divNotes.style.display = (divNotes.style.display == 'none') ? '' : 'none';
    }
}



/// Toggling sections with the notes window ///
function toggleNotesInput() {
    var divNotes = document.getElementById('notes');
    var divDisplay = document.getElementById('notesscroll');
    var divNoteAdd = document.getElementById('noteadddiv');
    var divInput = document.getElementById('notesinput');

    if (divNotes != null) {
        divDisplay.style.height = (divDisplay.style.height == '165px') ? '295px' : '165px';
        divNoteAdd.style.display = (divNoteAdd.style.display == 'none') ? '' : 'none';
        divInput.style.display = (divInput.style.display == 'none') ? '' : 'none';
    }
}



/// Delete a note ///
function deleteNote(NoteID) {
    document.notesform.notedelete.value = 'notedelete';
    document.notesform.noteID.value = NoteID;
    document.notesform.submit();
}



// DHTML date validation script.
// Courtesy of SmartWebby.com (http://www.smartwebby.com/dhtml/)
// Declaring valid date character, minimum year and maximum year

function isInteger(s) {
    var i;
    for (i = 0; i < s.length; i++) {
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

// find the number of decimal places in a number (x)
// use dec_sep for internationalization
function decimalPlaces(x, dec_sep) {
    var tmp = new String();
    tmp = x;

    if (tmp.indexOf(dec_sep) > -1)
        return tmp.length - tmp.indexOf(dec_sep) - 1;
    else
        return 0;
}

function stripCharsInBag(s, bag) {
    var i;
    var returnString = "";
    // Search through string's characters one by one.
    // If character is not in bag, append to returnString.
    for (i = 0; i < s.length; i++) {
        var c = s.charAt(i);
        if (bag.indexOf(c) == -1) returnString += c;
    }
    return returnString;
}

function daysInFebruary(year) {
    // February has 29 days in any year evenly divisible by four,
    // EXCEPT for centurial years which are not also divisible by 400.
    //return (((year % 4 == 0) && ( (!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28 );
    return (((year % 4 == 0) && ((!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28);
}

function DaysArray(n) {
    for (var i = 1; i <= n; i++) {
        this[i] = 31;
        if (i == 4 || i == 6 || i == 9 || i == 11) { this[i] = 30; }
        if (i == 2) { this[i] = 29; }
    }
    return this;
}

function isDate(dtStr) {
    var daysInMonth = DaysArray(12);
    var pos1 = dtStr.indexOf(dtCh);
    var pos2 = dtStr.indexOf(dtCh, pos1 + 1);
    var strMonth = dtStr.substring(0, pos1);
    var strDay = dtStr.substring(pos1 + 1, pos2);
    var strYear = dtStr.substring(pos2 + 1);
    var tokenValues = [];

    strYr = strYear;
    if (strDay.charAt(0) == "0" && strDay.length > 1) strDay = strDay.substring(1);
    if (strMonth.charAt(0) == "0" && strMonth.length > 1) strMonth = strMonth.substring(1);
    for (var i = 1; i <= 3; i++) {
        if (strYr.charAt(0) == "0" && strYr.length > 1) strYr = strYr.substring(1);
    }
    month = parseInt(strMonth);
    day = parseInt(strDay);
    year = parseInt(strYr);
    if (pos1 == -1 || pos2 == -1) {
        alert(termDateFormat);
        return false;
    }
    if (strMonth.length < 1 || month < 1 || month > 12) {
        alert(termValidMonth);
        return false;
    }
    if (strDay.length < 1 || day < 1 || day > 31 || (month == 2 && day > daysInFebruary(year)) || day > daysInMonth[month]) {
        alert(termValidDay);
        return false;
    }
    if (strYear.length != 4 || year == 0 || year < minYear || year > maxYear) {
        tokenValues = [minYear, maxYear];
        alert(detokenizeString(termValidYear, tokenValues));
        return false;
    }
    if (dtStr.indexOf(dtCh, pos2 + 1) != -1 || isInteger(stripCharsInBag(dtStr, dtCh)) == false) {
        alert(termValidDate);
        return false;
    }
    return true;
}

// Validate Date According to a Format Provided by the User
// If No userFormat is provided, default US format will be used
// User will be notified by a JavaScript Alert about any issues with the input date
function IsValidLocalizedDate(dtStr, userFormat) {
    return CalculateValidLocalizedDate(dtStr, userFormat, true);
}

/// Validate Date According to a Format Provided by the User
/// If No userFormat is provided, default US format will be used
/// User will not be notified what type of error exists in the input date string
function IsValidLocalizedDateNoAlert(dtStr, userFormat) {
    return CalculateValidLocalizedDate(dtStr, userFormat, false);
}

function CalculateValidLocalizedDate(dtStr, userFormat, EnableAlert) {
    userFormat = userFormat || 'mm/dd/yyyy', // default format in case user does not provide userFormat parameter
    userFormat = userFormat.toLowerCase(),
    delimiter = /[^mdy]/.exec(userFormat)[0],
    pos1 = dtStr.indexOf(delimiter),
      pos2 = dtStr.indexOf(delimiter, pos1 + 1),
    theFormat = userFormat.split(delimiter),
    theDate = dtStr.split(delimiter),
    isDate = function (date, format) {
        var m, d, y;
        for (var i = 0, len = format.length; i < len; i++) {
            if (/m/.test(format[i])) m = date[i]
            if (/d/.test(format[i])) d = date[i]
            if (/y/.test(format[i])) y = date[i]
        }
        var month = parseInt(m, 10);
        var day = parseInt(d, 10);
        var year = parseInt(y);
        var daysInMonth = DaysArray(12);
        if (pos1 == -1 || pos2 == -1) {
            alert(termDateFormat);
            return false;
        }
        if (m.length < 1 || month < 1 || month > 12) {
            if (EnableAlert) alert(termValidMonth);
            return false;
        }
        if (d.length < 1 || day < 1 || day > 31 || (month == 2 && day > daysInFebruary(year)) || day > daysInMonth[month]) {
            if (EnableAlert) alert(termValidDay);
            return false;
        }
        if (y.length != 4 || year == 0 || year < minYear || year > maxYear) {
            tokenValues = [minYear, maxYear];
            if (EnableAlert) alert(detokenizeString(termValidYear, tokenValues));
            return false;
        }
        if (dtStr.indexOf(delimiter, pos2 + 1) != -1 || isInteger(stripCharsInBag(dtStr, delimiter)) == false) {
            if (EnableAlert) alert(termValidDate);
            return false;
        }
        return true;
    }
    return isDate(theDate, theFormat)
}

/// Convert Date from a Specific Locale to ISO-8601 date format (yyyy/MM/dd) to be used by Javascript for various calculations and Comparisons
/// Function will return a blank value if Input date does not matches the pattern provided in Input Date format
function ConvertToISODate(dtStr, userFormat) {
    userFormat = userFormat.toLowerCase(),
    delimiter = /[^mdy]/.exec(userFormat)[0],
    pos1 = dtStr.indexOf(delimiter),
      pos2 = dtStr.indexOf(delimiter, pos1 + 1),
    theFormat = userFormat.split(delimiter),
    theDate = dtStr.split(delimiter),
    isDate = function (date, format) {
        var m, d, y;
        for (var i = 0, len = format.length; i < len; i++) {
            if (/m/.test(format[i])) m = date[i]
            if (/d/.test(format[i])) d = date[i]
            if (/y/.test(format[i])) y = date[i]
        }
        var month = parseInt(m);
        var day = parseInt(d);
        var year = parseInt(y);
        var daysInMonth = DaysArray(12);
        if (pos1 == -1 || pos2 == -1) {
            return '';
        }
        if (m.length < 1 || month < 1 || month > 12) {
            return '';
        }
        if (d.length < 1 || day < 1 || day > 31 || (month == 2 && day > daysInFebruary(year)) || day > daysInMonth[month]) {
            return '';
        }
        if (y.length != 4 || year == 0 || year < minYear || year > maxYear) {
            tokenValues = [minYear, maxYear];
            return '';
        }
        if (dtStr.indexOf(delimiter, pos2 + 1) != -1 || isInteger(stripCharsInBag(dtStr, delimiter)) == false) {
            return '';
        }
        return year + '/' + ('0' + month).slice(-2) + '/' + ('0' + day).slice(-2);
    }
    return isDate(theDate, theFormat);
}

function isDateNoAlert(txtDate) {
    var objDate,  // date object initialized from the txtDate string
        mSeconds, // txtDate in milliseconds
        day,      // day
        month,    // month
        year;     // year
    // date length should be 10 characters (no more no less)
    if (txtDate.length !== 10) {
        return false;
    }
    // third and sixth character should be '/'
    if (txtDate.substring(2, 3) !== '/' || txtDate.substring(5, 6) !== '/') {
        return false;
    }
    // extract month, day and year from the txtDate (expected format is mm/dd/yyyy)
    // subtraction will cast variables to integer implicitly (needed
    // for !== comparing)
    month = txtDate.substring(0, 2) - 1; // because months in JS start from 0
    day = txtDate.substring(3, 5) - 0;
    year = txtDate.substring(6, 10) - 0;
    // test year range
    if (year < 1000 || year > 3000) {
        return false;
    }
    // convert txtDate to milliseconds
    mSeconds = (new Date(year, month, day)).getTime();
    // initialize Date() object from calculated milliseconds
    objDate = new Date();
    objDate.setTime(mSeconds);
    // compare input date and parts from Date() object
    // if difference exists then date isn't valid
    if (objDate.getFullYear() !== year ||
        objDate.getMonth() !== month ||
        objDate.getDate() !== day) {
        return false;
    }
    // otherwise return true
    return true;
}

function ValidateOfferForm(userFormat) {
    var dt3 = document.mainform.form_ProdStartDate;
    var dt4 = document.mainform.form_ProdEndDate;
    if (document.getElementById('form_Name').value == "") {
        alert(termEnterName)
        document.getElementById('form_Name').focus();
        return false;
    }
    if (IsValidLocalizedDate(dt3.value, userFormat) == false) {
        dt3.focus();
        return false;
    }
    else if (IsValidLocalizedDate(dt4.value, userFormat) == false) {
        dt4.focus();
        return false;
    }
    if ((document.mainform.form_DispStartDate != undefined && document.mainform.form_DispStartDate != null) || (document.mainform.form_DispEndDate != undefined && document.mainform.form_DispEndDate != null)) {
        var dt5 = document.mainform.form_DispStartDate;
        var dt6 = document.mainform.form_DispEndDate;
        if (dt5.value.length != 0 && dt6.value.length != 0) {
            if (IsValidLocalizedDate(dt5.value, userFormat) == false) {
                dt5.focus();
                return false;
            }
            else if (IsValidLocalizedDate(dt6.value, userFormat) == false) {
                dt6.focus();
                return false;
            }
        }
    }
    return true;
}

function ValidateOfferConTimeForm() {
    var dt1 = document.mainform.start_hour;
    var dt2 = document.mainform.end_hour;
    var dt3 = document.mainform.start_minute;
    var dt4 = document.mainform.end_minute;

    if (dt1.value < 0 || dt1.value > 23) {
        alert(termValidStartHour);
        dt1.focus();
        return false;
    }
    else if (dt2.value < 0 || dt2.value > 23) {
        alert(termValidEndHour);
        dt2.focus();
        return false;
    }
    else if (dt3.value < 0 || dt3.value > 59) {
        alert(termStartMinute);
        dt3.focus();
        return false;
    }
    else if (dt4.value < 0 || dt4.value > 59) {
        alert(termEndMinute);
        dt4.focus();
        return false;
    }
    return true;
}


function ValidateCPEOfferForm() {
    var dt1 = document.form1.testingstart;
    var dt2 = document.form1.testingend;
    var dt3 = document.form1.eligibilitystart;
    var dt4 = document.form1.eligibilityend;
    var dt5 = document.form1.productionstart;
    var dt6 = document.form1.productionend;


    if (isDate(dt1.value) == false) {
        dt1.focus();
        return false;
    }
    else if (isDate(dt2.value) == false) {
        dt2.focus();
        return false;
    }
    else if (isDate(dt3.value) == false) {
        dt3.focus();
        return false;
    }
    else if (isDate(dt4.value) == false) {
        dt4.focus();
        return false;
    }
    else if (isDate(dt5.value) == false) {
        dt5.focus();
        return false;
    }
    else if (isDate(dt6.value) == false) {
        dt6.focus();
        return false;
    }
    else if (document.getElementById('form_name').value == "") {
        alert(termEnterName)
        document.getElementById('form_name').focus();
        return false;
    }

    return true;
}



function BoxStateUpdate(BoxID, AdminUserID, BoxOpen) {
    var xmlHttpReq = false;
    var self = this;

    // Mozilla/Safari/Ie7
    if (window.XMLHttpRequest) {
        self.xmlHttpReq = new XMLHttpRequest();
    }
        // IE 6
    else if (window.ActiveXObject) {
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    self.xmlHttpReq.open('GET', '/logix/UE/BoxStatusUpdate.aspx?boxid=' + BoxID + '&targetuser=' + AdminUserID + '&boxopen=' + BoxOpen, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.send();
}


/// Multi-language input toggle
var mlLastID = "";
var mlLastIDLocked = false;

function hideMultiLanguageInput(mlClickedID, event) {
    var mlClickedDiv = document.getElementById('ml_' + mlClickedID);
    var mlClickedWrap = document.getElementById('mlwrap_' + mlClickedID);
    var mlClickedMG = document.getElementById('mg_' + mlClickedID);
    var mlLastDiv = document.getElementById('ml_' + mlLastID);
    var mlLastWrap = document.getElementById('mlwrap_' + mlLastID);
    var stdInput = document.getElementById(mlClickedID);
    var defaultInput = document.getElementById('ml_' + mlClickedID + '_default').firstChild;
    var debugElem = document.getElementById('debug');

    if (mlLastID != "" && mlLastID != mlClickedID) {
        // A different div was the last one activated, so make sure it's hidden before proceeding
        mlLastDiv.style.display = 'none';
        mlLastWrap.style.zIndex = '0';
    }

    // Hide the box
    if (event.type == "mouseout" && mlLastIDLocked == true && mlClickedID == mlLastID) {
        // Input's locked open, so do nothing.
    } else {
        $(mlClickedDiv).hide(250);
        mlLastIDLocked = false;
    }

    // Set the last-clicked ID
    stdInput.value = defaultInput.value;
    mlLastID = mlClickedID;
}

function showMultiLanguageInput(mlClickedID, event, lock) {
    var mlClickedDiv = document.getElementById('ml_' + mlClickedID);
    var mlClickedWrap = document.getElementById('mlwrap_' + mlClickedID);
    var mlClickedMG = document.getElementById('mg_' + mlClickedID);
    var mlLastDiv = document.getElementById('ml_' + mlLastID);
    var mlLastWrap = document.getElementById('mlwrap_' + mlLastID);
    var stdInput = document.getElementById(mlClickedID);
    var defaultInput = document.getElementById('ml_' + mlClickedID + '_default').firstChild;
    var debugElem = document.getElementById('debug');

    if (mlLastID != "" && mlLastID != mlClickedID) {
        // A different div was the last one activated, so make sure it's hidden before proceeding
        mlLastDiv.style.display = 'none';
        mlLastWrap.style.zIndex = '0';
    }

    // Show the box
    if (event.type == "mouseover" && mlLastIDLocked == true && mlClickedID == mlLastID) {
        // Input's locked open, so do nothing.
    } else {
        mlClickedWrap.style.zIndex = '500';
        $(mlClickedDiv).show(250);
        defaultInput.focus();
        if (event.type == "mouseover" && mlClickedID != mlLastID) {
            mlLastIDLocked = false;
        }
    }

    // If lock=true, lock open the inputs
    if (lock == true) {
        mlLastIDLocked = true;
    }

    // Set the last-clicked ID
    stdInput.value = defaultInput.value;
    mlLastID = mlClickedID;
}

function placereciptmessage(mlClickedID, event) {
    var stdInput = document.getElementById(mlClickedID);
    var defaultInput = document.getElementById(mlLastID);
    // Set the Dropdown value in to MLI text control
    defaultInput.value = stdInput.value;
}

/// Resizer, used to collapse and expand box-class divs - and save their state///
function resizeBox(divName, imgName, altTitle, BoxID, AdminUserID) {
    var divElem = document.getElementById(divName);
    var imgElem = document.getElementById(imgName);
    var parentElem = divElem.parentNode;

    if (divElem != null) {
        divElem.style.display = (divElem.style.display == 'none') ? '' : 'none';
        imgElem.src = (divElem.style.display == 'none') ? '/images/arrowdown-off.png' : '/images/arrowup-off.png';
        imgElem.alt = (divElem.style.display == 'none') ? 'Show ' + altTitle : 'Hide ' + altTitle;
        imgElem.title = (divElem.style.display == 'none') ? 'Show ' + altTitle : 'Hide ' + altTitle;
        if (divElem.style.display == 'none') {
            BoxStateUpdate(BoxID, AdminUserID, 0);
        }
        else {
            BoxStateUpdate(BoxID, AdminUserID, 1);
        }
    }
}

/// Resizer, used to collapse and expand box-class divs ///
function resizeDiv(divName, imgName, altTitle) {
    var divElem = document.getElementById(divName);
    var imgElem = document.getElementById(imgName);
    var parentElem = divElem.parentNode;

    if (divElem != null) {
        divElem.style.display = (divElem.style.display == 'none') ? '' : 'none';
        imgElem.src = (divElem.style.display == 'none') ? '/images/arrowdown-off.png' : '/images/arrowup-off.png';
        imgElem.alt = (divElem.style.display == 'none') ? 'Show ' + altTitle : 'Hide ' + altTitle;
        imgElem.title = (divElem.style.display == 'none') ? 'Show ' + altTitle : 'Hide ' + altTitle;
    }
}

function handleResizeHover(bOver, divName, imgName) {
    var divElem = document.getElementById(divName);
    var imgElem = document.getElementById(imgName);
    var imgOn = "", imgOff = "";

    if (divElem != null && imgElem != null) {
        imgOn = (divElem.style.display == 'none') ? '/images/arrowdown-on.png' : '/images/arrowup-on.png';
        imgOff = (divElem.style.display == 'none') ? '/images/arrowdown-off.png' : '/images/arrowup-off.png';
        imgElem.src = (bOver) ? imgOn : imgOff;
    }
}

function createCookie(name, value, days) {
    if (days) {
        var date = new Date();
        date.setTime(date.getTime() + (days * 24 * 60 * 60 * 1000));
        var expires = "; expires=" + date.toGMTString();
    }
    else var expires = "";
    document.cookie = name + "=" + value + expires + "; path=/";
}

function readCookie(name) {
    var nameEQ = name + "=";
    var ca = document.cookie.split(';');
    for (var i = 0; i < ca.length; i++) {
        var c = ca[i];
        while (c.charAt(0) == ' ') c = c.substring(1, c.length);
        if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length, c.length);
    }
    return null;
}

function eraseCookie(name) {
    createCookie(name, "", -1);
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
        for (var i = 0; i < divElems.length; i++) {
            elem = document.getElementById(divElems[i]);
            if (elem != null) {
                divClosed = (elem.style.display == 'none');
                valPresent = (value & divVals[i]) > 0;
                if (divClosed && !valPresent) {
                    newValue += divVals[i];
                } else if (!divClosed && valPresent) {
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
        for (var i = 0; i < divElems.length; i++) {
            elem = document.getElementById(divElems[i]);
            if (elem != null) {
                valPresent = (value & divVals[i]) > 0;
                if (valPresent) {
                    elem.style.display = 'none';
                    imgElem = document.getElementById(divImages[i]);
                    if (imgElem != null) {
                        imgElem.src = '/images/arrowdown-off.png';
                    }
                }
            }
        }
    }
}

function handleNavAway(frm) {
    var saveChanges = false;

    if (IsFormChanged(frm)) {
        saveChanges = confirm(termPromptForSave);
        if (saveChanges) {
            if (frm.elements['Save'] == null) {
                saveElem = document.createElement("input");
                saveElem.type = 'hidden';
                saveElem.id = 'Save';
                saveElem.name = 'Save';
                saveElem.value = termSave;
                frm.appendChild(saveElem);
            }
            handleAutoFormSubmit();
        }
    }
}

function IsFormChanged(frm) {
    var result = false;
    var output = '';
    var optSelected = false;

    if (frm != null && frm.elements != null) {
        for (var i = 0, j = frm.elements.length; i < j; i++) {
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
                // check if any option was selected when page was loaded, if none then don't check as this 
                // will cause a false result for unsaved change.
                for (var k = 0, l = frm.elements[i].options.length; k < l; k++) {
                    optSelected = frm.elements[i].options[k].defaultSelected;
                    if (optSelected == true) break;
                }

                if (optSelected) {
                    for (var k = 0, l = frm.elements[i].options.length; k < l; k++) {
                        if (frm.elements[i].options[k].selected != frm.elements[i].options[k].defaultSelected) {
                            result = true;
                            break;
                        }
                    }
                }
            }
        }
    }

    return result;
}

function handleFormElements(frm, bDisabled) {
    if (frm != null) {
        for (var i = 0, j = frm.elements.length; i < j; i++) {
            if (frm.elements[i] != null) {
                frm.elements[i].disabled = bDisabled;
            }
        }
    }
}

function replaceAll(sString, sReplaceThis, sWithThis) {
    if (sReplaceThis != "" && sReplaceThis != sWithThis && sString != null && sString != "") {
        var counter = 0;
        var start = 0;
        var before = "";
        var after = "";

        while (counter < sString.length) {
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
//similar function to cleanRegExpString() to clean '^' in the string.
function cleanSpecialChar(str) {
    str = cleanRegExpString(str);
    str = replaceAll(str, "^", "\\^");

    return str;
}
//this function needs to be replaced by cleanSpecialChar() when the fix has been done in all the pages.
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

//limits the character that can be entered in a textarea
function limitText(Element, limitNum) {
    if (Element != null) {
        if (Element.value.length > limitNum) {
            Element.value = Element.value.substring(0, limitNum);
        }
    }
}


function updateBanners(bChecked, allID) {
    var elem = document.getElementById("banner");

    if (elem != null) {
        if (bChecked) {
            for (var i = 0; i < elem.options.length; i++) {
                elem.options[i].selected = false;
            }
            elem.disabled = true;
        } else {
            elem.disabled = false;
        }
    }

    // remove other All Banners
    var i = 1;
    var elemAll = document.getElementById("allbannersid" + i);

    while (elemAll != null) {
        if (i != allID) {
            elemAll.checked = false;
        }
        i++;
        elemAll = document.getElementById("allbannersid" + i)
    }

}


function handleMarkupEntries(elem) {
    var retVal = true;
    var lineFeedCt = 0;

    if (elem != null) {
        var pos = elem.value.search("<[a-zA-Z\/]");
        if (pos > -1) {
            retVal = false;
            alert(termMarkupTagWarning);
            lineFeedCt = substr_count(elem.value, "\n", 0, pos);
            selectSomeText(elem, pos, pos + 1, lineFeedCt);
        }
    }
    return retVal;
}

function selectSomeText(element, begin, end, posAdjust) {
    try {
        if (element.setSelectionRange) {
            element.focus();
            element.setSelectionRange(begin, end);
        } else if (element.createTextRange) {
            var range = element.createTextRange();
            range.collapse(true);
            range.moveEnd('character', end - posAdjust);
            range.moveStart('character', begin - posAdjust);
            range.select();
        }
    } catch (err) {
        // do nothing, selection is optional
    }
}

function substr_count(string, substring, start, length) {
    var c = 0;

    try {
        if (start) { string = string.substr(start); }
        if (length) { string = string.substr(0, length); }

        for (var i = 0; i < string.length; i++) {
            if (substring == string.substr(i, substring.length))
                c++;
        }
    } catch (err) {
        // do nothing, optional
    }
    return c;
}

function isInt(sNum) {
    return (sNum != "" && !isNaN(sNum) && (sNum / 1) == parseInt(sNum));
}


/// Script to allow one to get elements by class name ///
/*
Developed by Robert Nyman, http://www.robertnyman.com
Code/licensing: http://code.google.com/p/getelementsbyclassname/
*/
var getElementsByClassName = function (className, tag, elm) {
    if (document.getElementsByClassName) {
        getElementsByClassName = function (className, tag, elm) {
            elm = elm || document;
            var elements = elm.getElementsByClassName(className),
				nodeName = (tag) ? new RegExp("\\b" + tag + "\\b", "i") : null,
				returnElements = [],
				current;
            for (var i = 0, il = elements.length; i < il; i += 1) {
                current = elements[i];
                if (!nodeName || nodeName.test(current.nodeName)) {
                    returnElements.push(current);
                }
            }
            return returnElements;
        };
    } else if (document.evaluate) {
        getElementsByClassName = function (className, tag, elm) {
            tag = tag || "*";
            elm = elm || document;
            var classes = className.split(" "),
				classesToCheck = "",
				xhtmlNamespace = "http://www.w3.org/1999/xhtml",
				namespaceResolver = (document.documentElement.namespaceURI === xhtmlNamespace) ? xhtmlNamespace : null,
				returnElements = [],
				elements,
				node;
            for (var j = 0, jl = classes.length; j < jl; j += 1) {
                classesToCheck += "[contains(concat(' ', @class, ' '), ' " + classes[j] + " ')]";
            }
            try {
                elements = document.evaluate(".//" + tag + classesToCheck, elm, namespaceResolver, 0, null);
            }
            catch (e) {
                elements = document.evaluate(".//" + tag + classesToCheck, elm, null, 0, null);
            }
            while ((node = elements.iterateNext())) {
                returnElements.push(node);
            }
            return returnElements;
        };
    } else {
        getElementsByClassName = function (className, tag, elm) {
            tag = tag || "*";
            elm = elm || document;
            var classes = className.split(" "),
				classesToCheck = [],
				elements = (tag === "*" && elm.all) ? elm.all : elm.getElementsByTagName(tag),
				current,
				returnElements = [],
				match;
            for (var k = 0, kl = classes.length; k < kl; k += 1) {
                classesToCheck.push(new RegExp("(^|\\s)" + classes[k] + "(\\s|$)"));
            }
            for (var l = 0, ll = elements.length; l < ll; l += 1) {
                current = elements[l];
                match = false;
                for (var m = 0, ml = classesToCheck.length; m < ml; m += 1) {
                    match = classesToCheck[m].test(current.className);
                    if (!match) {
                        break;
                    }
                }
                if (match) {
                    returnElements.push(current);
                }
            }
            return returnElements;
        };
    }
    return getElementsByClassName(className, tag, elm);
};

function trimString(str) {
    var str = str.replace(/^\s\s*/, ''),
    ws = /\s/,
  i = str.length;
    while (ws.test(str.charAt(--i)));
    return str.slice(0, i + 1);
}

String.prototype.trim = function () {
    return this.replace(/^\s+|\s+$/g, '');
}

String.prototype.DecodeSingleQuotes = function () {
    return this.replace(new RegExp('&#39;', 'g'), '\'');
}

// Override default alert function to decode single quotes before getting displayed on UI.
window.alert = (function (original) {
    return function (str) {
        original(str.DecodeSingleQuotes());
    }
})(window.alert)

function addCommas(nStr) {
    nStr += '';
    x = nStr.split('.');
    x1 = x[0];
    x2 = x.length > 1 ? '.' + x[1] : '';
    var rgx = /(\d+)(\d{3})/;
    while (rgx.test(x1)) {
        x1 = x1.replace(rgx, '$1' + ',' + '$2');
    }
    return x1 + x2;
}

// ------------------------------------------------------------------
// Utility functions for parsing in getDateFromFormat()
// ------------------------------------------------------------------
function _isInteger(val) {
    var digits = "1234567890";
    for (var i = 0; i < val.length; i++) {
        if (digits.indexOf(val.charAt(i)) == -1) { return false; }
    }
    return true;
}
function _getInt(str, i, minlength, maxlength) {
    for (var x = maxlength; x >= minlength; x--) {
        var token = str.substring(i, i + x);
        if (token.length < minlength) { return null; }
        if (_isInteger(token)) { return token; }
    }
    return null;
}

// ------------------------------------------------------------------
// getDateFromFormat( date_string , format_string )
//
// This function takes a date string and a format string. It matches
// If the date string matches the format string, it returns the 
// getTime() of the date. If it does not match, it returns 0.
// ------------------------------------------------------------------
function getDateFromFormat(val, format) {
    val = val + "";
    format = format + "";
    var i_val = 0;
    var i_format = 0;
    var c = "";
    var token = "";
    var token2 = "";
    var x, y;
    var now = new Date();
    var year = now.getYear();
    var month = now.getMonth() + 1;
    var date = 1;
    var hh = now.getHours();
    var mm = now.getMinutes();
    var ss = now.getSeconds();
    var ampm = "";

    while (i_format < format.length) {
        // Get next token from format string
        c = format.charAt(i_format);
        token = "";
        while ((format.charAt(i_format) == c) && (i_format < format.length)) {
            token += format.charAt(i_format++);
        }
        // Extract contents of value based on format token
        if (token == "yyyy" || token == "yy" || token == "y") {
            if (token == "yyyy") { x = 4; y = 4; }
            if (token == "yy") { x = 2; y = 2; }
            if (token == "y") { x = 2; y = 4; }
            year = _getInt(val, i_val, x, y);
            if (year == null) { return null; }
            i_val += year.length;
            if (year.length == 2) {
                if (year > 70) { year = 1900 + (year - 0); }
                else { year = 2000 + (year - 0); }
            }
        }
        else if (token == "MMM" || token == "NNN") {
            month = 0;
            for (var i = 0; i < MONTH_NAMES.length; i++) {
                var month_name = MONTH_NAMES[i];
                if (val.substring(i_val, i_val + month_name.length).toLowerCase() == month_name.toLowerCase()) {
                    if (token == "MMM" || (token == "NNN" && i > 11)) {
                        month = i + 1;
                        if (month > 12) { month -= 12; }
                        i_val += month_name.length;
                        break;
                    }
                }
            }
            if ((month < 1) || (month > 12)) { return null; }
        }
        else if (token == "EE" || token == "E") {
            for (var i = 0; i < DAY_NAMES.length; i++) {
                var day_name = DAY_NAMES[i];
                if (val.substring(i_val, i_val + day_name.length).toLowerCase() == day_name.toLowerCase()) {
                    i_val += day_name.length;
                    break;
                }
            }
        }
        else if (token == "MM" || token == "M") {
            month = _getInt(val, i_val, token.length, 2);
            if (month == null || (month < 1) || (month > 12)) { return null; }
            i_val += month.length;
        }
        else if (token == "dd" || token == "d") {
            date = _getInt(val, i_val, token.length, 2);
            if (date == null || (date < 1) || (date > 31)) { return null; }
            i_val += date.length;
        }
        else if (token == "hh" || token == "h") {
            hh = _getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 1) || (hh > 12)) { return null; }
            i_val += hh.length;
        }
        else if (token == "HH" || token == "H") {
            hh = _getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 0) || (hh > 23)) { return null; }
            i_val += hh.length;
        }
        else if (token == "KK" || token == "K") {
            hh = _getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 0) || (hh > 11)) { return null; }
            i_val += hh.length;
        }
        else if (token == "kk" || token == "k") {
            hh = _getInt(val, i_val, token.length, 2);
            if (hh == null || (hh < 1) || (hh > 24)) { return null; }
            i_val += hh.length; hh--;
        }
        else if (token == "mm" || token == "m") {
            mm = _getInt(val, i_val, token.length, 2);
            if (mm == null || (mm < 0) || (mm > 59)) { return null; }
            i_val += mm.length;
        }
        else if (token == "ss" || token == "s") {
            ss = _getInt(val, i_val, token.length, 2);
            if (ss == null || (ss < 0) || (ss > 59)) { return null; }
            i_val += ss.length;
        }
        else if (token == "a") {
            if (val.substring(i_val, i_val + 2).toLowerCase() == "am") { ampm = "AM"; }
            else if (val.substring(i_val, i_val + 2).toLowerCase() == "pm") { ampm = "PM"; }
            else { return null; }
            i_val += 2;
        }
        else {
            if (val.substring(i_val, i_val + token.length) != token) { return null; }
            else { i_val += token.length; }
        }
    }
    // If there are any trailing characters left in the value, it doesn't match
    if (i_val != val.length) { return null; }
    // Is date valid for month?
    if (month == 2) {
        // Check for leap year
        if (((year % 4 == 0) && (year % 100 != 0)) || (year % 400 == 0)) { // leap year
            if (date > 29) { return null; }
        }
        else { if (date > 28) { return null; } }
    }
    if ((month == 4) || (month == 6) || (month == 9) || (month == 11)) {
        if (date > 30) { return null; }
    }
    // Correct hours value
    if (hh < 12 && ampm == "PM") { hh = hh - 0 + 12; }
    else if (hh > 11 && ampm == "AM") { hh -= 12; }

    var newdate = new Date(year, month - 1, date, hh, mm, ss);
    return newdate;
}


// replaces the tokens {#} in the str parameter with their corresponding tokenValues based on the order of the index array.
function detokenizeString(str, tokenValues) {
    var i = 0;

    if (str != null && tokenValues != null && (tokenValues instanceof Array)) {
        for (var i = 0; i < tokenValues.length; i++) {
            str = str.replace('{' + i + '}', tokenValues[i])
        }
    } else {
        str = '';
    }

    return str;
}
//Moved the function from CPEOffer-rew-discount.aspx to here to be used widely
function findHTMLTags(obj, alertText) {
    var receipText;
    receipText = document.getElementById(obj).value;
    if ((receipText.indexOf("<") > -1) || (receipText.indexOf(">") > -1)) {
        alert(alertText);
        document.getElementById(obj).focus();
        document.getElementById(obj).value = "";
        return false
    }
    return true
}
//Set Cursor position on textbox
function SetCursorPosition(oField, iCaretPos) {
    // IE Support
    if (document.selection) {

        // Set focus on the element
        oField.focus();

        // Create empty selection range
        var oSel = document.selection.createRange();

        // Move selection start and end to 0 position
        oSel.moveStart('character', -oField.value.length);

        // Move selection start and end to desired position
        oSel.moveStart('character', iCaretPos);
        oSel.moveEnd('character', 0);
        oSel.select();
    }

        // Firefox support
    else if (oField.selectionStart || oField.selectionStart == '0') {
        oField.selectionStart = iCaretPos;
        oField.selectionEnd = iCaretPos;
        oField.focus();
    }
}
//String Format function we can use like this : - confmsg.format([var1]); where var1 is value at {0} position
String.prototype.format = function (args) {
    var str = this;
    return str.replace(String.prototype.format.regex, function (item) {
        var intVal = parseInt(item.substring(1, item.length - 1));
        var replace;
        if (intVal >= 0) {
            replace = args[intVal];
        } else if (intVal === -1) {
            replace = "{";
        } else if (intVal === -2) {
            replace = "}";
        } else {
            replace = "";
        }
        return replace;
    });
};
String.prototype.format.regex = new RegExp("{-?[0-9]+}", "g");

//////---------------------------- Show and hide functions for the new MultilanguagePopup control ---------------
var openLocked = false;
var mlLastDivOpened;
var triggerringEventType;
function ShowMLI(clickedId, clickedDivId, clickedDivWrapperId, event) {

    var clickedInput = document.getElementById(clickedId);
    var clickedDiv = document.getElementById(clickedDivId);
    var clickedDivWrapper = document.getElementById(clickedDivWrapperId);
    var defaultSpanId = "#" + clickedId + "_default";
    if (mlLastDivOpened != null && mlLastDivOpened != clickedDiv) {
        mlLastDivOpened.style.display = 'none';
        mlLastDivOpened.style.zIndex = '0';
    }
    if (openLocked == false) {
        clickedDivWrapper.style.zIndex = 500;
        $(clickedDiv).show(250);

        triggerringEventType = event.type;

        mlLastDivOpened = clickedDiv;

        $("#" + clickedDivId).children().each(
        function () {
            if ($(this).attr('id')) {
                var defInput = $(this).attr('id').indexOf('_default');
                if (defInput > 0)
                    $(this).focus();
                //alert("Value:" + $(this).val());
            }
        });
    }
}
function HideMLI(clickedId, clickedDivId, event) {
    var clickedInput = document.getElementById(clickedId);
    var clickedDiv = document.getElementById(clickedDivId);
    $("#" + clickedDivId).children().each(
    function () {
        if ($(this).attr('id')) {
            var defInput = $(this).attr('id').indexOf('_default');
            if (defInput > 0)
                clickedInput.value = $(this).val();
        }
    });
    //alert(openLocked + ":" + event.type);
    if (triggerringEventType == "click" && event.type == "mouseout") {
        ;
    }
    else if (openLocked && event.type == "mouseout") {
            ;
    }
    else {
        $(clickedDiv).hide(250);
    }
    mlLastDivOpened = clickedDiv;
    openLocked = false;

    return false;
}

//AMS-2223: This function is usefull for loading a listbox with more items
function LoadItemsOnScroll(elementId, ajaxURL, inputJsonData) {
    var jqueryId = '#' + elementId;
    if (($(jqueryId).innerHeight() + $(jqueryId).scrollTop()) >= $(jqueryId)[0].scrollHeight && loadInProgress == false) {
        //var $notificationDiv =  $('#divNotification');
        //var $notificationLabel = $('#lblAjaxNotification');
        $.ajax({
            type: "POST",
            url: ajaxURL,
            data: inputJsonData,
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            beforeSend: BeforeSendSetup
        })
        .done(function (response) {
            OnLoadSuccess(response);
        })
        .fail(function (response) {
            OnLoadError(response);
        });
    }
}

//////---------------------------- End: Show and hide functions for the new MultilanguagePopup control ---------------
//Function to validate email
function ValidateEmail(email) {
    var emailRegex = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{1,}))$/;
    if (emailRegex.test(email)) {
        return true;
    }
    else {
        return false;
    }
}