/**
// version:7.3.1.138972.Official Build (SUSDAY10202)
version 1.5
December 4, 2005
Julian Robichaux -- http://www.nsftools.com/
[Modified for Logix]
*/

// default English phrases that can be overridden by the page.
var dayArray = new Array('Su,', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa');
var dateArray = new Array('0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
                          '10', '11', '12', '13', '14', '15', '16', '17', '18', '19',
                          '20', '21', '22', '23', '24', '25', '26', '27', '28', '29',
                          '30', '31');
var monthArray = new Array('January', 'February', 'March', 'April', 'May', 'June', 'July',
                           'August', 'September', 'October', 'November', 'December');
var calendarPhrase = 'Calendar';
var todayPhrase = 'Today';

// these variables define the date formatting we're expecting and outputting.
// If you want to use a different format by default, change the defaultDateSeparator
// and defaultDateFormat variables either here or on your HTML page.
var defaultDateSeparator = "/";  // common values would be "/" or "."
var defaultDateFormat = "mdy";     // valid values are "mdy", "dmy", and "ymd"
var dateSeparator = defaultDateSeparator;
var dateFormat = defaultDateFormat;
var CALENDAR_HEIGHT = 205;
var lastDayOfWeek = 6;
var firstDayOfWeek = 0;
var SysOpt158 = 0;

/**
This is the main function you'll call from the onClick event of a button.
Normally, you'll have something like this on your HTML page:

Start Date: <input name="StartDate">
<input type=button value="select" onclick="displayDatePicker('StartDate');">

That will cause the datepicker to be displayed beneath the StartDate field and
any date that is chosen will update the value of that field. If you'd rather have the
datepicker display beneath the button that was clicked, you can code the button
like this:

<input type=button value="select" onclick="displayDatePicker('StartDate', this);">

So, pretty much, the first argument (dateFieldName) is a string representing the
name of the field that will be modified if the user picks a date, and the second
argument (displayBelowThisObject) is optional and represents an actual node
on the HTML document that the datepicker should be displayed below.

In version 1.1 of this code, the dtFormat and dtSep variables were added, allowing
you to use a specific date format or date separator for a given call to this function.
Normally, you'll just want to set these defaults globally with the defaultDateSeparator
and defaultDateFormat variables, but it doesn't hurt anything to add them as optional
parameters here. An example of use is:

<input type=button value="select" onclick="displayDatePicker('StartDate', false, 'dmy', '.');">

This would display the datepicker beneath the StartDate field (because the
displayBelowThisObject parameter was false), and update the StartDate field with
the chosen value of the datepicker using a date format of dd.mm.yyyy
*/
function displayDatePicker(dateFieldName, evt, displayBelowThisObject, dtFormat, dtSep) {
    var targetDateField = document.getElementsByName(dateFieldName).item(0);

    if (targetDateField != null && targetDateField.disabled == true) {
        return;
    }

    // if we weren't told what node to display the datepicker beneath, just display it
    // beneath the date field we're updating
    if (!displayBelowThisObject)
        displayBelowThisObject = targetDateField;

    // if a date separator character was given, update the dateSeparator variable
    if (dtSep)
        dateSeparator = dtSep;
    else
        dateSeparator = defaultDateSeparator;

    // if a date format was given, update the dateFormat variable
    if (dtFormat)
        dateFormat = dtFormat;
    else
        dateFormat = defaultDateFormat;

    //var x = displayBelowThisObject.offsetLeft;
    //var y = displayBelowThisObject.offsetTop + displayBelowThisObject.offsetHeight ;

    var pos = findPos(displayBelowThisObject);

    mainOffsetX = 0;
    mainOffsetY = 0;
    if (isAncestorOfElement(displayBelowThisObject, "main")) {
        // adjust for the div id main
        var mainOffsetX = 0, mainOffsetY = 0;
        if (document.getElementById("main")) {
            var mainDivPos = findPos(document.getElementById("main"))
            mainOffsetX = mainDivPos[0];
            mainOffsetY = mainDivPos[1];
        }
    }
    x = pos[0] - mainOffsetX;
    y = pos[1] - mainOffsetY + displayBelowThisObject.offsetHeight;


    if (!IsDatePickerVisible(evt) && hasRoomAtTop(y)) {
        y -= CALENDAR_HEIGHT;
    }

    winH = 0;
    if (parseInt(navigator.appVersion) > 3) {
        if (navigator.appName == "Netscape") {
            winH = window.innerHeight;
        }
        if (navigator.appName.indexOf("Microsoft") != -1) {
            winH = document.body.offsetHeight;
        }
    }

    // deal with elements inside tables and such
    //  var parent = displayBelowThisObject;
    //  while (parent.offsetParent) {
    //     parent = parent.offsetParent;
    //     x += parent.offsetLeft;
    //     y += parent.offsetTop ;
    //  }


    drawDatePicker(targetDateField, x, y);
}

function findPos(obj) {
    var curleft = 0;
    var curtop = 0;

    if (obj.offsetParent) {
        curleft = obj.offsetLeft
        curtop = obj.offsetTop
        while (obj = obj.offsetParent) {
            curleft += obj.offsetLeft
            curtop += obj.offsetTop
        }
    }

    return [curleft, curtop];
}

// determines if the ancestorName string value is a parent (containing) element for the elem parameter
//         elem: object of what element you wish to check
// ancestorName: a string with the name of the id for the element used to determine whether it is an ancestor of elem. 
function isAncestorOfElement(elem, ancestorName) {
    var retVal = false;
    var ancestorElem = document.getElementById(ancestorName);

    if (elem != null && ancestorElem != null) {
        retVal = (elem.id == ancestorName);
        if (!retVal && elem.offsetParent) {
            while (elem = elem.offsetParent) {
                retVal = (elem.id == ancestorName);
                if (retVal) { break; }
            }
        }
    }

    return retVal;
}

function IsDatePickerVisible(evt) {
    var isVisible = true;

    var e = (window.event) ? window.event : evt;

    if (document.all) {
        isVisible = (parseInt(document.body.clientHeight) - (parseInt(e.clientY) + CALENDAR_HEIGHT)) >= 0;
    } else {
        isVisible = (parseInt(window.innerHeight) - (parseInt(e.pageY) + CALENDAR_HEIGHT)) >= 0;
    }

    return isVisible;
}

function hasRoomAtTop(y) {
    return (y - CALENDAR_HEIGHT) >= 0;
}

/**
Draw the datepicker object (which is just a table with calendar elements) at the
specified x and y coordinates, using the targetDateField object as the input tag
that will ultimately be populated with a date.

This function will normally be called by the displayDatePicker function.
*/
function drawDatePicker(targetDateField, x, y) {
    var dt = getFieldDate(targetDateField.value);

    // the datepicker table will be drawn inside of a <div> with an ID defined by the
    // global datePickerDivID variable. If such a div doesn't yet exist on the HTML
    // document we're working with, add one.
    if (!document.getElementById(datePickerDivID)) {
        // don't use innerHTML to update the body, because it can cause global variables
        // that are currently pointing to objects on the page to have bad references
        //document.body.innerHTML += "<div id='" + datePickerDivID + "' class='dpDiv'><\/div>";
        var newNode = document.createElement("div");
        newNode.setAttribute("id", datePickerDivID);
        newNode.setAttribute("class", "dpDiv");
        newNode.setAttribute("style", "visibility: hidden;");
        document.body.appendChild(newNode);
    }

    // move the datepicker div to the proper x,y coordinate and toggle the visiblity
    var pickerDiv = document.getElementById(datePickerDivID);
    pickerDiv.style.position = "absolute";
    pickerDiv.style.left = x + "px";
    pickerDiv.style.top = y + "px";
    pickerDiv.style.visibility = (pickerDiv.style.visibility == "visible" ? "hidden" : "visible");
    pickerDiv.style.display = (pickerDiv.style.display == "block" ? "none" : "block");
    pickerDiv.style.zIndex = 10000;

    // draw the datepicker table
    refreshDatePicker(targetDateField.name, dt.getFullYear(), dt.getMonth(), dt.getDate());

    // if this is IE6 then place an underlayment of an iframe to prevent select boxes from bleeding through the div
    var calFrame = document.getElementById('calendariframe');
    if (calFrame != null) {
        calFrame.style.position = "absolute"
        calFrame.style.left = x + "px";
        calFrame.style.width = parseInt(pickerDiv.clientWidth) + "px";
        calFrame.style.top = y + "px";
        calFrame.style.height = parseInt(pickerDiv.clientHeight) + "px";
        calFrame.style.visibility = (calFrame.style.visibility == "visible" ? "hidden" : "visible");
        calFrame.style.display = (calFrame.style.display == "block" ? "none" : "block");
        calFrame.style.zIndex = 5000;
    }

}


/**
This is the function that actually draws the datepicker calendar.
*/
function refreshDatePicker(dateFieldName, year, month, day) {
    // if no arguments are passed, use today's date; otherwise, month and year
    // are required (if a day is passed, it will be highlighted later)
    var thisDay = new Date();

    if ((month >= 0) && (year > 0)) {
        thisDay = new Date(year, month, 1);
    } else {
        day = thisDay.getDate();
        thisDay.setDate(1);
    }

    // the calendar will be drawn as a table
    // you can customize the table elements with a global CSS style sheet,
    // or by hardcoding style and formatting elements below
    var crlf = "\r\n";
    var TABLE = "<table cols='7' class='dpTable' summary='" + calendarPhrase + "'>" + crlf;
    var xTABLE = "<\/table>" + crlf;
    var TR = "<tr class='dpTR'>";
    var TR_title = "<tr class='dpTitleTR'>";
    var TR_days = "<tr class='dpDayTR'>";
    var TR_todaybutton = "<tr class='dpTodayButtonTR'>";
    var xTR = "<\/tr>" + crlf;
    var TD = "<td class='dpTD' onMouseOut='this.className=\"dpTD\";' onMouseOver=' this.className=\"dpTDHover\";' ";    // leave this tag open, because we'll be adding an onClick event
    var TD_title = "<td colspan='5' class='dpTitleTD'>";
    var TD_buttons = "<td class='dpButtonTD'>";
    var TD_todaybutton = "<td colspan='7' class='dpTodayButtonTD'>";
    var TD_days = "<td class='dpDayTD'>";
    var TD_selected = "<td class='dpDayHighlightTD' onMouseOut='this.className=\"dpDayHighlightTD\";' onMouseOver='this.className=\"dpTDHover\";' ";    // leave this tag open, because we'll be adding an onClick event
    var xTD = "<\/td>" + crlf;
    var DIV_title = "<div class='dpTitleText'>";
    var DIV_selected = "<div class='dpDayHighlight'>";
    var xDIV = "<\/div>";

    // start generating the code for the calendar table
    var html = TABLE;

    // this is the title bar, which displays the month and the buttons to
    // go back to a previous month or forward to the next month
    html += TR_title;
    html += TD_buttons + getButtonCode(dateFieldName, thisDay, -1, "&#9668;") + xTD;
    html += TD_title + DIV_title + monthArray[thisDay.getMonth()] + " " + thisDay.getFullYear() + xDIV + xTD;
    html += TD_buttons + getButtonCode(dateFieldName, thisDay, 1, "&#9658;") + xTD;
    html += xTR;

    // this is the row that indicates which day of the week we're on
    html += TR_days;
    for (i = firstDayOfWeek; i < dayArray.length; i++)
        html += TD_days + dayArray[i] + xTD;

    for (i = 0; i < firstDayOfWeek; i++)
        html += TD_days + dayArray[i] + xTD;

    html += xTR;

    // now we'll start populating the table with days of the month
    html += TR;

    // first, the leading blanks
    var blankCt = ((thisDay.getDay() - firstDayOfWeek + 7) % 7);
    for (i = 0; i < blankCt; i++)
        html += TD + "&nbsp;" + xTD;

    // now, the days of the month
    do {
        dayNum = thisDay.getDate();
        TD_onclick = " onclick=\"updateDateField('" + dateFieldName + "', '" + getDateString(thisDay) + "', '" + SysOpt158 + "');\">";

        if (dayNum == day)
            html += TD_selected + TD_onclick + DIV_selected + dateArray[dayNum] + xDIV + xTD;
        else
            html += TD + TD_onclick + dateArray[dayNum] + xTD;

        // if this is the last day of the week, start a new row
        if (thisDay.getDay() == lastDayOfWeek)
            html += xTR + TR;

        // increment the day
        thisDay.setDate(thisDay.getDate() + 1);
    } while (thisDay.getDate() > 1)

    // fill in any trailing blanks
    if (thisDay.getDay() > firstDayOfWeek) {
        for (i = lastDayOfWeek; i > thisDay.getDay(); i--)
            html += TD + "&nbsp;" + xTD;
    }
    html += xTR;

    // add a button to allow the user to easily return to today, or close the calendar
    var today = new Date();
    html += TR_todaybutton + TD_todaybutton;
    html += "<button class='dpTodayButton' id='todaybutton' onClick='refreshDatePicker(\"" + dateFieldName + "\");'>" + todayPhrase + "<\/button> ";
    html += xTD + xTR;

    // and finally, close the table
    html += xTABLE;

    document.getElementById(datePickerDivID).innerHTML = html;
}


/**
Convenience function for writing the code for the buttons that bring us back or forward
a month.
*/
function getButtonCode(dateFieldName, dateVal, adjust, label) {
    var newMonth = (dateVal.getMonth() + adjust) % 12;
    var newYear = dateVal.getFullYear() + parseInt((dateVal.getMonth() + adjust) / 12);
    if (newMonth < 0) {
        newMonth += 12;
        newYear += -1;
    }

    return "<button class='dpButton' onClick='refreshDatePicker(\"" + dateFieldName + "\", " + newYear + ", " + newMonth + ");'>" + label + "<\/button>";
}


/**
Convert a JavaScript Date object to a string, based on the dateFormat and dateSeparator
variables at the beginning of this script library.
*/
function getDateString(dateVal) {
    var dayString = "00" + dateVal.getDate();
    var monthString = "00" + (dateVal.getMonth() + 1);
    dayString = dayString.substring(dayString.length - 2);
    monthString = monthString.substring(monthString.length - 2);

    switch (dateFormat) {
        case "dmy":
            return dayString + dateSeparator + monthString + dateSeparator + dateVal.getFullYear();
        case "ymd":
            return dateVal.getFullYear() + dateSeparator + monthString + dateSeparator + dayString;
        case "mdy":
        default:
            return monthString + dateSeparator + dayString + dateSeparator + dateVal.getFullYear();
    }
}


/**
Convert a string to a JavaScript Date object.
*/
function getFieldDate(dateString) {
    var dateVal;
    var dArray;
    var d, m, y;
    var thisday = new Date();
    try {
        dArray = splitDateString(dateString);
        if (dArray) {
            switch (dateFormat) {
                case "dmy":
                    d = parseInt(dArray[0], 10);
                    m = parseInt(dArray[1], 10) - 1;
                    y = parseInt(dArray[2], 10);
                    break;
                case "ymd":
                    d = parseInt(dArray[2], 10);
                    m = parseInt(dArray[1], 10) - 1;
                    y = parseInt(dArray[0], 10);
                    break;
                case "mdy":
                default:
                    d = parseInt(dArray[1], 10);
                    m = parseInt(dArray[0], 10) - 1;
                    y = parseInt(dArray[2], 10);
                    break;
            }

            //  Year Must greater Than 1900 
            //  Month Range must be between 1 to 12
            if ((m >= 0 && m <= 12) && (y > 1900)) {
                dateVal = new Date(y, m, d);
            }
            else {
                d = thisday.getDate();
                dateVal = new Date(); 
                dateVal.setDate(d);
            }
            // dateVal = new Date(y, m, d);
        } else if (dateString) {
            dateVal = new Date(dateString);
        } else {
            dateVal = new Date();
        }
    } catch (e) {
        dateVal = new Date();
    }

    return dateVal;
}


/**
Try to split a date string into an array of elements, using common date separators.
If the date is split, an array is returned; otherwise, we just return false.
*/
function splitDateString(dateString) {
    var dArray;
    if (dateString.indexOf("/") >= 0)
        dArray = dateString.split("/");
    else if (dateString.indexOf(".") >= 0)
        dArray = dateString.split(".");
    else if (dateString.indexOf("-") >= 0)
        dArray = dateString.split("-");
    else if (dateString.indexOf("\\") >= 0)
        dArray = dateString.split("\\");
    else
        dArray = false;

    return dArray;
}


function updateDateField(dateFieldName, dateString, SysOpt158) {
    var calFrame = document.getElementById("calendariframe");
    var targetDateField = document.getElementsByName(dateFieldName).item(0);

    if (dateString) {
        targetDateField.value = dateString;

        // ensure that when the date field is changed that it fires an onchange event 
        fireEvent(targetDateField, "change");
    }

    var pickerDiv = document.getElementById(datePickerDivID);
    pickerDiv.style.visibility = "hidden";
    pickerDiv.style.display = "none";
    if (calFrame != null) {
        calFrame.style.visibility = "hidden";
    }

    targetDateField.focus();
    //Customer requested we introduce the datePickerClosed function to auto-populate other dates
    //We will call it datePickerAuto to avoid confusion 
    //if ((dateString) && (typeof(datePickerAuto) == "function"))
    //  datePickerAuto(targetDateField);
    if (SysOpt158 & ((dateString) && (typeof (datePickerClosed) == "function"))) {
        datePickerClosed(targetDateField);
    }
}

function autoDatePicker(dateFieldName, evt, extSysOpt158) {
    SysOpt158 = extSysOpt158;
    displayDatePicker(dateFieldName, evt);
}

function fireEvent(element, event) {
    if (document.createEventObject) {
        // dispatch for IE
        var evt = document.createEventObject();
        return element.fireEvent('on' + event, evt)
    }
    else {
        // dispatch for firefox + others
        var evt = document.createEvent("HTMLEvents");
        evt.initEvent(event, true, true); // event type,bubbling,cancelable
        return !element.dispatchEvent(evt);
    }
}
