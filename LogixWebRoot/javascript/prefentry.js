// The values of the variables below are overridden in the codebehind by text in the user's language
var termValueSelectOperation   = '';
var termValueNoUnused          = '';
var termValueSelect            = '';
var termValueEnter             = '';
var termValueAlreadySelected   = '';
var termValueOutsideRange      = '';
var termCurrentDate            = '';
var termEnterDateFormat        = '';
var termEnterAnniversaryFormat = '';
var termEnterValidMonth        = '';
var termEnterValidDay          = '';
var termEnterValidYear         = '';
var termEnterValidValue        = '';

function addValueEntry(tierlevel) {
  var elemOp = document.getElementById('optype_tier' + tierlevel);
  var elemVal = document.getElementById('valueentry_tier' + tierlevel);
  var elemSlct = document.getElementById('values_tier' + tierlevel);
  var optText = '';
  var optVal = '';
  var valIsSelectBox = false;

  if (isDataType(7)) {
    addDateValueEntry(tierlevel);
  } else {
    valIsSelectBox = isSelectFormField(elemVal);

    if (elemOp == null || elemOp.selectedIndex == -1) {
      alert(termValueSelectOperation);
      return;
    } else if (elemVal == null || (valIsSelectBox && elemVal.options.length == 0)) {
      alert(termValueNoUnused);
      return;
    } else if (elemVal == null || (valIsSelectBox && elemVal.selectedIndex == -1)) {
      alert(termValueSelect);
      return;
    } else if (elemVal == null || (!valIsSelectBox && trimString(elemVal.value)=='')) {
      alert(termValueEnter);
      return;
    } else if (valueExists(tierlevel)) {
      alert(termValueAlreadySelected);
      return;
    } else if ((isDataType(8) || isDataType(2)) && !isValidMinMax(tierlevel)) {
      alert(termValueOutsideRange);
      return;
    } else {
      optText = elemOp.options[elemOp.selectedIndex].text + ' ';
      optVal = elemOp[elemOp.selectedIndex].value + '|';

      if (valIsSelectBox) {
        optVal += elemVal[elemVal.selectedIndex].value;
        optText += elemVal.options[elemVal.selectedIndex].text;
        elemVal.options[elemVal.selectedIndex] = null;
      } else {
        optVal += trimString(elemVal.value);
        optText += trimString(elemVal.value);
        elemVal.value = '';
      }

      // add empty values for date fields
      optVal += "|||";

      elemSlct[elemSlct.options.length] = new Option(optText, optVal);

    }
  }

  adjustOperators(tierlevel);
}

function addDateValueEntry(tierlevel) {
  var elemOp = document.getElementById('optype_tier' + tierlevel);
  var elemDateOp = document.getElementById('dateoptype_tier' + tierlevel);
  var elemVal = document.getElementById('valueentry_tier' + tierlevel);
  var elemValType = document.getElementById('valtype_tier' + tierlevel);
  var elemValMod = document.getElementById('valmod_tier' + tierlevel);
  var elemSlct = document.getElementById('values_tier' + tierlevel);
  var optVal = ''
  var optText = '';
  var retObj = new PrefDateProperties();
  var offset = 0;

  retObj = isValidPrefDateEntry(parseInt(elemDateOp.value), parseInt(elemValType.value), elemVal.value, elemValMod.value); 
  if (retObj.isValid) {
    optVal = elemOp.value + '|' + retObj.valueString + '|' + elemDateOp.value + '|' + elemValType.value + '|' + retObj.valueModifier;
    optText = elemDateOp.options[elemDateOp.selectedIndex].text + ' '; 
    optText += elemOp.options[elemOp.selectedIndex].text + ' ';
    if (parseInt(elemValType.value) == 1) {
      optText += '[' + termCurrentDate + ']'
      if (retObj.valueModifier != '') {
        offset = parseInt(retObj.valueModifier, 10);
        optText += (offset >= 0) ? ' + ' : ' - ';
        optText += Math.abs(offset);
      }
    } else  {
      optText += retObj.valueString;
    }

    elemSlct[elemSlct.options.length] = new Option(optText, optVal);

    elemVal.value = '';
    elemValMod.value = '';
  }
}

function isValidPrefDateEntry(dateOp, valType, dateValue, modifier) {
  var retObj = new PrefDateProperties();

  switch (dateOp) {
    case 1:  // exact date
      if (valType == 0) {
        retObj = (isValidPrefDate(dateValue, true));
        if (!retObj.isValid) { retObj.errorMessage = termEnterDateFormat; }
      } else {
        retObj = validateModifier(modifier);
      }
      break;
    case 2: // anniversary
      if (valType == 0) {
        retObj = (isValidPrefDate(dateValue, false));
        if (!retObj.isValid) { retObj.errorMessage = termEnterAnniversaryFormat; }
      } else {
        retObj = validateModifier(modifier);
      }
      break;
    case 3: // month
      if (valType == 0) {
        retObj.isValid = (dateValue != null && trimString(dateValue) != '' && isSignedInteger(dateValue));
        if (!retObj.isValid) { retObj.errorMessage = termEnterValidMonth; }
        if (retObj.isValid && (parseInt(dateValue, 10) < 1 || parseInt(dateValue, 10) > 12)) {
          retObj.isValid = false;
          retObj.errorMessage = termEnterValidMonth;
        } else if (retObj.isValid) {
          retObj.valueString = parseInt(dateValue, 10);
        }
      } else {
        retObj = validateModifier(modifier);
      }
      break;
    case 4: // day
      if (valType == 0) {
        retObj.isValid = (dateValue != null && trimString(dateValue) != '' && isSignedInteger(dateValue));
        if (!retObj.isValid) { retObj.errorMessage = termEnterValidDay; }
        if (retObj.isValid && (parseInt(dateValue,10) < 1 || parseInt(dateValue,10) > 31)) {
          retObj.isValid = false;
          retObj.errorMessage = termEnterValidDay;
        } else if (retObj.isValid) {
          retObj.valueString = parseInt(dateValue, 10);
        }
      } else {
        retObj = validateModifier(modifier);
      }
      break;
    case 5: // year
      if (valType == 0) {
        retObj.isValid = (dateValue != null && trimString(dateValue) != '' && isSignedInteger(dateValue));
        if (!retObj.isValid) { retObj.errorMessage = termEnterValidYear; }
        if (retObj.isValid && parseInt(dateValue, 10) < 0) {
          retObj.isValid = false;
          retObj.errorMessage = termEnterValidYear;
        } else if (retObj.isValid && dateValue.length != 4) {
          retObj.isValid = false;
          retObj.errorMessage = termEnterValidYear;
        } else if (retObj.isValid) {
          retObj.valueString = parseInt(dateValue, 10);
        }
      } else {
        retObj = validateModifier(modifier);
      }
      break;
  }

  if (!retObj.isValid && retObj.errorMessage != '') {
    alert(retObj.errorMessage);
  }

  return retObj;
}

function validateModifier(modifier) {
  var retObj = new PrefDateProperties();

  retObj.isValid = (trimString(modifier)!= '' && isSignedInteger(modifier));
  if (retObj.isValid) {
    retObj.valueModifier = parseInt(modifier, 10);
  } else {
    retObj.valueModifier = '';
    retObj.errorMessage = termEnterValidValue;
  }

  return retObj;
}

function removeValueEntry(tierlevel) {
  var elemSlct = document.getElementById('values_tier' + tierlevel);
  var elemVal = document.getElementById('valueentry_tier' + tierlevel);
  var valIsSelectBox = false;
  var optText = '';
  var optVal = '';

  valIsSelectBox = isSelectFormField(elemVal);

  if (elemSlct != null && elemSlct.selectedIndex > -1) {
    if (valIsSelectBox && elemVal != null) {
      // get the value from the selected
      optText = substringFrom(elemSlct.options[elemSlct.selectedIndex].text, ' ');
      optVal = parseTokenValue(elemSlct, elemSlct.selectedIndex, 1);
      
      elemVal[elemVal.options.length] = new Option(optText, optVal);
    }
    elemSlct.options[elemSlct.selectedIndex] = null;
  }

  adjustOperators(tierlevel);
}

function valueExists(tierlevel) {
  var exists = false;
  var elemSlct = document.getElementById('values_tier' + tierlevel);
  var elemVal = document.getElementById('valueentry_tier' + tierlevel);
  var entryVal = '';

  var valIsSelectBox = isSelectFormField(elemVal);

  // find the value of the entry
  if (valIsSelectBox) {
    entryVal = elemVal.options[elemVal.selectedIndex].value
  } else {
    entryVal = trimString(elemVal.value);
  }

  // loop through the existing entries and see if any match the entry value.
  if (elemSlct != null && elemVal != null) {
    for (var i = 0; i < elemSlct.options.length && !exists; i++) {
      exists = (substringFrom(elemSlct.options[i].value,'|') == entryVal);
    }
  }

  return exists;
}

function handleValueTypeChange(tierlevel) {
  var elemVal = document.getElementById('valueentry_tier' + tierlevel);
  var elemValType = document.getElementById('valtype_tier' + tierlevel);
  var elemValMod = document.getElementById('valmod_tier' + tierlevel);
  var elemTrVal = document.getElementById('trVal_tier' + tierlevel);
  var elemTrValMod = document.getElementById('trValMod_tier' + tierlevel);

  if (elemVal != null && elemValType != null && elemValMod != null && elemTrVal != null && elemTrValMod != null) {
    switch (parseInt(elemValType.options[elemValType.selectedIndex].value)) {
      case 0:
        elemTrVal.style.display = '';
        elemTrValMod.style.display = 'none';
        elemValMod.value = '';
        break;
      case 1:
        elemTrVal.style.display = 'none';
        elemTrValMod.style.display = '';
        elemVal.value = '';
        break;    
    } 
  }
}

function handleAndClick(tierlevel) {
  if (!isMultiValuedPref()) {
    removeEqualsOperator(tierlevel);
  }
}

function handleOrClick(tierlevel) {
  addEqualsOperator(tierlevel);
}

function isDataType(id) {
  var elem = document.getElementById('selected');
  var tokenVal = parseTokenValue(elem, 0, 1);

  return (tokenVal==id)
   //retVal =  (elem.options[0].value.indexOf("|" + id) > -1)
}

function isMultiValuedPref() {
  var elem = document.getElementById('selected');

  return (parseTokenValue(elem, 0, 2) == '1');  
}

function parseTokenValue(elem, selIndex, tokenIndex) {
  var retVal = '';

  if (elem != null && elem.options.length > 0) {
    var val = elem.options[selIndex].value;
    var tokens = val.split('|');
    if (tokens.length >= tokenIndex) {
      retVal = tokens[tokenIndex];
    }
  }

  return retVal;
}

function removeEqualsOperator(tierlevel) {
  var elemOp = document.getElementById('optype_tier' + tierlevel);

  if (elemOp != null) {
    for (var i = 0; i < elemOp.options.length; i++) {
      if (elemOp.options[i].value == '1') {
        elemOp.options[i] = null;
        break;
      }      
    }
  }
}

function addEqualsOperator(tierlevel) {
  var elemOp = document.getElementById('optype_tier' + tierlevel);
  var addFound = false;

  if (elemOp != null) {
    for (var i = 0; i < elemOp.options.length; i++) {
      if (elemOp.options[i].value == '1') {
        addFound = true;  
      }
    }
    if (!addFound) {
      elemOp.options[elemOp.options.length] = new Option('=', '1')
      sortOptions(elemOp);
    }
  }

}

function adjustOperators(tierlevel) {
  var elemSlct = document.getElementById('values_tier' + tierlevel);
  var elemAnd = document.getElementById('valcboand' + tierlevel);
  var elemOr = document.getElementById('valcboor' + tierlevel);
  var optVal = 0;
  var foundEquals = false;

  if (elemSlct != null && !isMultiValuedPref()) {
    for (var i = 0; i < elemSlct.options.length; i++) {
      optVal = parseTokenValue(elemSlct, i, 0);
      if (optVal == '1') {
        foundEquals = true;
        break;
      }
    }

    if (foundEquals && elemAnd.checked) {
      removeEqualsOperator(tierlevel);
      if (elemAnd != null) elemAnd.disabled = true;
      if (elemOr != null) elemOr.checked = true;
    } else {
      if (elemAnd.checked == false) {
        addEqualsOperator(tierlevel);
      }
      if (elemAnd != null) elemAnd.disabled = foundEquals;
    }
    
  }

}

// Start general use functions

function isSelectFormField(elem) {
  var retVal = false;

  if (elem != null) {
    retVal = (elem.nodeName.toLowerCase() == 'select');
  }

  return retVal;
}

// returns everything to the right of the first specified fromStr found.
function substringFrom(str, fromStr) {
  var pos = str.indexOf(fromStr);

  if (pos > -1) {
    return str.substring(pos+1);
  } else {
    return str;
  }
}

function isValidMinMax(tierlevel) {
  var retVal = false;
  var elem = document.getElementById('valueentry_tier' + tierlevel);
  var elemMins = new Array();
  var elemMaxs = new Array();
  var val = 0;

  elemMins = document.getElementsByName('minvalue');
  elemMaxs = document.getElementsByName('maxvalue');

  if (elem != null) {
    if (!isNaN(elem.value) && ((elem.value) % 1 === 0)) {
      val = parseInt(elem.value);
      if (elemMins != null && elemMaxs != null) {
        for (var i = 0; i < elemMins.length; i++) {
          if (val >= parseInt(elemMins[i].value) && val <= parseInt(elemMaxs[i].value)) {
            retVal = true;
            break;
          }
        }
      }
    }
  }

  return retVal;
}

function isValidPrefDate(date_string, includeYear) {
  var days = [0,31,29,31,30,31,30,31,31,30,31,30,31];
  var year, month, day, date_parts = null;
  var temp_arr = null;
  var retObj = new PrefDateProperties();
  var yearOK = true;

  if (date_string.indexOf("/") > -1) {
    // formats: mm/dd/yyyy or mm/dd
    date_parts = date_string.split('/');
  } else if (date_string.indexOf("-") > -1) {
    // format: yyyy-mm--dd or yyyy-mm
    temp_arr = date_string.split('-');
    if (temp_arr.length >= 2) {
      date_parts = new Array();
      date_parts[0] = temp_arr[1];
      date_parts[1] = temp_arr[2];
    }
    if (temp_arr.length ==3) {   
      date_parts[2] = temp_arr[0];
    }
  }

  if (date_parts != null) {
    if (date_parts.length==3 || (date_parts.length==2 && !includeYear)) {
      month = trimString(date_parts[0]);
      day = trimString(date_parts[1]);
      if (includeYear) {
        year = trimString(date_parts[2]);
      }

      if (isInteger(month) && isInteger(day) && (!includeYear || (includeYear && isInteger(year)))) {
        month = parseInt(month, 10);
        day = parseInt(day, 10);
        if (includeYear) {
          year = parseInt(year, 10);
          yearOK = (year >= 1000)
          days[2] = (isLeapYear(parseInt(year))) ? 29 : 28;
        }

        retObj.isValid = (month >= 1 && month <= 12 && day >= 1 && day <= days[month] && yearOK);

        if (retObj.isValid) {
          retObj.valueString = month + '/' + day;
          retObj.valueString += (includeYear) ? '/' + year : '';
        } else {
          retObj.valueString = '';
        }
      }
    }
  }

  return retObj;
}

function isLeapYear(year) {
  return (year % 4 != 0 ? false : 
      ( year % 100 != 0 ? true: 
      ( year % 1000 != 0 ? false : true)));
}

function removeAllSpaces(s) {
  return s.replace(/ /gi, '');
}

function deleteOption(object, index) {
  object.options[index] = null;
}

function addOption(object, text, value) {
  var defaultSelected = false;
  var selected = false;
  var optionName = new Option(text, value, defaultSelected, selected)
  object.options[object.length] = optionName;
  object.options[object.length - 1].selected = false;

}

function sortOptions(what) {
  var copyOption = new Array();
  for (var i = 0; i < what.options.length; i++)
    copyOption[i] = new Array(what[i].value, what[i].text);

  copyOption.sort(function (a, b) { return a[0] - b[0]; });

  for (var i = what.options.length - 1; i > -1; i--)
    deleteOption(what, i);

  for (var i = 0; i < copyOption.length; i++)
    addOption(what, copyOption[i][1], copyOption[i][0])
}

// customer JS objects

function PrefDateProperties() {
  this.isValid = false;
  this.valueString = '';
  this.valueModifier = '';
  this.errorMessage = '';
}
