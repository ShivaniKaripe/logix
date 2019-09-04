// The value of the variable below are overridden in the codebehind by text in the user's language
var termInvalidDays   = '';

var optTexts;
var optVals;

function storeAllOperations(tierlevel) {
  var elemOp = document.getElementById('optype_tier' + tierlevel);

  optTexts = new Array();
  optVals = new Array();

  for (var i = 0; i < elemOp.length; i++) {
    optTexts[i] = elemOp.options[i].text;
    optVals[i] = elemOp.options[i].value;
  }
}

function syncOperatorsForDate(opTypeIDs, tierlevel) {
  var elemOp = document.getElementById('optype_tier' + tierlevel);
  var selVal = '';

  // save the currently selected one in a variable 
  // then clear it out and add the specified ones (opTypeIDs) and then select the previously selected one if it still is in the list
  selVal = elemOp.options[elemOp.selectedIndex].value;
  elemOp.options.length = 0;
  for (var i = 0; i < opTypeIDs.length; i++) {
    for (var j = 0; j < optVals.length; j++) {
      if (optVals[j] == opTypeIDs[i]) {
        elemOp.options[elemOp.options.length] = new Option(optTexts[j], optVals[j]);
        if (optVals[j] == selVal) { elemOp.options[elemOp.options.length - 1].selected = true; }
      }
    }
  }
}

function refreshDateValueBox(tierlevel) {
  var elemDateOp = document.getElementById('dateoptype_tier' + tierlevel);
  var elemOp = document.getElementById('optype_tier' + tierlevel);
  var elemValType = document.getElementById('valtype_tier' + tierlevel);
  var elemTrVal = document.getElementById('trVal_tier' + tierlevel);
  var elemValMod = document.getElementById('valmod_tier' + tierlevel);
  var elemTrValMod = document.getElementById('trValMod_tier' + tierlevel);
  var elemTrRange = document.getElementById('trRange_tier' + tierlevel);
  var opts;

  switch (parseInt(elemDateOp.options[elemDateOp.selectedIndex].value)) {
    case 1: // Exact Date
      // Operations: = != >= <= 
      opts = new Array(1, 2, 3, 4);
      syncOperatorsForDate(opts, tierlevel);
      // Value: shown only if Use Specified Date
      elemTrVal.style.display = (elemValType.value == '0') ? '' : 'none';
      // Value Modifier: Shown only if use current date
      elemTrValMod.style.display = (elemValType.value == '1') ? '' : 'none';
      // Range: hidden
      elemTrRange.style.display = 'none';
      break;
    case 2: // Anniversary
      // Operations: = != DateRange 
      opts = new Array(1, 2, 5);
      syncOperatorsForDate(opts, tierlevel);
      // Value: shown only if Use Specified Date
      elemTrVal.style.display = (elemValType.value == '0') ? '' : 'none';
      // Value Modifier: hidden
      elemTrValMod.style.display = 'none';
      // Range: shown only when DateRange Operator selected
      elemTrRange.style.display = (elemOp.value == '5') ? '' : 'none';
      break;
    default:  // month, day, year
      // Operations: = != >= <= 
      opts = new Array(1, 2, 3, 4);
      syncOperatorsForDate(opts, tierlevel);
      // Value: shown only if Use Specified Date
      elemTrVal.style.display = (elemValType.value == '0') ? '' : 'none';
      // Value Modifier: hidden
      elemTrValMod.style.display = 'none';
      // Range: hidden
      elemTrRange.style.display = 'none';
  }
    
}

function addDateValueEntry(tierlevel) {
  var elemOp = document.getElementById('optype_tier' + tierlevel);
  var elemDateOp = document.getElementById('dateoptype_tier' + tierlevel);
  var elemVal = document.getElementById('valueentry_tier' + tierlevel);
  var elemValType = document.getElementById('valtype_tier' + tierlevel);
  var elemValMod = document.getElementById('valmod_tier' + tierlevel);
  var elemSlct = document.getElementById('values_tier' + tierlevel);
  var elemBefore = document.getElementById('daysbefore_tier' + tierlevel);
  var elemAfter = document.getElementById('daysafter_tier' + tierlevel);
  var optVal = ''
  var optText = '';
  var retObj = new PrefDateProperties();
  var offset = 0;
  var daysBefore = 0;
  var daysAfter = 0;

  retObj = isValidPrefDateEntry(parseInt(elemDateOp.value), parseInt(elemValType.value), elemVal.value, elemValMod.value);
  
  if (!isValidDaysCount(parseInt(elemDateOp.value), tierlevel)) {
    alert(termInvalidDays);
  } else if (retObj.isValid) {
    // anniversary with date range operator selected is the only use for days before and after.
    if (parseInt(elemDateOp.value) == 2 && parseInt(elemOp.value)==5) {
     if (elemBefore.value != '' && isInteger(elemBefore.value)) { daysBefore = parseInt(elemBefore.value); }
     if (elemAfter.value != '' && isInteger(elemAfter.value)) { daysAfter = parseInt(elemAfter.value); }
    }

    optVal = elemOp.value + '|' + retObj.valueString + '|' + elemDateOp.value + '|' + elemValType.value + '|' + retObj.valueModifier + '|' + daysBefore + '|' + daysAfter;
    optText = elemDateOp.options[elemDateOp.selectedIndex].text + ' ';
    optText += elemOp.options[elemOp.selectedIndex].text + ' ';
    if (parseInt(elemValType.value) == 1) {
      optText += '[' + termCurrentDate + ']'
      if (retObj.valueModifier != '') {
        offset = parseInt(retObj.valueModifier, 10);
        optText += (offset >= 0) ? ' + ' : ' - ';
        optText += Math.abs(offset);
      }
    } else {
      optText += retObj.valueString;
    }

    if (daysBefore > 0 && daysAfter > 0) {
      optText += ' (-' + daysBefore + ' / +' + daysAfter + ')';
    } else if (daysBefore > 0 && daysAfter == 0) {
      optText += ' (-' + daysBefore + ')';
    } else if (daysBefore == 0 && daysAfter > 0) {
      optText += ' (+' + daysAfter + ')';
    }

    elemSlct[elemSlct.options.length] = new Option(optText, optVal);

    elemVal.value = '';
    elemValMod.value = '';
  }
}

function isValidDaysCount(dateOp, tierlevel) {
  var elemBefore = document.getElementById('daysbefore_tier' + tierlevel);
  var elemAfter = document.getElementById('daysafter_tier' + tierlevel);
  var validCt = true;

  if (elemBefore != null && elemAfter != null && dateOp == 2) {
    if (isNaN(elemBefore.value) || isNaN(elemAfter.value)) {
      validCt = false;
    } else if (parseInt(elemBefore.value) < 0 || parseInt(elemAfter.validCt) <0) {
      validCt = false;
    }
  }

  return validCt;
}

function isValidPrefDateEntry(dateOp, valType, dateValue, modifier) {
  var retObj = new PrefDateProperties();

  retObj.isValid = true;
  retObj.errorMessage = '';

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
      }
      break;
  }

  if (!retObj.isValid && retObj.errorMessage != '') {
    alert(retObj.errorMessage);
  }

  return retObj;
}

