// JS QuickTags version 1.2
// version:7.3.1.138972.Official Build (SUSDAY10202)
//
// Copyright (c) 2002-2005 Alex King
// http://www.alexking.org/
//
// Licensed under the LGPL license
// http://www.gnu.org/copyleft/lesser.html
//
// This JavaScript will insert the tags below at the cursor position in IE and 
// Gecko-based browsers (Mozilla, Camino, Firefox, Netscape). For browsers that 
// do not support inserting at the cursor position (Safari, OmniWeb) it appends
// the tags to the end of the content.
//
// The variable 'edCanvas' must be defined as the <textarea> element you want 
// to be editing in. See the accompanying 'index.html' page for an example.

var edButtons = new Array();
var edOpenTags = new Array();

function edButton(id, display, alt, tagStart, tagEnd, access, open, set) {
	this.id = id;			// used to name the toolbar button
	this.display = display;		// label on button
	this.alt = alt;			// alt/title text for the button
	this.tagStart = tagStart; 	// open tag
	this.tagEnd = tagEnd;		// close tag
	this.access = access;		// set to -1 if tag does not need to be closed
	this.open = open;		// set to -1 if tag does not need to be closed
	this.set = set;			// used to group similar sets of buttons
}

edButtons.push(
	new edButton(
		'ed_bold'
		,'B'
		,'Bold'
		,'|B|'
		,''
		,-1
		,-1
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_italic'
		,'I'
		,'Italic'
		,'|I|'
		,''
		,-1
		,-1
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_ul'
		,'U'
		,'Underline'
		,'|U|'
		,''
		,'u'
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_high'
		,'High'
		,'Double-high text'
		,'|HIGH|'
		,''
		,-1
		,-1
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_wide'
		,'Wide'
		,'Double-wide text'
		,'|WIDE|'
		,''
		,-1
		,-1
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_center'
		,'Center'
		,'Centered text'
		,'|CENTER|'
		,''
		,-1
		,-1
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_right'
		,'Right'
		,'Right-justified text'
		,'|RIGHT|'
		,''
		,-1
		,-1
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_center'
		,'Normal'
		,'Returns all print characteristics to normal'
		,'|NORMAL|'
		,''
		,-1
		,-1
		,1
	)
);

edButtons.push(
	new edButton(
		'ed_logo'
		,'Logo'
		,'Pre-stored logo'
		,'|LOGO|'
		,''
		,-1
		,-1
		,2
	)
);

edButtons.push(
	new edButton(
		'ed_line'
		,'Line'
		,'Horizontal rule (line)'
		,'|LINE|\n'
		,''
		,-1
		,-1
		,2
	)
);

edButtons.push(
	new edButton(
		'ed_stars'
		,'Stars'
		,'Horizontal rule (stars)'
		,'|STARS|\n'
		,''
		,-1
		,-1
		,2
	)
);

edButtons.push(
	new edButton(
		'ed_cut100'
		,'Cut'
		,'Cuts the paper'
		,'|CUT|\n'
		,''
		,-1
		,-1
		,2
	)
);

edButtons.push(
	new edButton(
		'ed_cut85'
		,'Cut85'
		,'Cuts the paper by 85%'
		,'|CUT85|\n'
		,''
		,-1
		,-1
		,2
	)
);

edButtons.push(
	new edButton(
		'ed_upca'
		,'UPCA'
		,'UPCA barcode'
		,''
		,''
		,'m'
		,-1
		,2
	)
); // special case

edButtons.push(
	new edButton(
		'ed_ean13'
		,'EAN13'
		,'EAN13 barcode'
		,''
		,''
		,'m'
		,-1
		,2
	)
); // special case

edButtons.push(
	new edButton(
		'ed_code39'
		,'Code39'
		,'Code39 barcode'
		,''
		,''
		,'m'
		,-1
		,2
	)
); // special case

edButtons.push(
	new edButton(
		'ed_customerid'
		,'Customer ID'
		,'Customer ID'
		,'|CUSTOMERID|'
		,''
		,-1
		,-1
		,3
	)
);

edButtons.push(
	new edButton(
		'ed_firstname'
		,'Customer first name'
		,'Customer first name'
		,'|FIRSTNAME|'
		,''
		,-1
		,-1
		,3
	)
);

edButtons.push(
	new edButton(
		'ed_lastname'
		,'Customer last name'
		,'Customer last name'
		,'|LASTNAME|'
		,''
		,-1
		,-1
		,3
	)
);

edButtons.push(
	new edButton(
		'ed_netamt'
		,'Net#'
		,'Net amount (count)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case

edButtons.push(
	new edButton(
		'ed_initialamt'
		,'Initial#'
		,'Initial amount (count)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case

edButtons.push(
	new edButton(
		'ed_earnedamt'
		,'Earned#'
		,'Earned amount (count)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case

edButtons.push(
	new edButton(
		'ed_redeemedamt'
		,'Redeemed#'
		,'Redeemed amount (count)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case

edButtons.push(
	new edButton(
		'ed_netdol'
		,'Net$'
		,'Net amount (dollars)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case

edButtons.push(
	new edButton(
		'ed_initialdol'
		,'Initial$'
		,'Initial amount (dollars)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case

edButtons.push(
	new edButton(
		'ed_earneddol'
		,'Earned$'
		,'Earned amount (dollars)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case

edButtons.push(
	new edButton(
		'ed_redeemeddol'
		,'Redeemed$'
		,'Redeemed amount (dollars)'
		,''
		,''
		,'m'
		,-1
		,4
	)
); // special case


//
//
//
// Functions

function edShowButton(button) {
	if (button.access) {
		var accesskey = ' accesskey = "' + button.access + '"'
	}
	else {
		var accesskey = '';
	}
	switch (button.id) {
		case 'ed_upca':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertUPCA(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_ean13':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertEAN13(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_code39':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertCode39(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_netamt':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertNetAmt(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_initialamt':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertInitialAmt(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_earnedamt':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertEarnedAmt(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_redeemedamt':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertRedeemedAmt(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_net$':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertNet$(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_initial$':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertInitial$(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_earned$':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertEarned$(edCanvas);" value="' + button.display + '" />');
			break;
		case 'ed_redeemed$':
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertRedeemed$(edCanvas);" value="' + button.display + '" />');
			break;
		default:
			document.write('<input type="button" id="' + button.id + '" ' + accesskey + ' class="ed_button" title="' + button.alt + '" onclick="edInsertTag(edCanvas, ' + i + ');" value="' + button.display + '"  />');
			break;
	}
}

function edAddTag(button) {
	if (edButtons[button].tagEnd != '') {
		edOpenTags[edOpenTags.length] = button;
		document.getElementById(edButtons[button].id).value = '/' + document.getElementById(edButtons[button].id).value;
	}
}

function edRemoveTag(button) {
	for (i = 0; i < edOpenTags.length; i++) {
		if (edOpenTags[i] == button) {
			edOpenTags.splice(i, 1);
			document.getElementById(edButtons[button].id).value = document.getElementById(edButtons[button].id).value.replace('/', '');
		}
	}
}

function edCheckOpenTags(button) {
	var tag = 0;
	for (i = 0; i < edOpenTags.length; i++) {
		if (edOpenTags[i] == button) {
			tag++;
		}
	}
	if (tag > 0) {
		return true; // tag found
	}
	else {
		return false; // tag not found
	}
}

function edToolbar() {
	document.write('<div id="ed_toolbar">')
	document.write('<span id="tools">');
	for (i = 0; i < 8; i++) {
		edShowButton(edButtons[i]);
	}
	document.write('<br />');
	for (i = 8; i < 16; i++) {
		edShowButton(edButtons[i]);
	}
	document.write('<br />');
	document.write('<br style="line-height:4px;" />');
	for (i = 16; i < 19; i++) {
		edShowButton(edButtons[i]);
	}
	document.write('<br />');
	document.write('<br style="line-height:4px;" />');
	for (i = 19; i < 23; i++) {
		edShowButton(edButtons[i]);
	}
	document.write('<br />');
	for (i = 23; i < 27; i++) {
		edShowButton(edButtons[i]);
	}
	document.write('<br />');
	document.write('<br style="line-height:4px;" />');
	document.write('</span>');
	document.write('<input type="button" class="ed_button" id="showtoolbar" onclick="edShowToolbar()" title="Show toolbar" value="▼" />');
	document.write('<input type="button" class="ed_button" id="hidetoolbar" onclick="edHideToolbar()" title="Hide toolbar" value="▲" />');
	document.write('</div>');
	document.getElementById('showtoolbar').style.display = 'inline';
	document.getElementById('hidetoolbar').style.display = 'none';
	document.getElementById('tools').style.display = 'none';
}

function edShowToolbar() {
	document.getElementById('showtoolbar').style.display = 'none';
	document.getElementById('hidetoolbar').style.display = 'inline';
	document.getElementById('tools').style.display = 'block';
}
function edHideToolbar() {
	document.getElementById('tools').style.display = 'none';
	document.getElementById('showtoolbar').style.display = 'inline';
	document.getElementById('hidetoolbar').style.display = 'none';
}



//
//
//
// Insertion code

function edInsertTag(myField, i) {
	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		if (sel.text.length > 0) {
			sel.text = edButtons[i].tagStart + sel.text + edButtons[i].tagEnd;
		}
		else {
			if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
				sel.text = edButtons[i].tagStart;
				edAddTag(i);
			}
			else {
				sel.text = edButtons[i].tagEnd;
				edRemoveTag(i);
			}
		}
		myField.focus();
	}
	//MOZILLA/NETSCAPE support
	else if (myField.selectionStart || myField.selectionStart == '0') {
		var startPos = myField.selectionStart;
		var endPos = myField.selectionEnd;
		var cursorPos = endPos;
		var scrollTop = myField.scrollTop;
		if (startPos != endPos) {
			myField.value = myField.value.substring(0, startPos)
			              + edButtons[i].tagStart
			              + myField.value.substring(startPos, endPos) 
			              + edButtons[i].tagEnd
			              + myField.value.substring(endPos, myField.value.length);
			cursorPos += edButtons[i].tagStart.length + edButtons[i].tagEnd.length;
		}
		else {
			if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
				myField.value = myField.value.substring(0, startPos) 
				              + edButtons[i].tagStart
				              + myField.value.substring(endPos, myField.value.length);
				edAddTag(i);
				cursorPos = startPos + edButtons[i].tagStart.length;
			}
			else {
				myField.value = myField.value.substring(0, startPos) 
				              + edButtons[i].tagEnd
				              + myField.value.substring(endPos, myField.value.length);
				edRemoveTag(i);
				cursorPos = startPos + edButtons[i].tagEnd.length;
			}
		}
		myField.focus();
		myField.selectionStart = cursorPos;
		myField.selectionEnd = cursorPos;
		myField.scrollTop = scrollTop;
	}
	else {
		if (!edCheckOpenTags(i) || edButtons[i].tagEnd == '') {
			myField.value += edButtons[i].tagStart;
			edAddTag(i);
		}
		else {
			myField.value += edButtons[i].tagEnd;
			edRemoveTag(i);
		}
		myField.focus();
	}
}

function edInsertContent(myField, myValue) {
	//IE support
	if (document.selection) {
		myField.focus();
		sel = document.selection.createRange();
		sel.text = myValue;
		myField.focus();
	}
	//MOZILLA/NETSCAPE support
	else if (myField.selectionStart || myField.selectionStart == '0') {
		var startPos = myField.selectionStart;
		var endPos = myField.selectionEnd;
		var scrollTop = myField.scrollTop;
		myField.value = myField.value.substring(0, startPos)
		              + myValue 
                      + myField.value.substring(endPos, myField.value.length);
		myField.focus();
		myField.selectionStart = startPos + myValue.length;
		myField.selectionEnd = startPos + myValue.length;
		myField.scrollTop = scrollTop;
	} else {
		myField.value += myValue;
		myField.focus();
	}
}

function edInsertUPCA(myField) {
	var myValue = prompt('Enter the 11-digit UPCA code', '');
	if (myValue) {
		myValue = '|UPCA[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertEAN13(myField) {
	var myValue = prompt('Enter the 13-digit EAN13 code', '');
	if (myValue) {
		myValue = '|EAN13[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertCode39(myField) {
	var myValue = prompt('Enter the Code39 code', '');
	if (myValue) {
		myValue = '|CODE39[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertNetAmt(myField) {
	var myValue = prompt('Enter the points program Var#:', '');
	if (myValue) {
		myValue = '|NETAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertInitialAmt(myField) {
	var myValue = prompt('Enter the points program Var#:', '');
	if (myValue) {
		myValue = '|INITIALAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertEarnedAmt(myField) {
	var myValue = prompt('Enter the points program Var#:', '');
	if (myValue) {
		myValue = '|EARNEDAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertRedeemedAmt(myField) {
	var myValue = prompt('Enter the points program Var#:', '');
	if (myValue) {
		myValue = '|REDEEMEDAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertNet$(myField) {
	var myValue = prompt('Enter the discount Var#:', '');
	if (myValue) {
		myValue = '|NETAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertInitial$(myField) {
	var myValue = prompt('Enter the discount Var#:', '');
	if (myValue) {
		myValue = '|INITIALAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertEarned$(myField) {
	var myValue = prompt('Enter the discount Var#:', '');
	if (myValue) {
		myValue = '|EARNEDAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}

function edInsertRedeemed$(myField) {
	var myValue = prompt('Enter the discount Var#:', '');
	if (myValue) {
		myValue = '|REDEEMEDAMT[' + myValue + ']|';
		edInsertContent(myField, myValue);
	}
}