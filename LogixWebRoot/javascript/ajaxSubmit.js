// No error status is returned with an error page, so there needs to be some
// other way to determine an error response.
//
// Parameters
//   str - A string typically returned from an AJAX call.
//
function isErrorMessage( msg )
{
  // Note: This is an extremely fragile check.
  // @todo: Strip comments. Commented HTML could mess with this check.
  var lower_msg = msg.toLowerCase();
  var idx = lower_msg.search("<!doctype");
  if( idx > -1 )
  {
    idx += 9;
    while ( lower_msg.charAt(idx) == ' ' || lower_msg.charAt(idx) == '\t' || lower_msg.charAt(idx) == '\n' || lower_msg.charAt(idx) == '\r' )
    {
      ++idx;
    }

    if( lower_msg.substr(idx, 4) != "html" )
    {
      return false;
    }
    idx += 4
  }
  else
  {
    idx = 0
  }

  lower_msg = lower_msg.substr( idx );
  idx = lower_msg.search("<html");
  if( idx == -1 )
  {
    return false;
  }
  idx += 5
  
  // The opening HTML tag must be a tag.
  if( ! (lower_msg.charAt(idx) == ' ' || lower_msg.charAt(idx) == '\t' || lower_msg.charAt(idx) == '\r' || lower_msg.charAt(idx) == '\n' || lower_msg.charAt(idx) == '>' ) )
  {
    return false;
  }
  idx += 1

  lower_msg = lower_msg.substr( idx );
  return ( lower_msg.search("</html>") != -1 )
}


// Swap the current document with the HTML string provided.
// This is primarily used to display error pages.
function swapDoc( html )
{
  var doc = document.open("text/html");
  doc.write(html);
  doc.close();
}


// strURL = the page that will receive the data
// formname = the NAME of the form you wish to submit
// responsediv = the id of the DIV taht will display the "wait message" and the form response.
// responsemsg = the waiting message (should be HTML code)

function xmlhttpPost(strURL,formname,responsediv,responsemsg) {
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
    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function() {
      if (self.xmlHttpReq.readyState == 4) {
        var responseText = self.xmlHttpReq.responseText;
        if( isErrorMessage(responseText) )
        {
          swapDoc( responseText );
        }
        else
        {
          updatepage(responseText,responsediv,true);
        }
      }
      else {
        updatepage(responsemsg,responsediv,false);
      }
    }
    self.xmlHttpReq.send(getquerystring(formname));
}

// strURL = the page that will receive the data
// formdata = the actual form data you wish to submit
// responsediv = the id of the DIV taht will display the "wait message" and the form response.
// callback = the javascript function that the client wishes to have called when the response
//            is received.  Send empty string ('') if no callback is needed.

function xmlhttpPostData(strURL, formdata, responsediv, callback) {
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
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function () {
    if (self.xmlHttpReq.readyState == 4) {
      var responseText = self.xmlHttpReq.responseText;
      if( isErrorMessage(responseText) )
      {
        swapDoc(responseText);
      }
      else
      {
        updateDiv(responseText, responsediv, true);
        if (callback != null) {
          setTimeout(callback, 10);
        }
      }
    }
  }
  self.xmlHttpReq.send(formdata);
}


// strURL = the page that will receive the data
// formname = the NAME of the form you wish to submit
// callback = the javascript function that the client wishes to have called when the response
//            is received.  Send null if no callback is needed.

function xmlhttpPostForm(strURL, formname, callback) {
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
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function () {
    if (self.xmlHttpReq.readyState == 4) {
      setTimeout(function () {
        // Verify that the response was reasonable and that the response text
        //   is of the expected form.
        var responseText = self.xmlHttpReq.responseText;
        if( isErrorMessage(responseText) )
        {
          swapDoc( responseText );
        }
        else
        {
          callback( responseText );
        }
      }, 10);
    }
  }
  self.xmlHttpReq.send(getquerystring(formname));
}

// strURL = the page that will receive the data
// formdata = the actual form data you wish to submit
// callback = the javascript function that the client wishes to have called when the response
//            is received.  Send empty string ('') if no callback is needed.

function xmlhttpPostDataCallback(strURL, formdata, callback) {
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
  self.xmlHttpReq.open('POST', strURL, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function () {
    if (self.xmlHttpReq.readyState == 4) {
      setTimeout(function () {
        if( isErrorMessage( self.xmlHttpReq.responseText ) )
        {
          swapDoc( self.xmlHttpReq.responseText );
        }
        else
        {
          callback(self.xmlHttpReq.responseText);
        }
      }, 10);
    }
  }
  self.xmlHttpReq.send(formdata);
}


function getquerystring(formname) {
    var form = document.forms[formname];
	var qstr = "";

function GetElemValue(name, value) {
        qstr += (qstr.length > 0 ? "&" : "")
            + escape(name).replace(/\+/g, "%2B") + "="
            + escape(value ? value : "").replace(/\+/g, "%2B");
			//+ escape(value ? value : "").replace(/\n/g, "%0D");
    }
	
	var elemArray = form.elements;
    for (var i = 0; i < elemArray.length; i++) {
        var element = elemArray[i];
        var elemType = element.type.toUpperCase();
        var elemName = element.name;
        if (elemName) {
            if (elemType == "TEXT"
                    || elemType == "TEXTAREA"
                    || elemType == "PASSWORD"
					|| elemType == "BUTTON"
					|| elemType == "RESET"
					|| elemType == "SUBMIT"
					|| elemType == "FILE"
					|| elemType == "IMAGE"
                    || elemType == "HIDDEN")
                GetElemValue(elemName, element.value);
            else if (elemType == "CHECKBOX" && element.checked)
                GetElemValue(elemName, 
                    element.value ? element.value : "On");
            else if (elemType == "RADIO" && element.checked)
                GetElemValue(elemName, element.value);
            else if (elemType.indexOf("SELECT") != -1)
                for (var j = 0; j < element.options.length; j++) {
                    var option = element.options[j];
                    if (option.selected)
                        GetElemValue(elemName,
                            option.value ? option.value : option.text);
                }
        }
    }
    return qstr;
}

function updatepage(str, responsediv, isready) {
  var reloadPageUrl = window.location.href;
  var elemResp = document.getElementById(responsediv);
  var respDataSent = false;

  if (responsediv != null) {
    if (isready) {
      var lines = str.split("\r\n");
      // return response: line 1 = status, line 2 = QueryString tokens (e.g. PKID=1), line 3 = Response message
      if (lines.length > 2) {
        if (lines[0] == 'OK') {
          respDataSent = (lines[1].indexOf('[DATA]:') > -1);

          if ((lines[1] == '' || respDataSent) && elemResp != null) {
            // send back a status message only when the QueryString is not specified.
            elemResp.style.backgroundColor = 'green';
            elemResp.style.display = 'block';
            elemResp.innerHTML = lines[2];

            if (respDataSent) {
              processAjaxResponseData(lines[1]);
            }
            return
          } else {
            // reload the page
            if (reloadPageUrl.indexOf('?') > -1) {
              reloadPageUrl = reloadPageUrl.substring(0, reloadPageUrl.indexOf('?'));
              reloadPageUrl = reloadPageUrl + lines[1];
            }
            window.location.href = reloadPageUrl
          }
        } else {
          document.getElementById(responsediv).style.display = 'block';
          document.getElementById(responsediv).style.backgroundColor = 'red';
          // get the response message returned from the ajax call
          str = '';
          for (var i = 2; i < lines.length; i++) {
            if (i > 2) {
              str += "<br />";
            }
            str += lines[i];          
          }
          
        }
      }
    } else {
      document.getElementById(responsediv).style.backgroundColor = 'green';
    }


    if (isready && str=='') {
      str = 'Error encountered.';
    }

    document.getElementById(responsediv).innerHTML = getCloseButtonHTML(str, responsediv);
  } else {
    alert(str);
  }

  function getCloseButtonHTML(str, divname) {
    return '<table style="width:100%;"><tr><td style="width:95%;">' + str + '</td><td style="text-align:center;vertical-align:top;cursor:hand;" ' +
           'onclick="javascript:document.getElementById(\'' + divname + '\').style.display=\'none\';">X</td></tr></table>'; 
  }
}

function updateDiv(str, responsediv, isready) {
  var elemResp = document.getElementById(responsediv);

  if (elemResp != null) {
    if (isready) {
      elemResp.innerHTML = str;
    }
  }
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
    self.xmlHttpReq.open('GET', 'BoxStatusUpdate.aspx?boxid='+BoxID+'&targetuser='+AdminUserID+'&boxopen='+BoxOpen, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.send();
}
