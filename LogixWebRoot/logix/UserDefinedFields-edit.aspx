<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>


<script src="/javascript/datePicker.js" type="text/javascript"></script>
<script src="/javascript/popup.js" type="text/javascript"></script>

<% 


  ' *****************************************************************************
  ' * FILENAME: UserDefinedFields-edit.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright � 2002 - 2009.  All rights reserved by:
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
  Dim UDFPK as Long
  Dim AdminUserID As Long
  Dim ExtID As String 
  Dim Description As String = ""
  Dim DataType as Integer = 0
  Dim PresentationStyleID as Integer

  Dim AdvSearch as Boolean = false
  Dim LastUpdate As String = ""
  Dim TempInt As Integer
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rstAssociated As DataTable = Nothing
  Dim row As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim Status as Integer
  Dim bNew as Boolean
  Dim saved as Boolean = false
    Dim rst2 As DataTable
    Dim udfvalidvals As DataTable
    Dim UDFPS_ID
    Dim numValues As Integer
    
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "userdefinedfields-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
' fill in if it was a get method
    If Request.RequestType = "GET" Then
	
            'UDFPK is the only thing that comes in on the querystring
	  UDFPK = IIf(Request.QueryString("UDFPK")="",0,MyCommon.Extract_Val(Request.QueryString("UDFPK")))
            'ExtID = htmlDecode(Logix.TrimAll(Request.QueryString("ExtID")))
	  
'ReasonFlag = IIf(Request.QueryString("ReasonFlag") = "", -1, MyCommon.Extract_Val(Request.QueryString("ReasonFlag")))
            'Description = Logix.TrimAll(Request.QueryString("Description"))
            'DataType = IIf(Request.QueryString("DataType")="",0,MyCommon.Extract_Val(Request.QueryString("DataType")))
            'numValues = IIf(Request.QueryString("numValues") = "", 0, MyCommon.Extract_Val(Request.QueryString("numValues")))

            'AdvSearch = IIf(Request.QueryString("AdvSearch") = "", False, True)
            'UDFPS_ID = IIf(Request.QueryString("ddPresentationStyle") = "", 0, MyCommon.Extract_Val(Request.QueryString("ddPresentationStyle")))
            
            
      If Request.QueryString("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.QueryString("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.QueryString("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    Else
            UDFPK = IIf(Request.Form("UDFPK") = "", 0, MyCommon.Extract_Val(Request.Form("UDFPK")))
            UDFPS_ID = IIf(Request.Form("ddPresentationStyle")="",0,MyCommon.Extract_Val(Request.Form("ddPresentationStyle")))
'ReasonFlag = IIf(Request.Form("ReasonFlag") = "", -1, MyCommon.Extract_Val(Request.Form("UDFPK")))
      If UDFPK <= 0 Then
        UDFPK = IIf(Request.QueryString("UDFPK") = "", 0, MyCommon.Extract_Val(Request.QueryString("UDFPK")))
      End If
	             
            ExtID = htmlDecode(Logix.TrimAll(Request.Form("ExtID")))
            DataType = IIf(Request.Form("DataType") = "", 0, MyCommon.Extract_Val(Request.Form("DataType")))
            numValues = IIf(Request.Form("numValues") = "", 0, MyCommon.Extract_Val(Request.Form("numValues")))
            AdvSearch = IIf(Request.Form("AdvSearch") = "", False, True)
      Description = htmlDecode(Logix.TrimAll(Request.Form("Description")))
'Response.Redirect(ReasonFlag)
      If Request.Form("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.Form("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.Form("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    End If
    
    Send_HeadBegin("term.userdefinedfield", , IIF(UDFPK>=0, UDFPK, ""))
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
var datePickerDivID = "datepicker";
  function toggleDropdown() {
		if (document.getElementById("actionsmenu") != null) {
		  bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
		  if (bOpen) {
			document.getElementById("actionsmenu").style.visibility = 'visible';
			document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
		  } else {
			document.getElementById("actionsmenu").style.visibility = 'hidden';
			document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
		  }
		}
	}

	
	function submitDesc() {

	//	var text = document.getElementById("desc").value;
		//text = htmlEncode(text);
	//	document.getElementById("Description").value = text;
		//document.getElementById("desc").disabled = true;

	}


	function deleteUDF(offercount) {
	    var del;
	    if (offercount > 0)
	        del = confirm('<%Sendb(Copient.PhraseLib.Lookup("confirm.deleteudf", LanguageID)) %>');
	    else
	        del = confirm('<%Sendb(Copient.PhraseLib.Lookup("confirm.delete", LanguageID)) %>');

	    return del;
	}

	
  // deleteRow is the id of the <tr> that needs to be deleted
  // form data will come across as a comma deleted row of udfvpk ids
  

	function deleteUDFValue(deleteRow) {

      
        //parse deleteRow name
       
       s="<input type=\"hidden\" name=\"udfvaluedelete_" +deleteRow+ "\" value=\""+deleteRow+"\"/>";
       
       //if the value is negative, the item has not been saved to database, so no reason to add it to the itemsToDelete list
        if(parseInt(deleteRow)>=0)
        {
            $("#itemsToDelete").append(s);
        }
       
        $("#tr_" + deleteRow + "").remove();//remove the display row        
        $("#tr_" + deleteRow + "H").remove();//remove the hidden input row
        resetRowOrder();
        
        if( ($("#valueTableBody > tr").length/2) > 0   && (document.getElementById("ddPresentationStyle").value == 6) 
                    && (document.getElementById("DataType").value != 5)
        )
        {
            disableAddNewValue(true);
        }
        else
        {
            disableAddNewValue(false);
        }
    }

    function resetRowOrder()
    {
        newRowOrder=0;//reset global order variable
        $("#valueTableBody > tr").each(function()
        {          
            if($(this)[0].style.display == "none")            
            {
                //use find to skip over <td> element
                $(this).find("input").each(function()
                {
                    var inputName = $(this).attr("name");
                    if(inputName != null && inputName.indexOf("order_") != -1)
                    {                            
                        $(this).val(newRowOrder++);
                    }
                });
            }
            
         });
        
         return;


	}

    /*
    reference : http://www.javascriptkit.com/script/script2/validatedate.shtml
    */
function checkdate(input){
        var validformat1=/^\d{2}\/\d{2}\/\d{4}$/ //Basic check for format validity
        var validformat2=/^\d{1}\/\d{2}\/\d{4}$/ 
        var validformat3=/^\d{2}\/\d{1}\/\d{4}$/ 
        var validformat4=/^\d{1}\/\d{1}\/\d{4}$/ 
        var returnval=false

        var test1 = validformat1.test(input.trim());
        var test2 = validformat2.test(input.trim());
        var test3 = validformat3.test(input.trim());
        var test4 = validformat4.test(input.trim());
        if (! (test1 || test2 || test3 || test4))
        {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invaliddateordateformat", LanguageID)) %>');
        }
        else{ //Detailed check for valid date ranges
            var monthfield=input.split("/")[0]
            var dayfield=input.split("/")[1]
            var yearfield=input.split("/")[2]
            var dayobj = new Date(yearfield, monthfield-1, dayfield)
            if ((dayobj.getMonth()+1!=monthfield)||(dayobj.getDate()!=dayfield)||(dayobj.getFullYear()!=yearfield))
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invaliddateordateformat", LanguageID)) %>');
            else
            returnval=true
        }
        return returnval
}

    var valueLimit = -1;
    var newRowCount=-1; //newly added items that have not been comitted to the database will have a negative ID
    var newRowOrder = 0;
    function addUDFValue(strValue,isRange)
    {        
        /*
            valueID = UDFVPK | newRowCount

        <tr id/name=tr_valueID>
            <td><input X button></td>
            <td>value</td>
            <td><input radio button   value=valueID></td>
            <td><input button move up></td>
            <td><input button move down></td>
            </tr>
            <tr display:none>
            <td><input type='hidden' name=valueID value=[valueID]></td>
            <td><input type='hidden' name=valueID_value value=strValue></td>
            <td><input type='hidden' name=valueID_order value=order></td>
            <td></td>
            <td></td>
            </tr>
        */

        //adding empty value is invalid
        if(strValue.trim() == "")
        {
            return;
        }


        //determine if it's a range of a single digit
        if (isRange) 
        {        
            var parts = strValue.split(":");
            if(parts.length==2)
            {
                parts[0] = parts[0].replace("{", "");
                parts[1] = parts[1].replace("}", "");

                if (parts[0].length == 0 && parts[1].length > 0) 
                {
                    strValue = parts[1];
                    isRange = false;
                }
                else if (parts[0].length > 0 && parts[1].length == 0) 
                {
                    strValue = parts[0];
                    isRange = false;
                }
                else if(parts[0].length > 0 && parts[1].length > 0) 
                {
                    isRange = true;
                    //correct if values are entered like {100:5}
                    var minEntered = parseInt(parts[0]);
                    var maxEntered = parseInt(parts[1]);
                    if (maxEntered < minEntered)
                    {
                        strValue = "{"+maxEntered+":"+minEntered+"}"
                    }
                }
                else
                {
                    return;
                }
            }
        }

        if(document.getElementById("DataType").value == '2')
        {
            //document.getElementById("dateValue").value= document.getElementById("dateValue").value.replace(/\\/g,"/");
            strValue= strValue.replace(/\\/g,"/");
            if(checkdate(strValue)==false)
            {
                    alert('<% Sendb(Copient.PhraseLib.Lookup("error.invaliddateordateformat", LanguageID)) %>');
                    return;
            }

        }

        //validate integer is in valid range
        if(document.getElementById("DataType").value == '1')
        {
            if((parseInt(strValue) > 2147483647) || (parseInt(strValue) < -2147483648))
            {
                alert('<% Sendb(Copient.PhraseLib.Lookup("error.integeroutofrange", LanguageID)) %>');
                return;
            }

            //remove preceding zeros from int            
            strValue=parseInt(strValue,10).toString();
        }


        
        var valueID=newRowCount;
        var newRowOrder = $("#valueTableBody > tr").length;
        newRowOrder = newRowOrder/2;
        var newRow = "<tr id=\"tr_"+valueID + "\"  name=\"tr_"+valueID+"\">";
        newRow += "<td> <input type=\"button\" value=\"X\" title=\"" + document.getElementById('deleteText').value + "\" name=\"ex_" + valueID + "\" id=\"ex_" + valueID + "\" class=\"ex\" onclick=\"deleteUDFValue('"+valueID+"')\" ></td>";
        if(document.getElementById("DataType").value == '7')
        {
          var strValue1 = "show-image.aspx?caller=udf&src=" + strValue
          newRow += "<td><span style=\"display: inline-block; width: 180px;min-width: 20px; max-width: 130px; overflow-x: scroll; \">"+HTMLEncode(strValue)+"</span></td>";
          newRow += "<td><img align=\"right\" src=\"" + HTMLEncode(strValue1) + "\" id=\"Image-" + valueID + "\" width=\"50\" height=\"50\" alt=\"Image not Found\" title=\"Click to view full-sized image\" onclick=\"showFullSizedImage('" + HTMLEncode(strValue1) + "');\" /></td>";
        }
        else
        {
          newRow += "<td><span style=\"display: inline-block; width: 180px;min-width: 20px; max-width: 180px; overflow: hidden;   text-overflow: ellipsis;\">"+HTMLEncode(strValue)+"</span></td>";
        }

        if(isRange==true)
        {
            newRow += "<td></td>";
        }
        else
        {
          if(document.getElementById("ddPresentationStyle").value == '6')
          {
            newRow += "<td><input type=\"radio\" name=\"defaultUDFValue\" value=\""+valueID+"\" checked ></td>";
          } else {
            newRow += "<td><input type=\"radio\" name=\"defaultUDFValue\" value=\""+valueID+"\"></td>";
          }
        }

        newRow +=  "<td><input type=\"button\" value=\"▲\" title=\"" + document.getElementById('moveupText').value + "\" name=\"mvu_" + valueID + "\" id=\"mvu_" + valueID + "\" onclick=\"moverow('" + valueID + "',-1);\"/></td>";
        newRow +=  "<td><input type=\"button\" value=\"▼\" title=\"" + document.getElementById('movedownText').value + "\" name=\"mvd_" + valueID + "\" id=\"mvd_" + valueID + "\" onclick=\"moverow('" + valueID + "',1);\"/></td>";
        
        
        newRow += "</tr>"; 
        newRow += "<tr id=\"tr_"+valueID + "H\"  name=\"tr_"+valueID+"H\" style=\"display:none;\"  ><td><input type='hidden' name='valueID_"+valueID+"' value='"+ valueID+"'></td>";
        newRow += "<td><input type='hidden' id='value_"+valueID+"' name='value_"+valueID+"' value='"+ HTMLEncode(strValue)+"' ></td>";
        newRow += "<td><input type='hidden' id='order_"+valueID+"' name='order_"+valueID+"' value='"+ newRowOrder+"'></td>";
        newRow += "<td></td>";
        newRow += "<td></td>";
        newRow += "</tr>";

        newRowCount--;
        $("#valueTableBody").append(newRow);
        
        
        var mvu = document.getElementById('mvu_'+valueID);
	if(mvu != null)
	{
        	mvu.onclick=function(){moverow(valueID,-1)};
	}	
        var mvd = document.getElementById('mvd_'+valueID);
        if(mvd!=null)
	{
		mvd.onclick=function(){moverow(valueID,1)};
	}

        var del = document.getElementById('ex_'+valueID);
        if(mvd!=null)
	{
		del.onclick=function(){deleteUDFValue(valueID)};
	}
        
        checkLimits();

        document.getElementById('ValueText').value = "";
        document.getElementById('dateValue').value = "";
        document.getElementById('ValueTextMin').value = "";
        document.getElementById('ValueTextMax').value = "";
        return valueID;
    }



    /*
    strValue : value ID
    direction : [1|-1]  -1 = move up, 1 = move down
    */
    function moverow(strValue,direction)
    {
    
    //house keeping ,remove any rogue nodes, IE was inserting empty text nodes, I need to remove these to get the indexing to work
        var i;
        var vtb = document.getElementById("valueTableBody");
        var maxI = vtb.childNodes.length;
        for(i=maxI-1; i>=0;i--)
        {
            var vtbc = vtb.childNodes[i];
            if(vtbc.tagName === undefined || vtbc.tagName.toLowerCase() != "tr")
            {
                vtbc.parentNode.removeChild(vtbc);
            }
        }


        var numRows = ($("#valueTableBody > tr").length)/2;

        //get current row position
        var curRowOrder = parseInt($("#order_"+strValue).val());

        //if move up and position is already at top
        if(direction == -1 && curRowOrder == 0)
        {
            return;
        }


        //if move down and position is already last
        if(direction == 1 && curRowOrder == (numRows-1))
        {
            return;
        }
        
        var thisRowIndex = curRowOrder*2;
        var swapRowIndex = (curRowOrder+direction)*2;
        var thisRow = document.getElementById("valueTableBody").childNodes[thisRowIndex];
        var thisRowH = document.getElementById("valueTableBody").childNodes[thisRowIndex + 1];
            
        var swapRow = document.getElementById("valueTableBody").childNodes[swapRowIndex];
        var swapRowH = document.getElementById("valueTableBody").childNodes[swapRowIndex + 1];

        if(direction==-1)
        {
            vtb.removeChild(thisRow);
            vtb.removeChild(thisRowH);
            swapRow.parentNode.insertBefore(thisRowH,swapRow);
            swapRow.parentNode.insertBefore(thisRow,thisRowH);
        }
        if(direction==1)
        {
            vtb.removeChild(swapRow);
            vtb.removeChild(swapRowH);
            thisRow.parentNode.insertBefore(swapRowH,thisRow);
            thisRow.parentNode.insertBefore(swapRow,swapRowH);
        }        
        
        resetRowOrder();
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


    function isNumber(control,evt) {
        
        if($("#DataType").val()!=1 && $("#DataType").val()!=5)
        {
            return;
        }

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
            actualkey == "-")// ||
            //actualkey == ".")
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
    //reference  http://stackoverflow.com/questions/784586/convert-special-characters-to-html-in-javascript
    function HTMLEncode(str)
    {
        var i = str.length,
        aRet = [];

        while (i--) 
        {
            var iC = str.charAt(i).charCodeAt();
            if (iC < 65 || iC > 127 || (iC>90 && iC<97)) 
            {
                aRet[i] = '&#'+iC+';';
            }   
            else 
            {
                aRet[i] = str.charAt(i);
            }
        }
        return aRet.join('');    
    }


	window.onload = function() {
		document.getElementById("ExtID").focus();
	};
    
    //in chrome, i could just hide/show the selections in the drop down
    //however, that did not work in IE, so I'm creating a cache of the options, and I'll dynamically add/remove
    var ps_array = new Array();
    <%
        Dim rstPS as DataTable
        Dim rowPS as datarow
        MyCommon.QueryStr = "Select * from UserDefinedFieldsPresentationStyles"        
        MyCommon.Open_LRTsp()
		rstPS = MyCommon.LRT_Select
        
        For Each rowPS In rstPS.Rows
            Send("ps_array[" & rowPS.Item("UDFPS_ID") & "] = new Array('ps_" + rowPS.Item("PresentationStyleID") + "', '" & rowPS.Item("UDFPS_ID") & "','" & rowPS.Item("PresentationStyle") & "');")
        Next
     %>
    
    var dt_array = new Array();
    <%
        Dim rstDT as DataTable
        MyCommon.QueryStr = "Select * from UserDefinedFieldsTypes"
			rstDT = MyCommon.LRT_Select
            For Each row2 In rstDT.Rows
            Send("dt_array[" & row2.Item("UDFTypeID") & "] = new Array('" & row2.Item("UDFTypeID") & "','" & row2.Item("DisplayText") & "');")
            Next
     %>

    var dt_ps_xref = new Array();
    <%
        'this results in an array that will have UDFTypeID as the index, the value of dt_ps_xref[udftypeid] will be an array of valid udfps_ids
        Dim rstVPS As DataTable
        Dim tempArray as new SortedList
        MyCommon.QueryStr = "Select * from UserDefinedFields_ValidPresentationStyles order by UDFTypeID,DisplayOrder"        
        MyCommon.Open_LRTsp()
		rstVPS = MyCommon.LRT_Select
        For Each rowVPS In rstVPS.Rows
            ' UDFTypeID, UDFPS_ID
            Dim udfTypeID as Integer
            Dim ludfPS_ID as Integer
            udfTypeID = rowVPS.Item("UDFTypeID")
            ludfPS_ID = rowVPS.Item("UDFPS_ID")
            if tempArray.ContainsKey(udfTypeID) then
                tempArray(udfTypeID).Add(ludfPS_ID)                
            else
                tempArray.Add(udfTypeID, new System.Collections.ArrayList)
                tempArray(udfTypeID).Add(ludfPS_ID)
            end if
        Next

        for each xKey in tempArray.Keys
            Dim s as String 
            s=""
            for each xItem in tempArray(xKey)
                if s.length > 0 then
                    s+=",'"+CType(xItem,String)+"'"
                else
                    s+="'"+CType(xItem,String)+"'"
                end if
            next
            Send("dt_ps_xref["+CType(xKey,String)+"]= new Array(" + s + ");")
        next
     %>
     
     function dataTypeChange(clearFields)
     {  
        if(typeof(clearFields)==='undefined') clearFields = true;
		document.getElementById("datepicker").style.visibility="hidden";
		document.getElementById("datepicker").style.display="none";
        var dtElement = document.getElementById("DataType")
        var psElement = document.getElementById("ddPresentationStyle")
        var currentSelectedpsElement = psElement.value;
        var picks = dt_ps_xref[dtElement.value];


        //rebuild the presentation style drop down
        $("#ddPresentationStyle > option").remove();
        
        var index = 0;
        for(index=0; index < picks.length;index++)
        {
            var newOpt = new Option();
            newOpt.id=ps_array[picks[index]][0];
            newOpt.value=ps_array[picks[index]][1];
            newOpt.innerHTML = ps_array[picks[index]][2];

            if(currentSelectedpsElement == ps_array[picks[index]][1])
            {
                newOpt.selected=true;
            }
            psElement.appendChild(newOpt);
        }
        prepList(false);
        return;
     }
      
     function addBooleanRows()
     {
        var newValueID;
        $("#valueTableBody > tr").remove();
        newValueID = addUDFValue("true",false);
        $("#ex_"+newValueID).attr('disabled','disabled');
        newValueID = addUDFValue("false",false);
        $("#ex_"+newValueID).attr('disabled','disabled');
        disableAddNewValue(true);
     }
        
     function disableAddNewValue(on_off)
     {        
        $("#ValueText").attr('disabled',on_off);
        $("#addNewValue").attr('disabled',on_off);

        $("#ValueTextMin").attr('disabled',on_off);
        $("#ValueTextMax").attr('disabled',on_off);
        $("#addNewValueNR").attr('disabled',on_off);

        $("#dateValue").attr('disabled',on_off);
        $("#udf-datevalue-picker").attr('disabled',on_off);
        $("#addNewDateValue").attr('disabled',on_off);

     }

     function checkLimits()
     {
     
        var dtElement = document.getElementById("DataType")
         if(($("#valueTableBody > tr").length/2) >= valueLimit)
        {   
            disableAddNewValue(true);
        }
        else
        {
            disableAddNewValue(false);
        }
     }
     /*
     this function modifies the values list if datatype/presentation style changes
     */
     function prepList(isLoading)
     {
        if(document.readyState != 'complete')
        {   
            return;
        }
        var valueTextElement = document.getElementById("ValueText");
        valueTextElement.value = "";
        
        var dtElement = document.getElementById("DataType");
        var psElement = document.getElementById("ddPresentationStyle");
        
        if(dtElement.value == '6')//Likert 
        {
            valueLimit=<%
        Dim rstDTL as DataTable
        MyCommon.QueryStr = "Select * from UserDefinedFieldsTypes where UDFTypeID=6"
			rstDTL = MyCommon.LRT_Select
            if rstDTL.Rows.Count = 1 then
                Sendb(rstDTL.Rows(0).Item("MaxValue"))
            end if
     %>;
        }
        else if(psElement.value == '6' && dtElement.value != "5") //textbox
        {
            valueLimit = 1;
        }
        else if(psElement.value == '5') //checkbox
        {
            valueLimit=0;
        }
        else
        {
            valueLimit=1000000;
        }


        //different presentation styles and data types imply a maximum number of values
        //if the current list exceeds this limit, we'll delete all
        //also will delete for integer, date, boolean, numeric range, so, we don't have to validate the data that's there :-(
        if(
        ($("#valueTableBody > tr").length/2) > valueLimit   ||
        (dtElement.value == '1' && !isLoading) || //integer
        (dtElement.value == '2' && !isLoading) || //date
        (dtElement.value == '3' && !isLoading) || //boolean
        (dtElement.value == '5' && !isLoading)    //numeric range
        )
        {   
            deleteAllValues();
            $("#valueTableBody > tr").remove();
        }
        else
        {
            $("#valueTableBody > tr > td > input").each(function()
            {
                var inputName = $(this).attr("name");
                if(inputName != null && inputName.indexOf("ex_") != -1)
                {                                
                    if(dtElement.value != '3') //if it's boolean, we want it to stay disabled
                    {                
                        $(this).attr('disabled','');
                    }
                }
            });           
        }

        //force values for boolean, non-checkbox
        if(dtElement.value == '3' &&  ps_array[psElement.value][0]!='ps_CheckBox')
        {
            if(isLoading==true)
            {
                disableAddNewValue(true);
            }
            else
            {
                addBooleanRows();
            }
        }
        else
        {
            disableAddNewValue(false);            
            checkLimits();
        }


        //enable/disable input elements
        if(dtElement.value == '5')//numeric range
        {   
            document.getElementById("numericRangeValues").style.visibility="visible";
            document.getElementById("addNewValueNR").style.visibility="visible";
            document.getElementById("ValueText").style.visibility="hidden";
            document.getElementById("addNewValue").style.visibility="hidden";
            document.getElementById("dateDiv").style.visibility="hidden";
            document.getElementById("addNewDateValue").style.visibility="hidden";

            
            document.getElementById("numericRangeValues").style.display="";
            document.getElementById("addNewValueNR").style.display="";
            document.getElementById("ValueText").style.display="none";
            document.getElementById("addNewValue").style.display="none";
            document.getElementById("dateDiv").style.display="none";
            document.getElementById("addNewDateValue").style.display="none";
        }
        else if(dtElement.value == '2')//date
        {
            document.getElementById("numericRangeValues").style.visibility="hidden";
            document.getElementById("addNewValueNR").style.visibility="hidden";
            document.getElementById("ValueText").style.visibility="hidden";
            document.getElementById("addNewValue").style.visibility="hidden";
            document.getElementById("dateDiv").style.visibility="visible";
            document.getElementById("addNewDateValue").style.visibility="visible";

            
            document.getElementById("numericRangeValues").style.display="none";
            document.getElementById("addNewValueNR").style.display="none";
            document.getElementById("ValueText").style.display="none";
            document.getElementById("addNewValue").style.display="none";
            document.getElementById("dateDiv").style.display="";
            document.getElementById("addNewDateValue").style.display="";

        /*
        <td><input class="short" id="udfVal-17" name="udfVal-17" maxlength="10" type="text" value=" "><img src="/images/calendar.png" class="calendar" id="udf-datevalue-picker" alt="Date picker" title="Date picker" onclick="displayDatePicker('udfVal-17', event);">  </td>
        */
        }
        else if(dtElement.value == '1')//integer
        {
            document.getElementById("numericRangeValues").style.visibility="hidden";
            document.getElementById("addNewValueNR").style.visibility="hidden";
            document.getElementById("ValueText").style.visibility="visible";
            document.getElementById("ValueText").maxLength="11";
            document.getElementById("addNewValue").style.visibility="visible";
            document.getElementById("dateDiv").style.visibility="hidden";
            document.getElementById("addNewDateValue").style.visibility="hidden";

            
            document.getElementById("numericRangeValues").style.display="none";
            document.getElementById("addNewValueNR").style.display="none";
            document.getElementById("ValueText").style.display="";
            document.getElementById("addNewValue").style.display="";
            document.getElementById("dateDiv").style.display="none";
            document.getElementById("addNewDateValue").style.display="none";

        /*
        <td><input class="short" id="udfVal-17" name="udfVal-17" maxlength="10" type="text" value=" "><img src="/images/calendar.png" class="calendar" id="udf-datevalue-picker" alt="Date picker" title="Date picker" onclick="displayDatePicker('udfVal-17', event);">  </td>
        */
        }
        else
        {            
            document.getElementById("numericRangeValues").style.visibility="hidden";
            document.getElementById("addNewValueNR").style.visibility="hidden";
            document.getElementById("ValueText").style.visibility="visible";
            document.getElementById("ValueText").maxLength="1000";
            document.getElementById("addNewValue").style.visibility="visible";
            document.getElementById("dateDiv").style.visibility="hidden";
            document.getElementById("addNewDateValue").style.visibility="hidden";

            
            document.getElementById("numericRangeValues").style.display="none";
            document.getElementById("addNewValueNR").style.display="none";
            document.getElementById("ValueText").style.display="";
            document.getElementById("addNewValue").style.display="";
            document.getElementById("dateDiv").style.display="none";
            document.getElementById("addNewDateValue").style.display="none";
        }        
     }

     function myIndexOf(someObj,someValue)
     {
        var index = -1;
        
        for(var j=0; j<someObj.length;j++)
        {
            if(someObj[j]==someValue)
            {
                return j;
            }
        }
        return index;
     }
     function presentationStyleChange(clearFields)
     {
        if(typeof(clearFields)==='undefined') clearFields = true;
        
        var dtElement = document.getElementById("DataType");
        var psElement = document.getElementById("ddPresentationStyle");
        var currentSelecteddtElement = dtElement.value;
                
        //rebuild the DataType dropdown
        $("#DataType > option").remove();
        var index=0;
        var validDT = "";
        for(var key in dt_array)
        {
            //loop through dt_array
            //if key (UDFTypeID) dt_ps_xref[key] contains psElement.value add new option
            //var test = dt_ps_xref[key].indexOf( psElement.value);
            
            //if(dt_ps_xref[key].indexOf( psElement.value) != -1)
            if(myIndexOf(dt_ps_xref[key],psElement.value) != -1)
            {
                //var newOpt = new Option(dt_array[key][1],dt_array[key][0]);
                var newOpt = new Option();
                newOpt.value = dt_array[key][0];
                newOpt.innerHTML=dt_array[key][1];
                if(dt_array[key][0]==currentSelecteddtElement)
                {
                    newOpt.selected=true;
                }
                dtElement.appendChild(newOpt);
            }
            
        }
        prepList(false);
        return;
     }

     function deleteAllValues()
     {
        newRowOrder=0;//reset global order variable
        $("#valueTableBody > tr").each(function()
        {        
            var abc = $(this).attr("style");
            if(abc==null)//find the non hidden data row
            {
                var tagid = $(this).attr("id");
            }
            
         });
        
         return;
	}
    
    /*
        If presentation style is Drop Down List, Vertical Rabio Buttons, Horizontal Radio Buttons, or List Box
        there must be values
        also if data type is numeric range, user must enter value.
    */
    function validateSave()
    {        
        var psElement = document.getElementById("ddPresentationStyle");
        var ddElement = document.getElementById("DataType");
        if(psElement.value == 1 || psElement.value == 2 || psElement.value == 3 || psElement.value == 4 || ddElement.value == 5)
        {
            if($("#valueTableBody > tr").length == 0)
            {
                alert('<% Sendb(Copient.PhraseLib.Lookup("error.mustenterudfvalues", LanguageID)) %>');
                return false;
            }
        }
        return true;
    }

    $(window).load(function() {
    // executes when complete page is fully loaded, including all frames, objects and images
        prepList(true);
});
    
  // displays a 'imagepopup' div with an image centered within it
  function showFullSizedImage(imagesrc) {
   var elemImg = document.getElementById('fullSizedImage');
   var elemWin = document.getElementById('imagepopup');

   if (elemImg != null) {
     elemImg.src = imagesrc;
     var iw = parseInt(elemImg.width);
     var ih = parseInt(elemImg.height);
     var aspect = (iw*1.0)/(ih*1.0);
     if (iw > 600) {
       iw = 600;
       ih = iw / aspect;
     }
     if (ih > 600) {
       ih = 600;
       iw = ih * aspect
     }
     elemImg.setAttribute("width",iw);
     elemImg.setAttribute("height", ih);
   }
   if (elemWin != null) { 
     elemWin.style.display = '';
   }
 }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 4)
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration") 
    GoTo done
  End If
  
  If (Request.QueryString("new") <> "" or Request.Form("new") <> "") Then
        Response.Redirect("UserDefinedFields-edit.aspx")
	UDFPK = 0
  End If


    
    If bSave Then
        
    'Dim Status as Integer
    If Description = "" Then
      infoMessage = Copient.PhraseLib.Lookup("error.nodescription", LanguageID)
    ElseIf ExtID = "" Then
      infoMessage = Copient.PhraseLib.Lookup("error.noextid", LanguageID)
    Else

      If UDFPK = 0 Then 'insert new udf
        MyCommon.QueryStr = "dbo.pt_UserDefinedFields_insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ExtID", SqlDbType.NVarChar, 50).Value = ExtID
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 100).Value = htmlDecode(Description)
        MyCommon.LRTsp.Parameters.Add("@DataType", SqlDbType.Int).Value = DataType
        MyCommon.LRTsp.Parameters.Add("@AdvSearch", SqlDbType.Bit).Value = AdvSearch
        MyCommon.LRTsp.Parameters.Add("@UDFPS_ID", SqlDbType.BigInt).Value = UDFPS_ID
        MyCommon.LRTsp.Parameters.Add("@UDFPK", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        UDFPK = MyCommon.NZ(MyCommon.LRTsp.Parameters("@UDFPK").Value, -1)
        Status = MyCommon.LRTsp.Parameters("@Status").Value
        MyCommon.Close_LRTsp()
        If Status = -1 Then
          infoMessage = Copient.PhraseLib.Lookup("error.extidinuse", LanguageID)
        ElseIf Status = -2 Then
          infoMessage = Copient.PhraseLib.Lookup("error.descriptioninuse", LanguageID)
        End If
      Else 'update udf
        MyCommon.QueryStr = "dbo.pt_UserDefinedFields_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@UDFPK", SqlDbType.BigInt).Value = UDFPK
        MyCommon.LRTsp.Parameters.Add("@ExtID", SqlDbType.NVarChar, 50).Value = ExtID
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 100).Value = htmlDecode(Description)
        MyCommon.LRTsp.Parameters.Add("@DataType", SqlDbType.Int).Value = DataType
        MyCommon.LRTsp.Parameters.Add("@AdvSearch", SqlDbType.Bit).Value = AdvSearch
        MyCommon.LRTsp.Parameters.Add("@UDFPS_ID", SqlDbType.BigInt).Value = UDFPS_ID
        MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        Status = MyCommon.LRTsp.Parameters("@Status").Value
        MyCommon.Close_LRTsp()
        If Status = -1 Then
          infoMessage = Copient.PhraseLib.Lookup("error.extidinuse", LanguageID)
        ElseIf Status = -2 Then
          infoMessage = Copient.PhraseLib.Lookup("error.descriptioninuse", LanguageID)
        End If
      End If
            
            
      ''handle udf values, logic is the same whether we're saveing new UDF or updating existing so, this can be outside of the above blocks
      For Each formItem In Request.Form.AllKeys()
        If formItem.Contains("udfvaluedelete_") Then
          Dim delUDFVPK As Integer
          delUDFVPK = System.Convert.ToInt64(Request.Form(formItem))
          MyCommon.QueryStr = "dbo.pt_UserDefinedFieldsValidValues_Delete"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@UDFVPK", SqlDbType.BigInt).Value = delUDFVPK
          MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LRTsp.ExecuteNonQuery()
          Status = MyCommon.LRTsp.Parameters("@Status").Value
          MyCommon.Close_LRTsp()
        End If
      Next
           
      Dim lol As String
      lol = ""
      'parse the form data and put it in a sorted list
      'sorted list allows us to properly order the list
      Dim udfvalues As New SortedList
      For Each formItem In Request.Form.AllKeys()
        If formItem.Contains("order_") Then
          Dim getID As String()
          Dim id As String
          Dim udfvalueinfo(2) As Object
                        
          getID = formItem.Split("_")
          id = getID(1)
          udfvalueinfo(0) = Request.Form("order_" + id)
          udfvalueinfo(1) = Server.HtmlDecode(Request.Form("value_" + id))
          udfvalueinfo(2) = System.Convert.ToInt64(Request.Form("valueID_" & id))
          udfvalues.Add(System.Convert.ToInt32(udfvalueinfo(0)), udfvalueinfo)
          lol &= "order_"
          lol &= udfvalueinfo(0)
          lol &= "value_"
          lol &= udfvalueinfo(1)
          lol &= "valueID_"
          lol &= udfvalueinfo(2)
        End If
      Next
                
      For Each itemOrder In udfvalues.Keys
        Dim udfvalueItem(2) As Object
        udfvalueItem = udfvalues.Item(itemOrder)
        Dim valueToStore = udfvalueItem(1)
        If DataType = 1 Then
          valueToStore = Convert.ToString(Convert.ToInt32(valueToStore))
        End If
                
        If udfvalueItem(2) < 0 Then
          'insert
          MyCommon.QueryStr = "dbo.pt_UserDefinedFieldsValidValues_Insert"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@UDFPK", SqlDbType.BigInt).Value = UDFPK
          MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 1000).Value = valueToStore
          MyCommon.LRTsp.Parameters.Add("@IsDefault", SqlDbType.Bit).Value = If(Request.Form("defaultUDFValue") = udfvalueItem(2), 1, 0)
          MyCommon.LRTsp.Parameters.Add("@DisplayOrder", SqlDbType.Int).Value = udfvalueItem(0)
          MyCommon.LRTsp.Parameters.Add("@UDFVPK", SqlDbType.BigInt).Direction = ParameterDirection.Output
          MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LRTsp.ExecuteNonQuery()
          Status = MyCommon.LRTsp.Parameters("@Status").Value
          MyCommon.Close_LRTsp()
        Else
          'update
          MyCommon.QueryStr = "dbo.pt_UserDefinedFieldsValidValues_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@UDFPK", SqlDbType.BigInt).Value = UDFPK
          MyCommon.LRTsp.Parameters.Add("@UDFVPK", SqlDbType.BigInt).Value = udfvalueItem(2)
          MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 1000).Value = valueToStore
          MyCommon.LRTsp.Parameters.Add("@IsDefault", SqlDbType.Bit).Value = If(Request.Form("defaultUDFValue") = udfvalueItem(2), 1, 0)
          MyCommon.LRTsp.Parameters.Add("@DisplayOrder", SqlDbType.Int).Value = itemOrder
          MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LRTsp.ExecuteNonQuery()
          Status = MyCommon.LRTsp.Parameters("@Status").Value
          MyCommon.Close_LRTsp()
        End If
      Next
            
      If UDFPK > 0 Then Response.Redirect("userdefinedfields-edit.aspx?UDFPK=" & UDFPK)

    End If
  ElseIf bDelete Then
    MyCommon.QueryStr = "delete from UserDefinedField_ValidValues where UDFPK = " & UDFPK
    MyCommon.LRT_Execute()
        
    MyCommon.QueryStr = "delete from UserDefinedFieldsValues where UDFPK = " & UDFPK
    MyCommon.LRT_Execute()
        
    MyCommon.QueryStr = "delete from UserDefinedFields where UDFPK = " & UDFPK
    MyCommon.LRT_Execute()
        
        
    Response.Redirect("userdefinedfields-list.aspx")
  End If
  
  LastUpdate = ""
  
  If Not bCreate Then
    ' No one clicked anything
    MyCommon.QueryStr = "select ExternalID, Description, DataType, AdvancedSearch, LastUpdate,UDFPS_ID " & _
                    "from UserDefinedFields with (NoLock) " & _
                    "where UDFPK=" & UDFPK & ";"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      Description = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
      ExtID = MyCommon.NZ(rst.Rows(0).Item("ExternalID"), "")
      PresentationStyleID = MyCommon.NZ(rst.Rows(0).Item("UDFPS_ID"), -1)
      AdvSearch = MyCommon.NZ(rst.Rows(0).Item("AdvancedSearch"), False)
      DataType = MyCommon.NZ(rst.Rows(0).Item("DataType"), 0)
      UDFPS_ID = MyCommon.NZ(rst.Rows(0).Item("UDFPS_ID"), 0)
      If (IsDBNull(rst.Rows(0).Item("LastUpdate"))) Then
        LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        LastUpdate = Logix.ToLongDateTimeString(rst.Rows(0).Item("LastUpdate"), MyCommon)
      End If
    ElseIf (UDFPK > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.predefinedtriggercodes", LanguageID) & " #" & UDFPK & "</h1>") '********************
      Send("</div>")
      Send("<div id=""main"">")
      Send("  <div id=""infobar"" class=""red-background"">")
      Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
      Send("  </div>")
      Send("</div>")
      GoTo done
    End If
  End If
  
%>
<form action="UserDefinedFields-edit.aspx" id="mainform" name="mainform" method="POST"
onload="document.mainform.ExtID.focus();">
<!--<div>-->
<%
	If UDFPK >0 Then
        Send("<input type=""hidden"" id=""UDFPK""   name=""UDFPK""      value=""" & UDFPK & """ />")
        Send("<input type=""hidden"" id=""UDFPS_ID"" name=""UDFPS_ID""   value=""" & UDFPS_ID & """ />")
    End If
%>
<div id="intro">
  <%
  	MyCommon.QueryStr = "select O.Name 'Name',  o.ProdEndDate 'EndDate', udf.OfferID from userdefinedfieldsvalues as udf " & _
											"join offers as o on udf.offerid = o.offerid where udf.UDFPK = " & UDFPK & _
											"UNION " & _
											"select cpe_i.IncentiveName 'Name', cpe_i.EndDate, udf.OfferID from userdefinedfieldsvalues as udf " & _
													"join cpe_incentives as cpe_i on udf.offerid = cpe_i.incentiveid where udf.UDFPK = " & UDFPK
	dst = MyCommon.LRT_Select 
    Sendb("<h1 id=""title"">")
    If UDFPK = 0 Then
     Sendb(Copient.PhraseLib.Lookup("term.newuserdefinedfield", LanguageID) )
    Else
      Sendb(Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) &"  # "  & UDFPK)
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If UDFPK =0 Then
      If (Logix.UserRoles.AddUserDefinedFields) Then
              Send_Save(" onclick=""return validateSave();"" ")
      End If
    Else
      ShowActionButton = (Logix.UserRoles.EditUserDefinedFields OR Logix.UserRoles.AddUserDefinedFields OR Logix.UserRoles.DeleteUserDefinedFields)''''''''''''''''''''''''''''''
      If (ShowActionButton) Then
        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
        If (Logix.UserRoles.EditUserDefinedFields) Then
                  Send_Save(" onClick=""return validateSave();"" ")
		End If
		If (Logix.UserRoles.DeleteUserDefinedFields) Then
         
		  Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""return deleteUDF(" & dst.Rows.Count & ");"" value=""Delete"">")
		End If
		If (Logix.UserRoles.AddUserDefinedFields) Then
          Send_New()
        End If
        Send("</div>")
      End If

    End If
    Send("</div>")
  %>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="column1">
    <div class="box" id="identification">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<table>")
        Send("<tr>")
        Send("<td><label for=""ExtID"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ":</label></td>")
        Send("<td><input type=""text"" class=""long"" id=""ExtID"" name=""ExtID""  maxlength=""50"" " & "value=""" & ExtID & """ /></td>")
        Send("</tr><tr>")
        Send("<td><label for=""Description"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</label></td>")
        Send("<td><input type=""text"" class=""long"" id=""Description"" name=""Description"" maxlength=""100"" " & "value=""" & HttpUtility.HtmlEncode(Description) & """/></td>")
        Send("</tr><tr>")
        Send("<td><lable for=""DataType"">" & Copient.PhraseLib.Lookup("term.datatype", LanguageID) & ":</label></td>")
        If UDFPK > 0 Then
          Send("<td><select class=""medium"" id=""DataType"" name=""DataType"" >")
          MyCommon.QueryStr = "Select * from UserDefinedFieldsTypes where UDFTypeID = (select DataType from UserDefinedFields where UDFPK=" + Convert.ToString(UDFPK) + ") "
        Else
          Send("<td><select class=""medium"" id=""DataType"" name=""DataType""  onchange=""dataTypeChange();"">")
          MyCommon.QueryStr = "Select * from UserDefinedFieldsTypes"
        End If
        rst2 = MyCommon.LRT_Select
        For Each row2 In rst2.Rows
          If (DataType = row2.Item("UDFTypeID")) Then
            Sendb("<option value=""" & row2.Item("UDFTypeID") & """ selected=""selected"">" & row2.Item("DisplayText") & "</option>")
          Else
            Sendb("<option value=""" & row2.Item("UDFTypeID") & """>" & row2.Item("DisplayText") & "</option>")
          End If
        Next
        Send("</select></td>")
        Send("</tr><tr>")
        Send("<td><label for=""AdvSearch"" >" & Copient.PhraseLib.Lookup("term.enableadvsearch", LanguageID) & "</label></td>")
        Send("<td><input type=""checkbox"" id=""AdvSearch"" name = ""AdvSearch"" " & IIf(AdvSearch, "checked", "") & "/></td>")
			
        Send("</tr></table>")
        'Send("<br />")
        Send("<br class=""half"" />")
    

		If UDFPK > 0 Then
              Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & LastUpdate)
              'Send("<br />")
          End If
      %>
      <hr class="hidden" />
    </div>
    <br />
   <div class="box" id="presentationstylediv">
      <h2>
        <span>
          <% 
            Sendb(Copient.PhraseLib.Lookup("term.presentation", LanguageID))
          %>
        </span>
      </h2>
      <%
        Send("<table>")
        Send("<tr>")

          If UDFPK > 0 Then
              ''if we're pulling up an existing UDF, we'll only get the presentation style assigned to that UDF, once the UDF is saved, we don't want the user to be able to change
              MyCommon.QueryStr = "select * from UserDefinedFieldsPresentationStyles where UDFPS_ID = (select UDFPS_ID from UserDefinedFields where UDFPK=" + Convert.ToString(UDFPK) + ")"
              Send("<td><select class=""medium"" id=""ddPresentationStyle"" name=""ddPresentationStyle"" >")
          Else
              ''This query will only pull valid presentation styles for the given datatype
              MyCommon.QueryStr = "select * from UserDefinedFieldsPresentationStyles where UDFPS_ID in (select distinct UDFPS_ID FROM UserDefinedFields_ValidPresentationStyles where UDFTypeID = " + Convert.ToString(DataType) + ")"
              Send("<td><select class=""medium"" id=""ddPresentationStyle"" name=""ddPresentationStyle"" onchange=""presentationStyleChange();"">")
          End If
          rst2 = MyCommon.LRT_Select
          For Each row2 In rst2.Rows
              If (PresentationStyleID = row2.Item("UDFPS_ID")) Then
                  Sendb("<option id=""ps_" + row2.Item("PresentationStyleID") + """ value=""" & row2.Item("UDFPS_ID") & """ selected=""selected"">" & row2.Item("PresentationStyle") & "</option>")
              Else
                  
                  'if existing presentation style is -1, which will happen when importing an old offer, set default presentation style correctly
                  Dim setSelected As Boolean = False
                  If PresentationStyleID = -1 And DataType = 3 And Convert.ToInt32(row2.Item("UDFPS_ID")) = 5 Then
                      setSelected = True
                  ElseIf PresentationStyleID = -1 And (DataType = 0 Or DataType = 1 Or DataType = 2) And Convert.ToInt32(row2.Item("UDFPS_ID")) = 6 Then
                      setSelected = True
                  End If
                  
                  Sendb("<option id=""ps_" + row2.Item("PresentationStyleID") + """ value=""" & row2.Item("UDFPS_ID") & """ " & IIf(setSelected, " selected=""selected"" ", "") & "     >" & row2.Item("PresentationStyle") & "</option>")
              End If
          Next
          'End If
          
          Send("</select></td>")
          Send("</tr><tr>")
          Send("</tr></table>")
          Send("<br class=""half"" />")
          %>
      <hr class="hidden" />
    </div>
    <div id="datepicker" class="dpDiv">
      </div>
      <%
        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>
      <br />
   <div class="box" id="Values">
      <h2>
        <span>
          <% 
               Sendb(Copient.PhraseLib.Lookup("term.values", LanguageID))
          %>
        </span>
      </h2>
      <%
          Send("<div>")
          Send("<table>")
               'Text box for value 
               'button for "Add new  value"
               Send("<tr>")
          Send("<td>") ' was mediumlong
          
          Dim displayDate As String = "display:none"
          Dim displayNumericRange As String = "display:none"
          Dim displayOther As String = ""
          
          If DataType = 5 Then
              displayDate = "display:none"
              displayOther = "display:none"
              displayNumericRange = ""
          ElseIf DataType = 2 Then
              displayDate = ""
              displayOther = "display:none"
              displayNumericRange = "display:none"
          End If
          
          Sendb("<input type=""text"" style=""" + displayOther + """ class=""mediumlong"" id=""ValueText"" name=""ValueText"" maxlength=""1000"" onkeypress=""return isNumber(this,event);"" value="""">")
          
          Sendb("<div id=""numericRangeValues"" style=""" + displayNumericRange + """ >")
          Sendb("<input type=""text"" class=""short"" id=""ValueTextMin"" name=""ValueTextMin"" maxlength=""1000"" onkeypress=""return isNumber(this,event);"" value="""">")
          Send("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
          Sendb("<input type=""text"" class=""short"" id=""ValueTextMax"" name=""ValueTextMax"" maxlength=""1000"" onkeypress=""return isNumber(this,event);"" value="""">")
          Sendb("</div>")
          
          Sendb("<div id=""dateDiv"" style=""" + displayDate + """ >")
          Sendb("<input class=""short"" id=""dateValue"" name=""dateValue"" maxlength=""10"" type=""text"" value="" ""><img src=""/images/calendar.png"" class=""calendar"" id=""udf-datevalue-picker"" alt=""Date picker"" title=""Date picker"" onclick=""displayDatePicker('dateValue', event);"">")
          Sendb("</div>")
          
          Send("</td>")
          Send("<td>")
          Sendb("<input type=""button"" style=""" + displayOther + """ id=""addNewValue"" name=""addNewValue"" value=""Add new value"" onclick=""addUDFValue(document.getElementById('ValueText').value,false);"">")
          Sendb("<input type=""button"" style=""" + displayNumericRange + """ id=""addNewValueNR"" name=""addNewValueNR"" value=""Add new value"" onclick=""addUDFValue('{' + document.getElementById('ValueTextMin').value + ':' + document.getElementById('ValueTextMax').value + '}',true);"">")
          Sendb("<input type=""button"" style=""" + displayDate + """ id=""addNewDateValue"" name=""addNewDateValue"" value=""Add new value"" onclick=""addUDFValue(document.getElementById('dateValue').value,false);"">")
          Send("</td>")
          Send("</tr>")
          Send("</table>")

          'current value Header
          'Del    Value       Default
          Send("<input type=""hidden"" id=""deleteText"" name=""deleteText"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """/>")
          Send("<input type=""hidden"" id=""moveupText"" name=""moveupText"" value=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """/>")
          Send("<input type=""hidden"" id=""movedownText"" name=""movedownText"" value=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """/>")
          Send("<div id=""itemsToDelete""></div>")
          Send("<table id=""valueTable"">")
          Send("<thead id=""udfvaluehead"">")
          Send("  <tr>")
          Send("    <th style=""width:32px;"">" & Left(Copient.PhraseLib.Lookup("term.delete", LanguageID), 3) & "</th>")
          Send("    <th class=""th-condition"">" & Copient.PhraseLib.Lookup("term.value", LanguageID) & "</th>")
          If DataType = 7 Then
            Send("    <th class=""th-condition"">" & Copient.PhraseLib.Lookup("term.image", LanguageID) & "</th>")
          End If
          Send("    <th class=""th-object"">" & Copient.PhraseLib.Lookup("term.default", LanguageID) & "</th>")
          Send("    <th class=""th-icon"">" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & "</th>")
          Send("    <th class=""th-icon"">" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & "</th>")
          Send("  </tr>")
          Send(" </thead>")
          Send("<tbody id=""valueTableBody"">")
          'foreach value
          MyCommon.QueryStr = "select * from UserDefinedField_ValidValues where UDFPK = " & Convert.ToString(UDFPK) & " order by DisplayOrder asc"
          udfvalidvals = MyCommon.LRT_Select
          For Each udfvalidval In udfvalidvals.Rows
              Dim strUDFVPK As String
              Dim strValue As String
              Dim displayOrder As String
              Dim isDefault As Boolean
              Dim isRange As Boolean
              
              strUDFVPK = System.Convert.ToString(udfvalidval.Item("UDFVPK")).Trim()
              strValue = HttpUtility.HtmlEncode(System.Convert.ToString(udfvalidval.Item("Value")).Trim())
              displayOrder = System.Convert.ToString(udfvalidval.Item("DisplayOrder")).Trim()
              isDefault = System.Convert.ToBoolean(udfvalidval.Item("IsDefault"))
              isRange = False
              'Dim parts As String()
              
           
              If System.Text.RegularExpressions.Regex.Match(strValue, "\{\-*[0-9]+:\-*[0-9]+\}").Captures.Count > 0 Then
                  isRange = True
              End If
              'parts = strValue.Split(":")
              'If parts.Count = 2 Then
              '    isRange = True
              'End If
              
              'delete button, value, default radio
              '   "▼""▲"
              Send("<tr id=""tr_" + strUDFVPK + """ name=""tr_" + strUDFVPK + """ >")
              If DataType = 3 Then
                  Send("<td><input disabled="""" type=""button"" value=""X"" title=""" + Copient.PhraseLib.Lookup("term.delete", LanguageID) + """ name=""ex_" + strUDFVPK + """ id=""ex_" + strUDFVPK + """ class=""ex"" onclick=""deleteUDFValue('" + strUDFVPK + "')""></td>")
              Else
                  Send("<td><input type=""button"" value=""X"" title=""" + Copient.PhraseLib.Lookup("term.delete", LanguageID) + """ name=""ex_" + strUDFVPK + """ id=""ex_" + strUDFVPK + """ class=""ex"" onclick=""deleteUDFValue('" + strUDFVPK + "')""></td>")
              End If

              If DataType = 7 Then
                Dim strValue1 As String = "show-image.aspx?caller=udf&src=" & strValue
                Send("<td> <span style=""display: inline-block; width: 180px;min-width: 20px; max-width: 130px;   overflow-x: scroll; "">" + strValue + "</span></td>")
                Send("<td> <img align=""right"" src=""" + strValue1 + """ id=""Image-" + strUDFVPK + """ width=""50"" height=""50"" alt=""Image not Found"" title=""Click to view full-sized image"" onclick=""showFullSizedImage('" + strValue1 + "');"" /></td>")
              Else
                Send("<td> <span style=""display: inline-block; width: 180px;min-width: 20px; max-width: 180px;   overflow: hidden;   text-overflow: ellipsis;"">" + strValue + "</span></td>")
              End If
              
              If isRange Then
                  Send("<td></td>")
              Else
                  Send("<td><input type=""radio"" name=""defaultUDFValue"" value=""" + strUDFVPK + """ " + If(isDefault, "checked", "") + "></td>")
              End If
              
              Send("<td><input type=""button"" value=""▲"" title=""" + Copient.PhraseLib.Lookup("term.moveup", LanguageID) + """name=""mvu_" + strUDFVPK + """ id=""mvu_" + strUDFVPK + """ onclick=""moverow('" + strUDFVPK + "',-1)""/></td>")
              Send("<td><input type=""button"" value=""▼"" title=""" + Copient.PhraseLib.Lookup("term.movedown", LanguageID) + """name=""mvd_" + strUDFVPK + """ id=""mvd_" + strUDFVPK + """ onclick=""moverow('" + strUDFVPK + "',1)""/></td>")
              Send("</tr>")
              Send("<tr id=""tr_" + strUDFVPK + "H""  name=""tr_" + strUDFVPK + "H"" style=""display:none;""  ><td><input type=""hidden"" name=""valueID_" + strUDFVPK + """ value=""" + strUDFVPK + """></td>")
              Send("<td><input type='hidden' id='value_" + strUDFVPK + "' name='value_" + strUDFVPK + "' value='" + strValue + "' ></td>")
              Send("<td><input type='hidden' id='order_" + strUDFVPK + "' name='order_" + strUDFVPK + "' value='" + displayOrder + "'></td>")
              Send("<td></td>")
              Send("<td></td>")
              Send("</tr>")
          Next
          
          'button clear default
          Send("</tbody>")
          Send("</table>")
          Send("</div>")

      %>
      <hr class="hidden" />
    </div>
  </div>
  <div id="gutter">
  </div>
	<% If UDFPK>0 Then%>
    <div id="column2">
    <div class="box" id="identification">
	<h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
        </span>
      </h2>
	   <div class="boxscroll">
        <%
          If (UDFPK > -1) Then
            If dst.Rows.Count > 0 Then
              For Each row In dst.Rows
                If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                  Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</a>")
                Else
                 Sendb(MyCommon.NZ(row.Item("Name"), ""))
                End If
                If (MyCommon.NZ(row.Item("EndDate"), Today) < Today) Then
                  Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                End If
                Send("<br />")
              Next
            Else
              Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          Else
            Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
          End If
        %>
      </div>
	</div>
	</div>
	<%End If%>
</div>
 </form> 

<%
  If MyCommon.Fetch_SystemOption(156) = "1" Then
    Send("<div id=""imagepopup"" style=""display:none;"">")
    Send("  <div style=""float:right;"">")
    Send("    <a href=""#"" onclick=""javascript:document.getElementById('imagepopup').style.display='none';"">" & Copient.PhraseLib.Lookup("term.close", LanguageID, "Close") & "</a>")
    Send("  </div>")
    Send("  <div id=""imagebody"">")
    Send("    <table id=""centertable""><tr><td><img id=""fullSizedImage"" src="""" /></td></tr></table>")
    Send("  </div>")
    Send("</div>")
  End If
%>  

<script runat="server">
Function htmlDecode(ByVal str as String) As String

    str.replace("&lt;", "<")
    str.replace("&gt;", ">")
	return str
End Function
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

<%

done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "Description")
MyCommon = Nothing
Logix = Nothing
%>
