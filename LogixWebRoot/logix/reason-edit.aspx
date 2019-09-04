<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: reason-edit.aspx 
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
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim ReasonID As Integer
  Dim ReasonDescription As String
    Dim Program As String = ""
  Dim LastUpdate As String
  Dim Deleted As Boolean = False
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim row As DataRow
  Dim row2 As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim FocusField As String = "ReasonDescription"
  Dim ReasonNameTitle As String = ""
  Dim SizeOfData As Integer
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim disabledattribute As String = ""
  Dim NewReason As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "reason-edit.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
            ReasonID = MyCommon.Extract_Val(Request.QueryString("ReasonID"))
      ReasonDescription = Request.QueryString("ReasonDescription")
	  Program = Request.QueryString("Program")
      NewReason = (Request.QueryString("NewReason") <> "")
      If ReasonID = 0 Then
        ReasonID = MyCommon.Extract_Val(Request.QueryString("ReasonID"))
      End If
      
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
      ReasonID = Request.Form("ReasonID")
      
      If ReasonID = 0 Then
        ReasonID = MyCommon.Extract_Val(Request.QueryString("ReasonID"))
      End If
	  
      NewReason = (Request.QueryString("NewReason") <> "")
      ReasonDescription = Request.Form("ReasonDescription")
	  Program = Request.Form("Program")
      
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
    
    
    Send_HeadBegin("term.reason", , ReasonID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
 
    Send_Scripts()
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
    
    If (Request.QueryString("new") <> "") Then
      Response.Redirect("reason-edit.aspx?NewReason=True")
    End If
    
    
    If bSave Then
    
      MyCommon.QueryStr = "dbo.pt_AdjustmentReasons_Merge"
      MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@ReasonID", SqlDbType.Int).Value = ReasonID
      MyCommon.LXSsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = ReasonDescription
	    MyCommon.LXSsp.Parameters.Add("@Program", SqlDbType.NVarChar, 12).Value = IIf(Program is Nothing,"",Program)
      ReasonDescription = MyCommon.Parse_Quotes(ReasonDescription)
      MyCommon.LXSsp.Parameters.Add("@Delete", SqlDbType.bit).Value = 0
      If (ReasonDescription = "") Then
        infoMessage = Copient.PhraseLib.Lookup("error.nodescription", LanguageID)
      ElseIf (ReasonID < 11) Then
        infoMessage = Copient.PhraseLib.Lookup("reason.badid", LanguageID)
      ElseIf(NewReason) Then
			MyCommon.QueryStr = "SELECT ReasonID FROM AdjustmentReasons with (NoLock) WHERE ReasonID="& ReasonID &" AND UserDefined=1 "
			rst = MyCommon.LXS_Select
			If (rst.Rows.Count > 0) Then
				infoMessage = Copient.PhraseLib.Lookup("reason.idexists", LanguageID)
			Else
			  MyCommon.LXSsp.ExecuteNonQuery()
			  MyCommon.Close_LXSsp()
                    ' MyCommon.Activity_Log(21, ReasonID, AdminUserID, Copient.PhraseLib.Lookup("history.reason-create", LanguageID))
			End If
	
	   Else
		  MyCommon.LXSsp.ExecuteNonQuery()
		  MyCommon.Close_LXSsp()
                ' MyCommon.Activity_Log(21, ReasonID, AdminUserID, Copient.PhraseLib.Lookup("history.reason-create", LanguageID))
		  
	  End If  
	  
      If infoMessage = "" Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "reason-edit.aspx?ReasonID=" & ReasonID & "&ReasonDescription=" & ReasonDescription)
        'Response.AddHeader("Location", "reason-edit.aspx?")
      End If
    ElseIf bDelete Then
      
        MyCommon.QueryStr = "dbo.pt_AdjustmentReasons_Merge"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ReasonID", SqlDbType.Int).Value = ReasonID
        MyCommon.LXSsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = ""
		MyCommon.LXSsp.Parameters.Add("@Program", SqlDbType.NVarChar, 12).Value = ""
        MyCommon.LXSsp.Parameters.Add("@Delete", SqlDbType.bit).Value = 1
        MyCommon.LXSsp.ExecuteNonQuery()
        MyCommon.Close_LXSsp()
            'MyCommon.Activity_Log(21, ReasonID, AdminUserID, Copient.PhraseLib.Lookup("history.reason-delete", LanguageID))
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "reasons-list.aspx")
        ReasonID = 0
        ReasonDescription = ""    
		Program = ""	
    End If
    
    LastUpdate = ""
    
    REM If Not bCreate Then
      REM ' no one clicked anything
      REM MyCommon.QueryStr = "select ReasonID, Description, LastUpdate " & _
                          REM " from AdjustmentReasons with (nolock) " & _
                          REM "where ReasonID=" & ReasonID & " and Enabled=1 and UserDefined=1"
      REM rst = MyCommon.LXS_Select()
      REM If (rst.Rows.Count > 0) Then
        REM For Each row In rst.Rows
          REM If (ReasonDescription = "") Then
            REM If Not row.Item("Description").Equals(System.DBNull.Value) Then
              REM ReasonDescription = row.Item("Description")
            REM End If
          REM End If
          REM If (LastUpdate = "") Then
            REM If Not row.Item("LastUpdate").Equals(System.DBNull.Value) Then
              REM LastUpdate = row.Item("LastUpdate")
            REM End If
          REM End If
        REM Next
      REM ElseIf (ReasonID > 0) Then
        REM Send("")
        REM Send("<div id=""intro"">")
        REM Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.reason", LanguageID) & " #" & ReasonID & "</h1>")
        REM Send("</div>")
        REM Send("<div id=""main"">")
        REM Send("    <div id=""infobar"" class=""red-background"">")
        REM Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        REM Send("    </div>")
        REM Send("</div>")
        REM GoTo done
      REM End If
    REM End If

%>

<script type="text/javascript">
    
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
	
	
}

</script>

<form action="reason-edit.aspx" id="mainform" name="mainform" onsubmit="return saveForm();">

  <div id="intro">
    <%
      Sendb("<h1 id=""title"">")
      NewReason = (Request.QueryString("NewReason") <> "")
      ReasonID = MyCommon.Extract_Val(Request.QueryString("ReasonID"))
      If NewReason Then
        Sendb(Copient.PhraseLib.Lookup("term.new", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reason", LanguageID))
      Else
        Sendb(Copient.PhraseLib.Lookup("term.reason", LanguageID) & " #" & ReasonID & ": ")
        MyCommon.QueryStr = "SELECT Description FROM AdjustmentReasons with (NoLock) WHERE ReasonID = " & ReasonID & " and UserDefined=1;"
        rst2 = MyCommon.LXS_Select
		MyCommon.QueryStr = "SELECT Program FROM AdjustmentReasons with (NoLock) WHERE ReasonID = " & ReasonID & " and UserDefined=1;"
        rst3 = MyCommon.LXS_Select
        If (rst2.Rows.Count > 0) Then
          'ReasonNameTitle = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
          ReasonDescription = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
        End If
		If (rst3.Rows.Count > 0) Then
			Program = MyCommon.NZ(rst3.Rows(0).Item("Program"), "")
        End If
        Sendb(MyCommon.TruncateString(ReasonNameTitle, 40))
      End If
      Sendb("</h1>")
    %>
	
	
	
    <div id="controls">
	
	
      <%
        If Not Deleted Then
          If (NewReason) Then
            'If (Logix.UserRoles.EditReasons) Then
              Sendb("<input type=""hidden"" class=""longest"" id=""NewReason"" name=""NewReason"" maxlength=""50"" value=""True"" />")
			  Send_Save()
			  
            'End If
          Else
            ShowActionButton = Logix.UserRoles.EditReasonCodes
            If (ShowActionButton) Then
              Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
              Send("<div class=""actionsmenu"" id=""actionsmenu"">")
     
                Send_Save()           
                Send_Delete()             
                Send_New()
              
              Send("</div>")
            End If
            If MyCommon.Fetch_SystemOption(75) Then
              If (Logix.UserRoles.AccessNotes) Then
                Send_NotesButton(26, ReasonID, AdminUserID)
              End If
            End If
          End If
        End If
      %>
	  
    </div>
  </div>
  <div id="main">
    <%
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      If Deleted Then
        GoTo DeleteSkip
      End If
    %>
    <div id="column1">
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <span id="TermCodeSpan">
          <label id="lblReasonID" for="ReasonID" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>:</label><br style="line-height: 0.1;" />
          <%
			If(Not NewReason) Then
                  Sendb("<input type=""text"" class=""longest"" id=""ReasonID"" name=""ReasonID"" maxlength=""9"" value=""" & IIf(ReasonID <> 0, ReasonID, "") & """readonly />")
			Else
                  Sendb("<input type=""text"" class=""longest"" id=""ReasonID"" name=""ReasonID"" maxlength=""9"" value=""" & IIf(ReasonID <> 0, ReasonID, "") & """ />")
			End If	
			%>
          <br class="half" />
        </span>
        <br />
        <br class="half" />
        <label for="desc" style="position: relative;"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" cols="48" rows="3" id="desc" name="ReasonDescription" value="<%IIF(ReasonDescription<> "", ReasonDescription, "")%>"  oninput="limitText(this,100);"><% Sendb(ReasonDescription)%></textarea><br />
        <br class="half" />
        <br class="half" />
		<strong>Applies to:</strong><br>
        <%
          
		If ((Program="Points")) Then
                Sendb("<input type=""radio""  name=""Program"" value=""Points"" checked=""checked"">"& Copient.PhraseLib.Lookup("term.points",LanguageID) &"<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""Stored Value"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""All"">" & Copient.PhraseLib.Lookup("term.all", LanguageID) & "<br></br>")
		Else if(Program="Stored Value")
                Sendb("<input type=""radio""  name=""Program"" value=""Points"" >" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""Stored Value""checked=""checked"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""All"">" & Copient.PhraseLib.Lookup("term.all", LanguageID) & "<br></br>")
		Else if(Program="All")
                Sendb("<input type=""radio""  name=""Program"" value=""Points"" >" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""Stored Value"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""All"" checked=""checked"">" & Copient.PhraseLib.Lookup("term.all", LanguageID) & "<br></br>")
		Else	
			Sendb("<input type=""radio""  name=""Program"" value=""Points"" >"& Copient.PhraseLib.Lookup("term.points",LanguageID) &"<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""Stored Value"">" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "<br>")
                Sendb("<input type=""radio"" name=""Program"" value=""All"" >" & Copient.PhraseLib.Lookup("term.all", LanguageID) & "<br></br>")
		End If
		%>
        
          
           
                        
          
       <%--  <br class="half" />
       <%
          MyCommon.QueryStr = "select ActivityDate from ActivityLog with (NoLock) where ActivityTypeID='21' and LinkID='" & ReasonID & "' order by ActivityDate asc;"
          dst = MyCommon.LRT_Select
          SizeOfData = dst.Rows.Count
          If SizeOfData > 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
            Send(Logix.ToLongDateTimeString(dst.Rows(0).Item("ActivityDate"), MyCommon))
            Send("<br />")
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
            Send(Logix.ToLongDateTimeString(dst.Rows(SizeOfData - 1).Item("ActivityDate"), MyCommon))
          End If
        %>
        <hr class="hidden" />--%>
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    
      
   
    <br clear="all" />
    <% DeleteSkip:%>
  </div>
</form>

<script type="text/javascript">

function saveForm(){
    var Pselected = document.getElementById('Pselected');
    return true;
}	
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

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function PhandleKeyUp(maxNumToShow) {
  var selectObj, textObj, PfunctionListLength;
  var i, numShown;
  var searchPattern;
  var selectedList;
  var elem = document.getElementById("Pfunctionselect");
    
  if (elem != null)
  {

    elem.size = "10";
    
    // Set references to the form elements
    selectObj = document.forms[0].Pfunctionselect;
    textObj = document.forms[0].Pfunctioninput;
    selectedList = document.getElementById("Pselected");

    // Remember the function list length for loop speedup
    PfunctionListLength = Pfunctionlist.length;
    
    // Set the search pattern depending
    if(document.forms[0].Pfunctionradio[0].checked == true)
    {
        searchPattern = "^"+textObj.value;
    }
    else
    {
        searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);

    // Create a regular expression

    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;

    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < PfunctionListLength; i++)
    {
        if(Pfunctionlist[i].search(re) != -1)
        {
            if (Pvallist[i] != "" && (selectedList.options.length < 1 || Pvallist[i] != selectedList.options[0].value) ) {
                selectObj[numShown] = new Option(Pfunctionlist[i],Pvallist[i]);
                numShown++;
            }
        }
        // Stop when the number to show is reached
        if(numShown == maxNumToShow)
        {
            break;
        }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1)
    {
        selectObj.options[0].selected = true;
    }
  }
}

function PremoveUsed()
{
    PhandleKeyUp(99999);
    // this function will remove items from the functionselect box that are used in 
    // selected and excluded boxes

    var funcSel = document.getElementById('Pfunctionselect');
    var elSel = document.getElementById('Pselected');
    var i,j;
  
    for (i = elSel.length - 1; i>=0; i--) {
        for(j=funcSel.length-1;j>=0; j--) {
            if(funcSel.options[j].value == elSel.options[i].value){
                funcSel.options[j] = null;
            }
        }
    }
}


// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function PhandleSelectClick(itemSelected)
{
    textObj = document.forms[0].Pfunctioninput;
     
    selectObj = document.forms[0].Pfunctionselect;
    selectedValue = document.getElementById("Pfunctionselect").value;
    if(selectedValue != ""){ selectedText = selectObj[document.getElementById("Pfunctionselect").selectedIndex].text; }
    
    selectboxObj = document.forms[0].Pselected;
    selectedboxValue = document.getElementById("Pselected").value;
    if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("Pselected").selectedIndex].text; }
    
    if(itemSelected == "Pselect1") {
        if(selectedValue != ""){
            // add items to selected box
            if(selectedValue == 1) {
                document.getElementById('Pselect1').disabled=true;
                // someone's adding all customers we need to empty the select box
                for (i = selectboxObj.length - 1; i>=0; i--) {
                    selectboxObj.options[i] = null;
                }
            }
            document.getElementById('Pdeselect1').disabled=false;
            document.getElementById('Pselect1').disabled=true;
            selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
        }
    }
    
    if(itemSelected == "Pdeselect1") {
        if(selectedboxValue != ""){
            // remove items from selected box
            document.getElementById("Pselected").remove(document.getElementById("Pselected").selectedIndex)
            if(selectedboxValue == 1) {
                document.getElementById('Pselect1').disabled=false;
            }
            if(selectboxObj.length == 0) {
                // nothing in the select box so disable deselect
                document.getElementById('Pdeselect1').disabled=true;
            }
        }
        if (document.getElementById("Pselected").options.length == 0) {
          document.getElementById('Pselect1').disabled=false;
        }
    }
    
    // remove items from large list that are in the other lists
    PremoveUsed();
    return true;
}

</script>

<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
    PhandleKeyUp(99999);
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (ReasonID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(26, ReasonID, AdminUserID)
    End If
  End If
  
done:
Finally
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
End Try
Send_BodyEnd("mainform", "ReasonDescription")
MyCommon = Nothing
Logix = Nothing
%>