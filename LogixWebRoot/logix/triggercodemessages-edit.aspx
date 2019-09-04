<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: triggercodemessages-edit.aspx 
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
  Dim TriggerCodeID as Long
  Dim AdminUserID As Long
  Dim ReasonFlag As Long = -1
  Dim Description As String = ""
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
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "triggercodemessages-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
	
	  TriggerCodeID = IIf(Request.QueryString("TriggerCodeID")="",0,MyCommon.Extract_Val(Request.QueryString("TriggerCodeID")))
            ReasonFlag = ValidateReasonFlag(Request.QueryString("ReasonFlag"))
      Description = htmlDecode(Logix.TrimAll(Request.QueryString("Description")))
	  
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
	  
      TriggerCodeID = IIf(Request.Form("TriggerCodeID") = "", 0, MyCommon.Extract_Val(Request.Form("TriggerCodeID")))
            ReasonFlag = ValidateReasonFlag(Request.Form("ReasonFlag"))
      If TriggerCodeID <= 0 Then
        TriggerCodeID = IIf(Request.QueryString("TriggerCodeID") = "", 0, MyCommon.Extract_Val(Request.QueryString("TriggerCodeID")))
      End If
	  
      Description = Logix.TrimAll(Request.Form("Description"))
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
    
    Send_HeadBegin("term.triggercodemessage", , IIF(ReasonFlag>=0, ReasonFlag, ""))
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
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
	
	function submitDesc() {
		
		var text = document.getElementById("desc").value;
		//text = htmlEncode(text);
		document.getElementById("Description").value = text;
		document.getElementById("desc").disabled = true;

	}
	function htmlEncode(str) {
		return String(str)
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
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
    Response.Redirect("triggercodemessages-edit.aspx")
	TriggerCodeID = 0
  End If


  
  If bSave Then

	If  (ReasonFlag < 0 or ReasonFlag >=100) Then 'only allows 2 characters
	  infoMessage = Copient.PhraseLib.Lookup("error.noreasonflag", LanguageID)
	Else If (Description = "") Then
      infoMessage = Copient.PhraseLib.Lookup("error.nodescription", LanguageID) 
	Else
      MyCommon.QueryStr = "select * from predefinedtriggercodemessages where Deleted = 0 and ReasonFlag = " & ReasonFlag
	  rst = MyCommon.LRT_Select()
	  
	 If rst.Rows.Count > 0 and TriggerCodeID =0 Then
		infoMessage =  Copient.PhraseLib.Lookup("error.reasonflaginuse",LanguageID)
	  Else
		  MyCommon.QueryStr = "dbo.pt_PredefinedTriggerCodeMessages_Insert"
		  MyCommon.Open_LRTsp()
		  MyCommon.LRTsp.Parameters.Add("@ReasonFlag", SqlDbType.tinyint).Value = ReasonFlag
		  MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 330).Value = Description
		  MyCommon.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
		  MyCommon.LRTsp.Parameters.Add("@TriggerCodeID", SqlDbType.Int).Direction = ParameterDirection.Output
		  MyCommon.LRTsp.ExecuteNonQuery()
		  TriggerCodeID = MyCommon.LRTsp.Parameters("@TriggerCodeID").Value
		  MyCommon.Close_LRTsp()
		  'MyCommon.Activity_Log(16, OfferCategoryID, AdminUserID, Copient.PhraseLib.Lookup("history.category-create", LanguageID)) ****************************

		  Response.Redirect("triggercodemessages-edit.aspx?TriggerCodeID=" & TriggerCodeID)
	  End If
    End If
    
  ElseIf bDelete Then
    MyCommon.QueryStr = "dbo.pt_PredefinedTriggerCodeMessages_Delete"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@ReasonFlag", SqlDbType.tinyint).Value = ReasonFlag
    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.int).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    Status = MyCommon.LRTsp.Parameters("@Status").Value
    MyCommon.Close_LRTsp()
    Response.Redirect("triggercodemessages-list.aspx")
  End If
  
  LastUpdate = ""
  
  If Not bCreate Then
    ' No one clicked anything
    MyCommon.QueryStr = "select TriggerCodeID, ReasonFlag, Description, LastUpdate " & _
                        "from PredefinedTriggerCodeMessages with (NoLock) " & _
                        "where TriggerCodeID=" & TriggerCodeID & ";"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      Description = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
	  ReasonFlag = MyCommon.NZ(rst.Rows(0).Item("ReasonFlag"),-1)
      If (IsDBNull(rst.Rows(0).Item("LastUpdate"))) Then
        LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        LastUpdate = Logix.ToLongDateTimeString(rst.Rows(0).Item("LastUpdate"), MyCommon)
      End If
    ElseIf (TriggerCodeID > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.predefinedtriggercodes", LanguageID) & " #" & ReasonFlag & "</h1>") 
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
<form action="#" id="mainform" name="mainform" method="post" >
<%
	If TriggerCodeID >0 Then
		Send("<input type=""hidden"" id=""ReasonFlag"" name=""ReasonFlag"" value="""& ReasonFlag &""" />")
	End if
	Send("<input type=""hidden"" id=""TriggerCodeID"" name=""TriggerCodeID"" value="""& TriggerCodeID &""" />")
	Send("<input type=""hidden"" id=""Description"" name=""Description""  />")
%>
<div id="intro">
  <%
    Sendb("<h1 id=""title"">")
    If TriggerCodeID = 0 Then
      Sendb(Copient.PhraseLib.Lookup("term.newtriggercodemessage", LanguageID))
    Else
      Sendb(Copient.PhraseLib.Lookup("term.triggercodemessage", LanguageID) & " #" & ReasonFlag)
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If TriggerCodeID =0 Then
      If (Logix.UserRoles.EditTriggerCodeMessages) Then
        Send_Save("onclick=""submitDesc()""")
      End If
    Else
      ShowActionButton = (Logix.UserRoles.EditTriggerCodeMessages)
      If (ShowActionButton) Then
        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
        If (Logix.UserRoles.EditTriggerCodeMessages) Then
          Send_Save("onclick=""submitDesc()""")
          Send_Delete()
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
		If TriggerCodeID = 0 Then
			Send("<label for=""ReasonFlag"">" & Copient.PhraseLib.Lookup("term.reasonflag", LanguageID) & ":</label><br />") 
			Send("<input type=""text"" class=""longest"" id=""ReasonFlag"" name=""ReasonFlag"" maxlength=""2"" " & IIf(ReasonFlag>-1, "value="""&ReasonFlag & """","") & "/>")
			Send("<br />")
			Send("<br class=""half"" />")
		End if
		Send("<label for=""desc"">" & Copient.PhraseLib.Lookup("term.Description", LanguageID) & ":</label><br />")
       If (Description Is Nothing) Then
          Description = ""
        End If
        Sendb("<textarea  type=""text"" class=""longest"" id=""desc"" name=""desc"" rows=""8""  maxlength=""330"">" & Description & "</textarea>")
        Send("<br />")
        Send("<br class=""half"" />")
      
		If TriggerCodeID > 0 Then
          
          Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & LastUpdate)
          Send("<br />")
        End If
      %>
      <hr class="hidden" />
    </div>
  </div>
  
  <div id="gutter">
  </div>
  
</div>
</form>
<script runat="server">
Function htmlDecode(ByVal str as String) As String
	str.replace("&amp;", "&")
    str.replace("&quot;", """")
    str.replace("&lt;", "<")
    str.replace("&gt;", ">")
	return str
End Function
    Private Function ValidateReasonFlag(ByVal Reason As String) As Integer
        Dim val As Integer
        If Integer.TryParse(Reason, val) Then
            Return val
        Else
            Return -1
        End If
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
