<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-supplemental-edit.aspx 
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
  Dim FieldID As Integer = 0
  Dim Name As String = ""
  Dim ExtFieldID As String = ""
  Dim FieldTypeID As Integer = 0
  Dim FieldTypeName As String = ""
  Dim Length As Integer = 0
  Dim Visible As Boolean = False
  Dim Editable As Boolean = False
  Dim Deleted As Boolean = False
  Dim LastUpdate As String = ""
    Dim MyCommon As New Copient.CommonInc
    Dim MyCryptLib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim dt As DataTable
  Dim dt2 As DataTable
  Dim dtAssociated As DataTable = Nothing
  Dim row As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim DeleteStatusID As Integer = 0
  Dim i As Integer = 0
  Dim Shaded As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.AppName = "customer-supplemental-edit.aspx"
  Response.Expires = 0
  
  Try
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      FieldID = IIf(Request.QueryString("FieldID") = "", 0, MyCommon.Extract_Val(Request.QueryString("FieldID")))
      ExtFieldID = Left(Logix.TrimAll(Request.QueryString("ExtFieldID")), 50)
      Name = Left(Logix.TrimAll(Request.QueryString("Name")), 100)
      FieldTypeID = IIf(Request.QueryString("FieldTypeID") = "", 0, MyCommon.Extract_Val(Request.QueryString("FieldTypeID")))
      Length = IIf(Request.QueryString("Length") = "", 0, MyCommon.Extract_Val(Request.QueryString("Length")))
      Visible = IIf(Request.QueryString("Visible") = "1", True, False)
      Editable = IIf(Request.QueryString("Editable") = "1", True, False)
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
      FieldID = IIf(Request.Form("FieldID") = "", 0, MyCommon.Extract_Val(Request.Form("FieldID")))
      If FieldID <= 0 Then
        FieldID = IIf(Request.QueryString("FieldID") = "", 0, MyCommon.Extract_Val(Request.QueryString("FieldID")))
      End If
      ExtFieldID = Left(Logix.TrimAll(Request.Form("ExtFieldID")), 50)
      Name = Left(Logix.TrimAll(Request.Form("Name")), 100)
      FieldTypeID = IIf(Request.Form("FieldTypeID") = "", 0, MyCommon.Extract_Val(Request.Form("FieldTypeID")))
      Length = IIf(Request.Form("Length") = "", 0, MyCommon.Extract_Val(Request.Form("Length")))
      Visible = IIf(Request.Form("Visible") = "1", True, False)
      Editable = IIf(Request.Form("Editable") = "1", True, False)
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

    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    
    Send_HeadBegin("term.customersupplemental", , FieldID)
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

  function toggleLength() {
    var type = document.getElementById("FieldTypeID");
    var typespan = document.getElementById("typespan");
    var length = document.getElementById("Length");
    var lengthspan = document.getElementById("lengthspan");

    if (type != null && typespan != null && length != null && lengthspan != null) {
      if (type.value == 1 || type.value == 4) {
        length.value = 0;
        lengthspan.style.display = 'none';
      } else {
        lengthspan.style.display = 'inline';
      }
    }
  }
  
  function toggleVisibility(selectedField) {
    var visible = document.getElementById("Visible");
    var editable = document.getElementById("Editable");

    if (visible != null && editable != null) {
      if (visible.checked == false && editable.checked == true && selectedField == 'Editable') {
        editable.checked = true;
        visible.checked = true;
      } else if (visible.checked == false && editable.checked == true && selectedField == 'Visible') {
        editable.checked = false;
        visible.checked = false;
      }
    }
  }
      
  function isDangerousString()  {
    var stringToCheck = document.getElementById("Name").value;
	var strsupText = document.getElementById("Name")
	var savebutton = document.getElementById("save")
         if ((stringToCheck.indexOf("<") > -1) || (stringToCheck.indexOf(">") > -1))  {
          alert('<% Sendb(Copient.PhraseLib.Lookup("categories.invalidname", LanguageID))%>');
		  strsupText.focus();
		  savebutton.disabled = true;
		  return false;
		} else {
		   savebutton.disabled = false;
           return true;
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
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("customer-supplemental-edit.aspx")
  End If
  
  'Customer usage
  Dim FieldUseCount As Integer = 0
  MyCommon.QueryStr = "select count(CustomerPK) as Customers from CustomerSupplemental as CS with (NoLock) " & _
                      "where FieldID=" & FieldID & " and Deleted=0;"
  dt2 = MyCommon.LXS_Select
  FieldUseCount = dt2.Rows(0).Item("Customers")
  'Details for the top 100 associated customers
  'Mask the AltID last four digits
        MyCommon.QueryStr = "select top 100 CS.CustomerPK , CID.ExtCardID from CustomerSupplemental as CS with (NoLock) " & _
                         "left join CardIDs as CID with (NoLock) on CID.CustomerPK = CS.CustomerPK " & _
                         "where FieldID=" & FieldID & " and Deleted=0 order by CustomerPK;"
  
    dtAssociated = MyCommon.LXS_Select
    Dim rowCnt As Integer = 0
    For Each row In dtAssociated.Rows
        Dim e_ExtCardID As String = ""
        e_ExtCardID = MyCryptLib.SQL_StringDecrypt(dtAssociated.Rows(rowCnt)("ExtCardID").ToString())
        If MyCommon.Fetch_SystemOption(144) Then
            'Mask the AltID last four digits
            'CASE WHEN CID.CardTypeID=3 THEN LEFT(CID.ExtCardID,LEN(CID.ExtCardID)-4) ELSE CID.ExtCardID END AS ExtCardID             
            If (CStr(dtAssociated.Rows(rowCnt)("CardTypeID")) = "3") Then
                e_ExtCardID = e_ExtCardID.Substring(0, e_ExtCardID.Length - 4)
            End If
        End If
        dtAssociated.Rows(rowCnt)("ExtCardID") = e_ExtCardID
        rowCnt = rowCnt + 1
    Next
    
    
    
  
  If bSave Then
    If (Name = "") Then
      infoMessage = Copient.PhraseLib.Lookup("categories.noname", LanguageID)
    ElseIf (Length <= 0) AndAlso (FieldTypeID <> 1 AndAlso FieldTypeID <> 4) Then
      infoMessage = Copient.PhraseLib.Lookup("customer-supplemental-edit.MaxLength", LanguageID)
    Else
      If (FieldID = 0) Then
        MyCommon.QueryStr = "SELECT FieldID FROM CustomerSupplementalFields with (NoLock) WHERE Deleted=0 and Name='" & MyCommon.Parse_Quotes(Name) & "';"
        dt = MyCommon.LXS_Select
        If (dt.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("categories.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "dbo.pt_CustomerSupplementalFields_Insert"
          MyCommon.Open_LXSsp()
          MyCommon.LXSsp.Parameters.Add("@ExtFieldID", SqlDbType.NVarChar, 50).Value = ExtFieldID
          MyCommon.LXSsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = Name
          MyCommon.LXSsp.Parameters.Add("@FieldTypeID", SqlDbType.Int).Value = FieldTypeID
          MyCommon.LXSsp.Parameters.Add("@Length", SqlDbType.Int).Value = Length
          MyCommon.LXSsp.Parameters.Add("@Visible", SqlDbType.Bit).Value = IIf(Visible, 1, 0)
          MyCommon.LXSsp.Parameters.Add("@Editable", SqlDbType.Bit).Value = IIf(Editable, 1, 0)
          MyCommon.LXSsp.Parameters.Add("@FieldID", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LXSsp.ExecuteNonQuery()
          FieldID = MyCommon.LXSsp.Parameters("@FieldID").Value
          MyCommon.Close_LXSsp()
          MyCommon.Activity_Log(47, FieldID, AdminUserID, Copient.PhraseLib.Lookup("history.customersupplemental-create", LanguageID))
          Response.Redirect("customer-supplemental-edit.aspx?FieldID=" & FieldID)
        End If
      Else
        ' update the existing field
        MyCommon.QueryStr = "SELECT FieldID FROM CustomerSupplementalFields WITH (NoLock) WHERE Deleted=0 AND Name='" & MyCommon.Parse_Quotes(Name) & "' AND FieldID<>" & FieldID & ";"
        dt = MyCommon.LXS_Select
        If (dt.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("categories.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "UPDATE CustomerSupplementalFields WITH (RowLock) SET " & _
                              "ExtFieldID='" & MyCommon.Parse_Quotes(ExtFieldID) & "', " & _
                              "Name='" & MyCommon.Parse_Quotes(Name) & "', " & _
                              "FieldTypeID=" & FieldTypeID & ", " & _
                              "Length=" & Length & ", " & _
                              "Visible=" & IIf(Visible, 1, 0) & ", " & _
                              "Editable=" & IIf(Editable, 1, 0) & ", " & _
                              "LastUpdate=getdate() " & _
                              "WHERE FieldID=" & FieldID & ";"
          MyCommon.LXS_Execute()
          MyCommon.Activity_Log(47, FieldID, AdminUserID, Copient.PhraseLib.Lookup("history.customersupplemental-edit", LanguageID))
          Response.Redirect("customer-supplemental-edit.aspx?FieldID=" & FieldID)
        End If
      End If
    End If
    
  ElseIf bDelete Then
    If FieldUseCount > 0 Then
      infoMessage = Copient.PhraseLib.Lookup("categories.inuse", LanguageID)
    Else
      MyCommon.QueryStr = "dbo.pt_CustomerSupplementalFields_Delete"
      MyCommon.Open_LXSsp()
      MyCommon.LXSsp.Parameters.Add("@FieldID", SqlDbType.BigInt).Value = FieldID
      MyCommon.LXSsp.ExecuteNonQuery()
      MyCommon.Close_LXSsp()
      MyCommon.Activity_Log(47, FieldID, AdminUserID, Copient.PhraseLib.Lookup("history.customersupplemental-delete", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "customer-supplemental-list.aspx")
    End If
  End If
  
  LastUpdate = ""
  
  If Not bCreate Then
    ' No one clicked anything
    MyCommon.QueryStr = "select FieldID, ExtFieldID, Name, FieldTypeID, Length, Visible, Editable, LastUpdate " & _
                        "from CustomerSupplementalFields as CS with (NoLock) " & _
                        "where Deleted=0 and FieldID=" & FieldID & ";"
    dt = MyCommon.LXS_Select()
    If (dt.Rows.Count > 0) Then
      FieldID = MyCommon.NZ(dt.Rows(0).Item("FieldID"), 0)
      ExtFieldID = MyCommon.NZ(dt.Rows(0).Item("ExtFieldID"), "")
      Name = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
      FieldTypeID = MyCommon.NZ(dt.Rows(0).Item("FieldTypeID"), 0)
      Length = MyCommon.NZ(dt.Rows(0).Item("Length"), 0)
      Visible = MyCommon.NZ(dt.Rows(0).Item("Visible"), False)
      Editable = MyCommon.NZ(dt.Rows(0).Item("Editable"), False)
      If (IsDBNull(dt.Rows(0).Item("LastUpdate"))) Then
        LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        LastUpdate = Logix.ToLongDateTimeString(dt.Rows(0).Item("LastUpdate"), MyCommon)
      End If
    ElseIf (FieldID > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.customersupplemental", LanguageID) & " #" & FieldID & "</h1>")
      Send("</div>")
      Send("<div id=""main"">")
      Send("  <div id=""infobar"" class=""red-background"">")
      Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
      Send("  </div>")
      Send("</div>")
      GoTo done
    End If
  End If
  
  If FieldTypeID > 0 Then
    MyCommon.QueryStr = "select Name,PhraseID from CustomerSupplementalFieldTypes with (NoLock) WHERE FieldTypeID=" & FieldTypeID & ";"
    dt2 = MyCommon.LXS_Select
    If dt2.Rows.Count > 0 Then
      FieldTypeName = Copient.PhraseLib.Lookup(dt2.Rows(0).Item("PhraseID"), LanguageID)
    End If
  End If
%>
<form action="#" id="mainform" name="mainform">
<input type="hidden" id="FieldID" name="FieldID" value="<% Sendb(FieldID) %>" />
<div id="intro">
  <%
    Sendb("<h1 id=""title"">")
    If FieldID = 0 Then
      Sendb(Copient.PhraseLib.Lookup("term.new", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.customersupplementalfield", LanguageID), VbStrConv.Lowercase))
    Else
      Sendb(Copient.PhraseLib.Lookup("term.customersupplemental", LanguageID) & " #" & FieldID & ": " & MyCommon.TruncateString(Name, 40))
    End If
    Send("</h1>")
    Send("<div id=""controls"">")
    If (FieldID = 0) Then
      If (Logix.UserRoles.EditCustomerSupplementalFields) Then
        Send_Save()
      End If
    Else
      ShowActionButton = (Logix.UserRoles.EditCustomerSupplementalFields)
      If (ShowActionButton) Then
        Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
        Send_Save()
        Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ />")
        Send_New()
        Send("</div>")
      End If
      If MyCommon.Fetch_SystemOption(75) Then
        If (Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(41, FieldID, AdminUserID)
        End If
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
        Send("<label for=""Name"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
        If (Name Is Nothing) Then Name = ""
        Send("<input type=""text"" class=""longest"" id=""Name"" name=""Name"" maxlength=""100"" value=""" & Name.Replace("""", "&quot;") & """ onblur=""javascript:return isDangerousString();"" /><br />")
        Send("<br class=""half"" />")
        Send("<label for=""ExtFieldID"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ":</label><br />")
        If (ExtFieldID Is Nothing) Then ExtFieldID = ""
        Send("<input type=""text"" class=""medium"" id=""ExtFieldID"" name=""ExtFieldID"" maxlength=""50"" value=""" & ExtFieldID.Replace("""", "&quot;") & """ /><br />")
        Send("<br class=""half"" />")
        
        Send("<div style=""display:inline;float:left;position:relative;width:100px;"" id=""typespan"">")
        Send("<label for=""FieldTypeID"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label><br />")
        If FieldUseCount > 0 Then
          Send("<input type=""hidden"" id=""FieldTypeID"" name=""FieldTypeID"" value=""" & FieldTypeID & """ />")
          Send("<select disabled=""disabled"">")
        Else
          Send("<select id=""FieldTypeID"" name=""FieldTypeID"" onchange=""javascript:toggleLength();"">")
        End If
        MyCommon.QueryStr = "select FieldTypeID, Name, PhraseID from CustomerSupplementalFieldTypes with (NoLock);"
        dt2 = MyCommon.LXS_Select
        If dt2.Rows.Count > 0 Then
          For Each row In dt2.Rows
            Sendb("  <option value=""" & MyCommon.NZ(row.Item("FieldTypeID"), 0) & """" & IIf(MyCommon.NZ(row.Item("FieldTypeID"), 0) = FieldTypeID, " selected=""selected""", "") & ">")
            If MyCommon.NZ(row.Item("Name"), "") <> "" Then
              Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"),  LanguageID))
            End If
            Send("</option>")
          Next
        End If
        Send("</select>")
        Send("</div>")
        Send("<div style=""display:inline;float:left;position:relative;width:100px;" & IIf(FieldTypeID = 2 Or FieldTypeID = 3 Or FieldTypeID = 5, "", "display:none;") & """ id=""lengthspan"">")
        Send("<label for=""Length"">" & Copient.PhraseLib.Lookup("term.length", LanguageID) & ":</label><br />")
        If FieldUseCount > 0 Then
          Send("<input type=""hidden"" id=""Length"" name=""Length"" value=""" & Length & """ />")
          Send("<input type=""text"" value=""" & Length & """ style=""width:40px;"" disabled=""disabled"" />")
        Else
          Send("<input type=""text"" id=""Length"" name=""Length"" maxlength=""3"" value=""" & Length & """ style=""width:40px;"" />")
        End If
        Send("</div>")
        Send("<br clear=""left"" />")
        
        Send("<br />")
        Send("<input type=""checkbox"" id=""Visible"" name=""Visible""" & IIf(Visible, " checked=""checked""", "") & " value=""1"" onclick=""javascript:toggleVisibility('Visible');"" /><label for=""Visible"">" & Copient.PhraseLib.Lookup("term.visible", LanguageID) & "</label><br />")
        Send("<br class=""half"" />")
        Send("<input type=""checkbox"" id=""Editable"" name=""Editable""" & IIf(Editable, " checked=""checked""", "") & " value=""1"" onclick=""javascript:toggleVisibility('Editable');"" /><label for=""Editable"">" & Copient.PhraseLib.Lookup("term.editable", LanguageID) & "</label><br />")
        Send("<br class=""half"" />")
        If (FieldID > 0) Then
          Send("<br />")
          Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & LastUpdate)
          Send("<br />")
        End If
      %>
      <hr class="hidden" />
    </div>
  </div>
  
  <div id="gutter">
  </div>
  
  <div id="column2">
    <div class="box" id="customers"<% Sendb(IIf(FieldID=0, " style=""display:none;""", "")) %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.associatedcustomers", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <%
          If (FieldID > 0) Then
            If dtAssociated.Rows.Count > 0 Then
              Sendb(Copient.PhraseLib.Detokenize("customer-supplemental-edit.FieldInUse", LanguageID, dtAssociated.Rows.Count))  'Field is in use by {0} customer(s) with the following cards
              If (dtAssociated.Rows.Count > 100) Then
                Send("; " & Copient.PhraseLib.Lookup("customer-supplemental-edit.First100Shown", LanguageID) & ":<br />")
              Else
                Send(":<br />")
              End If
              Send("<br class=""half"" />")
              For i = 0 To (dtAssociated.Rows.Count - 1)
                If (i > 0) Then
                  If (dtAssociated.Rows(i).Item("CustomerPK") <> dtAssociated.Rows(i - 1).Item("CustomerPK")) Then
                    If Shaded = "shaded" Then
                      Shaded = ""
                    Else
                      Shaded = "shaded"
                    End If
                  End If
                Else
                  If Shaded = "shaded" Then
                    Shaded = ""
                  Else
                    Shaded = "shaded"
                  End If
                End If
                Sendb(" <p style=""margin-bottom:0;"" class=""" & Shaded & """><a href=""customer-general.aspx?CustPK=" & dtAssociated.Rows(i).Item("CustomerPK") & """>")
                Sendb(MyCommon.NZ(dtAssociated.Rows(i).Item("ExtCardID"), "[" & Copient.PhraseLib.Detokenize("customer-supplemental-edit.NoCard", LanguageID, dtAssociated.Rows(i).Item("CustomerPK")) & "]") & "</a></p>")
              Next
            Else
              Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          Else
            Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
          End If
        %>
      </div>
      <hr class="hidden" />
    </div>

    
  </div>
  <br clear="all" />
  
</div>
</form>

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
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(41, FieldID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
End Try
Send_BodyEnd("mainform", "Name")
MyCommon = Nothing
Logix = Nothing
%>
