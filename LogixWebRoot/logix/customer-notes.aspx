<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-notes.aspx 
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
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim MyLookup As New Copient.CustomerLookup
  Dim dt As DataTable
  Dim dt2 As DataTable
  Dim dtHistory As DataTable
  Dim row As DataRow
  Dim UserTable As New Hashtable(50)
  Dim CustomerPK As Long = 0
  Dim CardPK As Long = 0
  Dim ExtCardID As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim bIsErrorMsg As Boolean
  Dim sSearchQuery As String
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer
  Dim SortText As String = "CreatedDate"
  Dim SortDirection As String
  Dim Shaded As String = "shaded"
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim Note As String = ""
  Dim FullName As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim ActivityText As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-notes.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
  If CustomerPK = 0 Then
    CustomerPK = MyCommon.Extract_Val(Request.Form("CustPK"))
  End If
  
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  If CardPK = 0 Then
    CardPK = MyCommon.Extract_Val(Request.Form("CardPK"))
  End If
  If CardPK > 0 Then
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
  End If
  
  Note = MyCommon.Parse_Quotes(Request.Form("custnote"))
  Note = Logix.TrimAll(Note)
  If (Note = "") Then
    Note = Request.QueryString("custnote")
  End If
  
  
  
  If (Request.QueryString("SortText") <> "") Then
    SortText = Request.QueryString("SortText")
  End If
  
  If (Request.QueryString("pagenum") = "") Then
    If (Request.QueryString("SortDirection") = "ASC") Then
      SortDirection = "DESC"
    ElseIf (Request.QueryString("SortDirection") = "DESC") Then
      SortDirection = "ASC"
    Else
      SortDirection = "DESC"
    End If
  Else
    SortDirection = Request.QueryString("SortDirection")
  End If
  
  If (Request.Form("custnote") <> "") Then
    MyCommon.QueryStr = "select FirstName, LastName from AdminUsers where AdminUserID=" & AdminUserID & ";"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      FirstName = MyCommon.NZ(dt.Rows(0).Item("FirstName"), "")
      LastName = MyCommon.NZ(dt.Rows(0).Item("LastName"), "")
    End If
    If Note <> "" Then
      If Note.Length > 1000 Then
        bIsErrorMsg = True
        InfoMessage = Copient.Lookup("error.notelength", LanguageID)
      Else
        MyCommon.QueryStr = "insert into CustomerNotes with (RowLock) (CustomerPK, AdminUserID, CreatedDate, NoteTypeID, Note, FirstName, LastName, LanguageID) " & _
                            " values (" & CustomerPK & ", " & AdminUserID & ", getDate(), 1, '" & Note & "', '" & FirstName & "', '" & LastName & "'," & LanguageID & ");"
        MyCommon.LXS_Execute()
        ActivityText = Copient.PhraseLib.Lookup("history.customer-added-note", LanguageID) & ": " & Note
        If ActivityText.Length > 1000 Then
          ActivityText = Left(ActivityText, 997) & "..."
        End If
        MyCommon.Activity_Log2(25, 8, CustomerPK, AdminUserID, ActivityText)
      End If
    End If
  End If
  
  MyCommon.QueryStr = "select CustomerPK, CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
  dt = MyCommon.LXS_Select
  If (dt.Rows.Count > 0) Then
    IsHouseholdID = MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0) = 1
  End If
  
  sSearchQuery = "select CN.CustomerPK, CN.AdminUserID, CN.CreatedDate, CN.Note, CN.FirstName, CN.LastName " & _
                 "from CustomerNotes CN with (NoLock) " & _
                 "where NoteTypeID=1 and CustomerPK=" & CustomerPK
  MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection
  dtHistory = MyCommon.LXS_Select
  sizeOfData = dtHistory.Rows.Count
  i = linesPerPage * PageNum
  
  MyCommon.QueryStr = "select AdminUserID, FirstName, LastName from AdminUsers with (NoLock);"
  dt = MyCommon.LRT_Select
  For Each row In dt.Rows
    FullName = MyCommon.NZ(row.Item("FirstName"), "") & " "
    FullName += MyCommon.NZ(row.Item("LastName"), "")
    UserTable.Add(row.Item("AdminUserID"), FullName)
  Next
  
  Send_HeadBegin("term.customer", "term.notes", CustomerPK)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  Send("  if (opener != null && !opener.closed) { ")
  Send("    if (opener.location.href.indexOf('CAM') > -1) { ")
  Send("      opener.location = '/logix/CAM/CAM-customer-general.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "'; ")
  Send("    } else { ")
  Send("      opener.location = '/logix/customer-general.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "'; ")
  Send("    }")
  Send("  }")
  Send("}")
  Send("function handleSubmit() { ")
  Send("  var retVal = true;")
  Send("  var elem = document.getElementById('custnote');")
  Send("   ")
  Send("    if (elem != null) { ")
  Send("      retVal = (elem.value.length <= 1000); ")
  Send("      if (!retVal) { alert('" & Copient.PhraseLib.Lookup("customer-notes.MaximumLength", LanguageID) & "'); }")
  Send("    }")
  Send("  return retVal;")
  Send("}")
  Send("</script>")
  
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.ViewCustomerNotes = False) Then
    Send_Denied(2, "perm.customers-access-notes")
    GoTo done
  End If
%>
<form action="customer-notes.aspx" method="post" id="mainform" name="mainform" onsubmit="return handleSubmit();">
  <div id="intro">
    <%
      Send("<input type=""hidden"" id=""CustPK"" name=""CustPK"" value=""" & CustomerPK & """ />")
      If CardPK > 0 Then
        Send("<input type=""hidden"" id=""CardPK"" name=""CardPK"" value=""" & CardPK & """ />")
      End If
      Send("<h1 id=""title"">")
      If CardPK = 0 Then
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
        End If
      Else
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & ExtCardID)
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
        End If
      End If
      MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
      dt2 = MyCommon.LXS_Select
      If dt2.Rows.Count > 0 Then
        FullName = IIf(MyCommon.NZ(dt2.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, MyCommon.NZ(dt2.Rows(0).Item("Prefix"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, MyCommon.NZ(dt2.Rows(0).Item("FirstName"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(dt2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, MyCommon.NZ(dt2.Rows(0).Item("LastName"), ""), "")
        FullName &= IIf(MyCommon.NZ(dt2.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & MyCommon.NZ(dt2.Rows(0).Item("Suffix"), ""), "")
      End If
      If FullName <> "" Then
        Sendb(": " & MyCommon.TruncateString(FullName, 30))
      End If
      Send("</h1>")
    %>
    <div id="controls">
      <%
        If (Logix.UserRoles.AddCustomerNotes) Then
          Send_Save()
        End If
      %>
    </div>
  </div>
  <div id="main">
    <%
      If (InfoMessage <> "" And bIsErrorMsg) Then
        Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")
      ElseIf (InfoMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & InfoMessage & "</div>")
      End If
    %>
    <div id="column">
      <% If (Logix.UserRoles.AddCustomerNotes) Then%>
      <div class="box" id="newnote">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.addnote", LanguageID))%>
          </span>
        </h2>
        <br class="half" />
        <center>
          <textarea rows="7" cols="75" id="custnote" name="custnote"></textarea>
        </center>
        <hr class="hidden" />
      </div>
      <% End If%>
      <div class="box" id="notehistory">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.notehistory", LanguageID))%>
          </span>
        </h2>
        <div style="border:solid 0px; height:278px; overflow:scroll; overflow-x:hidden;">
          <%
            If (dtHistory.Rows.Count > 0) Then
              Send("<table summary=""" & Copient.PhraseLib.Lookup("term.notehistory", LanguageID) & """>")
              Send("  <thead>")
              Send("    <tr>")
              Send("      <th class=""th-datetime"" scope=""col"">")
              Send("        <a href=""customer-notes.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;SortText=CreatedDate&amp;SortDirection=" & SortDirection & """>")
              Send("          " & Copient.PhraseLib.Lookup("term.created", LanguageID))
              Send("        </a>")
              If SortText = "CreatedDate" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
              Send("      </th>")
              Send("      <th class=""th-author"" scope=""col"">")
              Send("        <a href=""customer-notes.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;SortText=FirstName&amp;SortDirection=" & SortDirection & """>")
              Send("          " & Copient.PhraseLib.Lookup("term.author", LanguageID))
              Send("        </a>")
              If SortText = "FirstName" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
              Send("      </th>")
              Send("      <th class=""th-note"" scope=""col"">")
              Send("        <a href=""customer-notes.aspx?CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "") & "&amp;SortText=Note&amp;SortDirection=" & SortDirection & """>")
              Send("          " & Copient.PhraseLib.Lookup("term.note", LanguageID))
              Send("        </a>")
              If SortText = "Note" Then
                If SortDirection = "ASC" Then
                  Sendb("<span class=""sortarrow"">&#9660;</span>")
                Else
                  Sendb("<span class=""sortarrow"">&#9650;</span>")
                End If
              Else
              End If
              Send("      </th>")
              Send("    </tr>")
              Send("  </thead>")
              Send("  <tbody>")
              For Each row In dtHistory.Rows
                If UserTable.ContainsKey(MyCommon.NZ(row.Item("AdminUserID"), "")) Then
                  FullName = UserTable.Item(MyCommon.NZ(row.Item("AdminUserID"), ""))
                Else
                  FullName = ""
                End If
                Note = MyCommon.NZ(row.Item("Note"), "&nbsp;")
                Note = Note.Replace(vbCrLf, "<br />")
                Send("    <tr class=""" & Shaded & """ >")
                If (Not IsDBNull(row.Item("CreatedDate"))) Then
                  Send("      <td>" & Logix.ToShortDateTimeString(row.Item("CreatedDate"), MyCommon) & "</td>")
                Else
                  Send("      <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                End If
                Send("      <td>" & MyCommon.SplitNonSpacedString(FullName, 25) & "</td>")
                Send("      <td>" & MyCommon.SplitNonSpacedString(Note, 40) & "</td>")
                Send("    </tr>")
                Shaded = IIf(Shaded = "shaded", "", "shaded")
              Next
              Send("  </tbody>")
              Send("</table>")
            Else
              Send("<i>" & Copient.PhraseLib.Lookup("customer.nonotesposted", LanguageID) & "</i>")
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>
<%
done:
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "custnote")
  Logix = Nothing
  MyCommon = Nothing
%>
