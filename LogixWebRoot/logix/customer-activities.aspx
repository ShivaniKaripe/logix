<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>

<%
  ' *****************************************************************************
    ' * FILENAME: customer-activities.aspx 
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
  Dim SortDirection As Boolean
  Dim Shaded As String = "shaded"
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim Note As String = ""
  Dim FullName As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim ActivityText As String = ""

  Dim CustNotes(-1) As Copient.CustomerNote
  Dim CustNote As New Copient.CustomerNote
  Dim rst4 As DataTable
  Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
 
  Dim InstalledEngines(-1) As Integer 
  Dim IsCMOnly As Boolean = False
  Dim CPEHHFilter As String = ""
  Dim HHPK As Integer = 0
  Dim HouseholdID As String = ""
  Dim HHCustIdList As New ArrayList(5)
  Dim PaddedExtID As String = StrDup(25, "0")
  Dim CustExtIdList As String = ""
  Dim CustPKs As String = ""
  Dim CustomerTypeID As Integer = 0
  Dim OfferID As Integer = 0
  Dim ClientUserID1 As String = ""
  Dim IDLength As Integer = 0
  Dim CustomerGroupIDs As String() = Nothing
  
  Dim iCmAutoHouseholdCustGrpOptionId As Integer = 24
  Dim bCmAutoHouseholdCustGrpEnabled As Boolean = False
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-activities.aspx"
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
  
  If (MyCommon.Extract_Val(Request.QueryString("CardPK")) > 0) Then
    CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
    ExtCardID = MyLookup.FindExtCardIDFromCardPK(CardPK)
  End If
    
    If (Request.QueryString("SortText") <> "") Then
        SortText = Request.QueryString("SortText")
    End If
  
    If (Request.QueryString("pagenum") = "") Then
        If (Request.QueryString("SortDirection") = True) Then
            SortDirection = False
        ElseIf (Request.QueryString("SortDirection") = False) Then
            SortDirection = True
        Else
            SortDirection = False
        End If
    Else
        SortDirection = Request.QueryString("SortDirection")
    End If
  
  MyCommon.QueryStr = "select CustomerPK, CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
  dt = MyCommon.LXS_Select
  If (dt.Rows.Count > 0) Then
    IsHouseholdID = MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0) = 1
  End If
  
  sSearchQuery = "select CN.CustomerPK, CN.AdminUserID, CN.CreatedDate, CN.Note, CN.FirstName, CN.LastName " & _
                 "from CustomerNotes CN with (NoLock) " & _
                 "where NoteTypeID=1 and CustomerPK=" & CustomerPK
    MyCommon.QueryStr = sSearchQuery
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
    
   
  Send_HeadBegin("term.customer", "term.consnotesandhist", CustomerPK)
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
<form action="customer-activities.aspx" method="post" id="mainform" name="mainform" onsubmit="return handleSubmit();">
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
      <div class="box" id="notehistory">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.consnotesandhist", LanguageID))%>
          </span>
        </h2>
        <div style="border:solid 0px; height:278px; overflow:scroll; overflow-x:hidden;">
          <%
              Dim UseNotesAndActivity As Integer = 0
              CustNotes = Nothing
              
              CustNotes = MyLookup.GetCustomerNotesAndActivity(CustomerPK, "CreatedDate", SortDirection, ReturnCode)
             
              If CustNotes isnot Nothing 
                   If CustNotes.Length > 0 Then
                  
                      Send(Copient.PhraseLib.Lookup("customer-inquiry.notestop30", LanguageID) & "<br />")
                      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.notes", LanguageID) & """>")
                      Send("  <thead>")
                      Send("      <th class=""th-datetime"" scope=""col"">")
                      Send("        <a href=""customer-activities.aspx?CustPK=" & CustomerPK & ")" & "&amp;SortText=CreatedDate&amp;SortDirection=" & SortDirection & """>")
                      Send("          " & Copient.PhraseLib.Lookup("term.created", LanguageID)) 
                      Send("       </a>")
                      If SortText = "CreatedDate" Then
                          If SortDirection = True Then
                              Sendb("<span class=""sortarrow"">&#9660;</span>")
                          Else
                              Sendb("<span class=""sortarrow"">&#9650;</span>")
                          End If
                      Else
                      End If
                      Send("      </th>")
                      Send("      <th class=""th-author"" scope=""col"">" & Copient.PhraseLib.Lookup("term.author", LanguageID) & "</th>")
                      Send("      <th class=""th-note"" scope=""col"">" & Copient.PhraseLib.Lookup("term.note", LanguageID) & "</th>")
                      Send("    </tr>")
                      Send("  </thead>")
                      Send("  <tbody>")
                      i = 0
                      For Each CustNote In CustNotes
                          If i > 29 Then
                              GoTo closenotes
                          Else
                          
                              Send("    <tr" & Shaded & ">")
                              Send("      <td>" & Logix.ToShortDateTimeString(MyCommon.NZ(CustNote.GetCreatedDate, New Date(1900, 1, 1)), MyCommon) & "</td>")
                              If CustNote.GetFirstName = "" OrElse CustNote.GetLastName = "" Then
                                  if (CustNote.GetNoteID = 2 ) then
                                     Send("      <td>Default User</td>") 
                                  else
                                      MyCommon.QueryStr = "select FirstName, LastName from AdminUsers where AdminUserID=" & CustNote.GetAdminUserID & ";"
                                      rst4 = MyCommon.LRT_Select()
                                      Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst4.Rows(0).Item("FirstName"), ""), 25) & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst4.Rows(0).Item("LastName"), ""), 25) & "</td>")
                                  end if  
                              Else
                                  Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(CustNote.GetFirstName, ""), 25) & " " & MyCommon.SplitNonSpacedString(MyCommon.NZ(CustNote.GetLastName, ""), 25) & "</td>")
                              End If
                              Note = MyCommon.SplitNonSpacedString(MyCommon.NZ(CustNote.GetNote, ""), 40)
                              Note = Note.Replace(vbCrLf, "<br />")
                              Send("      <td>" & Note & "</td>")
                              Send("    </tr>")
                              If Shaded = " class=""shaded""" Then
                                  Shaded = ""
                              Else
                                  Shaded = " class=""shaded"""
                              End If
                          End If
                          i = i + 1
                      Next
    closenotes:
                      Send("  </tbody>")
                      Send("</table>")
                  Else
                      Send(Copient.PhraseLib.Lookup("customer.nonotesposted", LanguageID) & "<br />")
                  End If
              Else
                  Send(Copient.PhraseLib.Lookup("customer.nonotesposted", LanguageID) & "<br />")
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
