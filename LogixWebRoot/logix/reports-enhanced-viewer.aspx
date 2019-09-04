<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%
  ' *****************************************************************************
  ' * FILENAME: reports-enhanced-viewer.aspx 
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
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim AdminUserID As Long
  Dim dt As System.Data.DataTable
  Dim shaded As String = ""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Response.Expires = 0
  MyCommon.AppName = "reports-enhanced-viewer.aspx"
  
  Send_HeadBegin("term.reports")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send("<style type=""text/css"">")
  Send("  html {")
  Send("    overflow: auto;")
  Send("  }")
  Send("  td {")
  Send("    padding-right:25px;")
  Send("    text-align:left;")
  Send("    vertical-align:top;")
  Send("  }")
  Send("</style>")
  Send_Scripts()
  Send_HeadEnd()
  Send_PageBegin(7)
  
  If (Logix.UserRoles.AccessReports = False) Then
    Send_Denied(1, "perm.admin-reports")
    GoTo done
  End If
  
  Dim ReportID As Integer = 0
  If (Request.QueryString("ReportID") <> "") Then
    ReportID = Request.QueryString("ReportID")
  End If
  Send("<div class=""noprint"" style=""position:absolute;top:1px;right:1px;width:120px;"">")
  Send("  <form action=""#"">")
  Send("    <input type=""button"" class=""regular"" id=""print"" name=""print"" value=""" & Copient.PhraseLib.Lookup("term.print", LanguageID) & """ onclick=""javascript:window.print();"" />")
  Send("  </form>")
  Send("</div>")
  
  MyCommon.QueryStr = "select R.*, RT.Name as ReportType, RT.PhraseID as ReportTypePhraseID, AU.UserName, AU.FirstName, AU.LastName " & _
                      "from Reports as R with (NoLock) " & _
                      "left join ReportTypes as RT on R.ReportTypeID=RT.ReportTypeID " & _
                      "left join AdminUsers as AU on R.AdminUserID=AU.AdminUserID " & _
                      "where R.ReportID=" & ReportID & ";"
  dt = MyCommon.LRT_Select
  If dt.Rows.Count = 0 Then
    'Report record cannot be found
    Sendb("<p>" & Copient.PhraseLib.Lookup("term.report", LanguageID) & " " & ReportID & ": ")
    Send(Copient.PhraseLib.Lookup("reports.norecord", LanguageID) & "</p>")
  Else
    If dt.Rows(0).Item("Deleted") = 1 Then
      'Report record exists but is marked as deleted
      Sendb("<p>" & Copient.PhraseLib.Lookup("term.report", LanguageID) & " " & ReportID & ": ")
      Send(Copient.PhraseLib.Lookup("reports.deletedrecord", LanguageID) & "</p>")
    Else
      Try
        'Report record available
        Dim FilePath As String = Trim(MyCommon.Fetch_SystemOption(114))
        If Not (Right(FilePath, 1) = "\") Then
          FilePath = FilePath & "\"
        End If
        FilePath &= MyCommon.NZ(dt.Rows(0).Item("FileName"), "")
        Dim FileInfo As System.IO.FileInfo = New System.IO.FileInfo(FilePath)
        If Not FileInfo.Exists Then
          'Report file cannot be found
          Sendb("<p>" & Copient.PhraseLib.Lookup("term.report", LanguageID) & " " & ReportID & ": ")
          Send(Copient.PhraseLib.Lookup("reports.nofile", LanguageID) & " (" & FilePath & ")</p>")
        Else
          'Report file is available, so generate a header from the Reports record
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.report", LanguageID) & " " & ReportID & " " & StrConv(Copient.PhraseLib.Lookup("term.header", LanguageID), VbStrConv.Lowercase) & """>")
          Send("  <tr>")
          Send("    <td colspan=""2""><h2>" & Copient.PhraseLib.Lookup(MyCommon.NZ(dt.Rows(0).Item("ReportTypePhraseID"), 0), LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.report", LanguageID), VbStrConv.Lowercase) & "</h2></td>")
          Send("  </tr>")
          Send("  <tr>")
          Send("    <td>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</td>")
          Send("    <td>" & MyCommon.NZ(dt.Rows(0).Item("Name"), Copient.PhraseLib.Lookup("term.unnamed", LanguageID)) & "</td>")
          Send("  </tr>")
          Send("  <tr>")
          Send("    <td>" & Copient.PhraseLib.Lookup("term.file", LanguageID) & ":</td>")
          Send("    <td>" & FilePath & "</td>")
          Send("  </tr>")
          Send("  <tr>")
          Send("    <td>" & Copient.PhraseLib.Lookup("term.generated", LanguageID) & ":</td>")
          Sendb("    <td>")
          If IsDBNull(dt.Rows(0).Item("Updated")) Then
            Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
          Else
            Sendb(Logix.ToLongDateTimeString(dt.Rows(0).Item("Updated"), MyCommon))
          End If
          Send("</td>")
          Send("  </tr>")
                    If MyCommon.NZ(dt.Rows(0).Item("ReportTypeID"), 0) = 2 Then
                        Dim engineID As Integer = MyCommon.NZ(dt.Rows(0).Item("EngineID"), 0)
                        Dim ct As System.Data.DataTable
                        Send("<tr>")
                        Send("    <td>" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</td>")
                        Sendb("    <td>")
                        MyCommon.QueryStr = "select PhraseID from PromoEngines  with (NoLock) where engineID=" & MyCommon.NZ(dt.Rows(0).Item("engineID"), -1) & ";"
                        ct = MyCommon.LRT_Select
                        If ct.Rows.Count > 0 Then
                            Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(ct.Rows(0).Item("PhraseID"), 0), LanguageID) & " ")
                        End If
                        Send("</td>")
                        Send("</tr>")
                        Send("<tr>")
                        Send("    <td>" & Copient.PhraseLib.Lookup("term.externalsource", LanguageID) & ":</td>")
                        Sendb("    <td>")
                        MyCommon.QueryStr = "select Name from ExtCRMInterfaces where ExtInterfaceID = " & MyCommon.NZ(dt.Rows(0).Item("ExtInterfaceID"), -1) & ";"
                        ct = MyCommon.LRT_Select
                        If ct.Rows.Count > 0 Then
                            Sendb(MyCommon.NZ(ct.Rows(0).Item("Name"), ""))
                        End If
                        Send("</td>")
                        Send("</tr>")
                        Send("</tr>")
                        Send("<tr>")
                        Send("    <td>" & Copient.PhraseLib.Lookup("term.Offer", LanguageID) & Copient.PhraseLib.Lookup("term.Status", LanguageID) & ":</td>")
                        Sendb("    <td>")
                        Sendb(MyCommon.NZ(dt.Rows(0).Item("OfferStatus"), -1))
                        Send("</td>")
                        Send("</tr>")
                        Send("  <tr>")
                        Send("    <td>" & Copient.PhraseLib.Lookup("term.daterange", LanguageID) & ":</td>")
                        Sendb("    <td>")
                        If IsDBNull(dt.Rows(0).Item("StartDate")) AndAlso IsDBNull(dt.Rows(0).Item("EndDate")) Then
                            Sendb(Copient.PhraseLib.Lookup("term.all", LanguageID))
                        Else
                            If Not IsDBNull(dt.Rows(0).Item("StartDate")) Then
                                Sendb(Logix.ToShortDateString(dt.Rows(0).Item("StartDate"), MyCommon))
                            Else
                                Sendb(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))
                            End If
                            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.through", LanguageID), VbStrConv.Lowercase) & " ")
                            If Not IsDBNull(dt.Rows(0).Item("EndDate")) Then
                                Sendb(Logix.ToShortDateString(dt.Rows(0).Item("EndDate"), MyCommon))
                            End If
                        End If
                        Send("</td>")
                        Send("  </tr>")
                End If
          If MyCommon.NZ(dt.Rows(0).Item("ReportTypeID"), 0) = 3 Or MyCommon.NZ(dt.Rows(0).Item("ReportTypeID"), 0) = 4 Then
            Send("  <tr>")
            Send("    <td>" & Copient.PhraseLib.Lookup("term.storegroups", LanguageID) & ":</td>")
            Sendb("    <td>")
            If MyCommon.NZ(dt.Rows(0).Item("AllLocations"), False) Then
              Sendb(Copient.PhraseLib.Lookup("term.all", LanguageID))
            Else
              Sendb(MyCommon.NZ(dt.Rows(0).Item("LocationGroupIDs"), ""))
            End If
            Send("</td>")
            Send("  </tr>")
          End If
          If MyCommon.NZ(dt.Rows(0).Item("ReportTypeID"), 0) = 5 Then
              Send("  <tr>")
              Send("    <td>" & Copient.PhraseLib.Lookup("term.Customer Groups", LanguageID) & ":</td>")
              Sendb("    <td>")
              If MyCommon.NZ(dt.Rows(0).Item("allCustomerGroups"), False) Then
                  Sendb(Copient.PhraseLib.Lookup("term.all", LanguageID))
              Else
                  Sendb(MyCommon.NZ(dt.Rows(0).Item("CustomerGroupIDs"), ""))
              End If
              Send("</td>")
              Send("  </tr>")
              Dim engineID As Integer = MyCommon.NZ(dt.Rows(0).Item("EngineID"), 0)
              Dim ct As System.Data.DataTable
              Send("<tr>")
              Send("    <td><span class=""title"">" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</span></td>")
              Sendb("    <td>")
              MyCommon.QueryStr = "select PhraseID from PromoEngines  with (NoLock) where engineID=" & MyCommon.NZ(dt.Rows(0).Item("engineID"), -1) & ";"
              ct = MyCommon.LRT_Select
              If ct.Rows.Count > 0 Then
                  Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(ct.Rows(0).Item("PhraseID"), 0), LanguageID) & " ")
              End If
              Send("</td>")
              Send("</tr>")
          End If
          If MyCommon.NZ(dt.Rows(0).Item("ReportTypeID"), 0) <> 2 Then
            Send("  <tr>")
            Send("    <td>" & Copient.PhraseLib.Lookup("term.daterange", LanguageID) & ":</td>")
            Sendb("    <td>")
            If IsDBNull(dt.Rows(0).Item("StartDate")) AndAlso IsDBNull(dt.Rows(0).Item("EndDate")) Then
              Sendb(Copient.PhraseLib.Lookup("term.all", LanguageID))
            Else
              If Not IsDBNull(dt.Rows(0).Item("StartDate")) Then
                Sendb(Logix.ToShortDateString(dt.Rows(0).Item("StartDate"), MyCommon))
              Else
                Sendb(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))
              End If
              Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.through", LanguageID), VbStrConv.Lowercase) & " ")
              If Not IsDBNull(dt.Rows(0).Item("EndDate")) Then
                Sendb(Logix.ToShortDateString(dt.Rows(0).Item("EndDate"), MyCommon))
              End If
            End If
            Send("</td>")
            Send("  </tr>")
          End If
          If MyCommon.NZ(dt.Rows(0).Item("ReportTypeID"), 0) = 1 Then
            Send("  <tr>")
            Send("    <td>" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & ":</td>")
            Sendb("    <td>")
            Dim ct As System.Data.DataTable
            MyCommon.QueryStr = "select TypeID, Description, PhraseID from CustomerTypes with (NoLock) where TypeID=" & MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0) & ";"
            ct = MyCommon.LXS_Select
            If ct.Rows.Count > 0 Then
              Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(ct.Rows(0).Item("PhraseID"), 0), LanguageID) & " ")
            End If
            Sendb(MyCommon.NZ(dt.Rows(0).Item("CustomerID"), ""))
            Send("</td>")
            Send("  </tr>")
          End If
          Send("</table>")
          'Read the CSV and generate the main report content
          Dim Contents As String = File.ReadAllText(FilePath)
          Dim rows() As String
          Dim cols() As String
          Dim i As Integer = 0
          Dim j As Integer = 0
          rows = Contents.Split(ControlChars.CrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
          If (rows.Length > 0) Then
            cols = rows(0).Split(",".ToCharArray, StringSplitOptions.None)
            Send("<hr />")
            Send("<table style=""width:100%;"" summary=""" & Copient.PhraseLib.Lookup("term.report", LanguageID) & " " & ReportID & """>")
          If MyCommon.NZ(dt.Rows(0).Item("ReportTypeID"), 0) = 5 Then
            Send("  <thead>")
            Send("    <tr>")
            For j = 0 To cols.GetUpperBound(0)
                 If cols(j) = "OfferName" Then
                    Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("reports.offeridoffername", LanguageID) & "</th>")
                 Else
                    Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term." & cols(j), LanguageID) & "</th>")
                 End If
            Next
            Send("    </tr>")
            Send("  </thead>")
            Send("  <tbody>")
            Dim PrevExtProdID As String = ""
            For i = 1 To rows.GetUpperBound(0)
                cols = rows(i).Split(",".ToCharArray, StringSplitOptions.None)
                Send("    <tr" & shaded & ">")
            
                For j = 0 To cols.GetUpperBound(0)
                    If (cols(j).Contains("|")) Then
                        cols(j) = cols(j).Replace("|", "<br/>")
                    End If
                    If ((j = 0 AndAlso PrevExtProdID = cols(j))) Then
                        Send("      <td></td>")
                    Else
                        Send("        <td>" & cols(j) & "</td>")
                    End If
                    If (j = 0) Then
                        PrevExtProdID = cols(j)
                    End If
                Next
                Send("    </tr>")
                If shaded = "" Then
                    shaded = " class=""shaded"""
                Else
                    shaded = ""
                End If
            Next
            Send("  </tbody>")
           Else
              'All other report types (CD, OM, OS)
              Send("  <thead>")
              Send("    <tr>")
              For j = 0 To cols.GetUpperBound(0)
                Send("      <th scope=""col"">" & Copient.PhraseLib.Lookup("term." & cols(j), LanguageID) & "</th>")
              Next
              Send("    </tr>")
              Send("  </thead>")
              Send("  <tbody>")
              For i = 1 To rows.GetUpperBound(0)
                cols = rows(i).Split(",".ToCharArray, StringSplitOptions.None)
                Send("    <tr" & shaded & ">")
                For j = 0 To cols.GetUpperBound(0)
                  Send("      <td>" & cols(j) & "</td>")
                Next
                Send("    </tr>")
                If shaded = "" Then
                  shaded = " class=""shaded"""
                Else
                  shaded = ""
                End If
              Next
              Send("  </tbody>")
            End If
            Send("</table>")
        End If
	   End If
      Catch pathEX As PathTooLongException
        MyCommon.Error_Processor("The file path found in report system option 'Report file path'(114) is too long." & vbCrLf, pathEX.ToString() & vbCrLf & pathEX.StackTrace, "reports-enhanced-viewer.aspx", , )
      Catch ex As Exception
        MyCommon.Error_Processor(ex.Message, ex.ToString(), "reports-enhanced-viewer.aspx", , )
      End Try
    End If
  End If
  
done:
  Send_PageEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>