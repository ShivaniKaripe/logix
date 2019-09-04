<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: help.aspx 
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
  Dim dst As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim installedEngines As String = ""
  Dim pageSections As String = ""
  Dim editEnabled As Boolean = False
  Dim i As Integer = 0
  
  Dim bookPhrase As String = ""
  Dim partPhrase As String = ""
  Dim chapterPhrase As String = ""
  Dim sectionPhrase As String = ""
  Dim dstBooks As System.Data.DataTable
  Dim bookRow As System.Data.DataRow
  Dim dstParts As System.Data.DataTable
  Dim partRow As System.Data.DataRow
  Dim dstChapters As System.Data.DataTable
  Dim chapterRow As System.Data.DataRow
  Dim dstSections As System.Data.DataTable
  Dim sectionRow As System.Data.DataRow
  Dim bookCounter As Integer = 0
  Dim partCounter As Integer = 0
  Dim chapterCounter As Integer = 0
  Dim sectionCounter As Integer = 0
  
  Dim AdminUserID As Long
  Dim DefaultEngine As Integer = 0
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim Popup As Boolean = False
  Dim FileName As String = ""
  Dim UIPageID As Integer = 0
  Dim Address As String = ""
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "help.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)
  FileName = IIf(Request.QueryString("FileName") <> "", Request.QueryString("FileName"), "")
  MyCommon.QueryStr = "select UIPageID from UIPages with (NoLock) where FileName='" & FileName & "';"
  dst = MyCommon.LRT_Select
  If dst.Rows.Count > 0 Then
    UIPageID = MyCommon.NZ(dst.Rows(0).Item("UIPageID"), 0)
  End If
  
  ' See if user should be allowed to directly edit the master manual text (currently only Huw on Dev2)
  Address = Left(Request.Url.ToString, Request.Url.ToString.LastIndexOf("/") + 1)
  If (InStr(Address, "copienttech.com")) Then
    MyCommon.QueryStr = "select UserName from AdminUsers with (NoLock) where AdminUserID=" & AdminUserID & ";"
    dst = MyCommon.LRT_Select
    If dst.Rows.Count > 0 Then
      If MyCommon.NZ(dst.Rows(0).Item("UserName"), "") = "huw" Then
        editEnabled = True
      End If
    End If
  End If
  
  ' Find the default engine
  MyCommon.QueryStr = "select EngineID,DefaultEngine from PromoEngines with (NoLock) where DefaultEngine=1;"
  dst = MyCommon.LRT_Select
  If dst.Rows.Count > 0 Then
    DefaultEngine = MyCommon.NZ(dst.Rows(0).Item("EngineID"), 0)
  End If
  
  ' Build a comma-delimited list of installed engines, to be used in the queries that pull the text
  MyCommon.QueryStr = "select EngineID from PromoEngines where Installed=1;"
  dst = MyCommon.LRT_Select
  If dst.Rows.Count > 0 Then
    i = 1
    For Each row In dst.Rows
      installedEngines &= row.Item("EngineID")
      If i < dst.Rows.Count Then
        installedEngines &= ","
      End If
      i += 1
    Next
  End If
  
  ' Also, if a UIPageID exists, build a comma-delimited list of sections associated to that page
  If UIPageID > 0 Then
    MyCommon.QueryStr = "select SectionID from UserManualPages with (NoLock) " & _
                        "where UIPageID=" & UIPageID & " order by DisplayOrder;"
    dst = MyCommon.LRT_Select
    If dst.Rows.Count > 0 Then
      i = 1
      For Each row In dst.Rows
        pageSections &= row.Item("SectionID")
        If i < dst.Rows.Count Then
          pageSections &= ","
        End If
        i += 1
      Next
    End If
  End If
  
  Send_HeadBegin("term.usermanual")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  If Popup Then
    Send_BodyBegin(3)
  Else
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 1)
    Send_Subtabs(Logix, 1, 3)
  End If
%>
<div id="intro">
  <h1>
    <% Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))%> 5 <% Sendb(Copient.PhraseLib.Lookup("term.usermanual", LanguageID))%>
  </h1>
  <div id="controls">
    <%
      If MyCommon.Fetch_SystemOption(75) AndAlso (Popup = False) Then
        If (Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(2, 0, AdminUserID)
        End If
      End If
    %>
  </div>
</div>

<div id="main" class="forcescroll">
  <%
    ' Load the names of the book parts
    MyCommon.QueryStr = "select SectionTypeID, Description, PhraseID from UserManualSectionTypes "
    dst = MyCommon.LRT_Select
    bookPhrase = Copient.PhraseLib.Lookup(dst.Rows(0).Item("PhraseID"), LanguageID)
    partPhrase = Copient.PhraseLib.Lookup(dst.Rows(1).Item("PhraseID"), LanguageID)
    chapterPhrase = Copient.PhraseLib.Lookup(dst.Rows(2).Item("PhraseID"), LanguageID)
    sectionPhrase = Copient.PhraseLib.Lookup(dst.Rows(3).Item("PhraseID"), LanguageID)
    
    ' Get the book
    MyCommon.QueryStr = "select * from UserManual with (NoLock) " & _
                        "where SectionTypeID=1 and Visible=1 and SectionID=1 " & _
                        "and (EngineID is NULL or EngineID in (" & installedEngines & ")) " & _
                        "order by Sequence;"
    dstBooks = MyCommon.LRT_Select
    If dstBooks.Rows.Count > 0 Then
      bookCounter = 1
      For Each bookRow In dstBooks.Rows
        Send("<div class=""book"" id=""um" & bookRow.Item("SectionID") & """>")
        ' Within the book, get all available parts
        MyCommon.QueryStr = "select * from UserManual with (NoLock) " & _
                            "where SectionTypeID=2 and Visible=1 and ParentSectionID=" & bookRow.Item("SectionID") & " " & _
                            "and (EngineID is NULL or EngineID in (" & installedEngines & ")) " & _
                            "order by Sequence;"
        dstParts = MyCommon.LRT_Select
        If dstParts.Rows.Count > 0 Then
          partCounter = 0
          i = 0
          For Each partRow In dstParts.Rows
            If partRow.Item("NumberRestart") Then
              partCounter = 1
            End If
            Send("    ")
            Send("    <div class=""part"" id=""um" & partRow.Item("SectionID") & """>")
            Send("      <a id=""a" & partRow.Item("SectionID") & """></a>")
            If partRow.Item("DisplayTitle") Then
              Send("      <h1>" & IIf(partRow.Item("Numbered"), partPhrase & " " & MyCommon.ToRomanNumeral(partCounter) & ": ", "") & MyCommon.NZ(partRow.Item("Title"), "") & "</h1>")
              Send("      <div class=""helplinks"">")
              If i = 1 Then
                Send("        <a href=""#a" & dstParts.Rows(i + 1).Item("SectionID") & """ title=""" & Copient.PhraseLib.Lookup("term.next", LanguageID) & """>&#9660;</a>")
              ElseIf i < (dstParts.Rows.Count - 1) Then
                Sendb("        <a href=""#a" & dstParts.Rows(i - 1).Item("SectionID") & """ title=""" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & """>&#9650;</a>")
                Send("<a href=""#a" & dstParts.Rows(i + 1).Item("SectionID") & """ title=""" & Copient.PhraseLib.Lookup("term.next", LanguageID) & """>&#9660;</a>")
              Else
                Send("        <a href=""#a" & dstParts.Rows(i - 1).Item("SectionID") & """ title=""" & Copient.PhraseLib.Lookup("term.previous", LanguageID) & """>&#9650;</a>")
              End If
              Send("      </div>")
            End If
            If MyCommon.NZ(partRow.Item("ContentPhraseID"), 0) > 0 Then
              Send("      " & Copient.PhraseLib.Lookup(MyCommon.NZ(partRow.Item("ContentPhraseID"), 0), LanguageID))
            End If
            ' Within the part, get all available chapters/appendices
            MyCommon.QueryStr = "select * from UserManual with (NoLock) " & _
                                "where SectionTypeID=3 and Visible=1 and ParentSectionID=" & partRow.Item("SectionID") & " " & _
                                "and (EngineID is NULL or EngineID in (" & installedEngines & ")) " & _
                                "order by Sequence;"
            dstChapters = MyCommon.LRT_Select
            If dstChapters.Rows.Count > 0 Then
              'chapterCounter = 1 -- commenting out, to keep chapter numbering continuous across parts
              For Each chapterRow In dstChapters.Rows
                If chapterRow.Item("NumberRestart") Then
                  chapterCounter = 1
                End If
                Send("      ")
                Send("      <div class=""chapter" & IIf(chapterRow.Item("Contents"), " contents", "") & """ id=""um" & chapterRow.Item("SectionID") & """>")
                Send("        <a id=""a" & chapterRow.Item("SectionID") & """></a>")
                If chapterRow.Item("DisplayTitle") Then
                  If chapterRow.Item("Appendix") Then
                    Send("        <h2>" & IIf(chapterRow.Item("Numbered"), Copient.PhraseLib.Lookup("term.appendix", LanguageID) & " " & ChrW(chapterCounter + 64) & ". ", "") & MyCommon.NZ(chapterRow.Item("Title"), "") & "</h2>")
                  Else
                    Send("        <h2>" & IIf(chapterRow.Item("Numbered"), chapterCounter & ". ", "") & MyCommon.NZ(chapterRow.Item("Title"), "") & "</h2>")
                  End If
                End If
                If (Not IsDBNull(chapterRow.Item("ContentPhraseID"))) AndAlso (editEnabled) Then
                  Send("        <span class=""phraselink""><a href=""http://dev2.copienttech.com/cgi-bin/Connectors/PhraseMgmt.aspx?PhraseID=" & chapterRow.Item("ContentPhraseID") & """ target=""phraselib"">" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "</a></span>")
                End If
                If MyCommon.NZ(chapterRow.Item("ContentPhraseID"), 0) > 0 Then
                  Send("        " & Copient.PhraseLib.Lookup(MyCommon.NZ(chapterRow.Item("ContentPhraseID"), 0), LanguageID))
                End If
                If MyCommon.NZ(chapterRow.Item("Contents"), False) Then
                  ' Special case: the chapter is the "contents" chapter, so build the table of contents
                  BuildContents(bookRow.Item("SectionID"))
                End If
                ' Within the chapter, get all available sections
                MyCommon.QueryStr = "select * from UserManual with (NoLock) " & _
                                    "where SectionTypeID=4 and Visible=1 and ParentSectionID=" & chapterRow.Item("SectionID") & " " & _
                                    "and (EngineID is NULL or EngineID in (" & installedEngines & ")) " & _
                                    "order by Sequence;"
                dstSections = MyCommon.LRT_Select
                If dstSections.Rows.Count > 0 Then
                  sectionCounter = 1
                  For Each sectionRow In dstSections.Rows
                    If sectionRow.Item("NumberRestart") Then
                      sectionCounter = 1
                    End If
                    Send("        <div class=""section"" id=""um" & sectionRow.Item("SectionID") & """>")
                    Send("          <a id=""a" & sectionRow.Item("SectionID") & """></a>")
                    If sectionRow.Item("DisplayTitle") Then
                      If chapterRow.Item("Appendix") Then
                        Send("          <h3>" & IIf(sectionRow.Item("Numbered"), ChrW(chapterCounter + 64) & "." & sectionCounter & ". ", "") & MyCommon.NZ(sectionRow.Item("Title"), "") & "</h3>")
                      Else
                        Send("          <h3>" & IIf(sectionRow.Item("Numbered"), chapterCounter & "." & sectionCounter & ". ", "") & MyCommon.NZ(sectionRow.Item("Title"), "") & "</h3>")
                      End If
                    End If
                    If (Not IsDBNull(sectionRow.Item("ContentPhraseID"))) AndAlso (editEnabled) Then
                      Send("          <span class=""phraselink""><a href=""http://dev2.copienttech.com/cgi-bin/Connectors/PhraseMgmt.aspx?PhraseID=" & sectionRow.Item("ContentPhraseID") & """ target=""phraselib"">" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "</a></span>")
                    End If
                    If MyCommon.NZ(sectionRow.Item("ContentPhraseID"), 0) > 0 Then
                      Send("          " & Copient.PhraseLib.Lookup(MyCommon.NZ(sectionRow.Item("ContentPhraseID"), 0), LanguageID))
                    End If
                    Send("        </div> <!-- End " & sectionRow.Item("SectionID") & " (Section " & sectionCounter & ") -->")
                    sectionCounter += 1
                  Next
                End If
                Send("      </div> <!-- End " & chapterRow.Item("SectionID") & " (Chapter " & chapterCounter & ") ~~~~~~~~~~~~~~~~~~~~~ -->")
                chapterCounter += 1
              Next
            End If
            Send("    ")
            Send("    </div> <!-- End " & partRow.Item("SectionID") & " (Part " & partCounter & ") ========================================== -->")
            Send("    ")
            partCounter += 1
            i += 1
          Next
        End If
        Send("  </div> <!-- End book -->")
        bookCounter += 1
        
        partCounter = 0
        chapterCounter = 0
        sectionCounter = 0
      Next
    End If
  %>
  ◦
</div>

<script runat="server">
  Dim MyCommon As New Copient.CommonInc
  
  Sub BuildContents(ByVal BookID As Integer)
    Dim sqlBuf As New StringBuilder()
    
    Dim dstParts As System.Data.DataTable
    Dim partRow As System.Data.DataRow
    Dim dstChapters As System.Data.DataTable
    Dim chapterRow As System.Data.DataRow
    
    Dim partCounter As Integer = 0
    Dim chapterCounter As Integer = 0
    Dim sectionCounter As Integer = 0
    
    MyCommon.Open_LogixRT()
    
    MyCommon.QueryStr = "select * from UserManual with (NoLock) " & _
                        "where SectionTypeID=2 and Visible=1 and ParentSectionID=" & BookID & " " & _
                        "order by Sequence;"
    dstParts = MyCommon.LRT_Select
    If dstParts.Rows.Count > 0 Then
      partCounter = 0
      For Each partRow In dstParts.Rows
        If (IsDBNull(partRow.Item("EngineID"))) OrElse (IsDBNull(partRow.Item("EngineID")) = False AndAlso MyCommon.IsEngineInstalled(partRow.Item("EngineID"))) Then
          If partRow.Item("NumberRestart") Then
            partCounter = 1
          End If
          If MyCommon.NZ(partRow.Item("DisplayTitle"), True) Then
            Send("      <a class=""tocPart"" href=""#a" & partRow.Item("SectionID") & """>" & MyCommon.NZ(partRow.Item("Title"), "") & "</a><br />")
          End If
          MyCommon.QueryStr = "select * from UserManual with (NoLock) " & _
                              "where SectionTypeID=3 and Visible=1 and Contents=0 and ParentSectionID=" & partRow.Item("SectionID") & " " & _
                              "order by Sequence;"
          dstChapters = MyCommon.LRT_Select
          If dstChapters.Rows.Count > 0 Then
            For Each chapterRow In dstChapters.Rows
              If (IsDBNull(chapterRow.Item("EngineID"))) OrElse (IsDBNull(chapterRow.Item("EngineID")) = False AndAlso MyCommon.IsEngineInstalled(chapterRow.Item("EngineID"))) Then
                If chapterRow.Item("NumberRestart") Then
                  chapterCounter = 1
                End If
                If MyCommon.NZ(chapterRow.Item("DisplayTitle"), True) Then
                  Sendb("      <a class=""tocChapter"" href=""#a" & chapterRow.Item("SectionID") & """>")
                  Sendb(IIf(chapterRow.Item("Appendix"), Copient.PhraseLib.Lookup("term.appendix", LanguageID) & " ", ""))
                  If chapterRow.Item("Numbered") Then
                    If chapterRow.Item("Appendix") Then
                      Sendb(ChrW(chapterCounter + 64) & ". ")
                    Else
                      Sendb(chapterCounter & ". ")
                    End If
                  End If
                  Send(MyCommon.NZ(chapterRow.Item("Title"), "") & "</a><br />")
                End If
                chapterCounter += 1
              End If
            Next
          End If
          partCounter += 1
        End If
      Next
    End If
    
    Send(sqlBuf.ToString)
  End Sub
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(2, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
