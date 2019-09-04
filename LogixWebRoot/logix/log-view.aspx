<%@ Page Language="vb" Debug="true" CodeFile="logixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%
    ' *****************************************************************************
    ' * FILENAME: log-view.aspx 
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
%>
<script runat="server">
    Dim Common As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim AgentSubLogFolder As String = ""
	  Dim bCreateSeperateAgentLogFolder As Boolean = IIf(Common.Fetch_SystemOption(201) = "1", True, False)
    '----------------------------------------------------------------------------------
  
    Sub Get_Params(ByRef FileType As String, ByRef FileYear As String, ByRef FileMonth As String, ByRef FileDay As String, ByRef LocalServerID As Integer)
        FileType = ""
        FileYear = ""
        FileMonth = ""
        FileDay = ""
        LocalServerID = 0
        FileType = Server.HtmlEncode(GetCgiValue("filetype"))
        If FileType = "" Then FileType = "-1"
        FileType = Common.Extract_Val(FileType)
        FileYear = Server.HtmlEncode(GetCgiValue("fileyear"))
        If FileYear = "" Then FileYear = Microsoft.VisualBasic.DateAndTime.Year(Microsoft.VisualBasic.DateAndTime.Now)
        FileMonth = Server.HtmlEncode(GetCgiValue("filemonth"))
        If FileMonth = "" Then FileMonth = Microsoft.VisualBasic.DateAndTime.Month(Microsoft.VisualBasic.DateAndTime.Now)
        FileDay = Server.HtmlEncode(GetCgiValue("fileday"))
        If FileDay = "" Then FileDay = Microsoft.VisualBasic.DateAndTime.Day(Microsoft.VisualBasic.DateAndTime.Now)
        LocalServerID = Common.Extract_Val(Server.HtmlEncode(GetCgiValue("localserverid")))
    End Sub
  
    '----------------------------------------------------------------------------------
  
    Sub Send_Main()
        Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
        Dim CopientFileVersion As String = "7.3.1.138972"
        Dim CopientProject As String = "Copient Logix"
        Dim CopientNotes As String = ""
    
        Dim FileType As String = ""
        Dim FileYear As String = ""
        Dim FileMonth As String = ""
        Dim FileDay As String = ""
        Dim LocalServerID As Integer = 0
        Dim Handheld As Boolean = False
    
        If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
            Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
        End If
    
        Get_Params(FileType, FileYear, FileMonth, FileDay, LocalServerID)
    
        Send_HeadBegin("term.logfileviewer")
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas(0, "IE=8")
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        Send("<frameset rows=""45, *"">")
        Send("  <frame src=""log-view.aspx?mode=selectframe&amp;filetype=" & FileType & "&amp;localserverid=" & LocalServerID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & """ name=""SelectFrame"" scrolling=""no"" />")
        Send("  <frame src=""log-view.aspx?mode=logframe&amp;filetype=" & FileType & "&amp;localserverid=" & LocalServerID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & "&amp;filename=" & Server.HtmlEncode(GetCgiValue("filename")) & """ name=""LogFrame"" />")
        Send("  <noframes>")
        Send("    <body>")
        Send("      " & Copient.PhraseLib.Lookup("error.noframes", LanguageID))
        Send("    </body>")
        Send("  </noframes>")
        Send("</frameset>")
        Send("</html>")
    End Sub
  
    '--------------------------------------------------------------------------------------
  
    Sub LogFrame()

        Dim LogFileID As String= ""
        Dim FileType As String = ""
        Dim FileYear As String = ""
        Dim FileMonth As String = ""
        Dim FileDay As String = ""
        Dim FileName As String
        Dim FilePrefix As String
        Dim di As DirectoryInfo
        Dim FICollection As FileInfo()
        Dim fi As FileInfo
        Dim LocationID As Integer = 0
        Dim LocationPart As String
        Dim dst As DataTable
        Dim row As DataRow
        Dim SpecificFile As String = ""
        Dim PrefixSeparatorChar As String = ""
        Dim failOverLogPath As String = String.Empty
      
        Send_HeadBegin("term.logfileviewer")
        Send_HeadEnd()
    
        Send("<body style=""background-color:#ffffff; font-family:sans-serif; font-size:14px;"">")
        Send("<a id=""top"" name=""top""></a>")
        Send("<span style=""font-size:16px;"">")
    
    
        SpecificFile = Server.HtmlEncode(GetCgiValue("filename"))
     
        Get_Params(FileType, FileYear, FileMonth, FileDay, LocationID)
        FilePrefix = ""
        LocationPart = ""
        Send("<!-- SpecificFile=" & SpecificFile & " -->")
        Send("<!-- FileType=" & FileType & " -->")
    
        Common.QueryStr = "Select LogFileID, isnull(LogFilePrefix, '') as LogFilePrefix, isnull(LogByLocation, 0) as LogByLocation " &
                             "from LogFiles with (NoLock) where LogFileID=" & FileType & ";"
        dst = Common.LRT_Select
        If Not (dst.Rows.Count = 0) Then
            FilePrefix = Common.NZ(dst.Rows(0).Item("LogFilePrefix"), "")
            If Not (dst.Rows(0).Item("LogByLocation") = 0) Then
                LocationPart = Common.Leading_Zero_Fill(LocationID.ToString, 5)
            End If
            LogFileID = Common.NZ(dst.Rows(0).Item("LogFileID"), "")
        End If
        row = Nothing
        dst = Nothing
    
        Send("<!-- FilePrefix=" & FilePrefix & " -->")
    
        If FilePrefix = "" Then
            Send(Copient.PhraseLib.Lookup("log.select", LanguageID))
        Else
            If Not (SpecificFile = "") Then
                FileName = SpecificFile
            Else
                PrefixSeparatorChar = "."
                If Not (LocationPart = "") Then PrefixSeparatorChar = "-"
                FileName = FilePrefix & PrefixSeparatorChar & IIf(LocationPart = "", "", LocationPart & ".") & Common.Leading_Zero_Fill(FileYear, 4) & Common.Leading_Zero_Fill(FileMonth, 2) & Common.Leading_Zero_Fill(FileDay, 2) & ".txt"
            End If
            
            
            If (bCreateSeperateAgentLogFolder) Then
                Dim subLogFolder As String = ""
                                                   
                If (Not String.IsNullOrEmpty(LogFileID)) Then
                    Common.QueryStr = " SELECT LoggingSubPath From LogFiles as logf With (NoLock) " & _
                                             " where logf.LogFileID= " & Convert.ToInt32(LogFileID) & " ;"
                  
                    Dim dst1 As DataTable = Common.LRT_Select

                    If dst1.Rows.Count > 0 Then
                        subLogFolder = Common.NZ(dst1.Rows(0).Item("LoggingSubPath"),"")
                    End If

                    If Not String.IsNullOrEmpty(subLogFolder) AndAlso Not Common.LogPath.Contains(subLogFolder) Then
                        Common.Set_LogPath()
                        
                        If Not (Right(Common.LogPath, 1) = "\") Then
                            Common.LogPath = Common.LogPath & "\" & subLogFolder & "\"
                        Else
                            Common.LogPath = Common.LogPath & subLogFolder & "\"
                        End If
                                          
                        AgentSubLogFolder = IIf(Not String.IsNullOrEmpty(subLogFolder), subLogFolder, "")
                    Else
                        Common.Set_LogPath()
                    End If
            Else
                Common.Set_LogPath()
            End If
        Else
            Common.Set_LogPath()
        End If
            
            FileName = Common.LogPath & FileName
            Send("<!-- FileName=" & FileName & " -->")
            
            If Common.Exist(FileName) OrElse _
                Common.Exist(Path.Combine(Common.Fetch_SystemOption(183), Path.GetFileName(FileName))) OrElse _
                Common.Exist(Path.Combine(Path.Combine(Common.InstallPath, "Logs"), Path.GetFileName(FileName))) OrElse _
                Common.Exist(Path.Combine(Common.Fetch_SystemOption(42), Path.GetFileName(FileName))) Then
                PrimaryLogs(FileName)
                SecondaryLog(FileName)
                If Common.Fetch_SystemOption(42) = "" Then
                    If Common.Fetch_SystemOption(183) = "" Then
                        DefaultLocation(FileName)
                    End If
                End If
            Else
                Send(Copient.PhraseLib.Lookup("log.notfound", LanguageID) & " '" & FileName & "'<br />")
                Send("</span>")
                Send("&nbsp;<br />")
                Dim primaryLogPath = Common.Fetch_SystemOption(42)
                If (Not String.IsNullOrWhiteSpace(primaryLogPath)) Then
                    di = New DirectoryInfo(Common.Fetch_SystemOption(42))
                    FICollection = di.GetFiles(FilePrefix & "*")
                    If (FICollection.Length > 0) Then
                        Send(String.Format(Copient.PhraseLib.Lookup("log.availableat", LanguageID) & ":""{0}""<br />", Copient.PhraseLib.Lookup("settings.42", LanguageID)))
                        For Each fi In FICollection
                            FileYear = Year(fi.LastWriteTime)
                            FileMonth = Month(fi.LastWriteTime)
                            FileDay = Day(fi.LastWriteTime)
                            Send("<a href=""log-view.aspx?filetype=" & FileType & "&amp;locationid=" & LocationID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & "&amp;filename=" & fi.Name & """ target=""_top"">" & fi.Name & "</a> (" & FileMonth & "/" & FileDay & "/" & FileYear & ")<br />")
                        Next
                    End If
                    If (bCreateSeperateAgentLogFolder AndAlso Not String.IsNullOrEmpty(AgentSubLogFolder)) Then
                        If (My.Computer.FileSystem.DirectoryExists(Common.Fetch_SystemOption(42) & "\" & AgentSubLogFolder)) Then
                          di = New DirectoryInfo(Common.Fetch_SystemOption(42) & "\" & AgentSubLogFolder)
                          FICollection = di.GetFiles(FilePrefix & "*")
                          If (FICollection.Length > 0) Then
                              Send(String.Format(Copient.PhraseLib.Lookup("log.availableat", LanguageID) & ":""{0}""<br />", Copient.PhraseLib.Lookup("settings.42", LanguageID)))
                              For Each fi In FICollection
                                  FileYear = Year(fi.LastWriteTime)
                                  FileMonth = Month(fi.LastWriteTime)
                                  FileDay = Day(fi.LastWriteTime)
                                  Send("<a href=""log-view.aspx?filetype=" & FileType & "&amp;locationid=" & LocationID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & "&amp;filename=" & fi.Name & """ target=""_top"">" & fi.Name & "</a> (" & FileMonth & "/" & FileDay & "/" & FileYear & ")<br />")
                              Next
                          End If
                    End If
                End If
              End If
                
                ' if logs failover path is set retrieves all logs from that path.
                failOverLogPath = Common.Fetch_SystemOption(183)
                If (Not String.IsNullOrWhiteSpace(failOverLogPath)) Then
                    di = New DirectoryInfo(failOverLogPath)
                    FICollection = di.GetFiles(FilePrefix & "*")
                    If (FICollection.Length > 0) Then
                        Send(String.Format("<br/>" & Copient.PhraseLib.Lookup("log.availableat", LanguageID) & ":""{0}""<br />", Copient.PhraseLib.Lookup("settings.183", LanguageID)))
                        For Each fi In FICollection
                            FileYear = Year(fi.LastWriteTime)
                            FileMonth = Month(fi.LastWriteTime)
                            FileDay = Day(fi.LastWriteTime)
                            Send("<a href=""log-view.aspx?filetype=" & FileType & "&amp;locationid=" & LocationID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & "&amp;filename=" & fi.Name & """ target=""_top"">" & fi.Name & "</a> (" & FileMonth & "/" & FileDay & "/" & FileYear & ")<br />")
                        Next
                    End If
                    If (bCreateSeperateAgentLogFolder AndAlso Not String.IsNullOrEmpty(AgentSubLogFolder)) Then
                        If (My.Computer.FileSystem.DirectoryExists(Path.Combine(failOverLogPath, AgentSubLogFolder))) Then
                          di = New DirectoryInfo(Path.Combine(failOverLogPath, AgentSubLogFolder))
                          FICollection = di.GetFiles(FilePrefix & "*")
                          If (FICollection.Length > 0) Then
                              Send(String.Format("<br/>" & Copient.PhraseLib.Lookup("log.availableat", LanguageID) & ":""{0}""<br />", Copient.PhraseLib.Lookup("settings.183", LanguageID)))
                              For Each fi In FICollection
                                  FileYear = Year(fi.LastWriteTime)
                                  FileMonth = Month(fi.LastWriteTime)
                                  FileDay = Day(fi.LastWriteTime)
                                  Send("<a href=""log-view.aspx?filetype=" & FileType & "&amp;locationid=" & LocationID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & "&amp;filename=" & fi.Name & """ target=""_top"">" & fi.Name & "</a> (" & FileMonth & "/" & FileDay & "/" & FileYear & ")<br />")
                              Next
                          End If
                    End If
                End If
              End If
                
                'check logs at logix install path
                di = New DirectoryInfo(Path.Combine(Common.InstallPath, "Logs"))
                FICollection = di.GetFiles(FilePrefix & "*")
                If (FICollection.Length > 0) Then
                    Send(String.Format("<br/>" & Copient.PhraseLib.Lookup("log.availableat", LanguageID) & ":""{0}""<br />", Path.Combine(Common.InstallPath, "Logs")))
                    For Each fi In FICollection
                        FileYear = Year(fi.LastWriteTime)
                        FileMonth = Month(fi.LastWriteTime)
                        FileDay = Day(fi.LastWriteTime)
                        Send("<a href=""log-view.aspx?filetype=" & FileType & "&amp;locationid=" & LocationID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & "&amp;filename=" & fi.Name & """ target=""_top"">" & fi.Name & "</a> (" & FileMonth & "/" & FileDay & "/" & FileYear & ")<br />")
                    Next
                End If
                If (bCreateSeperateAgentLogFolder AndAlso Not String.IsNullOrEmpty(AgentSubLogFolder)) Then
                    If (My.Computer.FileSystem.DirectoryExists(Path.Combine(Common.InstallPath, "Logs", AgentSubLogFolder))) Then
                      di = New DirectoryInfo(Path.Combine(Common.InstallPath, "Logs", AgentSubLogFolder))
                      FICollection = di.GetFiles(FilePrefix & "*")
                      If (FICollection.Length > 0) Then
                          Send(String.Format("<br/>" & Copient.PhraseLib.Lookup("log.availableat", LanguageID) & ":""{0}""<br />", Path.Combine(Common.InstallPath, "Logs", AgentSubLogFolder)))
                          For Each fi In FICollection
                              FileYear = Year(fi.LastWriteTime)
                              FileMonth = Month(fi.LastWriteTime)
                              FileDay = Day(fi.LastWriteTime)
                              Send("<a href=""log-view.aspx?filetype=" & FileType & "&amp;locationid=" & LocationID & "&amp;fileyear=" & FileYear & "&amp;filemonth=" & FileMonth & "&amp;fileday=" & FileDay & "&amp;filename=" & fi.Name & """ target=""_top"">" & fi.Name & "</a> (" & FileMonth & "/" & FileDay & "/" & FileYear & ")<br />")
                          Next
                      End If
                End If
            End If
        End If
     End If
            
        Send("<a id=""bottom"" name=""bottom""></a>")
        Send("</body>")
        Send("</html>")
    End Sub
    Sub PrimaryLogs(LogFile As String)
        Dim FileNum As Integer
        Dim TempStr As String
        Dim FileName As String = Path.GetFileName(LogFile)
        Dim PrimaryFileName As String
        Dim PrimaryLocation As String = Common.Fetch_SystemOption(42)
        PrimaryFileName = Path.Combine(PrimaryLocation, FileName)
        If File.Exists(PrimaryFileName) Then
            Send(Copient.PhraseLib.Lookup("log.viewing", LanguageID) & " " & Copient.PhraseLib.Lookup("term.from", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("settings.42", LanguageID).ToLower() & ": " & Path.Combine(PrimaryLocation, AgentSubLogFolder, FileName))
            Send("</span>")
            Send("<pre>")
            FileNum = FreeFile()
            FileOpen(FileNum, PrimaryFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
			
            While Not (EOF(FileNum))
                TempStr = LineInput(FileNum)
                If InStr(TempStr, "IDL-") = 0 Then
                    TempStr = Replace(TempStr, "<", "&lt;", 1)
                    TempStr = Replace(TempStr, ">", "&gt;", 1)
                End If
                Send(TempStr)
				
            End While
            FileClose(FileNum)
            Send("</pre>")
  
        End If
        If bCreateSeperateAgentLogFolder AndAlso Not String.IsNullOrEmpty(AgentSubLogFolder) Then
            PrimaryFileName = Path.Combine(PrimaryLocation, AgentSubLogFolder, FileName)
            If File.Exists(PrimaryFileName) Then
                Send(Copient.PhraseLib.Lookup("log.viewing", LanguageID) & " " & Copient.PhraseLib.Lookup("term.from", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("settings.42", LanguageID).ToLower() & ": " & Path.Combine(PrimaryLocation, AgentSubLogFolder, FileName))
                Send("</span>")
                Send("<pre>")
                FileNum = FreeFile()
                FileOpen(FileNum, PrimaryFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
			
                While Not (EOF(FileNum))
                    TempStr = LineInput(FileNum)
                    If InStr(TempStr, "IDL-") = 0 Then
                        TempStr = Replace(TempStr, "<", "&lt;", 1)
                        TempStr = Replace(TempStr, ">", "&gt;", 1)
                    End If
                    Send(TempStr)
				
                End While
                FileClose(FileNum)
                Send("</pre>")
            End If
        End If
    End Sub
    Sub SecondaryLog(LogFile As String)
        Dim FileNum As Integer
        Dim TempStr As String
        Dim FileName As String = Path.GetFileName(LogFile)
        Dim FailoverFileName As String
        Dim FailOverLocation As String = Common.Fetch_SystemOption(183)
        FailoverFileName = Path.Combine(FailOverLocation, FileName)
        If File.Exists(FailoverFileName) Then
		
            Send("<span style=""font-size:16px;"">")
            Send(Copient.PhraseLib.Lookup("log.viewing", LanguageID) & " " & Copient.PhraseLib.Lookup("term.from", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("settings.183", LanguageID).ToLower() & ": " & FailOverLocation)
            Send("</span>")
            Send("<pre>")
            FileNum = FreeFile()
            FileOpen(FileNum, FailoverFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            While Not (EOF(FileNum))
                TempStr = LineInput(FileNum)
                If InStr(TempStr, "IDL-") = 0 Then
                    TempStr = Replace(TempStr, "<", "&lt;", 1)
                    TempStr = Replace(TempStr, ">", "&gt;", 1)
                End If
                Send(TempStr)
            End While
            FileClose(FileNum)
            Send("</pre>")
        End If
        If bCreateSeperateAgentLogFolder AndAlso Not String.IsNullOrEmpty(AgentSubLogFolder) Then
            FailOverLocation = Path.Combine(FailOverLocation, AgentSubLogFolder, FileName)
            If File.Exists(FailoverFileName) Then
                Send("<span style=""font-size:16px;"">")
                Send(Copient.PhraseLib.Lookup("log.viewing", LanguageID) & " " & Copient.PhraseLib.Lookup("term.from", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.installation", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.path", LanguageID).ToLower() & ": " & Path.Combine(FailOverLocation, AgentSubLogFolder, FileName))
            
                Send("</span>")
                Send("<pre>")
                FileNum = FreeFile()
                FileOpen(FileNum, FailoverFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                While Not (EOF(FileNum))
                    TempStr = LineInput(FileNum)
                    If InStr(TempStr, "IDL-") = 0 Then
                        TempStr = Replace(TempStr, "<", "&lt;", 1)
                        TempStr = Replace(TempStr, ">", "&gt;", 1)
                    End If
                    Send(TempStr)
                End While
                FileClose(FileNum)
                Send("</pre>")
            End If
        End If
    End Sub
    
    Sub DefaultLocation(LogFile As String)
        Dim FileNum As Integer
        Dim TempStr As String
        Dim FileName As String = Path.GetFileName(LogFile)
        Dim DefaultFileName As String
        DefaultFileName = Path.Combine(Common.InstallPath, "Logs", FileName)
        If File.Exists(DefaultFileName) Then
            Send("<span style=""font-size:16px;"">")
            Send(Copient.PhraseLib.Lookup("log.viewing", LanguageID) & " " & Copient.PhraseLib.Lookup("term.from", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.installation", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.path", LanguageID).ToLower() & ": " & DefaultFileName)
            Send("</span>")
            Send("<pre>")
            FileNum = FreeFile()
            FileOpen(FileNum, DefaultFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
            While Not (EOF(FileNum))
                TempStr = LineInput(FileNum)
                If InStr(TempStr, "IDL-") = 0 Then
                    TempStr = Replace(TempStr, "<", "&lt;", 1)
                    TempStr = Replace(TempStr, ">", "&gt;", 1)
                End If
                Send(TempStr)
            End While
            FileClose(FileNum)
            Send("</pre>")
        End If
        If bCreateSeperateAgentLogFolder AndAlso Not String.IsNullOrEmpty(AgentSubLogFolder) Then
            DefaultFileName = Path.Combine(Common.InstallPath, "Logs", AgentSubLogFolder, FileName)
            If File.Exists(DefaultFileName) Then
                Send("<span style=""font-size:16px;"">")
                Send(Copient.PhraseLib.Lookup("log.viewing", LanguageID) & " " & Copient.PhraseLib.Lookup("term.from", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.installation", LanguageID).ToLower() & " " & Copient.PhraseLib.Lookup("term.path", LanguageID).ToLower() & ": " & Path.Combine(Common.InstallPath, "Logs", AgentSubLogFolder, FileName))
                Send("</span>")
                Send("<pre>")
                FileNum = FreeFile()
                FileOpen(FileNum, DefaultFileName, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
                While Not (EOF(FileNum))
                    TempStr = LineInput(FileNum)
                    If InStr(TempStr, "IDL-") = 0 Then
                        TempStr = Replace(TempStr, "<", "&lt;", 1)
                        TempStr = Replace(TempStr, ">", "&gt;", 1)
                    End If
                    Send(TempStr)
                End While
                FileClose(FileNum)
                Send("</pre>")
            End If
        End If
    End Sub
    '-------------------------------------------------------------------
    Sub SelectFrame()
        Dim LogFileID As String
        Dim FileYear As String = ""
        Dim FileMonth As String = ""
        Dim FileDay As String = ""
        Dim LocalServerID As String = "0"
        Dim dst As DataTable
        Dim row As DataRow
        Dim ShowLocationSelector As Integer
        Dim EngineId As Integer
        
        LogFileID = ""
        Get_Params(LogFileID, FileYear, FileMonth, FileDay, LocalServerID)
    
        Send("<!-- LogFileID=" & LogFileID & " -->")
        Send("<!-- LocalServerID=" & LocalServerID & " -->")
    
        Send_HeadBegin("term.logfileviewer")
        Send_HeadEnd()
        Send("<body style=""background-color:#ffffff; font-family:sans-serif; font-size:14px; margin-top:10px;"">")
        Send("<form action=""log-view.aspx"" target=""_top"" method=""post"" id=""mainform"" name=""mainform"">")
        Send("<label for=""filetype"">" & Copient.PhraseLib.Lookup("log.filetype", LanguageID) & ":</label>")
        ShowLocationSelector = 0

        Common.QueryStr = "dbo.pa_LogView_Select"
        Common.Open_LRTsp()
        dst = Common.LRTsp_select
        Common.Close_LRTsp()
        If dst.Rows.Count > 0 Then
            Send("<select id=""filetype"" name=""filetype"">")
    
            For Each row In dst.Rows
                Send("<!--" & row.Item("LogFileID") & "  " & row.Item("LogByLocation") & "-->")
                Sendb(" <option value=""" & row.Item("LogFileID") & """")
                If LogFileID = row.Item("LogFileID") Then
                    If row.Item("LogByLocation") Then
                        ShowLocationSelector = 1
                        EngineId = row.Item("EngineId")
                    End If
                    Sendb(" selected=""selected""")
                End If
                Sendb(">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID, row.Item("Name")))
                Send("</option>")
            Next
            Send("</select>")
        End If  'dst.rows.count>0
        dst = Nothing
        row = Nothing
        Send("<!-- ShowLocationSelector=" & ShowLocationSelector & " -->")
   
        If ShowLocationSelector = 1 Then
            Select Case LogFileID
                Case 2 '(CMConnector)
                    Send("<label for=""localserverid"">" & Copient.PhraseLib.Lookup("term.location", LanguageID) & ":</label>")
                    Send("<select id=""localserverid"" name=""localserverid"">")
                    Send("<option value=""0"">" & Copient.PhraseLib.Lookup("term.unspecified", LanguageID) & "</option>")
                    Common.QueryStr = "select LS.LocalServerID, L.ExtLocationCode " & _
                                      "from Locations as L with (NoLock) Inner Join LocalServers as LS with (NoLock) on L.LocationID=LS.LocationID " & _
                                      "where L.EngineID=0 order by L.ExtLocationCode;"
                    dst = Common.LRT_Select
                    For Each row In dst.Rows
                        Sendb("<option value=""" & row.Item("LocalServerID") & """")
                        If row.Item("LocalServerID") = LocalServerID Then
                            Sendb(" selected=""selected""")
                        End If
                        Send(">" & Common.NZ(row.Item("ExtLocationCode"), "!Unknown!") & "</option>")
                    Next
                    row = Nothing
                    dst = Nothing
                    Send("</select>&nbsp;&nbsp;")
                Case Else
                    Send("<label for=""localserverid"">" & Copient.PhraseLib.Lookup("term.location", LanguageID) & ":</label>")
                    Send("<select id=""localserverid"" name=""localserverid"">")
                    Send("<option value=""0"">" & Copient.PhraseLib.Lookup("term.unspecified", LanguageID) & "</option>")
                    If EngineId = InstalledEngines.CPE Then
                    Common.QueryStr = "select LS.LocalServerID, L.ExtLocationCode " & _
                                        "from Locations as L with (NoLock) Inner Join LocalServers as LS with (NoLock) on L.LocationID=LS.LocationID " & _
                                        "where L.EngineID=2 order by L.ExtLocationCode;"
                    ElseIf EngineId = InstalledEngines.UE Then
                    Common.QueryStr = "select LSL.LocalServerID, L.ExtLocationCode " & _
                                        "from Locations as L with (NoLock) Inner Join LocalServerLocations as LSL with (NoLock) on L.LocationID=LSL.LocationID " & _
                                        "where L.EngineID=9 AND LSL.AssociationTypeID=1 order by L.ExtLocationCode;"
                    ElseIf EngineId = -1 Then
                    Common.QueryStr = "select LS.LocalServerID, L.ExtLocationCode " & _
                                       "from Locations as L with (NoLock) Inner Join LocalServers as LS with (NoLock) on L.LocationID=LS.LocationID " & _
                                        "where L.EngineID in (0,2,9) order by L.ExtLocationCode;"
					Else
                    Common.QueryStr = "select LSL.LocalServerID, L.ExtLocationCode " & _
                                        "from Locations as L with (NoLock) Inner Join LocalServerLocations as LSL with (NoLock) on L.LocationID=LSL.LocationID " & _
                                         "where L.EngineID Not in (2,9) AND LSL.AssociationTypeID=1 order by L.ExtLocationCode;"
                    End If
                    dst = Common.LRT_Select
                    For Each row In dst.Rows
                        Sendb("<option value=""" & row.Item("LocalServerID") & """")
                        If row.Item("LocalServerID") = LocalServerID Then
                            Sendb(" selected=""selected""")
                        End If
                        Send(">" & Common.NZ(row.Item("ExtLocationCode"), "!Unknown!") & "</option>")
                    Next
                    row = Nothing
                    dst = Nothing
                    Send("</select>&nbsp;&nbsp;")
            End Select
        End If 'ShowLocationSelector=1
    
    
        Send("<label for=""filemonth"">" & Copient.PhraseLib.Lookup("term.month", LanguageID) & ":</label> <input type=""text"" id=""filemonth"" name=""filemonth"" maxlength=""2"" size=""2"" value=""" & FileMonth & """ />&nbsp;&nbsp;")
        Send("<label for=""fileday"">" & Copient.PhraseLib.Lookup("term.day", LanguageID) & ":</label> <input type=""text"" id=""fileday"" name=""fileday"" maxlength=""2"" size=""2"" value=""" & FileDay & """ />&nbsp;&nbsp;")
        Send("<label for=""fileyear"">" & Copient.PhraseLib.Lookup("term.year", LanguageID) & ":</label> <input type=""text"" id=""fileyear"" name=""fileyear"" maxlength=""4"" size=""4"" value=""" & FileYear & """ />&nbsp;&nbsp;")
        Send("<input type=""submit"" id=""go"" name=""go"" value=""" & Copient.PhraseLib.Lookup("term.go", LanguageID) & """ style=""width:50px;"" />")
        Send("</form>")
        Send("</body>")
        Send("</html>")
    End Sub
  
    '-------------------------------------------------------------------
</script>
<%
    Dim Mode As String = ""
    Response.Expires = 0
    Common.AppName = "log-view.aspx"
  
    On Error GoTo ErrorTrap
  
    If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
    AdminUserID = Verify_AdminUser(Common, Logix)
  
  
    'Mode = ""
    'Mode = Server.HtmlEncode(Request.Form("mode"))
    If Mode = "" Then Mode = Server.HtmlEncode(Request.QueryString("mode"))
    Select Case UCase(Mode)
        Case "SELECTFRAME"
            SelectFrame()
        Case "LOGFRAME"
            LogFrame()
        Case Else
            Send_Main()
    End Select
  
    GoTo AllDone
  
ErrorTrap:
    Common.Error_Processor()
  
AllDone:
    If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
    Common = Nothing
    Logix = Nothing
    Response.End()
%>
