<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="System.Net.Http" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS.Utilities" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="CMS.AMS" %>
<%
    ' *****************************************************************************
    ' * FILENAME: health-resolutions.aspx 
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
    Dim rst As System.Data.DataTable
    Dim dst As System.Data.DataTable = New DataTable()
    Dim row As System.Data.DataRow
    Dim dst2 As System.Data.DataTable = New DataTable()
    Dim row2 As System.Data.DataRow
    Dim shaded As Boolean
    Dim i As Integer = 0
    Dim ResolutionID As Integer = 0
    Dim lastEntryOrder As Integer = 0
    Dim swapEntryOrder As Integer = 0
    Dim ServerType As Integer = Request.QueryString("SrvType")
    Dim ErrorID As Integer = Request.QueryString("Err")
    Dim ParamID As String = Request.QueryString("ParamID")
    Dim FromServerHealth As Integer = IIf(String.IsNullOrEmpty(Request.QueryString("FromServerHealth")), 0, Request.QueryString("FromServerHealth"))
    Dim FullName As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim HealthServiceURL = MyCommon.Fetch_UE_SystemOption(184).Trim("/")
    Dim MultiLanguageEnabled As Boolean = False
    Dim UpArrowDisabled As Boolean = False
    Dim DownArrowDisabled As Boolean = False
    Dim headers As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))
    Dim RESTServiceHelper As IRestServiceHelper
   
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "health-resolutions.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixWH()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    If MyCommon.Fetch_SystemOption(124) = 1 Then
        MultiLanguageEnabled = True
    End If
    Send_HeadBegin("term.resolutions")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
    function LoadDocument(url) {
        location = url;
    }
    
    function AddResolution() {
    <%If MultiLanguageEnabled Then %>
    var userLanguage = document.getElementById("UserLanguage").value;
    var elem=document.forms['newform']['restitle'+userLanguage];
    if (elem != null && elem.disabled == false) {
            elem.focus();
        }
        <%ELSE %>
        var elem = document.forms['newform'].restitle;
        if (elem != null && elem.disabled == false) {
            elem.focus();
        }
        <%END IF %>
    }
    function editResolution(number) {
        if (document.getElementById("editres" + number) != null) {
            document.getElementById("savedres" + number).style.display = 'none'
            document.getElementById("editres" + number).style.display = ''
        }
    }
    function cancelResolution(number) {
        if (document.getElementById("editres" + number) != null) {
            document.getElementById("savedres" + number).style.display = ''
            document.getElementById("editres" + number).style.display = 'none'
            document.getElementById("restitle" + number).value = document.getElementById("title" + number).innerText
            document.getElementById("restext" + number).value = document.getElementById("text" + number).innerText
        }
    }
</script>
<%
    Send_HeadEnd()
    CurrentRequest.Resolver.AppName="health-resolutions.aspx"
    RESTServiceHelper = CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IRestServiceHelper)()
    
   
    If (Request.QueryString("Save") <> "") Then
        If MultiLanguageEnabled Then
            If (Logix.TrimAll(Request.QueryString.Item("restitle" & Request.QueryString.Item("LogixDefaultLanguage"))) <> "") AndAlso (Logix.TrimAll(Request.QueryString.Item("restext" & Request.QueryString.Item("LogixDefaultLanguage"))) <> "") Then
                Dim resolutionTexts As New List(Of ResolutionText)()
                Dim installedLangTbl As New DataTable()
                'get all the installed languages
                MyCommon.QueryStr = "Select LanguageID from Languages where InstalledForUI=1"
                installedLangTbl = MyCommon.LRT_Select()
                For Each row In installedLangTbl.Rows
                    'get the resolution text and title for each language
                    Dim resolutionText As New ResolutionText()
                    resolutionText.Title = Logix.TrimAll(Request.QueryString.Item("restitle" & row.Item("LanguageID")))
                    resolutionText.Description = Logix.TrimAll(Request.QueryString.Item("restext" & row.Item("LanguageID")))
                    resolutionText.LanguageID = row.Item("LanguageID")
                    If resolutionText.Title <> "" AndAlso resolutionText.Description <> "" Then
                        resolutionTexts.Add(resolutionText)
                    End If
                Next
                Dim resolution As CMS.AMS.Models.Resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                                                                  .LastUpdate = (DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds,
                                                                                                  .ParamID = ParamID,
                                                                                                  .ResolutionID = 0,
                                                                                                  .UserID = AdminUserID,
                                                                                                  .DisplayOrder = 0}
                If resolution.ResolutionTexts.Count > 0 Then
                    If FromServerHealth = 1 Then
                        Dim result As KeyValuePair(Of Integer, String) = RESTServiceHelper.CallService(Of Integer)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Post, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
                        infoMessage = result.Value
                        resolution.ResolutionID = result.Key
                    Else
                        MyCommon.QueryStr = "select top 1 max(EntryOrder) as EntryOrder from HealthResolutions where ServerTypeID=" & ServerType & " and ErrorID=" & ErrorID & " and Deleted=0 order by EntryOrder desc;"
                        dst = MyCommon.LWH_Select
                        If dst.Rows.Count > 0 Then
                            lastEntryOrder = MyCommon.NZ(dst.Rows(0).Item("EntryOrder"), 0)
                        Else
                            lastEntryOrder = 0
                        End If
                        MyCommon.QueryStr = "insert into HealthResolutions (ErrorID, ServerTypeID, AdminUserID, CreatedDate, LastUpdate, EntryOrder, Title, ResolutionText) " & _
                                            "values(" & ErrorID & ", " & ServerType & ", " & AdminUserID & ", getDate(), getDate(), " & lastEntryOrder + 1 & ", N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restitle")) & "', N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restext")) & "');"

                        MyCommon.LWH_Execute()
                    End If

                Else
                    infoMessage = Copient.PhraseLib.Lookup("serverhealth.resolutionerror", LanguageID)
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("serverhealth.resolutionerror", LanguageID)
            End If
        Else
            If (Logix.TrimAll(Request.QueryString.Item("restitle")) <> "") AndAlso (Logix.TrimAll(Request.QueryString.Item("restext")) <> "") Then
                Dim resolutionTexts As New List(Of ResolutionText)()
                Dim resolutiontext As New ResolutionText()
                resolutiontext.Description = Request.QueryString.Item("restext")
                resolutiontext.Title = Request.QueryString.Item("restitle")
                resolutiontext.LanguageID = LanguageID
                resolutionTexts.Add(resolutiontext)
                Dim resolution As CMS.AMS.Models.Resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                                                   .LastUpdate = (DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds,
                                                                                   .ParamID = ParamID,
                                                                                   .ResolutionID = 0,
                                                                                   .UserID = AdminUserID,
                                                                                   .DisplayOrder = 0}
                If FromServerHealth = 1 Then
                    Dim result As KeyValuePair(Of Integer, String) = RESTServiceHelper.CallService(Of Integer)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Post, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
                    infoMessage = result.Value
                    resolution.ResolutionID = result.Key
                Else
                    MyCommon.QueryStr = "select top 1 max(EntryOrder) as EntryOrder from HealthResolutions where ServerTypeID=" & ServerType & " and ErrorID=" & ErrorID & " and Deleted=0 order by EntryOrder desc;"
                    dst = MyCommon.LWH_Select
                    If dst.Rows.Count > 0 Then
                        lastEntryOrder = MyCommon.NZ(dst.Rows(0).Item("EntryOrder"), 0)
                    Else
                        lastEntryOrder = 0
                    End If
                    MyCommon.QueryStr = "insert into HealthResolutions (ErrorID, ServerTypeID, AdminUserID, CreatedDate, LastUpdate, EntryOrder, Title, ResolutionText) " & _
                   "values(" & ErrorID & ", " & ServerType & ", " & AdminUserID & ", getDate(), getDate(), " & lastEntryOrder + 1 & ", N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restitle")) & "', N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restext")) & "');"

                    MyCommon.LWH_Execute()
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("serverhealth.resolutionerror", LanguageID)
            End If
        End If
    ElseIf (Request.QueryString("Edit") <> "") Then
        If MultiLanguageEnabled Then
            If (Logix.TrimAll(Request.QueryString.Item("restitle" & Request.QueryString.Item("LogixDefaultLanguage"))) <> "") AndAlso (Logix.TrimAll(Request.QueryString.Item("restext" & Request.QueryString.Item("LogixDefaultLanguage"))) <> "") Then
                Dim resolutionTexts As New List(Of ResolutionText)()
                Dim installedLangTbl As New DataTable()
                'get all the installed languages
                MyCommon.QueryStr = "Select LanguageID from Languages where InstalledForUI=1"
                installedLangTbl = MyCommon.LRT_Select()
                For Each row In installedLangTbl.Rows
                    'get the resolution text and title for each language
                    Dim resolutionText As New ResolutionText()
                    resolutionText.Title = Logix.TrimAll(Request.QueryString.Item("restitle" & row.Item("LanguageID")))
                    resolutionText.Description = Logix.TrimAll(Request.QueryString.Item("restext" & row.Item("LanguageID")))
                    resolutionText.LanguageID = row.Item("LanguageID")
                    resolutionTexts.Add(resolutionText)
                Next
                Dim resolution As CMS.AMS.Models.Resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                                                                 .LastUpdate = (DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds,
                                                                                                 .ParamID = ParamID,
                                                                                                 .ResolutionID = Request.QueryString("ResolutionID"),
                                                                                                 .UserID = AdminUserID,
                                                                                                 .DisplayOrder = Request.QueryString.Item("DisplayOrder")}
                If resolution.ResolutionTexts.Count > 0 Then
                    If FromServerHealth = 1 Then
                        Dim result As KeyValuePair(Of String, String) = RESTServiceHelper.CallService(Of String)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Put, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
                        infoMessage = result.Value
                        resolution.ResolutionID = result.Key
                    Else
                        MyCommon.QueryStr = "update HealthResolutions set " & _
                                  "Title=N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restitle")) & "', " & _
                                  "ResolutionText=N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restext")) & "', " & _
                                  "LastUpdate=getDate(), " & _
                                  "AdminUserID=" & AdminUserID & " " & _
                                  "where ResolutionID=" & Request.QueryString("ResolutionID") & ";"
                        MyCommon.LWH_Execute()
                    End If
                Else
                    infoMessage = Copient.PhraseLib.Lookup("serverhealth.resolutionerror", LanguageID)
                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("serverhealth.resolutionerror", LanguageID)
            End If
        Else
            If (Logix.TrimAll(Request.QueryString.Item("restitle")) <> "") AndAlso (Logix.TrimAll(Request.QueryString.Item("restext")) <> "") Then
                Dim resolutionTexts As New List(Of ResolutionText)()
                resolutionTexts.Add(New ResolutionText With {.Description = Logix.TrimAll(Request.QueryString.Item("restext")),
                                                             .Title = Logix.TrimAll(Request.QueryString.Item("restitle")),
                                                             .LanguageID = LanguageID})
                Dim resolution As CMS.AMS.Models.Resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                                                                  .LastUpdate = (DateTime.UtcNow - New DateTime(1970, 1, 1)).TotalMilliseconds,
                                                                                                  .ParamID = ParamID,
                                                                                                  .ResolutionID = Request.QueryString("ResolutionID"),
                                                                                                  .UserID = AdminUserID,
                                                                                                  .DisplayOrder = Request.QueryString.Item("DisplayOrder")}
                If FromServerHealth = 1 Then
                    Dim result As KeyValuePair(Of String, String) = RESTServiceHelper.CallService(Of String)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Put, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
                    infoMessage = result.Value
                Else
                    MyCommon.QueryStr = "update HealthResolutions set " & _
                                  "Title=N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restitle")) & "', " & _
                                  "ResolutionText=N'" & MyCommon.Parse_Quotes(Request.QueryString.Item("restext")) & "', " & _
                                  "LastUpdate=getDate(), " & _
                                  "AdminUserID=" & AdminUserID & " " & _
                                  "where ResolutionID=" & Request.QueryString("ResolutionID") & ";"
                    MyCommon.LWH_Execute()

                End If
            Else
                infoMessage = Copient.PhraseLib.Lookup("serverhealth.resolutionerror", LanguageID)
            End If
        End If
    ElseIf (Request.QueryString("mode") = "Delete") Then
        If FromServerHealth = 1 Then
            Dim result As KeyValuePair(Of String, String) = RESTServiceHelper.CallService(Of String)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions/" + Request.QueryString("ResolutionID"), LanguageID, HttpMethod.Delete, String.Empty, False, headers)
            infoMessage = result.Value
        Else
            MyCommon.QueryStr = "update HealthResolutions with (RowLock) set Deleted=1" & _
                                " where ResolutionID=" & Request.QueryString("ResolutionID")
            MyCommon.LWH_Execute()

        End If
    ElseIf (Request.QueryString("Up") <> "") Then
        If FromServerHealth = 1 Then
            Dim resolutionTexts As New List(Of ResolutionText)()
            Dim resolution As CMS.AMS.Models.Resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                                                              .ParamID = ParamID,
                                                                                              .ResolutionID = Request.QueryString("ResolutionID"),
                                                                                              .DisplayOrder = Request.QueryString.Item("DisplayOrderPrev")}

            Dim result As KeyValuePair(Of String, String) = RESTServiceHelper.CallService(Of String)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Put, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
            infoMessage = result.Value

            resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                            .ParamID = ParamID,
                                                            .ResolutionID = Request.QueryString("ResolutionIDPrev"),
                                                            .DisplayOrder = Request.QueryString.Item("DisplayOrder")}
            result = RESTServiceHelper.CallService(Of String)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Put, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
            infoMessage = result.Value

        Else
            MyCommon.QueryStr = "select EntryOrder from HealthResolutions where ResolutionID=" & MyCommon.Extract_Val(Request.QueryString("PrevID")) & ";"
            dst = MyCommon.LWH_Select
            If dst.Rows.Count > 0 Then
                swapEntryOrder = MyCommon.NZ(dst.Rows(0).Item("EntryOrder"), 1)
                MyCommon.QueryStr = "update HealthResolutions with (RowLock) set EntryOrder=" & swapEntryOrder & _
                                    " where ResolutionID=" & MyCommon.Extract_Val(Request.QueryString("ResolutionID")) & ";"
                MyCommon.LWH_Execute()
                MyCommon.QueryStr = "update HealthResolutions with (RowLock) set EntryOrder=" & MyCommon.Extract_Val(Request.QueryString("EntryOrder")) & _
                                    " where ResolutionID=" & MyCommon.Extract_Val(Request.QueryString("PrevID")) & ";"
                MyCommon.LWH_Execute()
            End If
        End If
    ElseIf (Request.QueryString("Down") <> "") Then
        If FromServerHealth = 1 Then
            Dim resolutionTexts As New List(Of ResolutionText)()
            Dim resolution As CMS.AMS.Models.Resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                                                                  .ParamID = ParamID,
                                                                                                  .ResolutionID = Request.QueryString("ResolutionID"),
                                                                                                  .DisplayOrder = Request.QueryString.Item("DisplayOrderNext")}


            Dim result As KeyValuePair(Of String, String) = RESTServiceHelper.CallService(Of String)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Put, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
            infoMessage = result.Value
            resolution = New CMS.AMS.Models.Resolution With {.ResolutionTexts = resolutionTexts,
                                                                                                 .ParamID = ParamID,
                                                                                                 .ResolutionID = Request.QueryString("ResolutionIDNext"),
                                                                                                 .DisplayOrder = Request.QueryString.Item("DisplayOrder")}
            result = RESTServiceHelper.CallService(Of String)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions", LanguageID, HttpMethod.Put, Newtonsoft.Json.JsonConvert.SerializeObject(New With {.resolution = resolution}), False, headers)
            infoMessage = result.Value
        Else
            MyCommon.QueryStr = "select EntryOrder from     HealthResolutions where ResolutionID=" & MyCommon.Extract_Val(Request.QueryString("NextID")) & ";"
            dst = MyCommon.LWH_Select
            If dst.Rows.Count > 0 Then
                swapEntryOrder = MyCommon.NZ(dst.Rows(0).Item("EntryOrder"), 1)
                MyCommon.QueryStr = "update HealthResolutions with (RowLock) set EntryOrder=" & swapEntryOrder & _
                                    " where ResolutionID=" & MyCommon.Extract_Val(Request.QueryString("ResolutionID")) & ";"
                MyCommon.LWH_Execute()
                MyCommon.QueryStr = "update HealthResolutions with (RowLock) set EntryOrder=" & Request.QueryString("EntryOrder") & _
                                    " where ResolutionID=" & MyCommon.Extract_Val(Request.QueryString("NextID")) & ";"
                MyCommon.LWH_Execute()
            End If
        End If
    End If

    Send_BodyBegin(3)

    If (Logix.UserRoles.ViewHealthResolution = False) Then
        Send_Denied(1, "perm.view-resolution")
        GoTo done
    End If

    If FromServerHealth = 1 Then
        Dim resolutions As CMS.AMS.Models.Resolutions

        dst.Columns.Add("ResolutionID", GetType(Integer))
        dst.Columns.Add("AdminUserID", GetType(Integer))
        dst.Columns.Add("CreatedDate", GetType(DateTime))
        dst.Columns.Add("LastUpdate", GetType(DateTime))
        dst.Columns.Add("EntryOrder", GetType(Integer))
        dst.Columns.Add("ResolutionText", GetType(String))
        dst.Columns.Add("Title", GetType(String))
        dst.Columns.Add("languageId", GetType(Integer))

        Dim result As KeyValuePair(Of CMS.AMS.Models.Resolutions, String) = RESTServiceHelper.CallService(Of CMS.AMS.Models.Resolutions)(RESTServiceList.ServerHealthService, HealthServiceURL + "/resolutions/" + ParamID, LanguageID, HttpMethod.Get, String.Empty, False, headers)
        resolutions = result.Key
        infoMessage += result.Value
        ' Dim resolutionText As CMS.AMS.Models.ResolutionText 
        If MultiLanguageEnabled Then
            If Not resolutions Is Nothing Then
                For Each item As CMS.AMS.Models.Resolution In resolutions.ResolutionList
                    For Each resolutionTexts As ResolutionText In item.ResolutionTexts
                        dst.Rows.Add(item.ResolutionID, item.UserID, CMS.AMS.ExtentionMethods.ConvertToLocalDateTime(item.LastUpdate), CMS.AMS.ExtentionMethods.ConvertToLocalDateTime(item.LastUpdate), item.DisplayOrder, resolutionTexts.Description, resolutionTexts.Title, resolutionTexts.LanguageID)
                    Next

                Next
            End If
        Else
            If Not resolutions Is Nothing Then
                For Each item As CMS.AMS.Models.Resolution In resolutions.ResolutionList
                    Dim resolutionText = (From r In item.ResolutionTexts Where r.LanguageID = LanguageID Select r)
                    If resolutionText(0) Is Nothing Then
                        resolutionText = (From r In item.ResolutionTexts Where r.LanguageID = 1 Select r)
                    End If
                    If resolutionText(0) IsNot Nothing Then
                        dst.Rows.Add(item.ResolutionID, item.UserID, CMS.AMS.ExtentionMethods.ConvertToLocalDateTime(item.LastUpdate), CMS.AMS.ExtentionMethods.ConvertToLocalDateTime(item.LastUpdate), item.DisplayOrder, resolutionText(0).Description, resolutionText(0).Title, resolutionText(0).LanguageID)
                    End If

                Next
            End If
        End If

        dst2.Columns.Add("ErrorDescription", GetType(String))
        dst2.Columns.Add("ErrorCode", GetType(String))
        dst2.Columns.Add("Severity", GetType(String))
        dst2.Columns.Add("PhraseID", GetType(Integer))
        dst2.Rows.Add(CMS.AMS.ServerHealthHelper.GetErrorDefinition(ParamID, LanguageID), ParamID, "", 0)

    Else
        ' Query to get resolution postings regarding the error
        MyCommon.QueryStr = "select ResolutionID, AdminUserID, CreatedDate, LastUpdate, EntryOrder, ResolutionText, Title from HealthResolutions " & _
                            "where ErrorID=" & ErrorID & " and ServerTypeID=" & ServerType & " and Deleted=0 order by EntryOrder;"
        dst = MyCommon.LWH_Select

        ' Query to get the definition of the error
        MyCommon.QueryStr = "select ErrorDescription, ErrorCode, Severity, PhraseID from HealthErrors " & _
                            "where ErrorID=" & ErrorID & " and ServerTypeID=" & ServerType & ";"
        dst2 = MyCommon.LWH_Select
    End If

%>
<div id="intro">
    <h1 id="title">
        <%
            If ServerType = 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.localserver", LanguageID) & " " & Copient.PhraseLib.Lookup("term.error", LanguageID))
            ElseIf ServerType = 2 Then
                Sendb(Copient.PhraseLib.Lookup("term.centralserver", LanguageID) & " " & Copient.PhraseLib.Lookup("term.error", LanguageID))
            End If
            If ErrorID <> 0 Then
                Send(" #" & ErrorID)
            End If
        %>
    </h1>
    <div id="controls">
    </div>
</div>
<div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
        <div class="box" id="general">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.definition", LanguageID))%>
                </span>
            </h2>
            <%
                If dst2.Rows.Count > 0 Then
                    If Not IsDBNull(dst2.Rows(0).Item("PhraseID")) AndAlso dst2.Rows(0).Item("PhraseID") > 0 Then
                        Send(Copient.PhraseLib.Lookup(dst2.Rows(0).Item("PhraseID"), LanguageID) & "<br />")
                    Else
                        Send(MyCommon.NZ(dst2.Rows(0).Item("ErrorDescription"), "...") & "<br />")
                    End If
                Else
                    Send(Copient.PhraseLib.Lookup("store-health.nodefinition", LanguageID) & "<br />")
                End If
                If ServerType = 1 Then
                    Send("<br class=""half"" />")
                    Sendb(Copient.PhraseLib.Lookup("term.severity", LanguageID) & ": ")
                    If (dst2.Rows.Count > 0) Then
                        Select Case MyCommon.NZ(dst2.Rows(0).Item("Severity"), 0)
                            Case 1 ' High
                                Send("<span class=""red"">" & Copient.PhraseLib.Lookup("term.high", LanguageID) & "</span><br />")
                            Case 5 'Medium
                                Send("<span class=""darkred"">" & Copient.PhraseLib.Lookup("term.medium", LanguageID) & "</span><br />")
                            Case 10 ' Low
                                Send("<span>" & Copient.PhraseLib.Lookup("term.low", LanguageID) & "</span><br />")
                            Case 0 ' Zero or nothing
                                Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
                        End Select
                    Else
                        Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
                    End If
                End If
                If (Logix.UserRoles.CreateandEditHealthResolution) Then
                    Send("      <br class=""half"" />")
                    Send("      <b><a href=""#"" onclick=""javascript:AddResolution()"">" & Copient.PhraseLib.Lookup("health-resolutions.submit", LanguageID) & "</a></b><br />")
                End If
            %>
        </div>
        <div class="box" id="resolutions">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.resolutions", LanguageID))%>
                </span>
            </h2>
            <%
                If dst.Rows.Count > 0 Then
                    i = 0
                    While (i < dst.Rows.Count)
                        ResolutionID = MyCommon.NZ(dst.Rows(i).Item("ResolutionID"), 0)
                        If MyCommon.NZ(dst.Rows(i).Item("AdminUserID"), 0) <> 0 Then
                            MyCommon.QueryStr = "select FirstName, LastName from AdminUsers where AdminUserID=" & dst.Rows(i).Item("AdminUserID") & ";"
                            rst = MyCommon.LRT_Select
                            If rst.Rows.Count > 0 Then
                                FullName = MyCommon.NZ(rst.Rows(0).Item("FirstName"), "") & " " & MyCommon.NZ(rst.Rows(0).Item("LastName"), "")
                            Else
                                FullName = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
                            End If
                        Else
                            FullName = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
                        End If
                        'show only those resoluitions which are in user selected language or logix default language
                        Dim flagLang As Boolean = True
                        If FromServerHealth = 1 Then
                            If dst.Rows(i).Item("languageId") <> LanguageID Then
                                Dim data = (From r In dst.AsEnumerable() Where r.Field(Of Integer)("ResolutionID") = dst.Rows(i).Item("ResolutionID") And r.Field(Of Integer)("LanguageID") = LanguageID)
                                If data IsNot Nothing AndAlso data.Count > 0 Then
                                    'For Each row3 As DataRow In query.Rows
                                    '    If row3.Item("languageId") = LanguageID Then
                                    flagLang = False
                                    'End If
                                    '    Next
                                End If
                                If flagLang = True AndAlso dst.Rows(i).Item("languageId") <> 1 Then
                                    flagLang = False
                                End If
                            End If
                        End If
                        If flagLang Then
                            Send("<br class=""half"" />")
                            Send("<div id=""resolution" & ResolutionID & """ style=""padding-bottom:10px;" & IIf((i + 1) = dst.Rows.Count, "", "border-bottom:1px solid #999999;") & """>")
                            Send("  <form id=""editform" & ResolutionID & """ name=""editform" & MyCommon.NZ(dst.Rows(i).Item("ResolutionID"), 0) & """ action="""">")
                            Send("    <input type=""hidden"" id=""ResolutionID" & ResolutionID & """ name=""ResolutionID"" value=""" & ResolutionID & """ />")
                        
                            If FromServerHealth = 1 Then
                                Send("    <input type=""hidden"" id=""ParamID" & ResolutionID & """ name=""ParamID"" value=""" & ParamID & """ />")
                                Send("    <input type=""hidden"" id=""DisplayOrder" & ResolutionID & """ name=""DisplayOrder"" value=""" & MyCommon.NZ(dst.Rows(i).Item("EntryOrder"), 0) & """ />")
                                'find the previous resolution shown in ui
                                Dim j As Integer = 1
                                Dim indexOfPrevRow As Integer = 0
                                While (i - j >= 0)
                                    If ResolutionID = dst.Rows(i - j).Item("ResolutionID") Then
                                        j = j + 1
                                    Else
                                        If dst.Rows(i - j).Item("languageId") = dst.Rows(i).Item("languageId") Then
                                            indexOfPrevRow = i - j
                                            Exit While
                                        Else
                                            Dim recWithUserLanguageId = (From c In dst.AsEnumerable() Where c.Field(Of Integer)("ResolutionID") = dst.Rows(i - j).Item("ResolutionID") And c.Field(Of Integer)("languageId") = LanguageID)
                                            If recWithUserLanguageId IsNot Nothing AndAlso recWithUserLanguageId.Count > 0 Then
                                                j = j + 1
                                            Else
                                                If dst.Rows(i - j).Item("languageId") = 1 Then
                                                    indexOfPrevRow = i - j
                                                    Exit While
                                                Else
                                                    j = j + 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End While
                                If indexOfPrevRow = 0 AndAlso i <> j Then
                                    UpArrowDisabled = True
                                    Send("    <input type=""hidden"" id=""DisplayOrderPrev" & ResolutionID & """ name=""DisplayOrderPrev"" value=""0"" />")
                                    Send("    <input type=""hidden"" id=""ResolutionIDPrev" & ResolutionID & """ name=""ResolutionIDPrev"" value=""0""/>")
                                Else
                                    UpArrowDisabled = False
                                    Send("    <input type=""hidden"" id=""DisplayOrderPrev" & ResolutionID & """ name=""DisplayOrderPrev"" value=""" & MyCommon.NZ(dst.Rows(indexOfPrevRow).Item("EntryOrder"), 0) & """ />")
                                    Send("    <input type=""hidden"" id=""ResolutionIDPrev" & ResolutionID & """ name=""ResolutionIDPrev"" value=""" & MyCommon.NZ(dst.Rows(indexOfPrevRow).Item("ResolutionID"), 0) & """ />")
                                End If
                                'find the next resolution to be shown in ui
                                Dim k As Integer = 1
                                Dim indexOfNextRow As Integer = 0
                                While (i + k <= dst.Rows.Count - 1)
                                    If ResolutionID = dst.Rows(i + k).Item("ResolutionID") Then
                                        k = k + 1
                                    Else
                                        If dst.Rows(i + k).Item("languageId") = dst.Rows(i).Item("languageId") Then
                                            indexOfNextRow = i + k
                                            Exit While
                                        Else
                                            Dim recWithUserLanguageId = (From c In dst.AsEnumerable() Where c.Field(Of Integer)("ResolutionID") = dst.Rows(i + k).Item("ResolutionID") And c.Field(Of Integer)("languageId") = LanguageID)
                                            If recWithUserLanguageId IsNot Nothing AndAlso recWithUserLanguageId.Count > 0 Then
                                                k = k + 1
                                            Else
                                                If dst.Rows(i + k).Item("languageId") = 1 Then
                                                    indexOfNextRow = i + k
                                                    Exit While
                                                Else
                                                    k = k + 1
                                                End If
                                            End If
                                        End If
                                            
                                    End If
                                End While
                                If indexOfNextRow = 0 Then
                                    DownArrowDisabled = True
                                    Send("    <input type=""hidden"" id=""DisplayOrderNext" & ResolutionID & """ name=""DisplayOrderNext"" value=""0"" />")
                                    Send("    <input type=""hidden"" id=""ResolutionIDNext" & ResolutionID & """ name=""ResolutionIDNext"" value=""0"" />")
                                Else
                                    DownArrowDisabled = False
                                    Send("    <input type=""hidden"" id=""DisplayOrderNext" & ResolutionID & """ name=""DisplayOrderNext"" value=""" & MyCommon.NZ(dst.Rows(indexOfNextRow).Item("EntryOrder"), 0) & """ />")
                                    Send("    <input type=""hidden"" id=""ResolutionIDNext" & ResolutionID & """ name=""ResolutionIDNext"" value=""" & MyCommon.NZ(dst.Rows(indexOfNextRow).Item("ResolutionID"), 0) & """ />")
                                End If
                                Send("    <input type=""hidden"" id=""FromServerHealth" & ResolutionID & """ name=""FromServerHealth"" value=""" & FromServerHealth & """ />")
                            Else
                                Send("    <input type=""hidden"" id=""SrvType" & ResolutionID & """ name=""SrvType"" value=""" & ServerType & """ />")
                                Send("    <input type=""hidden"" id=""Err" & ResolutionID & """ name=""Err"" value=""" & ErrorID & """ />")
                                Send("    <input type=""hidden"" id=""EntryOrder" & ResolutionID & """ name=""EntryOrder"" value=""" & MyCommon.NZ(dst.Rows(i).Item("EntryOrder"), 0) & """ />")
                                If i = 0 Then
                                    Send("    <input type=""hidden"" id=""PrevID" & ResolutionID & """ name=""PrevID"" value=""0"" />")
                                Else
                                    Send("    <input type=""hidden"" id=""PrevID" & ResolutionID & """ name=""PrevID"" value=""" & MyCommon.NZ(dst.Rows(i - 1).Item("ResolutionID"), 0) & """ />")
                                End If
                                If i = dst.Rows.Count - 1 Then
                                    Send("    <input type=""hidden"" id=""NextID" & ResolutionID & """ name=""NextID"" value=""0"" />")
                                Else
                                    Send("    <input type=""hidden"" id=""NextID" & ResolutionID & """ name=""NextID"" value=""" & MyCommon.NZ(dst.Rows(i + 1).Item("ResolutionID"), 0) & """ />")
                                End If
                            End If
                      
                            Send("    <div id=""savedres" & ResolutionID & """>")
                            Send("      <div class=""shadedlight"" style=""margin-bottom:3px;"">")
                            If (Logix.UserRoles.CreateandEditHealthResolution) Then
                                Send("        <input type=""button"" class=""edit"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""editResolution(" & MyCommon.NZ(dst.Rows(i).Item("ResolutionID"), 0) & ");"" />")
                                If UpArrowDisabled Then
                                    Sendb("        <input class=""up"" type=""submit"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ disabled=""disabled"" />")
                                Else
                                    Sendb("        <input class=""up"" type=""submit"" value=""&#9650;"" name=""up"" title=""" & Copient.PhraseLib.Lookup("term.moveup", LanguageID) & """ />")
                                End If
                                If DownArrowDisabled Then
                                    Send("<input class=""down"" type=""submit"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ disabled=""disabled"" />")
                                Else
                                    Send("<input class=""down"" type=""submit"" value=""&#9660;"" name=""down"" title=""" & Copient.PhraseLib.Lookup("term.movedown", LanguageID) & """ />")
                                End If
                            End If
                            If (Logix.UserRoles.DeleteHealthResolution) Then
                                Send("        <input class=""ex"" type=""button"" value=""X"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("store-health.confirmdelete", LanguageID) & "')){LoadDocument('health-resolutions.aspx?mode=Delete&FromServerHealth=" & FromServerHealth & "&ParamID=" & ParamID & "&ResolutionID=" & dst.Rows(i).Item("ResolutionID") & "&SrvType=" & ServerType & "&Err=" & ErrorID & "')}else{return false}"" />")
                            End If
                            Send("<span id=""title" & ResolutionID & """ >")
                            Send("        <h3 style=""display:inline;"">" & MyCommon.NZ(dst.Rows(i).Item("Title"), "—") & "</h3></span>")
                            Send("      </div>")
                            Send("<span id=""text" & ResolutionID & """ >")
                            Send("      " & MyCommon.NZ(dst.Rows(i).Item("ResolutionText"), "") & "</span><br />")
                            Send("      — <i>" & FullName & ", " & Logix.ToShortDateTimeString(dst.Rows(i).Item("LastUpdate"), MyCommon) & "</i>")
                            Send("      <br />")
                            Send("    </div>")
                       
                            Send("    <div id=""editres" & MyCommon.NZ(dst.Rows(i).Item("ResolutionID"), 0) & """ style=""display:none;"">")
                            Send("      <div class=""shadedlight"" style=""margin-bottom:2px;"">")
                            Send("        <input type=""button"" class=""cancel"" id=""cancel" & MyCommon.NZ(dst.Rows(i).Item("ResolutionID"), 0) & """ name=""cancel"" value=""Cancel"" onclick=""cancelResolution(" & MyCommon.NZ(dst.Rows(i).Item("ResolutionID"), 0) & ");"" />")
                            Send("     <input type=""hidden"" id=""LogixDefaultLanguage"" name=""LogixDefaultLanguage"" value=""" & MyCommon.Fetch_SystemOption(1) & """ />")
                            Send("        <input type=""submit"" class=""edit"" id=""edit" & MyCommon.NZ(dst.Rows(i).Item("ResolutionID"), 0) & """ name=""edit"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
                            If MultiLanguageEnabled AndAlso FromServerHealth = 1 Then
                                Dim dstLang As New DataTable()
                                MyCommon.QueryStr = "select L.LanguageID,L.Name,L.PhraseTerm from Languages as L where InstalledForUI=1"
                                dstLang = MyCommon.LRT_Select()
                                Send("    <label for=""restitle"">" & Copient.PhraseLib.Lookup("term.title", LanguageID) & ":</label><br />")
                                Send("    <input type=""hidden"" id=""editUserLanguage"" name=""editUserLanguage"" value=""" & LanguageID & """ />")
                                For Each row In dstLang.Rows
                                    Dim resolutiontextOb = (From r As DataRow In dst Where r.Item("languageId") = row.Item("LanguageID") And r.Item("resolutionid") = dst.Rows(i).Item("resolutionid"))
                                    Send("    <label for=""LanguageName"">" & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID) & ":</label><br />")
                                    If resolutiontextOb(0) IsNot Nothing Then
                                        Send("        <input type=""text"" id=""restitle" & row.Item("LanguageID") & """ name=""restitle" & row.Item("LanguageID") & """ style=""width:65%;font-weight:bold;"" maxlength=""50"" value=""" & MyCommon.NZ(resolutiontextOb(0).Item("title"), 0) & """  /><br />")
                                    Else
                                        Send("        <input type=""text"" id=""restitle" & row.Item("LanguageID") & """ name=""restitle" & row.Item("LanguageID") & """ style=""width:65%;font-weight:bold;"" maxlength=""50"" value=""""  /><br />")
                                    End If
                                    
                                Next
                                Send("      </div>")
                                Send("    <label for=""restext"">" & Copient.PhraseLib.Lookup("term.text", LanguageID) & ":</label><br />")
                                For Each row In dstLang.Rows
                                    Dim resolutiontextOb = (From r As DataRow In dst Where r.Item("languageId") = row.Item("LanguageID") And r.Item("resolutionid") = dst.Rows(i).Item("resolutionid"))
                                    Send("    <label for=""LanguageName"">" & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID) & ":</label><br />")
                                    If resolutiontextOb(0) IsNot Nothing Then
                                        Send("      <textarea id=""restext" & row.Item("LanguageID") & """ name=""restext" & row.Item("LanguageID") & """ style=""height:120px;width:99%;"">" & MyCommon.NZ(resolutiontextOb(0).Item("ResolutionText"), 0) & "</textarea><br />")
                                    Else
                                        Send("      <textarea id=""restext" & row.Item("LanguageID") & """ name=""restext" & row.Item("LanguageID") & """ style=""height:120px;width:99%;""></textarea><br />")
                                    End If
                                Next
                                Send("    </div>")
                            Else
                                Send("        <input type=""text"" id=""restitle" & ResolutionID & """ name=""restitle"" style=""width:65%;font-weight:bold;"" maxlength=""50"" value=""" & MyCommon.NZ(dst.Rows(i).Item("Title"), 0) & """ /><br />")
                                Send("      </div>")
                                Send("      <textarea id=""restext" & ResolutionID & """ name=""restext"" style=""height:120px;width:99%;"">" & MyCommon.NZ(dst.Rows(i).Item("ResolutionText"), 0) & "</textarea><br />")
                                Send("    </div>")
                            
                            End If
                            Send("  </form>")
                            Send("</div>")
                        End If
                        i = i + 1
                        IIf(shaded, shaded = False, shaded = True)
                    End While
                Else
                    Send("<div id=""resolution"">")
                    Send("  <i>" & Copient.PhraseLib.Lookup("store-health.noresolutions", LanguageID) & "</i>")
                    Send("</div>")
                End If
            %>
        </div>
        <%
            'Create New Health Resolution
            If (Logix.UserRoles.CreateandEditHealthResolution) Then
                Send("<div class=""box"" id=""newresolution"">")
                Send("  <form id=""newform"" name=""newform"" action="""">")
                Send("    <input type=""submit"" class=""save"" id=""save"" name=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
                Send("    <h2>")
                Send("      <span>" & Copient.PhraseLib.Lookup("store-health.newresolution", LanguageID) & "</span>")
                Send("    </h2>")
                Send("    <input type=""hidden"" id=""SrvType"" name=""SrvType"" value=""" & ServerType & """ />")
                Send("    <input type=""hidden"" id=""Err"" name=""Err"" value=""" & ErrorID & """ />")
                Send("    <input type=""hidden"" id=""Err"" name=""Err"" value=""" & ErrorID & """ />")
                Send("    <input type=""hidden"" id=""ParamID"" name=""ParamID"" value=""" & ParamID & """ />")
                Send("    <input type=""hidden"" id=""UserLanguage"" name=""UserLanguage"" value=""" & LanguageID & """ />")
                Send("     <input type=""hidden"" id=""LogixDefaultLanguage"" name=""LogixDefaultLanguage"" value=""" & MyCommon.Fetch_SystemOption(1) & """ />")
                Send("    <input type=""hidden"" id=""FromServerHealth"" name=""FromServerHealth"" value=""" & FromServerHealth & """ />")
                If MultiLanguageEnabled AndAlso FromServerHealth = 1 Then
                    Dim dstLang As New DataTable()
                    MyCommon.QueryStr = "select L.LanguageID,L.Name,L.PhraseTerm from Languages as L where InstalledForUI=1"
                    dstLang = MyCommon.LRT_Select()
                    Send("    <label for=""restitle"">" & Copient.PhraseLib.Lookup("term.title", LanguageID) & ":</label><br />")
                    If dstLang IsNot Nothing AndAlso dstLang.Rows.Count > 0 Then
                        For Each row In dstLang.Rows
                            Send("    <label for=""LanguageName"">" & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID) & ":</label><br />")
                            Send("    <input type=""text"" id=""restitle" & row.Item("LanguageID") & """ name=""restitle" & row.Item("LanguageID") & """ style=""width:99%;"" maxlength=""50"" /><br />")
                        Next
                        Send("    <label for=""restext"">" & Copient.PhraseLib.Lookup("term.text", LanguageID) & ":</label><br />")
                        For Each row In dstLang.Rows
                            Send("    <label for=""LanguageName"">" & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID) & ":</label><br />")
                            Send("    <textarea id=""restext" & row.Item("LanguageID") & """ name=""restext" & row.Item("LanguageID") & """ style=""height:120px;width:99%;""></textarea><br />")
                        Next
                    End If
                Else
                    Send("    <label for=""restitle"">" & Copient.PhraseLib.Lookup("term.title", LanguageID) & ":</label><br />")
                    Send("    <input type=""text"" id=""restitle "" name=""restitle"" style=""width:99%;"" maxlength=""50"" /><br />")
                    Send("    <label for=""restext"">" & Copient.PhraseLib.Lookup("term.text", LanguageID) & ":</label><br />")
                    Send("    <textarea id=""restext"" name=""restext"" style=""height:120px;width:99%;""></textarea><br />")
                End If
                Send("  </form>")
                Send("</div>")
                
            End If
        %>
    </div>
</div>
<%
done:
    Send_BodyEnd()
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixWH()
    MyCommon = Nothing
    Logix = Nothing
%>
