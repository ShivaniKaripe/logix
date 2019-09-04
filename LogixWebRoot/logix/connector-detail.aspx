<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>

<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Linq" %>
<%@ Import Namespace="System.Xml.Linq" %>
<%@ Import Namespace="System.IO" %>

<%
  ' *****************************************************************************
  ' * FILENAME: connector-detail.aspx 
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
  Dim dt, dt2, dt3 As DataTable
  Dim row, row2 As System.Data.DataRow
  Dim ConnectorID As Integer = 0
    Dim ConnectorName As String = ""
    Dim NamePhraseID As Integer = 0
  Dim ConnectorDesc As String = ""
  Dim Path As String = ""
  Dim Installed As Boolean = True
  Dim Visible As Boolean = True
  Dim UsesGUIDs As Boolean = False
  Dim NewGUID As String = ""
  Dim DeleteGUID As String = ""
  Dim SaveGUID As String = ""
  Dim SaveGUIDDesc As String = ""
  Dim SaveGUIDChan As String = ""
  Dim SavePKID As Integer = 0  'The primary key of the ConnectorGUIDs row being saved
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim OptSelected As String = ""
  Dim TempStr As String = ""
  Dim OptionObj As Copient.SystemOption = Nothing
    Dim HistoryStr As String = ""
    Dim GuidStr As String = ""
  Dim OptionID As Integer
  Dim OldGUID As String = ""
    Dim ChannelName As String = ""
    Dim objcPattern As AMSResult(Of CouponPattern) = New AMSResult(Of CouponPattern)()
    Dim objcSettings As AMSResult(Of CouponConfig) = New AMSResult(Of CouponConfig)()
    Dim objCouponPattern As CouponPattern
    Dim objCouponSettings As CouponConfig
    Dim objcSett As CouponSettings
    Dim CreatedDate As String = ""
    CurrentRequest.Resolver.AppName = "connector-detail.aspx"
    Dim objPatternService As ICouponPatternService = CurrentRequest.Resolver.Resolve(Of ICouponPatternService)()
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "connector-detail.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
  If (Request.QueryString("ConnectorID") <> "") Then
    ConnectorID = MyCommon.Extract_Val(Request.QueryString("ConnectorID"))
  End If
    If (Request.QueryString("infoMessage") <> "") Then
        infoMessage = Request.QueryString("infoMessage")
    End If
    MyCommon.QueryStr = "select GUID from ConnectorGUIDs where ConnectorID=70"
    dt = MyCommon.LRT_Select()
    If dt.Rows.Count > 0 Then
        GuidStr = MyCommon.NZ(dt.Rows(0).Item("GUID"), "")
    End If
   
  If (Request.QueryString("NewGUID") <> "") Then
    If (ConnectorID = 53) Then
      MyCommon.QueryStr = "insert into ConnectorGUIDs (ConnectorID, GUID, Description, CreatedDate, LastUpdate,ExtInterfaceID) values (" & ConnectorID & ", '" & GUID() & "', N'" & MyCommon.Parse_Quotes(Request.QueryString("NewGUIDdesc")) & "', getdate(), getdate(),N'" & MyCommon.Parse_Quotes(Request.QueryString("drpExtInterfaceId")) & "');"
      MyCommon.LRT_Execute()
    ElseIf (ConnectorID = 58) Then  ' Channel connector
      MyCommon.QueryStr = "insert into ConnectorGUIDs (ConnectorID, GUID, Description, ChannelID, CreatedDate, LastUpdate) values (@ConnectorID, @GUID, @Description, @ChannelID, getdate(), getdate());"
      MyCommon.DBParameters.Add("@ConnectorID", SqlDbType.Int).Value = ConnectorID
      MyCommon.DBParameters.Add("@GUID", SqlDbType.NVarChar, 36).Value = GUID()
      MyCommon.DBParameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = Request.QueryString("NewGUIDdesc")
      Dim NewGUIDchan = 0
      Integer.TryParse(GetCgiValue("NewGUIDChan"), NewGUIDchan)
      MyCommon.DBParameters.Add("@ChannelID", SqlDbType.Int).Value = NewGUIDchan
      MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
    Else
      MyCommon.QueryStr = "insert into ConnectorGUIDs (ConnectorID, GUID, Description, CreatedDate, LastUpdate) values (" & ConnectorID & ", '" & GUID() & "', N'" & MyCommon.Parse_Quotes(Request.QueryString("NewGUIDdesc")) & "', getdate(), getdate());"
      MyCommon.LRT_Execute()
    End If
    MyCommon.Activity_Log(42, ConnectorID, AdminUserID, Copient.PhraseLib.Lookup("history.connector-guidadd", LanguageID))
    Response.Redirect("/logix/connector-detail.aspx?ConnectorID=" & ConnectorID)
       
  ElseIf (Request.QueryString("SaveGUID") <> "") Or (Request.QueryString("SaveGUIDDesc") <> "") Or (Request.QueryString("SaveGUIDChan") <> "") Then
    OldGUID = Request.QueryString("OldGUID")
    SaveGUID = Request.QueryString("SaveGUID")
    SaveGUIDDesc = Request.QueryString("SaveGUIDDesc")
    SaveGUIDChan = GetCgiValue("SaveGUIDChan")
    If Not (Integer.TryParse(GetCgiValue("PK"), SavePKID)) Then SavePKID = 0
               
    If Not (Regex.Match(SaveGUID, "^[a-fA-Z0-9]{8}-[a-fA-Z0-9]{4}-[a-fA-Z0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12}$").Success) Then
      infoMessage = Copient.PhraseLib.Lookup("term.Invalid-Guid", LanguageID)
    Else
      MyCommon.QueryStr = "select GUID, description, ChannelID from ConnectorGUIDs where PKID=@PKID;"
      MyCommon.DBParameters.Add("@PKID", SqlDbType.Int).Value = SavePKID
      dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
      If (dt.Rows.Count > 0) Then
        'first, make sure the GUID we are saving is not already in use for the same connector
        MyCommon.QueryStr = "select GUID from ConnectorGUIDs where ConnectorID=@ConnectorID and GUID=@SaveGUID and not(PKID=@PKID);"
        MyCommon.DBParameters.Add("@ConnectorID", SqlDbType.Int).Value = ConnectorID
        MyCommon.DBParameters.Add("@SaveGUID", SqlDbType.NVarChar, 36).Value = SaveGUID
        MyCommon.DBParameters.Add("@PKID", SqlDbType.Int).Value = SavePKID
        dt3 = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt3.Rows.Count > 0 Then
          infoMessage = Copient.PhraseLib.Lookup("term.duplicate-guid", LanguageID)
        Else
          MyCommon.QueryStr = "update ConnectorGUIDs set GUID=@SaveGUID, Description=@SaveGUIDDesc, LastUpdate=getdate()"
          MyCommon.DBParameters.Add("@SaveGUID", SqlDbType.NVarChar, 36).Value = SaveGUID
          MyCommon.DBParameters.Add("@SaveGUIDDesc", SqlDbType.NVarChar, 1000).Value = SaveGUIDDesc
          If ConnectorID = 58 Then
            MyCommon.QueryStr = MyCommon.QueryStr & ", ChannelID=@SaveGUIDChan"
            MyCommon.DBParameters.Add("SaveGUIDChan", SqlDbType.Int).Value = SaveGUIDChan
          End If
          MyCommon.QueryStr = MyCommon.QueryStr & " where PKID=@PKID"
          MyCommon.DBParameters.Add("@PKID", SqlDbType.Int).Value = SavePKID
          MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
          MyCommon.Activity_Log(42, ConnectorID, AdminUserID, Copient.PhraseLib.Lookup("history.connector-guidedit", LanguageID))
          Response.Redirect("/logix/connector-detail.aspx?ConnectorID=" & ConnectorID)
        End If
      Else
        'we can't save the GUID because it doesn't exist!
        infoMessage = Copient.PhraseLib.Lookup("term.guidnotupdated", LanguageID)  'GUID not updated - it does not exist
      End If
    End If
  ElseIf (Request.QueryString("DeleteGUID") <> "") Then
    DeleteGUID = Request.QueryString("DeleteGUID")
    MyCommon.QueryStr = "select GUID from ConnectorGUIDs where ConnectorID=" & ConnectorID & " and GUID='" & DeleteGUID & "';"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      MyCommon.QueryStr = "delete from ConnectorGUIDs where ConnectorID=" & ConnectorID & " and GUID='" & DeleteGUID & "';"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(42, ConnectorID, AdminUserID, Copient.PhraseLib.Lookup("history.connector-guiddelete", LanguageID))
      Response.Redirect("/logix/connector-detail.aspx?ConnectorID=" & ConnectorID)
    End If
  ElseIf (Request.QueryString("SaveOptions") <> "") Then
    MyCommon.QueryStr = "select OptionID, OptionName, OptionValue, PhraseID " & _
                    "from InterfaceOptions with (NoLock) " & _
                    "where ConnectorID = " & ConnectorID & " and Visible=1 " & _
                    "order by OptionName;"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      For Each row In dt.Rows
        TempStr = MyCommon.Parse_Quotes(Request.QueryString("option" & MyCommon.NZ(row.Item("OptionID"), 0)))
        TempStr = Logix.TrimAll(TempStr)

        OptionObj = New Copient.SystemOption(MyCommon.NZ(row.Item("OptionID"), 0), MyCommon.NZ(row.Item("OptionValue"), ""))
        OptionObj.SetNewValue(TempStr)

        If OptionObj.IsModified Then
          MyCommon.QueryStr = "Update InterfaceOptions with (RowLock) set OptionValue=N'" & OptionObj.GetNewValue() & "', LastUpdate=getdate() where OptionID=" & OptionObj.GetOptionID()
          MyCommon.LRT_Execute()
            
          If (MyCommon.RowsAffected > 0) Then
            HistoryStr = Copient.PhraseLib.Lookup("history.edit-connectorsetting", LanguageID) & " '" & MyCommon.NZ(row.Item("OptionName"), "") & "'" & _
                         " from: " & OptionObj.GetOldValue() & " to: " & OptionObj.GetNewValue()
                        MyCommon.Activity_Log(42, ConnectorID, AdminUserID, HistoryStr)
          End If
        End If
      Next
      
        End If
    ElseIf (Request.QueryString("SaveFilter") <> "") Then
        Dim tmpAttr As New List(Of [String])
        Dim MyArray() As String
      
        If Not (Request.QueryString("defaultFilters") = "") Then
            MyArray = Request.QueryString("defaultFilters").Split(",")
            tmpAttr.AddRange(MyArray.ToList())
        End If
        If Not (Request.QueryString("selFilterXML") = "") Then
            MyArray = Request.QueryString("selFilterXML").Split(",")
            tmpAttr.AddRange(MyArray.ToList())
        End If
        If (tmpAttr.Count > 0) Then
            tmpAttr = tmpAttr.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).Distinct().ToList()
            Dim NonfilteredXMlLst As New Dictionary(Of [String], [String])()
            For Each s As String In tmpAttr
                MyArray = s.Split("/")
                If (NonfilteredXMlLst.ContainsKey(MyArray(0))) Then
                    Dim val As String = NonfilteredXMlLst(MyArray(0))
                    val = IIf(String.IsNullOrEmpty(val), MyArray(1), val & "," & MyArray(1))
                    NonfilteredXMlLst(MyArray(0)) = val
                Else
                    NonfilteredXMlLst.Add(MyArray(0), MyArray(1))
                End If
            Next
            If (NonfilteredXMlLst.Count > 0) Then
                Dim FilteredString As New StringBuilder()
                For Each pair As KeyValuePair(Of [String], [String]) In NonfilteredXMlLst
                    FilteredString.Append(pair.Key).Append("-").Append(pair.Value).Append("|")
                Next
                If (FilteredString.Length > 0) Then
                    MyCommon.QueryStr = "pa_Update_FilterOutputColumns"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ReferenceId", SqlDbType.BigInt).Value = MyCommon.Extract_Val(Request.QueryString("ddlWebMethods"))
                    MyCommon.LRTsp.Parameters.Add("@FilteredAttributes", SqlDbType.VarChar).Value = FilteredString.ToString().Trim("|")
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If
            End If
        End If
    ElseIf Request.QueryString("savepattern") <> "" Then
        objCouponPattern = New CouponPattern()
        Dim temp As Integer
        If Not (Integer.TryParse(Request.QueryString("spinner"), temp)) Then
            infoMessage = Copient.PhraseLib.Lookup("error.invalid-couponlength", LanguageID)
        Else
            objCouponPattern.Length = Request.QueryString("spinner")
            If Request.QueryString("prefix") IsNot Nothing Then
                objCouponPattern.Prefix = Request.QueryString("prefix")
            Else
                objCouponPattern.Prefix = ""
            End If
            If Request.QueryString("suffix") IsNot Nothing Then
                objCouponPattern.Suffix = Request.QueryString("suffix")
            Else
                objCouponPattern.Suffix = ""
            End If
            If Request.QueryString("textstart") IsNot Nothing Then
                objCouponPattern.Start = Request.QueryString("textstart")
            Else
                objCouponPattern.Start = ""
            End If
            If Request.QueryString("textendrange") IsNot Nothing Then
                objCouponPattern.End = Request.QueryString("textendrange")
            Else
                objCouponPattern.End = ""
            End If
            If Request.QueryString("types") = 1 Then
                objCouponPattern.ContentType = "NUMERIC"
            Else
                objCouponPattern.ContentType = "ALPHANUMERIC"
            End If
            If Request.QueryString("sequence") = 1 Then
                objCouponPattern.Order = "RANDOM"
            Else
                objCouponPattern.Order = "SEQUENTIAL"
            End If
            If GuidStr <> "" Then
                Dim result As AMSResult(Of Boolean) = objPatternService.setCouponPattern(objCouponPattern, GuidStr, LanguageID)
                If result.ResultType <> AMSResultType.Success Then
                    If (result.MessageString.Contains("COUPON_")
                       ) Then
                        infoMessage = Copient.PhraseLib.Lookup(result.MessageString, LanguageID)
                    Else
                        infoMessage=result.MessageString
                    End If
                    
                Else
                    HistoryStr = Copient.PhraseLib.Lookup("history.edit-connectorsetting", LanguageID) & " : " & Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & Copient.PhraseLib.Lookup("term.pattern", LanguageID) & " " & Copient.PhraseLib.Lookup("term.settings", LanguageID)
                    MyCommon.Activity_Log(42, ConnectorID, AdminUserID, HistoryStr)
                End If
                End If
        End If
    
   ElseIf Request.QueryString("saveconfigs") <> "" Then
        objCouponSettings = New CouponConfig()
        objcSett = New CouponSettings()
        objCouponSettings.configs = New List(Of CouponSettings)()
        objcSett.Type = "MAX_COUPONS_ALLOWED"
        If Request.QueryString("maxCoupons") IsNot Nothing Then
            objcSett.Value = Request.QueryString("maxCoupons")
        Else
           objcSett.Value = ""
        End If
        objCouponSettings.configs.Add(objcSett)
        objcSett = New CouponSettings()
        objcSett.Type = "DEFAULT_NO_OF_COUPONS"
        If Request.QueryString("defaultCoupons") IsNot Nothing Then
            objcSett.Value = Request.QueryString("defaultCoupons")
        Else
            objcSett.Value = ""
        End If
        objCouponSettings.configs.Add(objcSett)
        objcSett = New CouponSettings()
        objcSett.Type = "THRESHOLD"
        If Request.QueryString("threshold") IsNot Nothing Then
            objcSett.Value = Request.QueryString("threshold")
        Else
            objcSett.Value = ""
        End If
        objCouponSettings.configs.Add(objcSett)
        
        If GuidStr <> "" Then
            Dim result As AMSResult(Of Boolean) = objPatternService.setCouponSettings(objCouponSettings, GuidStr, LanguageID)
            If result.ResultType <> AMSResultType.Success Then
                infoMessage = result.MessageString
            Else
                HistoryStr = Copient.PhraseLib.Lookup("history.edit-connectorsetting", LanguageID) & " : " & Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & Copient.PhraseLib.Lookup("term.coupon", LanguageID) & " " & Copient.PhraseLib.Lookup("term.settings", LanguageID)
                MyCommon.Activity_Log(42, ConnectorID, AdminUserID, HistoryStr)
                End If
        End If
    ElseIf (Request.QueryString("SaveFilter") <> "") Then
        Dim tmpAttr As New List(Of [String])
        Dim MyArray() As String
      
        If Not (Request.QueryString("defaultFilters") = "") Then
            MyArray = Request.QueryString("defaultFilters").Split(",")
            tmpAttr.AddRange(MyArray.ToList())
        End If
        If Not (Request.QueryString("selFilterXML") = "") Then
            MyArray = Request.QueryString("selFilterXML").Split(",")
            tmpAttr.AddRange(MyArray.ToList())
        End If
        If (tmpAttr.Count > 0) Then
            tmpAttr = tmpAttr.Where(Function(s) Not String.IsNullOrWhiteSpace(s)).Distinct().ToList()
            Dim NonfilteredXMlLst As New Dictionary(Of [String], [String])()
            For Each s As String In tmpAttr
                MyArray = s.Split("/")
                If (NonfilteredXMlLst.ContainsKey(MyArray(0))) Then
                    Dim val As String = NonfilteredXMlLst(MyArray(0))
                    val = IIf(String.IsNullOrEmpty(val), MyArray(1), val & "," & MyArray(1))
                    NonfilteredXMlLst(MyArray(0)) = val
                Else
                    NonfilteredXMlLst.Add(MyArray(0), MyArray(1))
                End If
            Next
            If (NonfilteredXMlLst.Count > 0) Then
                Dim FilteredString As New StringBuilder()
                For Each pair As KeyValuePair(Of [String], [String]) In NonfilteredXMlLst
                    FilteredString.Append(pair.Key).Append("-").Append(pair.Value).Append("|")
                Next
                If (FilteredString.Length > 0) Then
                    MyCommon.QueryStr = "pa_Update_FilterOutputColumns"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ReferenceId", SqlDbType.BigInt).Value = MyCommon.Extract_Val(Request.QueryString("ddlWebMethods"))
                    MyCommon.LRTsp.Parameters.Add("@FilteredAttributes", SqlDbType.VarChar).Value = FilteredString.ToString().Trim("|")
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If
            End If
        End If
  End If

        If ConnectorID = 70 AndAlso GuidStr <> "" Then
            objcPattern = objPatternService.getCouponPattern(GuidStr, LanguageID)
            If objcPattern.ResultType <> AMSResultType.Success Then
                infoMessage = objcPattern.MessageString
            End If
        
        objcSettings = objPatternService.getCouponSettings(GuidStr, LanguageID)
        'If infoMessage = "" AndAlso objcSettings.ResultType <> AMSResultType.Success Then
        '    infoMessage = objcSettings.MessageString
        'End If
        
        End If
  
        ' Load the connector
        MyCommon.QueryStr = "select Name, DescriptionPhraseID, Path, UsesGUIDs, NamePhraseID, Installed, Visible from Connectors as C with (NoLock) " & _
                            "where ConnectorID=" & ConnectorID & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
        ConnectorName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
        NamePhraseID = MyCommon.NZ(dt.Rows(0).Item("NamePhraseID"), "")
            If Not IsDBNull(dt.Rows(0).Item("DescriptionPhraseID")) Then
                ConnectorDesc = Copient.PhraseLib.Lookup(dt.Rows(0).Item("DescriptionPhraseID"), LanguageID)
            Else
                ConnectorDesc = ""
            End If
            Installed = MyCommon.NZ(dt.Rows(0).Item("Installed"), False)
            Visible = MyCommon.NZ(dt.Rows(0).Item("Visible"), False)
            Path = MyCommon.NZ(dt.Rows(0).Item("Path"), "")
            UsesGUIDs = MyCommon.NZ(dt.Rows(0).Item("UsesGUIDs"), False)
        End If
  
        Send_HeadBegin("term.connector", , ConnectorID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
%>
<style type="text/css">
  td
  {
    vertical-align: top;
  }
  .editguid, .saveguid, .cancelguid
  {
    font-size: 10px;
    width: 55px;
  }
  .select, .Deselect
  {
    font-size: 15px;
    font-weight :bold;
    width: 55px;
    height: 30px;
  }
  .newguid
  {
    font-size: 10px;
  }
  .descinput
  {
    font-size: 12px;
    width: 240px;
  }
  .desctext
  {
    font-size: 12px;
    width: 160px;
  }
  * html .descinput
  {
    width: 245px;
  }
  #NewGUIDdesc
  {
    color: #aaaaaa;
    font-size: 12px;
    width: 350px;
  }
</style>
<%
  Send_Scripts()
%>
<script src="../javascript/jquery.min.js" type="text/javascript"></script>
<script type="text/javascript">

    function onChangeWebmethod(webmethodid) {
        xmlhttpPost('/logix/XMLFeeds.aspx', 'GetResponseAttributes=1&ReferenceId=' + webmethodid + '&ConnectorId=' + '<%Sendb(ConnectorID)%>');
    }

    function xmlhttpPost(strURL, qryStr) {
        var processingPage = "<div class=\"loading\"><br /><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + 'Loading ... please wait.<\/div>';
        document.getElementById("DivFilters").innerHTML = processingPage;
        if (document.getElementById("SaveFilter") != null)
            document.getElementById("SaveFilter").disabled = true;
        var xmlHttpReq = false;
        var self = this;

        if (window.XMLHttpRequest) { // Mozilla/Safari
            self.xmlHttpReq = new XMLHttpRequest();
        } else if (window.ActiveXObject) { // IE
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
        }
        strURL += "?" + qryStr;
        self.xmlHttpReq.open('POST', strURL, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.send(qryStr);
        self.xmlHttpReq.onreadystatechange = function () {
            if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
                respTxt = self.xmlHttpReq.responseText;
                DisplayDiv(respTxt);
            }
        }
    }

    function DisplayDiv(respTxt) {
        document.getElementById("DivFilters").innerHTML = respTxt;
        if (document.getElementById("SaveFilter") != null)
            document.getElementById("SaveFilter").disabled = false;
        initializeSelectControls();
    }

    function SaveFilterClick() {
        $('#selFilterXML option').prop('selected', true);
    }

    function initializeSelectControls() {

        var $select1 = $('#SelFullXML');
        var $select2 = $('#selFilterXML');


        $('option').each(function () {
            $(this).data('optgroup', $(this).parent().attr('label'));
        });

        $('#select').click(function () {
            var $el = $('#SelFullXML option:selected');

            $el.each(function () {
                groupName = $(this).data('optgroup'),
                            $parent = $(this).parent(),
                            $optgroup = $select2.find('optgroup[label="' + groupName + '"]');
                if (!$optgroup.length) $optgroup = $('<optgroup label="' + $(this).data('optgroup') + '" />').appendTo($select2);
                $(this).appendTo($optgroup);
                if (!$parent.children().length) $parent.remove();
            });
            if ($('#SelFullXML').has('option').length == 0)
                $("#select").prop('disabled', true);
            else
                $("#select").prop('disabled', false);
            return false;
        });
        $('#Deselect').click(function () {
            var $el = $('#selFilterXML option:selected');

            $el.each(function () {
                groupName = $(this).data('optgroup'),
                                $parent = $(this).parent(),
                                $optgroup = $select1.find('optgroup[label="' + groupName + '"]');
                if (!$optgroup.length) $optgroup = $('<optgroup label="' + $(this).data('optgroup') + '" />').appendTo($select1);
                $(this).appendTo($optgroup);
                if (!$parent.children().length) $parent.remove();
            });
            if ($('#SelFullXML').has('option').length == 0)
                $("#select").prop('disabled', true);
            else
                $("#select").prop('disabled', false);
            return false;
        });

    }

    $(function () {
        initializeSelectControls();
    });


    $(document).ready(function () {
        $("#spinner").spinner({ min: 1, max: 20 });
    })
    var reA = /[^a-zA-Z]/g;
    var reN = /[^0-9]/g;
    function sortAlphaNum(a, b) {
        var aA = a.replace(reA, "");
        var bA = b.replace(reA, "");
        if (aA === bA) {
            var aN = parseInt(a.replace(reN, ""), 10);
            var bN = parseInt(b.replace(reN, ""), 10);
            return aN === bN ? 0 : aN > bN ? 1 : -1;
        } else {
            return aA > bA ? 1 : -1;
        }
    }

    function openpreview() {
        var prefix = document.getElementById("prefix").value;
        var suffix = document.getElementById("suffix").value;
        var textstart = document.getElementById("textstart").value;
        var textendrange = document.getElementById("textendrange").value;
        var couponType = document.getElementById("types").value;
        var length = document.getElementById("spinner").value;
        var previewpattern;
        var digits = length - (prefix.length + suffix.length);
        var regAlphaNum = new RegExp(/^[A-Z0-9]+$/);
        var regNum = new RegExp(/^[0-9]+$/);

        if (length < 1 || length > 20) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-couponlength", LanguageID))%>');
        }
        else if (!regNum.test(length)) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-couponlength", LanguageID))%>');
        }
        else if (prefix.length != 0 && !regAlphaNum.test(prefix)) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-prefix", LanguageID))%>');
        }
        else if (suffix.length != 0 && !regAlphaNum.test(suffix)) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-suffix", LanguageID))%>');
        }
        else if (couponType == 1 && textstart.length != 0 && !regNum.test(textstart)) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-start", LanguageID))%>');
        }
        else if (couponType == 1 && textendrange.length != 0 && !regNum.test(textendrange)) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-end", LanguageID))%>');
        }
        else if (couponType == 2 && textstart.length != 0 && !regAlphaNum.test(textstart)) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-startalpha", LanguageID))%>');
        }
        else if (couponType == 2 && textendrange.length != 0 && !regAlphaNum.test(textendrange)) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.end-alphanumeric", LanguageID))%>');
        }
        else if (prefix.length + suffix.length >= length) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalid-combination", LanguageID))%>');
        }
        else if (textstart.length > digits) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.start-exceeding", LanguageID))%>');
        }
        else if (textendrange.length > digits) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.end-exceeding", LanguageID))%>');
        }
        else if (couponType == 1 && textstart > textendrange) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalidstartend", LanguageID))%>');
        }
        else if (sortAlphaNum(textstart, textendrange) == 1) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalidstartend", LanguageID))%>');
        }
        else if (textstart.replace(/\b0+/g, '').length > textendrange.replace(/\b0+/g, '').length) {
            alert('<% Sendb(Copient.PhraseLib.Lookup("error.invalidstartend", LanguageID))%>');
        }
        else {
            textstart = pad(textstart, digits, 0);
            textendrange = pad(textendrange, digits, 0);
            previewpattern = prefix + textstart + suffix + " - " + prefix + textendrange + suffix;
            var myUrl = '/logix/UE/CouponPattern-Preview.aspx?pattern=' + previewpattern;
            openpatternprevPopup(myUrl);
        }
}
function deleteGuid(guid) {
    if (confirm('<%Sendb(Copient.PhraseLib.Lookup("connector-detail.ConfirmDeleteGUID", LanguageID))%>')) {
          document.getElementById("DeleteGUID").value = guid;
          document.mainform.submit();
      } else {
          return false;
      }
  }
  function pad(n, width, z) {
      z = z || '0';
      n = n + '';
      return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
  }
  function isValidLength() {
      var length = document.getElementById("spinner").value;
      if (length < 1 || length > 20) {
          alert('<%Sendb(Copient.PhraseLib.Lookup("error.invalid-couponlength", LanguageID))%>');
        }
        setStartEndRange();
    }
    function setStartEndRange() {
        var min, max;
        var prefix = document.getElementById("prefix").value;
        var suffix = document.getElementById("suffix").value;
        var length = document.getElementById("spinner").value;
        var contentType = document.getElementById("types").value;
        var digits = length - (prefix.length + suffix.length);
        if (length >= 1 && length <= 20 && ((prefix.length + suffix.length) < length)) {
            min = pad(0, digits);
            if (contentType == 1)
                max = pad(9, digits, 9);
            else
                max = pad('Z', digits, 'Z');
            document.getElementById("textstart").value = min;
            document.getElementById("textendrange").value = max;
        }
        else {
            document.getElementById("textstart").value = "";
            document.getElementById("textendrange").value = "";
        }

    }

    function clearNewGuidDesc() {
        document.getElementById("NewGUIDdesc").value = '';
        document.getElementById("NewGUIDdesc").style.color = '#000000';
    }

    function newGuid() {
        document.getElementById('NewGUID').value = '1';
        document.mainform.submit();
    }
    function CheckTimeOutValue(strName) {
        var inputBox = document.getElementById(strName).value;
        if (inputBox == '') {
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.numericvalue", LanguageID))%>');
        return false;
    }
    var retValue = IsNumeric(inputBox, false, false);
    if (retValue == false) {
        alert('<%Sendb(Copient.PhraseLib.Lookup("term.numericvalue", LanguageID))%>');
        return false;
    }
    return true;
}
</script>
<%
  Send("<script type=""text/javascript"">")
  
  Send("function saveGuid(guid) { ")
  Send("  document.getElementById(""OldGUID"").value = guid")
  Send("  document.getElementById(""SaveGUID"").value = document.getElementById(""ddescinput-"" + guid).value; ")
  Send("  document.getElementById(""SaveGUIDDesc"").value = document.getElementById(""descinput-"" + guid).value; ")
  Send("  document.getElementById(""PK"").value = document.getElementById(""guidpk-"" + guid).value; ")
  If ConnectorID = 58 Then
    Send("  var e=document.getElementById(""descchansel-"" + guid); ")
    Send("  document.getElementById(""SaveGUIDChan"").value = e.options[e.selectedIndex].value; ")
  End If
  Send("  document.mainform.submit(); ")
  Send("} ")

  Send("function toggleDescEdit(guid) { ")
  Send("  if (document.getElementById('descedit-' + guid).style.display == 'none') { ")
  Send("    document.getElementById('desc-' + guid).style.display = 'none'; ")
  Send("    document.getElementById('descedit-' + guid).style.display = 'block'; ")
  Send("    document.getElementById('descguid-' + guid).style.display = 'none'; ")
  Send("    document.getElementById('desceditguid-' + guid).style.display = 'block'; ")
  If ConnectorID = 58 Then
    Send("    document.getElementById('descchan-' + guid).style.display = 'none'; ")
    Send("    document.getElementById('desceditchan-' + guid).style.display = 'block'; ")
  End If
  Send("  } else { ")
  Send("    document.getElementById('desc-' + guid).style.display = 'block'; ")
  Send("    document.getElementById('descedit-' + guid).style.display = 'none'; ")
  Send("    document.getElementById('descguid-' + guid).style.display = 'block'; ")
  Send("    document.getElementById('desceditguid-' + guid).style.display = 'none'; ")
  If ConnectorID = 58 Then
    Send("    document.getElementById('descchan-' + guid).style.display = 'block'; ")
    Send("    document.getElementById('desceditchan-' + guid).style.display = 'none'; ")
  End If
  Send("    document.mainform.submit(); ")
  Send("  } ")
  Send("} ")
  Send("</scr" & "ipt>")

  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 3, , ConnectorID)
  
  If (Logix.UserRoles.AccessConnectors = False) Then
    Send_Denied(1, "perm.connectors-access")
    Send_BodyEnd()
    GoTo done
  End If
%>
<form action="connector-detail.aspx" method="get" id="mainform" name="mainform">
<%
  Send("<input type=""hidden"" id=""ConnectorID"" name=""ConnectorID"" value=""" & ConnectorID & """ />")
  Send("<input type=""hidden"" id=""NewGUID"" name=""NewGUID"" value="""" />")
  Send("<input type=""hidden"" id=""DeleteGUID"" name=""DeleteGUID"" value="""" />")
  Send("<input type=""hidden"" id=""SaveGUID"" name=""SaveGUID"" value="""" />")
  Send("<input type=""hidden"" id=""SaveGUIDDesc"" name=""SaveGUIDDesc"" value="""" />")
  Send("<input type=""hidden"" id=""PK"" name=""PK"" value="""" />")
  If ConnectorID = 58 Then
    Send("<input type=""hidden"" id=""SaveGUIDChan"" name=""SaveGUIDChan"" value="""" />")
  End If
  Send("<input type=""hidden"" id=""OldGUID"" name=""OldGUID"" value="""" />")
%>
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.connector", LanguageID) & " #" & ConnectorID & ": " & Copient.phraseLib.Lookup(NamePhraseID, LanguageID))
    %>
  </h1>
  <div id="controls">
    <%
      If MyCommon.Fetch_SystemOption(75) Then
        If (Logix.UserRoles.AccessNotes) Then
          Send_NotesButton(38, ConnectorID, AdminUserID)
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
  %>
  <div id="column">
    <div class="box" id="identification">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<table summary=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """>")
        Send("  <tr>")
        Send("    <td style=""width:90px;"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</td>")
        Send("    <td>" & Copient.phraseLib.Lookup(NamePhraseID, LanguageID) & "</td>")
        Send("  </tr>")

        If Not Installed OrElse Not Visible Then
          If Not Installed Then
            Send("  <tr>")
            Send("    <td>" & Copient.PhraseLib.Lookup("term.status", LanguageID) & ":</td>")
            Send("    <td><span class=""grey"">" & Copient.PhraseLib.Lookup("term.noninstalled", LanguageID) & "</span></td>")
            Send("  </tr>")
          ElseIf Not Visible Then
            Send("  <tr>")
            Send("    <td>" & Copient.PhraseLib.Lookup("term.status", LanguageID) & ":</td>")
            Send("    <td><span class=""grey"">" & Copient.PhraseLib.Lookup("term.hidden", LanguageID) & "</span></td>")
            Send("  </tr>")
          End If
        End If
        Send("  <tr>")
        Send("    <td>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</td>")
        Send("    <td>" & ConnectorDesc & "</td>")
        Send("  </tr>")
        Send("</table>")
      %>
    </div>
    <div class="box" id="guids" <% Sendb(IIf(UsesGUIDs, "", " style=""display:none;""")) %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.guids", LanguageID))%>
        </span>
      </h2>
      <%
        If (Logix.UserRoles.EditConnectors) Then
          If (ConnectorID = 53) Then
            MyCommon.QueryStr = "SELECT ExtInterfaceID, [Name] FROM ExtCRMInterfaces ORDER BY [Name]"
            dt2 = MyCommon.LRT_Select
                
            Send("<table summary=""" & Copient.PhraseLib.Lookup("term.guids", LanguageID) & """>")
            Send("<tr>")
            Send("<td style=""width:140px;"">" & Copient.PhraseLib.Lookup("connector-detail.NewGUIDDescription", LanguageID) & ":</td>")
            Send("<td ><input type=""text"" id=""NewGUIDdesc"" name=""NewGUIDdesc"" value=""" & Copient.PhraseLib.Lookup("connector-detail.NewGUIDDescription", LanguageID) & "..."" onclick=""javascript:clearNewGuidDesc();"" /></td>")
            Send("<td ><input type=""button"" class=""generateguid"" id=""generate"" name=""generate"" value=""" & Copient.PhraseLib.Lookup("connector-detail.GenerateNewGUID", LanguageID) & """ onclick=""javascript:newGuid();"" /></td>")
            Send("</tr>")
            Send("<tr>")
            Send("<td style=""width:140px;"">" & Copient.PhraseLib.Lookup("connector-detail.ExtInterfaceID", LanguageID) & ":</td>")
            If dt2.Rows.Count > 0 Then
              Send("<td><select id=""drpExtInterfaceId"" name=""drpExtInterfaceId"">")
              For Each row2 In dt2.Rows
                If MyCommon.NZ(row2.Item("Name"), "") = "None" Then
                  Send(" <option value=""" & MyCommon.NZ(row2.Item("ExtInterfaceID"), "") & """ selected=""selected"">" & MyCommon.NZ(row2.Item("ExtInterfaceID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "</option>")
                Else
                  Send(" <option value=""" & MyCommon.NZ(row2.Item("ExtInterfaceID"), "") & """>" & MyCommon.NZ(row2.Item("ExtInterfaceID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "</option>")
                End If
              Next
              Send("</select></td>")
            End If
            Send("</tr>")
            Send("</table>")
          Else
            Send("<input type=""text"" id=""NewGUIDdesc"" name=""NewGUIDdesc"" value=""" & Copient.PhraseLib.Lookup("connector-detail.NewGUIDDescription", LanguageID) & "..."" onclick=""javascript:clearNewGuidDesc();"" />")
            If (ConnectorID = 58) Then
              Send_Channel_Selector(MyCommon, "NewGUIDchan", 0)
            End If
            Send("<input type=""button"" class=""generateguid"" id=""generate"" name=""generate"" value=""" & Copient.PhraseLib.Lookup("connector-detail.GenerateNewGUID", LanguageID) & """ onclick=""javascript:newGuid();"" />")
            Send("<br class=""half"" />")
          End If
        End If
        Send("<br />")
        If (ConnectorID = 53) Then
          MyCommon.QueryStr = "select PKID, GUID, ExtInterfaceID, Description, CreatedDate from ConnectorGUIDs with (NoLock) " & _
                       "where ConnectorID=" & ConnectorID & ";"
        ElseIf (ConnectorID = 58) Then
          MyCommon.QueryStr = "select CG.PKID, CG.GUID, CG.Description, CG.CreatedDate, isnull(C.ChannelID, 0) as ChannelID, isnull(C.PhraseTerm, '') as PhraseTerm, isnull(C.Enabled, 0) as Enabled " & _
                              "from ConnectorGUIDs as CG with (NoLock) Left Join Channels as C on CG.ChannelID=C.ChannelID and C.Enabled=1 " & _
                              "where ConnectorID=" & ConnectorID & ";"
        Else
          MyCommon.QueryStr = "select PKID, GUID, Description, CreatedDate from ConnectorGUIDs with (NoLock) " & _
                              "where ConnectorID=" & ConnectorID & ";"
        End If
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
          Send("<table summary=""" & Copient.PhraseLib.Lookup("term.guids", LanguageID) & """>")
          Send("  <thead>")
          Send("    <tr>")
          Send("      <th style=""width:40px;text-align:center;"">")
          Send("        " & Left(Copient.PhraseLib.Lookup("term.delete", LanguageID), 3))
          Send("      </th>")
          Send("      <th style=""width:265px;"">")
          Send("        " & Copient.PhraseLib.Lookup("term.guid", LanguageID))
          Send("      </th>")
          If (ConnectorID = 53) Then
            Send("      <th>")
            Send("        " & Copient.PhraseLib.Lookup("connector-detail.ExtInterfaceID", LanguageID))
            Send("      </th>")
          ElseIf (ConnectorID = 58) Then
            Send("      <th>")
            Send("        " & Copient.PhraseLib.Lookup("term.channel", LanguageID))
            Send("      </th>")
          End If
          Send("      <th>")
          Send("        " & Copient.PhraseLib.Lookup("term.description", LanguageID))
          Send("      </th>")
          Send("    </tr>")
          Send("  </thead>")
          Send("  <tbody>")
          For Each row In dt.Rows
            Send("    <tr>")
            Send("      <td style=""text-align:center;"">")
            Send("        <input type=""button"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ name=""ex"" id=""ex-" & MyCommon.NZ(row.Item("GUID"), "") & """ class=""ex""" & IIf(Logix.UserRoles.EditConnectors, "", " disabled=""disabled""") & " onclick=""javascript:deleteGuid('" & MyCommon.NZ(row.Item("GUID"), "") & "')"" />")
            Send("      </td>")
                             
            Send("      <td>")
            Send("        <div id=""descguid-" & MyCommon.NZ(row.Item("GUID"), "") & """>")
            Send("        " & MyCommon.NZ(row.Item("GUID"), ""))
            Send("        </div>")
                   
            Send("        <div id=""desceditguid-" & MyCommon.NZ(row.Item("GUID"), "") & """ style=""display:none;"">")
            Send("          <input class=""descinput"" id=""ddescinput-" & MyCommon.NZ(row.Item("GUID"), "") & """ maxlength=""36"" value=""" & MyCommon.NZ(row.Item("GUID"), "") & """ />")
            Send("          <input type=""hidden"" id=""guidpk-" & MyCommon.NZ(row.Item("GUID"), "") & """ name=""guidpk-" & MyCommon.NZ(row.Item("GUID"), "") & """ value=""" & row.Item("PKID") & """ />")
            Send("        </div>")
            Send("      </td>")
            If (ConnectorID = 53) Then
              Send("      <td>")
              Send("        <div id=""descguid-" & MyCommon.NZ(row.Item("GUID"), "") & """>")
              Send("        " & MyCommon.NZ(row.Item("ExtInterfaceID"), ""))
              Send("        </div>")
              Send("      </td>")
            ElseIf (ConnectorID = 58) Then  'the Channel web service
              ChannelName = Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID, "")
              If row.Item("ChannelID") = 0 Then ChannelName = "<font color=""red"">GUID can't be used - select channel</font>"
              Send("      <td>")
              'send the selected channel associated with the GUID
              Send("        <div id=""descchan-" & MyCommon.NZ(row.Item("GUID"), "") & """>")
              Send("        " & ChannelName)
              Send("        </div>")
              Send("        <div id=""desceditchan-" & MyCommon.NZ(row.Item("GUID"), "") & """ style=""display:none;"">")
              'send the list of available channels that could be associated with the GUID
              Send_Channel_Selector(MyCommon, "descchansel-" & row.Item("GUID"), row.Item("ChannelID"))
              Send("        </div>")
              Send("      </td>")
            End If
            Send("      <td>")
            Send("        <div id=""desc-" & MyCommon.NZ(row.Item("GUID"), "") & """>")
            If Logix.UserRoles.EditConnectors Then
              Send("        <input type=""button"" class=""editguid"" id=""ed-" & MyCommon.NZ(row.Item("GUID"), "") & """ name=""ed"" onclick=""javascript:toggleDescEdit('" & MyCommon.NZ(row.Item("GUID"), "") & "');"" title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ style=""float:right;"" />")
              Send("        " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Description"), ""), 25))
              Send("        </div>")
            End If
            Send("        <div id=""descedit-" & MyCommon.NZ(row.Item("GUID"), "") & """ style=""display:none;"">")
            Send("          <input class=""desctext"" id=""descinput-" & MyCommon.NZ(row.Item("GUID"), "") & """ maxlength=""400"" value=""" & MyCommon.NZ(row.Item("Description"), "") & """ />")
            Send("          <input type=""button"" class=""saveguid"" id=""save-" & MyCommon.NZ(row.Item("GUID"), "") & """ name=""save"" title=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & IIf(Logix.UserRoles.EditConnectors, "", " disabled=""disabled""") & """ onclick=""javascript:saveGuid('" & MyCommon.NZ(row.Item("GUID"), "") & "');"" />")
            Send("          <input type=""button"" class=""cancelguid"" id=""cancel-" & MyCommon.NZ(row.Item("GUID"), "") & """ name=""cancel"" title=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & IIf(Logix.UserRoles.EditConnectors, "", " disabled=""disabled""") & """ onclick=""javascript:toggleDescEdit('" & MyCommon.NZ(row.Item("GUID"), "") & "');"" />")
            Send("        </div>")
            Send("      </td>")
            Send("    </tr>")
          Next
          Send("  </tbody>")
          Send("</table>")
        End If
      %>
      <hr class="hidden" />
    </div>
    <%
        Dim bEnableFilterResponse As Boolean = IIf(MyCommon.Fetch_SystemOption(317) = "1", True, False)
        If (ConnectorID = 3 AndAlso bEnableFilterResponse) Then
            MyCommon.QueryStr = "SELECT ReferenceId,ReferenceName,PhraseID FROM FilterOutputColumns WHERE ConnectorID=3"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then

                Send("   <div class=""box"" id=""WebmethodResponse"">")
                Send("   <h2>")
                Send("   <span>")
                Send(Copient.PhraseLib.Lookup("term.filterresponse", LanguageID))
                Send("   </span>")
                Send("   </h2>")
                Send("<table summary=""" & Copient.PhraseLib.Lookup("term.guids", LanguageID) & """>")
                Send("    <tr>")
                Send("   <td valign=""top"">")
                'Send("<br class=""half"">")
                Send("   <select id=""ddlWebMethods"" name=""ddlWebMethods"" onChange=""javascript:onChangeWebmethod(this.value);"" >")
                Dim i As Integer = 1
                Dim webMethodID As Integer = 0
                For Each row2 In dt.Rows
                    If i = 1 Then
                        webMethodID = MyCommon.NZ(row2.Item("ReferenceId"), 0)
                    End If
                    Send("<option value=""" & MyCommon.NZ(row2.Item("ReferenceId"), 0) & """" & IIf(i = 1, "selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), ""),LanguageID) & "</option>")
                    i = i + 1
                Next
                Send("</select>")
                Send("</td>")
                Send("      <td>")
                Send("        <div id=""saveconfig"">")
                If Logix.UserRoles.EditConnectors Then
                    Send("        <input type=""submit"" class=""editguid"" id=""SaveFilter"" name=""SaveFilter"" onclick=""javascript:SaveFilterClick()"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """  style=""float:right;"" />")
                    Send("        </div>")
                End If
                Send("</td>")
                Send("</tr>")
                Send("</table>")
               
                'Send("<br class=""half"">")
                Send("<div id=""DivFilters"">")
                 LoadFilterData(webMethodID, MyCommon, Logix.UserRoles.EditConnectors, ConnectorID)
                Send("</div>")
                Send("</div>")
            End If
        End If
            
      ' load up all the options for this connector; if none, then don't show the option box
      MyCommon.QueryStr = "select OptionID, OptionName, OptionValue, PhraseID " & _
                          "from InterfaceOptions with (NoLock) " & _
                          "where ConnectorID = " & ConnectorID & " and Visible=1 " & _
                          "order by OptionName;"
      dt = MyCommon.LRT_Select
    %>
    <div class="box" id="options" <% Sendb(IIf(dt.Rows.Count > 0, "", " style=""display:none;""")) %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.options", LanguageID))%>
        </span>
      </h2>
      <%
        If dt.Rows.Count > 0 Then
          For Each row In dt.Rows
            OptionID = MyCommon.NZ(row.Item("OptionID"), 0)

            Select Case OptionID
              Case 36, 37, 38 ' default chargeback departments
                ' defaults for chargeback departments are only editable when banners are not enable
                If MyCommon.Fetch_SystemOption(66) <> "1" Then
                  Send("<label for=""option" & OptionID & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("OptionName"), "")) & ":</label>")

                  MyCommon.QueryStr = "select distinct ExternalID, case when IsNull(ExternalID, '') = '' then Name " & _
                                      " else ExternalID + ' - ' + Name end as OptionText " & _
                                      "from ChargeBackDepts with (NoLock) where Deleted=0 "

                  If OptionID = 38 Then ' basket level
                    MyCommon.QueryStr &= " and ChargeBackDeptID<>0 "
                  ElseIf OptionID = 37 Then ' dept level
                    MyCommon.QueryStr &= " and ChargeBackDeptID<>0 and ChargeBackDeptID<>14 "
                  ElseIf OptionID = 36 Then 'item level
                    MyCommon.QueryStr &= " and ChargeBackDeptID<>10 "
                  End If

                  MyCommon.QueryStr &= " order by OptionText;"
                  
                  dt2 = MyCommon.LRT_Select
                  If dt2.Rows.Count = 0 Then
                    Send("<input type=""text"" id=""option" & OptionID & """ name=""option" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
                  Else
                    Send("<select id=""option" & OptionID & """ name=""option" & OptionID & """>")
                    Dim SelectedSet As Boolean = False
                                  For Each row2 In dt2.Rows
                                      If MyCommon.NZ(row2.Item("ExternalID"), "") = MyCommon.NZ(row.Item("OptionValue"), "") Then
                                          Send("  <option title=""" & MyCommon.NZ(row2.Item("OptionText"), "") & """ value=""" & MyCommon.NZ(row2.Item("ExternalID"), "") & """ selected=""selected"">" & TruncateWordAppendEllipsis(MyCommon.NZ(row2.Item("OptionText"), ""), 40) & "</option>")
                                          SelectedSet = True
                                      Else
                                          Send("  <option title=""" & MyCommon.NZ(row2.Item("OptionText"), "") & """ value=""" & MyCommon.NZ(row2.Item("ExternalID"), "") & """>" & TruncateWordAppendEllipsis(MyCommon.NZ(row2.Item("OptionText"), ""), 40) & "</option>")
                                      End If
                                  Next
                    If (Not SelectedSet) Then Send("  <option value="""" selected=""selected"">" & "</option>")
                    Send("</select>")
                  End If
                  Send("<br />")
                End If
              Case Else
                Send("<label for=""option" & OptionID & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row.Item("OptionName"), "")) & ":</label>")

                MyCommon.QueryStr = "select OptionValue, Description, PhraseID from InterfaceOptionValues with (NoLock) " & _
                                    "where OptionID=" & MyCommon.NZ(row.Item("OptionID"), 0)
                dt2 = MyCommon.LRT_Select
                If dt2.Rows.Count = 0 Then
                  Send("<input type=""text"" id=""option" & OptionID & """ name=""option" & OptionID & """ value=""" & MyCommon.NZ(row.Item("OptionValue"), "") & """ />")
                  Send("<br />")
                Else
                  Send("<select id=""option" & OptionID & """ name=""option" & OptionID & """" & IIf(Logix.UserRoles.EditConnectors, "", " disabled=""disabled""") & ">")
                  For Each row2 In dt2.Rows
                    If MyCommon.NZ(row2.Item("OptionValue"), "") = MyCommon.NZ(row.Item("OptionValue"), "") Then
                      Send("  <option value=""" & MyCommon.NZ(row2.Item("OptionValue"), "") & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Description"), "")) & "</option>")
                    Else
                      Send("  <option value=""" & MyCommon.NZ(row2.Item("OptionValue"), "") & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseID"), 0), LanguageID, MyCommon.NZ(row2.Item("Description"), "")) & "</option>")
                    End If
                  Next
                  Send("</select>")
                  Send("<br />")
                End If
            End Select
          Next
          Send("<br />")
          If (Logix.UserRoles.EditConnectors) Then
            Send("<input type=""submit"" name=""SaveOptions""" & " " & IIf(OptionID = 100, "OnClick =""  return CheckTimeOutValue('option" & OptionID & "');""", "") & " value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
          End If
        End If
      %>
      <hr class="hidden" />
    </div>
      
       <div class="box" id="couponOptions" <% Sendb(IIf(ConnectorID = 70 AndAlso GuidStr <> "", "", " style=""display:none;""")) %>>
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.pattern", LanguageID) & " " & Copient.PhraseLib.Lookup("term.settings", LanguageID))%>
        </span>
      </h2><br />
          <table id="patternSettings" <% Sendb(IIf(ConnectorID = 70, "", " style=""display:none;""")) %>>
               <tr><td style="width:150px;"><label id="couponLength"><% Sendb(Copient.PhraseLib.Lookup("coupon.length", LanguageID) & ":") %></label> </td> 
                   <td>
           <%
               If ConnectorID = 70 Then
                   If objcPattern.ResultType = AMSResultType.Success AndAlso objcPattern.Result IsNot Nothing Then
                       Send(" <input style=""width: 30px;"" id=""spinner"" name=""spinner"" value=""" & objcPattern.Result.Length & """ onblur=""isValidLength()""> </td></tr>")
                   Else
                                        Send(" <input style=""width: 30px;"" id=""spinner"" name=""spinner"" value=""18"" onblur=""isValidLength()""> </td></tr>")
                   End If%>
                       <tr><td> <label id="couponPrefix"><% Sendb(Copient.PhraseLib.Lookup("term.coupon", LanguageID) & " " & Copient.PhraseLib.Lookup("term.prefix", LanguageID) & ":") %></label> </td>
                   <td><%  
                           If objcPattern.ResultType = AMSResultType.Success AndAlso objcPattern.Result IsNot Nothing Then
                               Send("<input name=""prefix"" id=""prefix"" type=""text"" value=""" & objcPattern.Result.Prefix.ToString() & """ onblur=""setStartEndRange()""/>")
                           Else
                               Send("<input name=""prefix"" id=""prefix"" type=""text"" onblur=""setStartEndRange()""/>")
                           End If%>
                       </td>
               </tr>
             <tr><td> <label id="couponSuffix"><% Sendb(Copient.PhraseLib.Lookup("term.coupon", LanguageID) & " " & Copient.PhraseLib.Lookup("term.suffix", LanguageID) & ":") %></label> </td>
                   <td>
                       <% If objcPattern.ResultType = AMSResultType.Success AndAlso objcPattern.Result IsNot Nothing Then
                               Send("<input type=""text"" name=""suffix"" id=""suffix"" value=""" & objcPattern.Result.Suffix.ToString() & """ onblur=""setStartEndRange()""/></td>")
                           Else
                               Send("<input type=""text"" id=""suffix"" name=""suffix"" onblur=""setStartEndRange()""/></td>")
                           End If%>
                        </tr>
             <tr><td> <label id="couponType"><% Sendb(Copient.PhraseLib.Lookup("coupon.type", LanguageID) & ":")%></label> </td>
                   <td><select id="types" name="types" onblur="setStartEndRange()" <% Sendb(IIf(Logix.UserRoles.EditConnectors, "", " disabled=""disabled""")) %>>
                       <%If objcPattern.ResultType = AMSResultType.Success AndAlso objcPattern.Result IsNot Nothing Then
                               Send(" <option value=""1"" " & IIf(objcPattern.Result.ContentType = "NUMERIC", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.numeric", LanguageID) & " </option>")
                               Send(" <option value=""2"" " & IIf(objcPattern.Result.ContentType = "ALPHANUMERIC", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.alphanumeric", LanguageID) & " </option>")
                           Else
                               Send(" <option value=""1"" selected=""selected"">" & Copient.PhraseLib.Lookup("term.numeric", LanguageID) & " </option>")
                               Send(" <option value=""2"">" & Copient.PhraseLib.Lookup("term.alphanumeric", LanguageID) & " </option>")
                           End If%>
                       </select></td>
               </tr>
             <tr><td> <label id="startrange"><% Sendb(Copient.PhraseLib.Lookup("range-start.coupon", LanguageID) & ":") %></label> </td>
                   <td>
                       <% If objcPattern.ResultType = AMSResultType.Success AndAlso objcPattern.Result IsNot Nothing Then
                               Send("<input type=""text"" name=""textstart"" id=""textstart"" value=""" & objcPattern.Result.Start.ToString() & """ /></td></tr>")
                           Else
                                    Send("<input type=""text"" id=""textstart"" name=""textstart"" value=""000000000000000000""/></td></tr>")
                           End If%>
                        <tr><td> <label id="endrange"><% Sendb(Copient.PhraseLib.Lookup("range-end.coupon", LanguageID) & ":") %></label> </td>
                   <td>
                       <%If objcPattern.ResultType = AMSResultType.Success AndAlso objcPattern.Result IsNot Nothing Then
                               Send("<input type=""text"" name=""textendrange"" id=""textendrange"" value=""" & objcPattern.Result.End.ToString() & """ /></td></tr>")
                           Else
                                        Send("<input type=""text"" id=""textendrange"" name=""textendrange"" value=""999999999999999999""/></td></tr>")
                           End If%>
                        <tr><td> <label id="couponSeq"><% Sendb(Copient.PhraseLib.Lookup("coupon.sequence", LanguageID)) %></label> </td>
                   <td><select id="sequence" name="sequence" <% Sendb(IIf(Logix.UserRoles.EditConnectors, "", " disabled=""disabled""")) %>>
                       <%If objcPattern.ResultType = AMSResultType.Success AndAlso objcPattern.Result IsNot Nothing Then
                               Send(" <option value=""1"" " & IIf(objcPattern.Result.Order = "RANDOM", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.random", LanguageID) & " </option>")
                               Send(" <option value=""2"" " & IIf(objcPattern.Result.Order = "SEQUENTIAL", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.sequential", LanguageID) & " </option>")
                           Else
                               Send(" <option value=""1"" selected=""selected"">" & Copient.PhraseLib.Lookup("term.random", LanguageID) & " </option>")
                               Send(" <option value=""2"">" & Copient.PhraseLib.Lookup("term.sequential", LanguageID) & " </option>")
                           End If%>
                        </select></td>
               </tr>
                       <%
               End If
               %>
           </table><br />
           &nbsp;<input type="button" id="preview" name="preview" onclick="openpreview()" value=" <% Sendb(Copient.PhraseLib.Lookup("term.preview", LanguageID))%> "  />
     &nbsp;<input type="submit" id="save" name="savepattern" value=" <% Sendb(Copient.PhraseLib.Lookup("term.save", LanguageID))%> " <% Sendb(IIf(Logix.UserRoles.EditConnectors, "", " style=""display:none;""")) %> />
           </div>
      <div class="box" id="coupon_Options" <% Sendb(IIf(ConnectorID = 70 AndAlso GuidStr <> "", "", " style=""display:none;""")) %>>
          <h2> <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.options", LanguageID))%>
        </span></h2><br />
           <table id="coupon_options" <% Sendb(IIf(ConnectorID = 70, "", " style=""display:none;""")) %>>
               <tr>
                   <td style="width:310px;"><label id="maxCoupons"><%Sendb(Copient.PhraseLib.Lookup("term.maxcoupons", LanguageID) & " :")%></label> </td><td>
                       <%  If objcSettings.ResultType = AMSResultType.Success AndAlso objcSettings.Result IsNot Nothing Then
                               Send("<input type=""text"" id=""maxCoupons"" name=""maxCoupons"" value=""" & objcSettings.Result.configs.Item(0).Value & """/></td>")
                           Else
                               Send("<input type=""text"" id=""maxCoupons"" name=""maxCoupons"" value=""100""/></td>")
                           End If%>
                  
                   </tr><tr> <td><label id="defaultCoupons"><%Sendb(Copient.PhraseLib.Lookup("term.defaultcoupons", LanguageID) & " :")%></label> </td>
                       <td><%  If objcSettings.ResultType = AMSResultType.Success AndAlso objcSettings.Result IsNot Nothing Then
                                   Send("<input type=""text"" id=""defaultCoupons"" name=""defaultCoupons"" value=""" & objcSettings.Result.configs.Item(1).Value & """/></td>")
                               Else
                                   Send("<input type=""text"" id=""defaultCoupons"" name=""defaultCoupons"" value=""50""/></td>")
                               End If%>
                        </tr>
                
               <tr> <td>
                       <label id="threshold"><% Sendb(Copient.PhraseLib.Lookup("term.couponthreshold", LanguageID) & " :") %></label> </td> 
                   <td><%If ConnectorID = 70 Then
                               
                               If objcSettings.ResultType = AMSResultType.Success AndAlso objcSettings.Result IsNot Nothing Then
                                   Send("<input type=""text"" id=""threshold"" name=""threshold"" value=""" & objcSettings.Result.configs.Item(2).Value & """/> %</td>")
                               Else
                                   Send("<input type=""text"" id=""threshold"" name=""threshold"" value=""90""/> %</td>")
                               End If
                               %>
                          </td>
                  
                        </tr>
                             <%   End If%>
               </table>
                           
          <br />
     &nbsp;<input type="submit" id="saveconfigs" name="saveconfigs" value=" <% Sendb(Copient.PhraseLib.Lookup("term.save", LanguageID))%> " <% Sendb(IIf(Logix.UserRoles.EditConnectors, "", " style=""display:none;""")) %> />
          </div>
  </div>
</div>
</form>
<script runat="server">
  Function GUID() As String
    GUID = System.Guid.NewGuid().ToString()
    Return GUID
  End Function

  ' -------------------------------------------------------------------------------------------------------  
  
  Sub Send_Channel_Selector(Common As Copient.CommonInc, FormFieldID As String, SelectedChannelID As Integer)
    Dim dt As DataTable
    Dim row As DataRow
    Common.QueryStr = "select ChannelID, PhraseTerm from Channels where enabled=1;"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      Send("        <select class=""descinput"" style=""width: 110px;"" id=""" & FormFieldID & """ name=""" & FormFieldID & """>")
      For Each row In dt.Rows
        Send("        <option value=""" & row.Item("ChannelID") & """" & IIf(row.Item("ChannelID") = SelectedChannelID, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup(row.Item("PhraseTerm"), LanguageID) & "</option>")
      Next
      dt = Nothing
      Send("        </select>")
    End If
  End Sub
  
  ' -------------------------------------------------------------------------------------------------------
  
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (ConnectorID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(38, ConnectorID, AdminUserID)
    End If
    End If
  Send_BodyEnd("mainform", "name")
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
