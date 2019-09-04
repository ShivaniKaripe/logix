<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:6.0.1.114508.Official Build (SUSDAY10082) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Linq" %>
<%@ Import Namespace="System.Xml.Linq" %>
<%
  ' *****************************************************************************
    ' * FILENAME: reports-enhanced-configuration.aspx 
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
  ' * Version : 6.0.1.114508 
  ' *
    ' *****************************************************************************
    
    Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
    Dim CopientFileVersion As String = "6.0.1.114508"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""
    Dim Handheld As Boolean = False
    
    Dim AdminUserID As Long
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim dt As DataTable
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "reports-enhanced-configuration.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
      Try
          Send_HeadBegin("term.reportconfiguration")
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
        
        If (Request.QueryString("SaveFilter") <> "") Then
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
                        MyCommon.LRTsp.Parameters.Add("@ReferenceId", SqlDbType.BigInt).Value = MyCommon.Extract_Val(Request.QueryString("ddlReportDetails"))
                        MyCommon.LRTsp.Parameters.Add("@FilteredAttributes", SqlDbType.VarChar).Value = FilteredString.ToString().Trim("|")
                        MyCommon.LRTsp.ExecuteNonQuery()
                        MyCommon.Close_LRTsp()
                    End If
                End If
            End If
            End If 
%>
  <style type="text/css">
 .select, .Deselect
  {
    font-size: 15px;
    font-weight :bold;
    width: 55px;
    height: 30px;
  }
</style>
<%
  Send_Scripts()
%>
<script src="../javascript/jquery-1.10.2.min.js" type="text/javascript"></script>
<script type="text/javascript">

    function onChangeWebmethod(webmethodid) {
        xmlhttpPost('/logix/XMLFeeds.aspx', 'GetReportResponseAttributes=1&ReferenceId=' + webmethodid);
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
</script>
    <form action="#" id="WebmethodResponse" name="WebmethodResponse"  method="get">
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.reportconfiguration", LanguageID))%>
  </h1>
  <div id="controls">

        <%
            Send("<input type=""submit"" class=""editguid"" id=""SaveFilter"" name=""SaveFilter"" onclick=""javascript:SaveFilterClick()""   value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """  style=""float:right;"" />")
        %>
      <br />
 
    </div>
    </div>
    <div id="main">
    <%
        MyCommon.QueryStr = "SELECT ReferenceId,ReferenceName FROM FilterOutputColumns with (NoLock) WHERE ReferenceId=1 and ConnectorID = 0"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then

            Send("   <div class=""box"" id=""WebmethodResponse"">")
            Send("   <h2>")
            Send("   <span>")
            Send(Copient.PhraseLib.Lookup("term.reportconfiguration", LanguageID))
            Send("   </span>")
            Send("   </h2>")
            Send("   <td valign=""top"">")
            Send("<br class=""half"">")
            Send("   <select id=""ddlReportDetails"" name=""ddlReportDetails"" >")
            Dim i As Integer = 1
            Dim webMethodID As Integer = 0
            For Each row2 In dt.Rows
                If i = 1 Then
                    webMethodID = MyCommon.NZ(row2.Item("ReferenceId"), 0)
                End If
                Send("<option value=""" & MyCommon.NZ(row2.Item("ReferenceId"), 0) & """" & IIf(i = 1, "selected=""selected""", "") & ">" & MyCommon.NZ(row2.Item("ReferenceName"), "") & "</option>")
                i = i + 1
            Next
            Send("</select>")
            Send("</td>")
            Send("<br class=""half"">")
            Send("<br class=""half"">")
            Send("<div id=""DivFilters"">")
            LoadFilterData(webMethodID, MyCommon, True, 0)
            Send("</div>")
            Send("</div>")
        End If
    %>
  </div>
     </form>
<%
done:
      Finally
          MyCommon.Close_LogixRT()
          Logix = Nothing
          MyCommon = Nothing
      End Try
Send_BodyEnd("searchform", "searchterms")

%>

  

