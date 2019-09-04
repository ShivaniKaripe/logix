﻿<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %><%@ Import
    Namespace="CMS.AMS.Contract" %><%@ Import Namespace="System.IO" %><%@ Import Namespace="System.Xml" %>
<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Xml.Schema" %>
<%@ Import Namespace="System.Xml.XPath" %>
<%@ Import Namespace="CMS.AMS" %>
<%
    ' *****************************************************************************
    ' * FILENAME: offer-list.aspx 
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
    ' * NOTES   : $Id: offer-list.aspx 123981 2018-05-25 06:05:15Z kn250067 $
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
    Dim rst As DataTable
    Dim dst As System.Data.DataTable
    Dim dstOffersUniqueUserHasRights As System.Data.DataTable
    Dim OffersUniqueUserHasRightsRows As DataRow
    Dim Shaded As String = "shaded"
    Dim idNumber As Integer
    Dim idSearch As String = ""
    'added
    Dim buyeridSearch As String = ""
    Dim idSearchText As String = ""
    Dim PageNum As Integer = 0
    Dim MorePages As Boolean
    Dim linesPerPage As Integer = 20
    Dim sizeOfData As Integer
    Dim i As Integer = 0
    Dim File As HttpPostedFile
    Dim sSummaryPage As String
    Dim SearchMatchesROID As Boolean = False
    Dim RoidExtension As String = ""
    Dim WhereBuf As New StringBuilder()
    Dim AdvSearchSQL As String = ""
    Dim CriteriaMsg As String = ""
    Dim CriteriaTokens As String = ""
    Dim CriteriaError As Boolean = False
    Dim BannersEnabled As Boolean = False
    Dim IE6ScrollFix As String = ""
    Dim infoMessage As String = ""
    Dim CustomerInquiry As Boolean = False
    Dim TotalUsers As Integer = 0
    Dim SourceName As String = ""
    Dim restrictLinks As Boolean = False
    Dim Handheld As Boolean = False
    Dim FavoriteOption6 As Boolean = False
    Dim OffersUniqueUserHasRightsToList As New SortedList()
    Dim OffersInPage As String = ""
    Dim ListCouter As Long
    Dim CountQuery As String = ""
    Dim SelectQuery As String = ""
    Dim SelectQuery1 As String = ""
    Dim SelectQueryOrderBy As String = ""
    Dim SelectSortDirection As String = ""
    Dim StartPoint As Long
    Dim EndPoint As Long
    Dim EndPage As Long
    Dim bWorkflowActive As Boolean = False
    Dim bProductionSystem As Boolean = True
    Dim bTestSystem As Boolean = False
    Dim bArchiveSystem As Boolean = False
    Dim bCmInstalled As Boolean = False
    Dim bUeInstalled As Boolean = False
    Dim bAdvancedExternalSearch As Boolean = False
    Dim bSearchInternalAndExternal As Boolean = False
    Dim bStoreUser As Boolean = False
    Dim sValidLocIDs As String = ""
    Dim sValidSU As String = ""
    Dim wherestr As String = ""
    Dim sJoin As String = ""
    Dim iLen As Integer = 0
    Const FOLDER_NOT_IN_USE As String = "~FNIU~"
    Const SUCCESS As String = "Success!"
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
    Dim FilterEngine As Integer = 0 ' if System option #249 is enabled then CM is set as default option in the filterengine dropdown.
    Dim SqlInjectionFormList() As String = {"xidOption", "xid", "idSearch", "idOption", "buyeridSearch", "buyeridOption", "productionIdSearch", "productionIdOption", "offerName", "nameOption", "desc", "descOption", "roid", "roidOption", "createdby", "createdbyOption", "lastupdatedby", "lastupdatedbyOption", "engine", "engineOption", "banner", "bannerOption", "category", "categoryOption", "product", "productOption", "mclu", "mcluOption", "priority", "priorityOption", "createdDate1", "createdDate2", "createdOption", "startDate1", "startDate2", "startOption", "endDate1", "endDate2", "endOption", "sourceOption", " advSql", "favoriteOption", "banner", "allbannersid"}
    For Each querystring As String In Request.QueryString
        If (CMS.ExtentionMethods.IsSqlInjectioned(Request.QueryString(querystring))) Then
            Response.Redirect("~/logix/error-forbidden.aspx", True)
        End If
    Next
    For Each requestFormItem As String In SqlInjectionFormList
        Dim requestFormItemValue As String = Request.Form(requestFormItem)
        If requestFormItemValue IsNot Nothing AndAlso requestFormItemValue.Trim().Length > 0 Then
            If (CMS.ExtentionMethods.IsSqlInjectioned(requestFormItemValue)) Then
                Response.Redirect("~/logix/error-forbidden.aspx", True)
            End If
        End If
    Next

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "offer-list.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    CMS.AMS.CurrentRequest.Resolver.AppName = MyCommon.AppName

    ' Check the logged in user to see if they're to be restricted from listing some offers on this page
    MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                        "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                        "where AU.AdminUserID=" & AdminUserID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        If (rst.Rows(0).Item("prestrict") = True) Then
            restrictLinks = True
        End If
    End If

    ' Total users
    MyCommon.QueryStr = "select AdminUserID from AdminUsers;"
    rst = MyCommon.LRT_Select
    TotalUsers = rst.Rows.Count

    'Store User
    If (MyCommon.Fetch_CM_SystemOption(131) = "1") Then
        MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
        rst = MyCommon.LRT_Select
        iLen = rst.Rows.Count
        If iLen > 0 Then
            bStoreUser = True
            sValidSU = AdminUserID
            For i = 0 To (iLen - 1)
                If i = 0 Then
                    sValidLocIDs = rst.Rows(0).Item("LocationID")
                Else
                    sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
                End If
            Next

            MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
            rst = MyCommon.LRT_Select
            iLen = rst.Rows.Count
            If iLen > 0 Then
                For i = 0 To (iLen - 1)
                    sValidSU &= "," & rst.Rows(i).Item("UserID")
                Next
            End If
        End If
    End If

    bAdvancedExternalSearch = (MyCommon.Fetch_SystemOption(167) = "1")

    bCmInstalled = MyCommon.IsEngineInstalled(0)
    bUeInstalled = MyCommon.IsEngineInstalled(9)
    If bCmInstalled Then
        bWorkflowActive = (MyCommon.Fetch_CM_SystemOption(74) = "1")
        bTestSystem = (MyCommon.Fetch_CM_SystemOption(77) = "1")
        bArchiveSystem = (MyCommon.Fetch_CM_SystemOption(77) = "2")
        If bTestSystem Or bArchiveSystem Then
            bProductionSystem = False
        Else
            bProductionSystem = True
        End If
    End If

    CustomerInquiry = Request.QueryString("CustomerInquiry") <> ""
    PageNum = Request.QueryString("pagenum")
    If PageNum < 0 Then PageNum = 0
    MorePages = False

    If CustomerInquiry Then
        Send_HeadBegin("term.customerinquiry", "term.offers")
    Else
        Send_HeadBegin("term.offers")
    End If
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    If restrictLinks Then
        Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
    End If
    Send_Scripts(New String() {"datePicker.js", "popup.js"})
%>
<style type="text/css">
    body
    {
        overflow: visible;
    }
    
    #controls form
    {
        display: inline !important;
    }
    * html table
    {
        table-layout: auto !important;
    }
    * html #XIDcol
    {
        width: auto !important;
    }
    
    #performactions
    {
        position: absolute;
        z-index: 999;
        top: 10px; /* set top value */
        left: 600px; /* set left value */
        width: 400px; /* set width value */
        height: 250px;
    }
    
    #OfferfadeDiv
    {
        background-color: #e0e0e0;
        position: absolute;
        top: 0px;
        left: 0px;
        width: 100%;
        height: 100%;
        z-index: 1000;
        display: none;
        opacity: .4;
        filter: alpha(opacity=40);
    }
    
    #execduplicateoffer
    {
        overflow-y: auto;
    }
    #statusClose:hover
    {
        opacity: 1.0;
        filter: alpha(opacity=100); /* For IE8 and earlier */
    }
    #statusClose
    {
        display: block;
        float: right;
        position: relative;
        height: 15px;
        opacity: 0.7;
        filter: alpha(opacity=70); /* For IE8 and earlier */
    }
    .divideMe
    {
        width: 100px;
        word-wrap: break-word;
        -ms-word-break: break-all;
        -webkit-hyphens: auto;
        -moz-hyphens: auto;
        hyphens: auto;
    }
</style>
<%
    Send_Scripts()
%>
<script type="text/javascript">

  function doOnClick() {
        $('#hrefAdvSearchq').click();
  }

  function launchAdvSearch() {
    self.name = "OfferListWin";
    <%
      If CustomerInquiry Then
        Send("openPopup(""advanced-search.aspx?CustomerInquiry=1"");")
      Else
        Send("openPopup(""advanced-search.aspx"");")
      End If
    %>
  }
  function editSearchCriteria() {
    var tokenStr = document.frmIter.advTokens.value;
    
    self.name = "OfferListWin";
    <%
      If CustomerInquiry Then
        Send("openPopup(""advanced-search.aspx?CustomerInquiry=1&amp;tokens="" + tokenStr);")
      Else
        Send("openPopup(""advanced-search.aspx?tokens="" + tokenStr);")
      End If
    %>
  }
  function reInitializeSelectedItems(){
    selectedItems = new Array();
  }
//   function chooseFile() {
//      document.getElementById("fileInput").click();
//   }
//   function fileonclick()
//   {
//   var filename=document.getElementById("fileInput").value;
//    document.getElementById("lblfileupload").innerText = filename.replace("C:\\fakepath\\", "");
//  }
</script>
<script type="text/javascript" src="../javascript/jquery.min.js"></script>
<script type="text/javascript" src="../javascript/thickbox.js"></script>
<%


    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    If CustomerInquiry Then
        If (Not restrictLinks) Then
            Send_Tabs(Logix, 3)
            Send_Subtabs(Logix, 32, 2, , 0)
        Else
            Send_Subtabs(Logix, 91, 1, , 0)
        End If
    Else
        Send_Tabs(Logix, 2)
        Send_Subtabs(Logix, 20, 1)
    End If

    If (Logix.UserRoles.AccessOffers = False) Then
        Send_Denied(1, "perm.offers-access")
        GoTo done
    End If

    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

    If Request.QueryString("LargeFile") = "true" Then
        infoMessage = Copient.PhraseLib.Lookup("error.UploadTooLarge", LanguageID)
    End If

    If Request.Files.Count >= 1 Then

        'get the file data
        File = Request.Files.Get(0)
        'Response.Write("Got file upload request" & File.ContentType)
        'check that it's not too big

        If (File.ContentLength = 0 AndAlso File.FileName <> "") Then
            infoMessage = Copient.PhraseLib.Lookup("term.upload-file-not-found", LanguageID) & " (" & File.FileName & ")"
        ElseIf (File.ContentLength = 0 AndAlso File.FileName = "") Then
            infoMessage = Copient.PhraseLib.Lookup("term.nofileselected", LanguageID)
        ElseIf File.ContentType <> "text/xml" And File.ContentType <> "application/octet-stream" And File.ContentType <> "application/x-gzip" _
        And File.ContentType <> "application/x-gzip-compressed" And File.ContentType <> "application/gzip" And File.ContentType <> "application/x-tar" _
        And File.ContentType <> "text/plain" Then
            infoMessage = File.ContentType & vbCrLf & Copient.PhraseLib.Lookup("offer-list.notxml", LanguageID)
        Else
            'save file
            ' Response.Write("saving file")
            'open in out stream
            Dim sr As System.IO.StreamReader
            Dim UploadFileName As String
            Dim TimeStampStr As String
            TimeStampStr = MyCommon.Leading_Zero_Fill(Day(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Hour(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Minute(Date.Now), 2) & MyCommon.Leading_Zero_Fill(Second(Date.Now), 2)
            'UploadFileName = "Offer-" & TimeStampStr & ".xml"
            UploadFileName = MyCommon.Fetch_SystemOption(29) & "\Offer-import" & TimeStampStr & ".xml"
            File.SaveAs(UploadFileName)

            sr = New System.IO.StreamReader(File.InputStream)
            sr.Close()

            ' ok the file is there now time to write it out
            Dim bStatus As Boolean
            Dim EngineId As Integer = -1
            Dim sOfferId As String
            Dim sXml As String = ""
            Dim MyImportXml As New Copient.ImportXml(MyCommon)
            Dim CpeImport As New Copient.ImportXMLCPE
            CMS.AMS.CurrentRequest.Resolver.AppName = "Offer-List.aspx"
            Dim UEImport As Copient.ImportXMLUE = CMS.AMS.CurrentRequest.Resolver.Resolve(Of Copient.ImportXMLUE)()
            Dim sMsg As String
            Dim xsdName As String = "PromoUE.xsd"
            Dim xsdPath As String = MyCommon.Get_Install_Path & "AgentFiles\" & xsdName
            Dim LogFile As String = "ImportXmlUE." & Date.Now.ToString("yyyyMMdd") & ".txt"
            Dim sErrorMsg As String = ""
            EngineId = MyImportXml.GetOfferEngineId(UploadFileName, sXml)

            MyCommon.QueryStr = "SELECT Description FROM PromoEngines WITH (NoLock) WHERE EngineID=" & EngineId & " AND Installed=1; "
            dst = MyCommon.LRT_Select
            If dst.Rows.Count > 0 Then
                ' cpe 
                If (EngineId = 2 OrElse EngineId = 3 OrElse EngineId = 6) Then
                    If (MyCommon.Fetch_SystemOption(66) = "1") Then
                        CpeImport.SetBanners(GetBanners())
                    End If
                    CpeImport.ImportOffer(UploadFileName, sXml, AdminUserID, LanguageID, False)
                    sMsg = CpeImport.GetErrorMsg()
                    If (sMsg.Trim() = "") Then
                        MyCommon.Activity_Log(3, CpeImport.GetOfferId(), AdminUserID, Copient.PhraseLib.Lookup("offer.imported", LanguageID))
                        MarkOfferAsImported(CpeImport.GetOfferId, MyCommon)
                        Select Case CpeImport.GetEngineId()
                            Case 2
                                sSummaryPage = "CPEoffer-sum.aspx"
                            Case 3
                                sSummaryPage = "web-offer-sum.aspx"
                            Case 6
                                sSummaryPage = "/logix/CAM/CAM-offer-sum.aspx"
                            Case Else
                                sSummaryPage = "CPEoffer-sum.aspx"
                        End Select
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", sSummaryPage & "?OfferID=" & CpeImport.GetOfferId() & "&imported=1")
                    Else
                        infoMessage = sMsg
                    End If

                    ' ue 
                ElseIf (EngineId = 9) Then
                    If (MyCommon.Fetch_SystemOption(66) = "1") Then
                        UEImport.SetBanners(GetBanners())
                    End If
                    If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not Logix.UserRoles.CreateUEOffers) Then
                        infoMessage = Copient.PhraseLib.Lookup("term.unsupportedpromoengine", LanguageID)
                    Else
                        If Not IsValidXml(MyCommon, LogFile, xsdPath, sXml) Then
                            sMsg = GetErrorMsg()
                            infoMessage = sMsg
                        Else
                            UEImport.ImportOffer(UploadFileName, sXml, AdminUserID, LanguageID, False)
                            sMsg = UEImport.GetErrorMsg()
                            If (sMsg.Trim() = "") Then
                                MyCommon.Activity_Log(3, UEImport.GetOfferId(), AdminUserID, Copient.PhraseLib.Lookup("offer.imported", LanguageID))
                                MarkOfferAsImported(UEImport.GetOfferId, MyCommon)
                                Response.Status = "301 Moved Permanently"
                                Response.AddHeader("Location", "/logix/UE/UEoffer-sum.aspx?OfferID=" & UEImport.GetOfferId() & "&imported=1")
                            Else
                                infoMessage = sMsg
                            End If
                        End If
                    End If
                    ' ???
                ElseIf (EngineId = -1) Then
                    infoMessage = MyImportXml.GetStatusMsg
                    If infoMessage = "" Then
                        infoMessage = Copient.PhraseLib.Lookup("offer-list.notoffer", LanguageID)
                    End If

                Else
                    bStatus = MyImportXml.ImportOfferLoad(UploadFileName, sXml, AdminUserID, LanguageID)
                    sMsg = MyImportXml.GetStatusMsg
                    If sMsg.Length > 0 Then
                        If bStatus Then
                            ' display warning using sMsg
                            sOfferId = MyImportXml.GetOfferId
                            infoMessage = sMsg
                            AssignOfferToBanners(CInt(sOfferId), MyCommon)
                            MarkOfferAsImported(CLng(sOfferId), MyCommon)
                            Response.Status = "301 Moved Permanently"
                            Response.AddHeader("Location", "offer-gen.aspx?OfferID=" & sOfferId & "&imported=1&infoMessage=" & infoMessage)
                        Else
                            ' display error using sMsg
                            infoMessage = sMsg
                        End If
                    Else
                        If bStatus Then
                            sOfferId = MyImportXml.GetOfferId
                            AssignOfferToBanners(CInt(sOfferId), MyCommon)
                            MarkOfferAsImported(CLng(sOfferId), MyCommon)
                            Response.Status = "301 Moved Permanently"
                            Response.AddHeader("Location", "offer-gen.aspx?OfferID=" & sOfferId & "&imported=1")
                        End If
                    End If
                    If System.IO.File.Exists(UploadFileName) = True Then
                        System.IO.File.Delete(UploadFileName)
                    End If
                End If
            Else
                ' the engine in XML is not installed
                sMsg = MyImportXml.GetErrorMsg()
                If sMsg <> "" Then
                    infoMessage = sMsg
                Else

                    infoMessage = Copient.PhraseLib.Lookup("term.unsupportedpromoengine", LanguageID)

                End If
            End If

        End If
    End If

    ' handle an Advance Search Criteria
    If (Request.Form("mode") = "advancedsearch") Then
        Dim TempStr As String = ""
        Dim CritBuf As New StringBuilder()
        Dim CritTokenBuf As New StringBuilder()

        If (hasOption("xid") AndAlso Request.Form("xidOption") <> "") Then
            WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("xidOption")), Request.Form("xid"), "AOLV.ExtOfferID"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.xid", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("xidOption"))) & " '" & Request.Form("xid").Trim & "'")
            CritTokenBuf.Append("XID," & Integer.Parse(Request.Form("xidOption")) & "," & Request.Form("xid").Trim & ",|")
        End If

        If (hasOption("idSearch") AndAlso Request.Form("idOption") <> "") Then
            If IsValid_LongValue(Request.Form("idSearch")) Then
                WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("idOption")), Request.Form("idSearch"), "AOLV.OfferID"))
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                CritBuf.Append(Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("idOption"))) & " '" & Request.Form("idSearch").Trim & "'")
                CritTokenBuf.Append("ID," & Integer.Parse(Request.Form("idOption")) & "," & Request.Form("idSearch").Trim & ",|")
                bSearchInternalAndExternal = True
            Else
                If Not CriteriaError Then
                    CriteriaError = True
                    CriteriaMsg = Copient.PhraseLib.Lookup("customer-manual.InvalidOfferID", LanguageID) & " (" & Long.MaxValue & " " & Copient.PhraseLib.Lookup("offerid-maxvalue", LanguageID) & ")"
                End If
            End If
        End If

        'added
        If (hasOption("buyeridSearch") AndAlso Request.Form("buyeridOption") <> "") Then
            WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("buyeridOption")), Request.Form("buyeridSearch"), "AOLV.ExternalbuyerId"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.buyerid", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("buyeridOption"))) & " '" & Request.Form("buyeridSearch").Trim & "'")
            CritTokenBuf.Append("Buyer Id," & Integer.Parse(Request.Form("buyeridOption")) & "," & Request.Form("buyeridSearch").Trim & ",|")
            bSearchInternalAndExternal = True
        End If

        If (hasOption("productionIdSearch") AndAlso Request.Form("productionIdOption") <> "") Then
            WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("productionIdOption")), Request.Form("productionIdSearch"), "AOLV.ProductionID"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.productionid", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("productionIdOption"))) & " '" & Request.Form("productionIdSearch").Trim & "'")
            CritTokenBuf.Append("PRODUCTIONID," & Integer.Parse(Request.Form("productionIdOption")) & "," & Request.Form("productionIdSearch").Trim & ",|")
            bSearchInternalAndExternal = True
        End If
        If (hasOption("offerName") AndAlso Request.Form("nameOption") <> "") Then
            WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("nameOption")), Request.Form("offerName"), "AOLV.Name"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.name", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("nameOption"))) & " '" & Request.Form("offerName").Trim & "'")
            CritTokenBuf.Append("Name," & Integer.Parse(Request.Form("nameOption")) & "," & Request.Form("offerName").Trim & ",|")
        End If

        If (hasOption("desc") AndAlso Request.Form("descOption") <> "") Then
            WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("descOption")), Request.Form("desc"), "Convert(nvarchar(1000),OfferDescription)"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.description", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("descOption"))) & " '" & Request.Form("desc").Trim & "'")
            CritTokenBuf.Append("Desc," & Integer.Parse(Request.Form("descOption")) & "," & Request.Form("desc").Trim & ",|")
        End If

        If (hasOption("roid") AndAlso Request.Form("roidOption") <> "") Then
            If IsValid_LongValue(Request.Form("roid")) Then
                WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("roidOption")), Request.Form("roid"), "RewardOptionID"))
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                CritBuf.Append(Copient.PhraseLib.Lookup("term.roid", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("roidOption"))) & " '" & Request.Form("roid").Trim & "'")
                CritTokenBuf.Append("ROID," & Integer.Parse(Request.Form("roidOption")) & "," & Request.Form("roid").Trim & ",|")
            Else
                If Not CriteriaError Then
                    CriteriaError = True
                    CriteriaMsg = Copient.PhraseLib.Lookup("roid-numeric", LanguageID) & " (" & Long.MaxValue & " " & Copient.PhraseLib.Lookup("roid-maxvalue", LanguageID) & ")"
                End If
            End If
        End If

        If (hasOption("createdby") AndAlso Request.Form("createdbyOption") <> "") Then
            WhereBuf.Append("and AOLV.CreatedByAdminID IN (select Distinct AdminUserID from AdminUsers where ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("createdbyOption")), Request.Form("createdby"), "UserName"))
            WhereBuf.Append("or " & GetOptionString(MyCommon, Integer.Parse(Request.Form("createdbyOption")), Request.Form("createdby"), "FirstName"))
            WhereBuf.Append("or " & GetOptionString(MyCommon, Integer.Parse(Request.Form("createdbyOption")), Request.Form("createdby"), "LastName"))
            WhereBuf.Append(") ")
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.createdby", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("createdbyOption"))) & " '" & Request.Form("createdby").Trim & "'")
            CritTokenBuf.Append("CreatedBy," & Integer.Parse(Request.Form("createdbyOption")) & "," & Request.Form("createdby").Trim & ",|")
        End If

        If (hasOption("lastupdatedby") AndAlso Request.Form("lastupdatedbyOption") <> "") Then
            WhereBuf.Append(" and AOLV.LastUpdatedByAdminID IN (select Distinct AdminUserID from AdminUsers where ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("lastupdatedbyOption")), Request.Form("lastupdatedby"), "UserName"))
            WhereBuf.Append("or " & GetOptionString(MyCommon, Integer.Parse(Request.Form("lastupdatedbyOption")), Request.Form("lastupdatedby"), "FirstName"))
            WhereBuf.Append("or " & GetOptionString(MyCommon, Integer.Parse(Request.Form("lastupdatedbyOption")), Request.Form("lastupdatedby"), "LastName"))
            WhereBuf.Append(") ")
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.lastupdatedby", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("lastupdatedbyOption"))) & " '" & Request.Form("lastupdatedby").Trim & "'")
            CritTokenBuf.Append("LastUpdatedBy," & Integer.Parse(Request.Form("lastupdatedbyOption")) & "," & Request.Form("lastupdatedby").Trim & ",|")
        End If

        If (hasOption("engine") AndAlso Request.Form("engineOption") <> "") Then
            WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("engineOption")), Request.Form("engine"), "PromoEngine"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.engine", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("engineOption"))) & " '" & Request.Form("engine").Trim & "'")
            CritTokenBuf.Append("Engine," & Integer.Parse(Request.Form("engineOption")) & "," & Request.Form("engine").Trim & ",|")
            If (bEnableRestrictedAccessToUEOfferBuilder) Then
                FilterEngine = IIf(Request.Form("engine").Trim.ToLower.Equals("cm"), 0, 9)
            End If
        End If

        If (BannersEnabled) Then
            If (hasOption("banner") AndAlso Request.Form("bannerOption") <> "") Then
                WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("bannerOption")), Request.Form("banner"), "BAN.Name"))
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                CritBuf.Append(Copient.PhraseLib.Lookup("term.banner", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("bannerOption"))) & " '" & Request.Form("banner").Trim & "'")
                CritTokenBuf.Append("BAN.Name," & Integer.Parse(Request.Form("bannerOption")) & "," & Request.Form("banner").Trim & ",|")
            End If
        End If

        If (hasOption("category") AndAlso Request.Form("categoryOption") <> "") Then
            WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("categoryOption")), Request.Form("category"), "ODescription"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.category", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("categoryOption"))) & " '" & Request.Form("category").Trim & "'")
            CritTokenBuf.Append("Category," & Integer.Parse(Request.Form("categoryOption")) & "," & Request.Form("category").Trim & ",|")
        End If

        If (hasOption("product") AndAlso Request.Form("productOption") <> "") Then
            WhereBuf.Append(" and AOLV.OfferID in (" & GetProductOfferList(MyCommon, Request.Form("product"), Integer.Parse(Request.Form("productOption"))) & ") ")
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.product", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("productOption"))) & " '" & Request.Form("product").Trim & "'")
            CritTokenBuf.Append("Product," & Integer.Parse(Request.Form("productOption")) & "," & Request.Form("product").Trim & ",|")
        End If

        'Search MCLU  
        If (hasOption("mclu") AndAlso Request.Form("mcluOption") <> "") Then
            WhereBuf.Append(" and AOLV.OfferID in (" & GetMCLUOfferList(MyCommon, Request.Form("mclu"), Integer.Parse(Request.Form("mcluOption"))) & ") ")
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.mclu", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("mcluOption"))) & " '" & Request.Form("mclu").Trim & "'")
            CritTokenBuf.Append("MCLU," & Integer.Parse(Request.Form("mcluOption")) & "," & Request.Form("mclu").Trim & ",|")
        End If

        If (hasOption("priority") AndAlso Request.Form("priorityOption") <> "") Then
            WhereBuf.Append(" and AOLV.OfferID in (" & GetPriorityOfferList(MyCommon, Request.Form("priority"), Integer.Parse(Request.Form("priorityOption"))) & ") ")
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.cm-offer-priority", LanguageID) & " " & GetPriorityOptionType(Integer.Parse(Request.Form("priorityOption"))) & " '" & Request.Form("priority").Trim & "'")
            CritTokenBuf.Append("Priority," & Integer.Parse(Request.Form("priorityOption")) & "," & Request.Form("priority").Trim & ",|")
        End If

        If (Not CriteriaError AndAlso Validate_Date_Range(Request.Form("createdDate1"), Request.Form("createdDate2"), Request.Form("createdOption"), CriteriaMsg)) Then
            Try
                TempStr = GetDateOption(MyCommon, Integer.Parse(Request.Form("createdOption")), Request.Form("createdDate1"), Request.Form("createdDate2"), "AOLV.CreatedDate")
            Catch aex As ApplicationException
                CriteriaError = True
                TempStr = ""
                CriteriaMsg = aex.Message
            End Try

            If (TempStr <> "") Then WhereBuf.Append(" and " & TempStr)
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            If TempStr IsNot Nothing AndAlso (TempStr.IndexOf("between") > -1) Then
                CritBuf.Append(Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("createdOption"))) & " '" & Request.Form("createdDate1").Trim & "'")
                If hasOption("createdDate2") Then
                    CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("createdDate2").Trim & "'")
                End If
                CritTokenBuf.Append("Created," & Integer.Parse(Request.Form("createdOption")) & "," & Request.Form("createdDate1").Trim & "," & Request.Form("createdDate2").Trim & "|")
            Else
                CritBuf.Append(Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("createdOption"))) & " '" & Request.Form("createdDate1").Trim & "'")
                CritTokenBuf.Append("Created," & Integer.Parse(Request.Form("createdOption")) & "," & Request.Form("createdDate1").Trim & ",|")
            End If
        Else
            If (Not CriteriaError And CriteriaMsg <> "") Then
                CriteriaError = True
                CriteriaMsg = CriteriaMsg + " " + Copient.PhraseLib.Lookup("term.for", LanguageID).ToLower() + " '" + Copient.PhraseLib.Lookup("term.created", LanguageID) + "' " + Copient.PhraseLib.Lookup("term.option", LanguageID)
                TempStr = ""
            End If
        End If

        If (Not CriteriaError AndAlso Validate_Date_Range(Request.Form("startDate1"), Request.Form("startDate2"), Request.Form("startOption"), CriteriaMsg)) Then
            Try
                TempStr = GetDateOption(MyCommon, Integer.Parse(Request.Form("startOption")), Request.Form("startDate1"), Request.Form("startDate2"), "ProdStartDate")
            Catch aex As ApplicationException
                CriteriaError = True
                TempStr = ""
                CriteriaMsg = aex.Message
            End Try

            If (TempStr <> "") Then WhereBuf.Append(" and " & TempStr)

            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            If TempStr IsNot Nothing AndAlso (TempStr.IndexOf("between") > -1) Then
                CritBuf.Append(Copient.PhraseLib.Lookup("term.starts", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("startOption"))) & " '" & Request.Form("startDate1").Trim & "'")
                If hasOption("startDate2") Then
                    CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("startDate2").Trim & "'")
                End If
                CritTokenBuf.Append("Starts," & Integer.Parse(Request.Form("startOption")) & "," & Request.Form("startDate1").Trim & "," & Request.Form("startDate2").Trim & "|")
            Else
                CritBuf.Append(Copient.PhraseLib.Lookup("term.starts", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("startOption"))) & " '" & Request.Form("startDate1").Trim & "'")
                CritTokenBuf.Append("Starts," & Integer.Parse(Request.Form("startOption")) & "," & Request.Form("startDate1").Trim & ",|")
            End If
        Else
            If (Not CriteriaError AndAlso CriteriaMsg <> "") Then
                CriteriaError = True
                CriteriaMsg = CriteriaMsg + " " + Copient.PhraseLib.Lookup("term.for", LanguageID).ToLower() + " '" + Copient.PhraseLib.Lookup("term.starts", LanguageID) + "' " + Copient.PhraseLib.Lookup("term.option", LanguageID)
                TempStr = ""
            End If
        End If

        If (Not CriteriaError AndAlso Validate_Date_Range(Request.Form("endDate1"), Request.Form("endDate2"), Request.Form("endOption"), CriteriaMsg)) Then
            Try
                TempStr = GetDateOption(MyCommon, Integer.Parse(Request.Form("endOption")), Request.Form("endDate1"), Request.Form("endDate2"), "ProdEndDate")
            Catch aex As ApplicationException
                CriteriaError = True
                TempStr = ""
                CriteriaMsg = aex.Message
            End Try

            If (TempStr <> "") Then WhereBuf.Append(" and " & TempStr)

            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            If TempStr IsNot Nothing AndAlso (TempStr.IndexOf("between") > -1) Then
                CritBuf.Append(Copient.PhraseLib.Lookup("term.ends", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("endOption"))) & " '" & Request.Form("endDate1").Trim & "'")
                If hasOption("endDate2") Then
                    CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("endDate2").Trim & "'")
                End If
                CritTokenBuf.Append("Ends," & Integer.Parse(Request.Form("endOption")) & "," & Request.Form("endDate1").Trim & "," & Request.Form("endDate2").Trim & "|")
            Else
                CritBuf.Append(Copient.PhraseLib.Lookup("term.ends", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("endOption"))) & " '" & Request.Form("endDate1").Trim & "'")
                CritTokenBuf.Append("Ends," & Integer.Parse(Request.Form("endOption")) & "," & Request.Form("endDate1").Trim & ",|")
            End If
        Else
            If (Not CriteriaError AndAlso CriteriaMsg <> "") Then
                CriteriaError = True
                CriteriaMsg = CriteriaMsg + " " + Copient.PhraseLib.Lookup("term.for", LanguageID).ToLower() + " '" + Copient.PhraseLib.Lookup("term.ends", LanguageID) + "' " + Copient.PhraseLib.Lookup("term.option", LanguageID)
                TempStr = ""
            End If
        End If

        If (Request.Form("sourceOption") <> "0") Then
            If (Request.Form("sourceOption") = "-1") Then
                SourceName = Copient.PhraseLib.Lookup("term.allsources", LanguageID)
                WhereBuf.Append(" and InboundCRMEngineID in (" & GetExternalSourceOfferList(MyCommon) & ")")
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                CritBuf.Append(Copient.PhraseLib.Lookup("term.source", LanguageID) & " " & GetOptionType(6) & " '" & SourceName & "'")
                CritTokenBuf.Append("Source," & Request.Form("sourceOption") & "," & "" & ",|")
            Else
                MyCommon.QueryStr = "select Name from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=" & Request.Form("sourceOption") & ";"
                dst = MyCommon.LRT_Select
                If dst.Rows.Count > 0 Then
                    SourceName = MyCommon.NZ(dst.Rows(0).Item("Name"), "")
                End If
                WhereBuf.Append(" and " & GetOptionString(MyCommon, "6", Request.Form("sourceOption"), "InboundCRMEngineID"))
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                CritBuf.Append(Copient.PhraseLib.Lookup("term.source", LanguageID) & " " & GetOptionType(6) & " '" & SourceName & "'")
                CritTokenBuf.Append("Source," & Request.Form("sourceOption") & "," & "" & ",|")
            End If
            bSearchInternalAndExternal = True
        End If

        If (Request.Form("favoriteOption") <> "0") Then
            'WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("favoriteOption")), "1", "Favorite"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.favorite", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("favoriteOption"))) & " on")
            CritTokenBuf.Append("Favorite," & Integer.Parse(Request.Form("favoriteOption")) & "," & "1" & ",|")
        End If

        If MyCommon.Fetch_SystemOption(156) = "1" Then
            MyCommon.QueryStr = "select * from UserDefinedFields where AdvancedSearch = 1"
            dst = MyCommon.LRT_Select

            Dim RowCount As Integer = -1
            Dim row As DataRow
            Dim udfdst As DataTable
            Dim udfquery As New StringBuilder()
            If dst.Rows.Count <= 5 Then
                If dst.Rows.Count > 0 Then
                    RowCount = dst.Rows.Count - 1
                End If
            Else
                RowCount = 4
            End If
            For udfcount As Integer = 0 To RowCount
                MyCommon.QueryStr = "select Top 1 * from UserDefinedFields udf inner join UserDefinedFieldsTypes t on udf.DataType = t.UDFTypeID where udf.UDFPK = " & Request.Form("UDFDataType-" & udfcount)
                udfdst = MyCommon.LRT_Select
                row = udfdst.Rows(0)

                Dim UDFPK = row.Item("UDFPK")
                If (hasOption("udf-" & udfcount) And Request.Form("udf-" & udfcount) <> "") And row.Item("DataType") <> 3 Then

                    'If udfAdvSearch = False Then	udfAdvSearch = True

                    Select Case row.Item("DataType")
                        Case 0, 1, 4, 5, 6 'string, integer, listbox, likert

                            If udfquery.Length > 0 Then udfquery.Append(" Intersect ")

                            Dim testString As String
                            testString = "select OfferID from UserDefinedFieldsValues where " & GetOptionString(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), Request.Form("udf-" & udfcount), row.Item("ColumnName"))


                            udfquery.Append("select OfferID from UserDefinedFieldsValues where " & GetOptionString(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), Request.Form("udf-" & udfcount), row.Item("ColumnName")))
                            udfquery.Append(" and UDFPK = " & UDFPK & " and OfferID = AOLV.OfferID")
                            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                            CritBuf.Append("UDF-" & UDFPK & " " & GetOptionType(Integer.Parse(Request.Form("udfOption-" & udfcount))) & " '" & Request.Form("udf-" & udfcount).Trim & "'")
                            CritTokenBuf.Append("UDFRow-" & udfcount & ",UDF-" & UDFPK & "," & Integer.Parse(Request.Form("udfOption-" & udfcount)) & "," & Request.Form("udf-" & udfcount).Trim & ",|")
                        Case 2
                            Try
                                TempStr = "select OfferID from UserDefinedFieldsValues where " & GetDateOption(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), Request.Form("udf-" & udfcount), MyCommon.NZ(Request.Form("udfEnd-" & udfcount), ""), row.Item("ColumnName"))
                                TempStr &= " and UDFPK = " & UDFPK & " and OfferID = AOLV.OfferID"
                            Catch aex As ApplicationException
                                CriteriaError = True
                                TempStr = ""
                                CriteriaMsg = aex.Message
                            End Try
                            If (TempStr <> "") Then
                                If udfquery.Length > 0 Then udfquery.Append(" Intersect ")
                                udfquery.Append(TempStr)
                            End If
                            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                            If TempStr IsNot Nothing AndAlso (TempStr.IndexOf("between") > -1) Then
                                CritBuf.Append("UDF-" & UDFPK & " " & GetDateOptionType(Integer.Parse(Request.Form("udfOption-" & udfcount))) & " '" & Request.Form("udf-" & udfcount).Trim & "'")
                                If hasOption("udfEnd-" & udfcount) Then
                                    CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("udfEnd-" & udfcount).Trim & "'")
                                End If
                                CritTokenBuf.Append("UDFRow-" & udfcount & ",UDF-" & UDFPK & "," & Integer.Parse(Request.Form("udfOption-" & udfcount)) & "," & Request.Form("udf-" & udfcount).Trim & "," & Request.Form("udfEnd-" & udfcount).Trim & "|")
                            Else
                                CritBuf.Append("UDF-" & UDFPK & " " & GetDateOptionType(Integer.Parse(Request.Form("udfOption-" & udfcount))) & " '" & Request.Form("udf-" & udfcount).Trim & "'")
                                CritTokenBuf.Append("UDFRow-" & udfcount & ",UDF-" & UDFPK & "," & Integer.Parse(Request.Form("udfOption-" & udfcount)) & "," & Request.Form("udf-" & udfcount).Trim & ",|")
                            End If
                    End Select
                ElseIf Request.Form("udfOption-" & udfcount) <> "0" And row.Item("DataType") = "3" Then
                    'If udfAdvSearch = False Then	udfAdvSearch = True
                    If udfquery.Length > 0 Then udfquery.Append(" Intersect ")
                    udfquery.Append("select OfferID from UserDefinedFieldsValues where " & GetOptionString(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), "1", row.Item("ColumnName")))
                    udfquery.Append(" and UDFPK = " & UDFPK & " and OfferID = AOLV.OfferID")
                    If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                    CritBuf.Append("UDF-" & UDFPK & " " & GetOptionType(Integer.Parse(Request.Form("udfOption-" & udfcount))) & " on")
                    CritTokenBuf.Append("UDFRow-" & udfcount & ",UDF-" & UDFPK & "," & Integer.Parse(Request.Form("udfOption-" & udfcount)) & "," & "1" & ",|")

                End If
            Next
            If udfquery.Length > 0 Then WhereBuf.Append(" and AOLV.OfferID = (" & udfquery.ToString & ")")

        End If



        If (Not CriteriaError) Then
            CriteriaMsg &= CritBuf.ToString
        End If
        CriteriaTokens = CritTokenBuf.ToString
    End If
    Dim m_isOAWEnabled As Boolean = False
    Dim SortText As String = "AOLV.OfferID"
    Dim SortDirection As String = "DESC"
    Dim ShowExpired As String = ""
    Dim ShowActive As String = ""
    Dim PrctSignPos As Integer
    Dim FilterOffer As String
    Dim FilterUser As String
    Dim enableBuyerIdSearch As Boolean = IIf(MyCommon.Fetch_UE_SystemOption(169) = "1", True, False)
    Dim m_OAWService As IOfferApprovalWorkflowService = CurrentRequest.Resolver.Resolve(Of IOfferApprovalWorkflowService)()
    Dim BannerIds As Integer()
    FilterOffer = Server.HtmlEncode(Request.QueryString("filterOffer"))
    FilterUser = Server.HtmlEncode(Request.QueryString("filterUser"))
    If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Request.Form("mode") <> "advancedsearch") Then
        If (Not String.IsNullOrEmpty(Request.QueryString("filterengine"))) Then
            FilterEngine = Convert.ToInt16(Server.HtmlEncode(Request.QueryString("filterengine")))
            If (FilterEngine = 9) Then
                If (FilterOffer = "5" OrElse FilterOffer = "6" OrElse FilterOffer = "7" OrElse FilterOffer = "8") Then FilterOffer = 0
            End If
        End If
    End If
    If BannersEnabled Then
        BannerIds = GetBanners()
        m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabledForBanners(BannerIds).Result
    Else
        m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabled().Result
    End If
    If (FilterOffer = "") Then FilterOffer = "1"
    If (FilterUser = "") Then FilterUser = AdminUserID.ToString
    If (FilterOffer = "0" OrElse FilterOffer = "3" OrElse FilterOffer = "4") Then
        ShowExpired = " AOLV.deleted=0 and isnull(AOLV.InboundCRMEngineID,0) = 0 "
    ElseIf (FilterOffer = "1") Then
        ShowExpired = " AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and isnull(AOLV.InboundCRMEngineID,0) = 0 "
    Else
        ShowExpired = " AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and isnull(AOLV.InboundCRMEngineID,0) = 0 "
    End If

    If (Request.QueryString("SortText") <> "") Then
        SortText = HttpUtility.HtmlEncode(Request.QueryString("SortText"))
    End If


    If (Request.QueryString("Sort") <> "") Then
        If (Request.QueryString("SortDirection") = "ASC") Then
            SortDirection = "DESC"
        ElseIf (Request.QueryString("SortDirection") = "DESC") Then
            SortDirection = "ASC"
        Else
            SortDirection = "DESC"
        End If
    Else
        If (Request.QueryString("SortDirection") <> "") Then
            SortDirection = Request.QueryString("SortDirection")
        End If
    End If

    If bStoreUser Then
        sJoin = "Full Outer Join OfferLocUpdate olu with (NoLock) on AOLV.OfferID=olu.OfferID "
        wherestr = " (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) and "
    End If

    MyCommon.QueryStr = "from AllOffersListviewNoCAM AOLV with (NoLock) " & sJoin

    If (Request.Form("favoriteOption") <> "0" And Request.Form("favoriteOption") <> "") Then
        If (Request.Form("favoriteOption") = "6") Then
            MyCommon.QueryStr = "into #temp from AllOffersListviewNoCAM AOLV "
            MyCommon.QueryStr += " RIGHT JOIN AdminUserOffers AUO on AOLV.OfferID=AUO.OfferID "
            'MyCommon.QueryStr = MyCommon.QueryStr & WhereBuf.ToString & "where"
        Else
            WhereBuf.Append("and AOLV.OfferID not in (select OfferID from AdminUserOffers  group BY OfferID) ")
        End If

    End If
    If BannersEnabled Then
        MyCommon.QueryStr = MyCommon.QueryStr & "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID " & _
                            "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID "
    End If
    MyCommon.QueryStr = MyCommon.QueryStr & "where" & wherestr

    If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
        If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
        If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
        If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))

        MyCommon.QueryStr += ShowExpired & " and IsNull(AOLV.isTemplate,0)=0 " & WhereBuf.ToString & ShowActive
        If (Request.Form("favoriteOption") <> "0" And Request.Form("favoriteOption") <> "") Then
            If (Request.Form("favoriteOption") = "6") Then
                MyCommon.QueryStr += " ORDER BY AOLV.OfferID "
            End If
        End If
        'MyCommon.QueryStr += " order by " & SortText & " " & SortDirection
        AdvSearchSQL = WhereBuf.ToString
        If (FilterOffer <> "3") Then
            'If this option is enabled and selected, return a dropdown of all users with at least one banner in common
            If (FilterOffer = "4") Then
                'MyCommon.QueryStr &= " and AOLV.CreatedByAdminID = " & AdminUserID & " "
                MyCommon.QueryStr &= " and AOLV.CreatedByAdminID = " & FilterUser & " "
            End If
        End If
        If (BannersEnabled) Then
            MyCommon.QueryStr &= " and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                                 " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                                 "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                                 "                     where AUB.AdminUserID = " & AdminUserID & ") ) "
        End If
        If bAdvancedExternalSearch Then
            'Enable Advanced Search for External Offers   
            bSearchInternalAndExternal = True
        End If
    Else
        If (Request.QueryString("searchterms") <> "") Then
            If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
                idSearch = idNumber.ToString
            Else
                idSearch = "-1"
            End If
            idSearchText = MyCommon.Parse_Quotes(Server.HtmlDecode(Request.QueryString("searchterms")))
            PrctSignPos = idSearchText.IndexOf("%")
            If (PrctSignPos > -1) Then
                idSearch = "-1"
                idSearchText = idSearchText.Replace("%", "[%]")
            End If
            If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
            If bProductionSystem Then
                MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired & " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & _
                                           "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%'"
            Else
                MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired & " and (AOLV.OfferID=" & idSearch & " or AOLV.ProductionID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & _
                                    "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%'"
            End If
            ''If buyerid sysopt is enabled and ue engine is installed append your condition to include the results for that buyer id
            If (enableBuyerIdSearch) AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                'MyCommon.QueryStr = MyCommon.QueryStr & " or( AOLV.ExternalbuyerId like N'%" & idSearchText & "%' and Deleted=0)"
                MyCommon.QueryStr = MyCommon.QueryStr & " or AOLV.ExternalbuyerId like N'%" & idSearchText & "%')"
            Else
                MyCommon.QueryStr = MyCommon.QueryStr & ")"

            End If

            MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 " & ShowActive
            If idSearch > 0 Then
                bSearchInternalAndExternal = True
            End If
        Else
            MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired
            MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 " & ShowActive
        End If

        ' check if banners are enabled
        If (BannersEnabled) Then
            MyCommon.QueryStr &= " and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                                 " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                                 "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                                 "                     where AUB.AdminUserID = " & AdminUserID & _
                                 IIf(FilterOffer = "4", " or AUB.AdminUserID = " & FilterUser, "") & ") ) "
        End If

        If (FilterOffer <> "3") Then
            'If this option is enabled and selected, return a dropdown of all users with at least one banner in common
            If (FilterOffer = "4") Then
                MyCommon.QueryStr &= " and AOLV.CreatedByAdminID = " & FilterUser & " "
            End If
        End If
    End If

    ShowExpired = IIf(FilterOffer = "0" OrElse FilterOffer = "3" OrElse FilterOffer = "4", "TRUE", "FALSE")
    If (FilterOffer = "3") Then
        MyCommon.QueryStr &= " and AOLV.OfferID in (" & _
                             "select Distinct STI.IncentiveID as OfferID from CPE_ST_Incentives STI with (NoLock) " & _
                             "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = STI.IncentiveID " & _
                             "where STI.Deleted=0 and I.Deleted=0 and (getdate() between STI.StartDate and DateAdd(d, 1, STI.EndDate)) and (I.UpdateLevel <> STI.UpdateLevel or I.StatusFlag <> 0) " & _
                             "union " & _
                             "select Distinct STO.OfferID as OfferID from CM_ST_Offers STO with (NoLock) " & _
                             "inner join Offers O with (NoLock) on O.OfferID = STO.OfferID where STO.Deleted=0 and O.Deleted=0 " & _
                             "and (getdate() between STO.ProdStartDate and DateAdd(d, 1, STO.ProdEndDate)) and (O.UpdateLevel <> STO.UpdateLevel or O.StatusFlag <> 0)) and isnull(AOLV.InboundCRMEngineID,0) = 0 "

    ElseIf (FilterOffer = "2") Then
        If (BannersEnabled) Then
            If m_isOAWEnabled Then
                MyCommon.QueryStr = "from AllActiveOffersListView AOLV with (NoLock) " &
                        "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID and AOLV.EngineID<>6 " &
                        "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " &
                        " inner join CPE_Incentives CI on AOLV.OfferID = CI.IncentiveID " &
                        " where CI.StatusFlag=0 and AOLV.IsTemplate=0 and isnull(AOLV.InboundCRMEngineID,0) = 0 and (AOLV.OfferID not in (select OfferID from BannerOffers BO with (NoLock)) " &
                        " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " &
                        "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" &
                        "                     where AUB.AdminUserID = " & AdminUserID & " and isnull(AOLV.InboundCRMEngineID,0) = 0 ) ) "
            Else
                MyCommon.QueryStr = "from AllActiveOffersListView AOLV with (NoLock) " &
                "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID and AOLV.EngineID<>6 " &
                "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " &
                " where AOLV.IsTemplate=0 and isnull(AOLV.InboundCRMEngineID,0) = 0 and (AOLV.OfferID not in (select OfferID from BannerOffers BO with (NoLock)) " &
                " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " &
                "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" &
                "                     where AUB.AdminUserID = " & AdminUserID & " and isnull(AOLV.InboundCRMEngineID,0) = 0 ) ) "
            End If
        Else
            If m_isOAWEnabled Then
                MyCommon.QueryStr = "from AllActiveOffersListView AOLV inner join CPE_Incentives CI on AOLV.OfferID = CI.IncentiveID " &
                                     " where CI.StatusFlag=0 and AOLV.IsTemplate=0 And AOLV.PromoEngine<>'CAM' and AOLV.InboundCRMEngineID=0 "
            Else
                MyCommon.QueryStr = "from AllActiveOffersListView AOLV where AOLV.IsTemplate=0 and AOLV.PromoEngine<>'CAM' and InboundCRMEngineID=0 "
            End If

        End If
        If bWorkflowActive Then
            MyCommon.QueryStr &= " and not exists (select OfferId from Offers with (NoLock) where isnull(WorkflowStatus,0) > 0 and OfferId=AOLV.OfferID) "
        End If
        If (AdvSearchSQL <> "") Then
            MyCommon.QueryStr &= AdvSearchSQL
        Else
            If (Request.QueryString("searchterms") <> "") Then
                MyCommon.QueryStr &= " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
            End If
        End If


    ElseIf (FilterOffer = "5" Or FilterOffer = "6" Or FilterOffer = "7") Then
        MyCommon.QueryStr &= " and exists (select OfferId from Offers with (NoLock) where (isnull(WorkflowStatus,0) = " & (Integer.Parse(FilterOffer) - 4) & ") and OfferId=AOLV.OfferID) "

    ElseIf (FilterOffer = "8") Then
        'filtering to only display unexpired offers that have a status of testing and development. Includes Scheduled offers which are not deployed
        MyCommon.QueryStr = "from AllOffersListviewNoCAM AOLV with (NoLock) " & sJoin
        If (BannersEnabled) Then
            MyCommon.QueryStr &= "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID and AOLV.EngineID<>6 " & _
                                "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                                "where " & wherestr & " AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and  AOLV.DeploySuccessDate is null and isnull(AOLV.InboundCRMEngineID,0) = 0  " & _
                                      " and isnull(AOLV.isTemplate,0)=0  and not exists ( Select OfferID from [AllActiveOffersListView] AAOFLV where AOLV.OfferID = AAOFLV.OfferID)  " & _
                                      " and (AOLV.OfferID not in (select OfferID from BannerOffers BO with (NoLock))  or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                                      "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                                      "                     where AUB.AdminUserID = " & AdminUserID & " and isnull(AOLV.InboundCRMEngineID,0) = 0 ) ) "
        Else
            MyCommon.QueryStr &= " where AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and AOLV.DeploySuccessDate is null and " & _
                               " isnull(AOLV.InboundCRMEngineID,0) = 0  and isnull(AOLV.isTemplate,0)=0   " & _
                               " and not exists ( Select OfferID from [AllActiveOffersListView] AAOFLV where " & wherestr & " AOLV.OfferID = AAOFLV.OfferID) "
        End If
        If bWorkflowActive Then
            MyCommon.QueryStr &= " and not exists (select OfferId from Offers with (NoLock) where isnull(WorkflowStatus,0) > 0 and OfferId=AOLV.OfferID) "
        End If
        If (AdvSearchSQL <> "") Then
            MyCommon.QueryStr &= AdvSearchSQL
        Else
            If (Request.QueryString("searchterms") <> "") Then
                MyCommon.QueryStr &= " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
            End If
        End If
    End If
    'If UE Engine is installed and User is assoiated with any Buyer and if user is not having View Offer Regardless of Buyer Permission, list User-Buyer specific Offers
    If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer)) Then
        MyCommon.QueryStr = MyCommon.QueryStr & " and ( AOLV.BuyerId in (select BuyerId from BuyerRoleUsers where AdminUserID=" & AdminUserID & "))"
    End If


    'At this point, MyCommon.QueryStr contains the FROM and WHERE clauses of the query.  We need to build 2 versions of this query, one that will tell us the count of the total number of rows
    'and the second that will return the data for the sub (paginated) set of rows that we are going to return on the page
    'First we'll tack on what we need to query for the count of the total number of rows that meet the search & filter criteria  

    CountQuery = "select AOLV.OfferID " & MyCommon.QueryStr
    'Second we'll tack on what we need to query for the subset of data that needs to be displayed on this page
    'start by adding the names of the columns that we'll need for the page display.  This is not the completed SelectQuery, we'll add more to it later after we know the complete record count
    If BannersEnabled Then
        SelectQuery = "BAN.BannerID, BAN.Name as BannerName, AOLV.* " & MyCommon.QueryStr
    Else
        SelectQuery = "AOLV.* " & MyCommon.QueryStr
    End If


    If (Request.Form("favoriteOption") <> "0" And Request.Form("favoriteOption") <> "") Then
        If (Request.Form("favoriteOption") = "6") Then
            CountQuery = "select #temp.OfferID from #temp"
            SelectQuery = "Select distinct " & SelectQuery
            MyCommon.QueryStr = SelectQuery
            MyCommon.LRT_Execute()
        End If
    End If

    ' If advanced external search is on and searching for Offer ID
    ' Then search internal and external offers
    If bAdvancedExternalSearch And bSearchInternalAndExternal Then
        SelectQuery = SelectQuery.Replace("isnull(AOLV.InboundCRMEngineID,0) = 0", "isnull(AOLV.InboundCRMEngineID,0) > -1 ")
        CountQuery = CountQuery.Replace("isnull(AOLV.InboundCRMEngineID,0) = 0", "isnull(AOLV.InboundCRMEngineID,0) > -1 ")
    End If

    If (bEnableRestrictedAccessToUEOfferBuilder) Then
        If (Logix.UserRoles.CreateUEOffers AndAlso Logix.UserRoles.AccessTranslatedUEOffers) Then
            SelectQuery = SelectQuery & " and EngineID =" & FilterEngine & " "
            CountQuery = CountQuery & " and EngineID =" & FilterEngine & " "
        ElseIf (Logix.UserRoles.CreateUEOffers AndAlso Not Logix.UserRoles.AccessTranslatedUEOffers) Then
            SelectQuery = SelectQuery & " and EngineID =" & FilterEngine & "  and isnull(AOLV.InboundCRMEngineID,0) <> 10 "
            CountQuery = CountQuery & " and EngineID =" & FilterEngine & " and isnull(AOLV.InboundCRMEngineID,0) <> 10 "
        ElseIf (Not Logix.UserRoles.CreateUEOffers AndAlso Logix.UserRoles.AccessTranslatedUEOffers) Then
            If (FilterEngine = 0) Then
                SelectQuery = SelectQuery & " and EngineID = " & FilterEngine & " "
                CountQuery = CountQuery & " and EngineID = " & FilterEngine & " "
            ElseIf (FilterEngine = 9) Then
                SelectQuery = SelectQuery & " and EngineID = " & FilterEngine & " and isnull(AOLV.InboundCRMEngineID,0) = 10 "
                CountQuery = CountQuery & " and EngineID = " & FilterEngine & " and isnull(AOLV.InboundCRMEngineID,0) = 10 "
            End If
        ElseIf (Not Logix.UserRoles.CreateUEOffers AndAlso Not Logix.UserRoles.AccessTranslatedUEOffers) Then
            SelectQuery = SelectQuery & " and EngineID = 0  and isnull(AOLV.InboundCRMEngineID,0) <> 10 "
            CountQuery = CountQuery & " and EngineID = 0  and isnull(AOLV.InboundCRMEngineID,0) <> 10 "
        End If
    End If

    'before we run the CountQuery or the SelectQuery, we need to see if we are doing an export to Excel
    If (Request.QueryString("excel") <> "") Then
        MyCommon.QueryStr = "select " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
        dst = MyCommon.LRT_Select
        infoMessage = ExportListToExcel(dst, MyCommon, Logix)
        If infoMessage = "" Then
            GoTo done
        End If
    End If

    'Run the Count Query to determine the total number of rows that meet the search & filter criteria
    MyCommon.QueryStr = CountQuery
    dstOffersUniqueUserHasRights = MyCommon.LRT_Select

    If dstOffersUniqueUserHasRights.Rows.Count > 0 Then
        'Move into a SortedList to remove duplicate OfferIDs so you have a unique list and duplicates because of the BannerOffers join do not appear.
        OffersUniqueUserHasRightsToList.Clear()
        ListCouter = 0
        sizeOfData = 0
        For Each OffersUniqueUserHasRightsRows In dstOffersUniqueUserHasRights.Rows
            If (Not OffersUniqueUserHasRightsToList.ContainsKey(OffersUniqueUserHasRightsRows.Item("OfferID"))) Then
                ListCouter += 1
                OffersUniqueUserHasRightsToList.Add(OffersUniqueUserHasRightsRows.Item("OfferID"), ListCouter)
            End If
        Next
        sizeOfData = OffersUniqueUserHasRightsToList.Count
    End If

    If (sizeOfData - linesPerPage) < 0 Then
        StartPoint = 1
    Else
        StartPoint = (linesPerPage * PageNum) + 1
    End If
    EndPoint = sizeOfData
    dstOffersUniqueUserHasRights = Nothing
    dst = Nothing
    SelectQueryOrderBy = ""

    'Now that we know the total record count and updated that value in 'sizeOfData', we can determine how we should slice up the SelectQuery. 
    If (Request.Form("favoriteOption") <> "0" And Request.Form("favoriteOption") <> "") Then
        If (Request.Form("favoriteOption") = "6") Then
            FavoriteOption6 = True
        Else
            Send("<!-- building normal query -->")
            SelectQuery1 = "Select DISTINCT " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
        End If
    Else
        Send("<!-- building normal query -->")
        ' query for all results for all pages      
        SelectQuery1 = "Select DISTINCT " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
    End If


    'Run the query that returns the subset of data to be displayed on this page 
    If SelectQuery1 <> "" Then
        'If SelectQuery1 not empty then we know we need to limit the record set described in the SelectQuery to the number of records that can be displayed on this page
        OffersUniqueUserHasRightsToList.Clear()
        ListCouter = 0
        MyCommon.QueryStr = SelectQuery1
        dstOffersUniqueUserHasRights = MyCommon.LRT_Select
        If dstOffersUniqueUserHasRights.Rows.Count > 0 Then
            'Move into a SortedList to remove duplicate OfferIDs so you have a unique list and duplicates because of the BannerOffers join do not apprear.
            For Each OffersUniqueUserHasRightsRows In dstOffersUniqueUserHasRights.Rows
                If (Not OffersUniqueUserHasRightsToList.ContainsKey(OffersUniqueUserHasRightsRows.Item("OfferID"))) Then
                    ListCouter += 1
                    OffersUniqueUserHasRightsToList.Add(OffersUniqueUserHasRightsRows.Item("OfferID"), ListCouter)
                End If
            Next
        End If
        EndPage = StartPoint + linesPerPage - 1
        If EndPoint < EndPage Then EndPage = EndPoint
        OffersInPage = ""
        For pageIndex As Long = StartPoint To EndPage
            If (OffersInPage.Length > 0) Then
                OffersInPage = OffersInPage & ", " & OffersUniqueUserHasRightsToList.GetKey(OffersUniqueUserHasRightsToList.IndexOfValue(pageIndex)).ToString
            Else
                OffersInPage = OffersUniqueUserHasRightsToList.GetKey(OffersUniqueUserHasRightsToList.IndexOfValue(pageIndex)).ToString
            End If
        Next pageIndex
        dstOffersUniqueUserHasRights = Nothing
    End If

    If FavoriteOption6 Then
        ' Original Query:  SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by #temp.OfferID" & " " & SortDirection & ") as RowNumber, #temp.* from #temp" & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
        If (OffersInPage.Length > 0) Then
            SelectQuery = "select * from ( select #temp.* from #temp" & " ) as Table1 where #temp.OfferID in (" & OffersInPage & ")"
        Else
            SelectQuery = "select * from ( select #temp.* from #temp" & " ) as Table1"
        End If
    Else
        'Original Query:  SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
        If (OffersInPage.Length > 0) Then
            SelectQuery = "select * from ( select " & SelectQuery & " ) as Table1 where OfferID in (" & OffersInPage & ")"
        Else
            SelectQuery = "select * from (select  " & SelectQuery & " ) as Table1"
        End If
    End If
    MyCommon.QueryStr = SelectQuery
    'Send("<!-- Query=" & MyCommon.QueryStr & " -->")
    'Response.End()
    dst = MyCommon.LRT_Select
    If (BannersEnabled) Then
        dst = ConsolidateBanners(dst, SortText, SortDirection, MyCommon)
    Else
        dst = SortNoBanners(dst, SortText, SortDirection, MyCommon)
    End If

    MyCommon.QueryStr = "IF EXISTS (select  * from tempdb.dbo.sysobjects o where o.xtype in ('U') and o.id = object_id(N'tempdb..#temp')) begin drop table #temp end"
    MyCommon.LRT_Execute()


    If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
        If (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "CPE") Then
            sSummaryPage = "CPEoffer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
        ElseIf (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "Website") Then
            sSummaryPage = "web-offer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
        ElseIf (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "CAM") Then
            sSummaryPage = "CAM/CAM-offer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
        ElseIf (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "UE") Then
            sSummaryPage = "UE/UEoffer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
        Else
            sSummaryPage = "offer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
        End If
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", sSummaryPage)
    End If
    i = linesPerPage * PageNum


%>
<div id="intro">
    <h1 id="title">
        <%
            If CustomerInquiry Then
                Sendb(Copient.PhraseLib.Lookup("customer-inquiry.offerfavorites", LanguageID))
            Else
                Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))
            End If
        %>
    </h1>
    <div id="controls">
        <%
            If CustomerInquiry Then
                Send("<form id=""controlsform"" name=""controlsform"" action=""#"">")
                Send("</form>")
            Else
                If dst.Rows.Count > 0 Then
                    Send_ExportToExcel()
                End If
                If (Logix.UserRoles.ImportOffer) Then
                    Send_Import()
                End If
                Send("<form id=""controlsform"" name=""controlsform"" action=""offer-new.aspx"">")
                If (Logix.UserRoles.CreateOfferFromBlank) Then
                    Send_New()
                End If
                Send("</form>")
            End If
        %>
    </div>
</div>
<%
    If Request.Browser.Type = "IE6" Then
        IE6ScrollFix = " onscroll=""javascript:document.getElementById('importer').style.display='none';document.getElementById('importeriframe').style.display='none';"""
    End If
%>
<div id="main" <% Sendb(IE6ScrollFix) %>>
    <%
        If (infoMessage <> "") Then
            Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
        End If
        If (Application("MassOperaionStaus") IsNot Nothing AndAlso Application("MassOperaionStaus").ToString() <> String.Empty) Then
            Dim status = Application("MassOperaionStaus").ToString()
            'If status.StartsWith(FOLDER_NOT_IN_USE) Then
            status = status.Substring(status.LastIndexOf("~") + 1)
            If status <> String.Empty Then
                If status.Contains("Failed") OrElse status.Contains("Exception") OrElse status.Contains("Error") Then
                    Send("<div id=""FolderStatus"" style=""color: whitesmoke;"" class=""red-background"">")
                Else
                    Send("<div id=""FolderStatus"" style=""color: whitesmoke;"" class=""green-background"">")
                End If
                Send("<img src=""..\images\desktop\window\close-on.png""  id=""statusClose""/>")
                Send("<p>" & status & "</p>")
                Send("</div>")
                'End If
            End If
        End If
        If CustomerInquiry Then
            Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, ShowExpired, , , AdminUserID)
        Else
            Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, ShowExpired, , , AdminUserID)
        End If
        If (CriteriaMsg <> "") Then
            Session.Add("AdvSearchquery", SelectQuery1.ToString())
            Dim AdvSearchq As String = "True"
            Send("<div id=""criteriabar""" & IIf(CriteriaError, " style=""background-color:red;""", "") & ">" & CriteriaMsg & "<a href=""javascript:editSearchCriteria();"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "]</a><a href=""offer-list.aspx" & IIf(CustomerInquiry, "?CustomerInquiry=1", "") & """ class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.clear", LanguageID) & "]</a><a id= ""hrefAdvSearchq"" href= ""XMLFeeds.aspx?AdvSearchQuery=" & AdvSearchq & "&amp;height=400&amp;width=600 "" title=""Offers"" class=""thickbox"" style=""padding-left:15px;"" onclick=""javascript:reInitializeSelectedItems();"" >[Actions]</a></div>")
      
        End If
    %>
    <%
        If (FilterOffer = "4") Then
            Send("<br />")
        End If
    %>
    <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID)) %>">
        <thead>
            <tr>
                <% If CustomerInquiry Then%>
                <th align="left" class="th-xid" scope="col" id="XIDcol">
                    <% Sendb(Copient.PhraseLib.Lookup("term.favorite", LanguageID))%>
                </th>
                <% Else%>
                <th align="left" class="th-xid" scope="col" id="XIDcol">
                    <a id="xidLink" onclick="handleIter('xidLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=ExtOfferID&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID))%>
                    </a>
                    <%
                        If SortText = "ExtOfferID" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <% End If%>
                <th align="left" class="th-bigid" scope="col">
                    <a id="idLink" onclick="handleIter('idLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=AOLV.OfferID&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
                    </a>
                    <%
                        If SortText = "AOLV.OfferID" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-engine" scope="col">
                    <a id="engineLink" onclick="handleIter('engineLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=PromoEngine&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID))%>
                    </a>
                    <%
                        If SortText = "PromoEngine" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <%--buyer id display only when sysopt169 is enabled--%>
                <%  If (MyCommon.Fetch_UE_SystemOption(169) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
                <th align="left" class="th-name" scope="col">
                    <a id="buyeridLink" onclick="handleIter('idLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=AOLV.BuyerId&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.buyerid", LanguageID))%>
                    </a>
                    <%
                        If SortText = "AOLV.BuyerId" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <%End If%>
                <th align="left" class="th-name" scope="col">
                    <a id="nameLink" onclick="handleIter('nameLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=AOLV.Name&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
                    </a>
                    <%
                        If SortText = "AOLV.Name" OrElse SortText = "Name" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-date" scope="col" style="display: none;">
                    <a id="createLink" onclick="handleIter('createLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=CreatedDate&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
                    </a>
                    <%
                        If SortText = "CreatedDate" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-date" scope="col">
                    <a id="startLink" onclick="handleIter('startLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=ProdStartDate&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.starts", LanguageID))%>
                    </a>
                    <%
                        If SortText = "ProdStartDate" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-date" scope="col">
                    <a id="endLink" onclick="handleIter('endLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=ProdEndDate&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.ends", LanguageID))%>
                    </a>
                    <%
                        If SortText = "ProdEndDate" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-status" scope="col">
                    <a id="statusLink" onclick="handleIter('statusLink');" href="offer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=AOLV.StatusFlag&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine=" & FilterEngine, ""))%>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%></a>
                    <%
                        If Request.Params("SortText") = "AOLV.StatusFlag" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
            </tr>
        </thead>
        <tbody>
            <%
                Dim Statuses As New Hashtable(20)
                Dim OfferStatus As Copient.LogixInc.STATUS_FLAGS
                Dim arrList As New ArrayList(20)
                Dim OfferList() As String
                Dim j As Integer = 0
                Dim StartIdx As Integer = 0
                Dim Pagelength As Integer = dst.Rows.Count - 1

                For j = 0 To (dst.Rows.Count - 1)
                    arrList.Add(dst.Rows(j).Item("OfferID"))
                Next
                arrList.TrimToSize()
                ReDim OfferList(arrList.Count - 1)
                For j = 0 To arrList.Count - 1
                    OfferList(j) = arrList.Item(j).ToString
                Next
                Statuses = Logix.GetStatusForOffers(OfferList, LanguageID)
                If (Request.Params("SortText") = "AOLV.StatusFlag") Then
                    dst = Logix.SortOfferStatuses(dst, Statuses, SortDirection)
                    Pagelength = linesPerPage + (linesPerPage * PageNum) - 1
                    If Pagelength > dst.Rows.Count Then
                        Pagelength = dst.Rows.Count - 1
                    End If
                End If
                'While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)

                For i = StartIdx To Pagelength Step 1
                    If Statuses.Contains(dst.Rows(i).Item("OfferID").ToString) Then
                        OfferStatus = Statuses.Item(dst.Rows(i).Item("OfferID").ToString)
                    Else
                        OfferStatus = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
                    End If

                    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
                        SearchMatchesROID = (idNumber = MyCommon.NZ(dst.Rows(i).Item("RewardOptionID"), -2))
                    Else
                        SearchMatchesROID = False
                    End If
                    RoidExtension = IIf(SearchMatchesROID, "&nbsp;&nbsp;<span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.roid", LanguageID) & ": " & MyCommon.NZ(dst.Rows(i).Item("RewardOptionID") & ")</span>", ""), "")
                    Send("<tr class=""" & Shaded & """>")
                    If CustomerInquiry Then
                        MyCommon.QueryStr = "select AdminUserID from AdminUserOffers where OfferID=" & dst.Rows(i).Item("OfferID") & ";"
                        rst = MyCommon.LRT_Select
                        Send("  <td style=""text-align:center;""><a href=""javascript:openPopup('offer-favorite.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & "&amp;CustomerInquiry=1')"">" & rst.Rows.Count & "/" & TotalUsers & "</a></td>")
                    Else
                        Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("ExtOfferID"), "&nbsp;"), 9, "<br />") & "</td>")
                    End If
                    Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("OfferID"), 0) & "</td>")
                    Send("  <td>" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "")) & "</td>")
                    'Send("  <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0), LanguageID) & "</td>")
                    If (MyCommon.Fetch_UE_SystemOption(169) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                        'Dim externalbuyerid As String = MyCommon.GetExternalBuyerId(MyCommon.NZ(dst.Rows(i).Item("BuyerID"), ""))
                        Send(" <td>" & MyCommon.NZ(dst.Rows(i).Item("ExternalBuyerId"), "") & "</td>")
                    End If

                    Send("  <td class=""divideMe"">")
                    If restrictLinks Then
                        Sendb(MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension)
                    Else
                        If (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CPE") Then
                            Sendb("  <a href=""CPEoffer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & dst.Rows(i).Item("Name") & RoidExtension & "</a>")
                        ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Website") Then
                            Sendb("  <a href=""web-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & dst.Rows(i).Item("Name") & RoidExtension & "</a>")
                        ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Email") Then
                            Sendb("  <a href=""email-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & dst.Rows(i).Item("Name") & RoidExtension & "</a>")
                        ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CAM") Then
                            Sendb("  <a href=""CAM/CAM-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & dst.Rows(i).Item("Name") & RoidExtension & "</a>")
                        ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "UE") Then
                            Sendb("  <a href=""UE/UEoffer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & dst.Rows(i).Item("Name") & RoidExtension & "</a>")
                        Else
                            Sendb("  <a href=""offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & dst.Rows(i).Item("Name") & RoidExtension & "</a>")
                        End If
                    End If
                    If (MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) > 0 AndAlso MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) <= 10 AndAlso MyCommon.NZ(dst.Rows(i).Item("UpdateLevel"), 0) > 0 AndAlso CustomerInquiry = False) Then
                        Send("<br /><span class=""red"" style=""font-size:10px;font-weight:bold;"">(" & Copient.PhraseLib.Lookup("alert.offermodified", LanguageID) & ")</span></td>")
                    ElseIf MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) = 11 Or MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) = 12 Then
                        Send("<br /><span class=""green"" style=""font-size:10px;font-weight:bold;"">(" & Copient.PhraseLib.Lookup("alert.awaitingrecommendation", LanguageID) & ")</span></td>")
                    ElseIf MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) = 13 Or MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) = 14 Or MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) = 15 Then
                        Send("<br /><span class=""green"" style=""font-size:10px;font-weight:bold;"">(" & Copient.PhraseLib.Lookup("alert.awaitingapproval", LanguageID) & ")</span></td>")
                    Else
                        Send("</td>")
                    End If
                    If (Not IsDBNull(dst.Rows(i).Item("CreatedDate"))) Then
                        Send("  <td style=""display:none;"">" & Logix.ToShortDateString(dst.Rows(i).Item("CreatedDate"), MyCommon) & "</td>")
                    Else
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    End If
                    If (Not IsDBNull(dst.Rows(i).Item("ProdStartDate"))) And (dst.Rows(i).Item("ProdStartDate") > New Date(1900, 1, 1)) Then
                        Send("  <td>" & Logix.ToShortDateString(dst.Rows(i).Item("ProdStartDate"), MyCommon) & "</td>")
                    Else
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
                    End If
                    If (Not IsDBNull(dst.Rows(i).Item("ProdEndDate"))) And (dst.Rows(i).Item("ProdStartDate") > New Date(1900, 1, 1)) Then
                        Send("  <td>" & Logix.ToShortDateString(dst.Rows(i).Item("ProdEndDate"), MyCommon) & "</td>")
                    Else
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
                    End If
                    Send("  <td>" & Logix.GetOfferStatusHtml(OfferStatus, LanguageID) & "</td>")
                    Send("</tr>")
                    If Shaded = "shaded" Then
                        Shaded = ""
                    Else
                        Shaded = "shaded"
                    End If

                Next

            %>
        </tbody>
    </table>
</div>
<script runat="server">
    Protected Overrides Sub OnPreInit(e As EventArgs)
        MyBase.OnPreInit(e)
        CurrentRequest.Resolver.AppName = "offer-list.aspx"
        Dim xssEncoding As IXssEncoding = CMS.AMS.CurrentRequest.Resolver.Resolve(Of IXssEncoding)()
        
        xssEncoding.EncodeInputParams(Request)
    End Sub
    Dim sErrorMsg As String
    Private Function hasOption(ByRef optionName As String) As Boolean
    
        Dim val As String = Request.Form(optionName)
        Return val IsNot Nothing AndAlso val.Trim().Length > 0
        
    End Function
    Public Function GetErrorMsg() As String
        Return sErrorMsg
    End Function
    Function IsValidXml(ByRef Common As Copient.CommonIncConfigurable, ByVal LogFile As String, _
                                ByVal sXsdFileName As String, ByVal sXml As String) As Boolean

        Dim Settings As XmlReaderSettings
        Dim xr As XmlReader = Nothing
        Dim bValid As Boolean = True
        Dim sr As StringReader = Nothing
        Try
            Settings = New XmlReaderSettings()
            Settings.Schemas.Add(Nothing, sXsdFileName)
            Settings.ValidationType = ValidationType.Schema
            Settings.IgnoreComments = True
            Settings.IgnoreProcessingInstructions = True
            Settings.IgnoreWhitespace = True
            sr = New StringReader(sXml)
            xr = XmlReader.Create(sr, Settings)
            Do While (xr.Read())
                'Console.WriteLine("NodeType: " & xr.NodeType.ToString & " - " & xr.LocalName & " Depth: " & xr.Depth.ToString)
            Loop
            bValid = True
        Catch eXmlSch As XmlSchemaException
            sErrorMsg = BuildErrorMsg("(Xml Schema Validation Error Line: " & eXmlSch.LineNumber.ToString & " - Col: " & eXmlSch.LinePosition.ToString & ") " & eXmlSch.Message, False)
            bValid = False
        Catch eXml As XmlException
            sErrorMsg = BuildErrorMsg("(Xml Error Line: " & eXml.LineNumber.ToString & " - Col: " & eXml.LinePosition.ToString & ") " & eXml.Message, False)
            bValid = False
        Catch exApp As ApplicationException
            sErrorMsg = BuildErrorMsg("Application Error: " & exApp.ToString, False)
            bValid = False
        Catch ex As Exception
            sErrorMsg = BuildErrorMsg("Error: " & ex.ToString, False)
            bValid = False
        Finally
            If Not xr Is Nothing Then
                xr.Close()
            End If
        End Try

        ' Log Error if one exists
        If (sErrorMsg <> "") Then
            Common.Write_Log(LogFile, sErrorMsg)
        End If

        Return bValid
    End Function
    Private Function BuildErrorMsg(ByVal sMsg As String, ByVal bSystemError As Boolean) As String
        Dim sErrMsg As String = ""
        sErrMsg = "(ValidateOfferXML)" & " - " & sMsg
        Return sErrMsg
    End Function
  
  
  
    Function GetOptionString(ByRef MyCommon As Copient.CommonInc, ByVal OptionIndex As Integer, _
                             ByVal OptionValue As String, ByVal FieldName As String) As String
        Dim FieldBuf As New StringBuilder()
        FieldBuf.Append(FieldName & " ")
        Select Case OptionIndex
            Case 1 ' contains
                FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
            Case 2 ' exact
                FieldBuf.Append(" = '" & MyCommon.Parse_Quotes(OptionValue.Trim) & "' ")
            Case 3 ' starts with
                FieldBuf.Append(" like '" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
            Case 4 ' ends with
                FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "' ")
            Case 5 ' excludes
                FieldBuf = New StringBuilder()
                FieldBuf.Append(" (" & FieldName & " not like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' or " & FieldName & " is null) ")
            Case 6 ' is
                FieldBuf.Append(" = " & MyCommon.Parse_Quotes(OptionValue.Trim) & " ")
            Case 7 ' is not
                FieldBuf.Append(" <> " & MyCommon.Parse_Quotes(OptionValue.Trim) & " ")
            Case Else ' default to contains
                FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
        End Select
        Return FieldBuf.ToString
    End Function
  
    Function GetOptionType(ByVal OptionIndex As Integer) As String
        Dim OptionType As String = "contains"
        Select Case OptionIndex
            Case 1 ' contains
                OptionType = Copient.PhraseLib.Lookup("term.contains", LanguageID)
            Case 2 ' exact
                OptionType = "="
            Case 3 ' starts with
                OptionType = Copient.PhraseLib.Lookup("term.startswith", LanguageID)
            Case 4 ' ends with
                OptionType = Copient.PhraseLib.Lookup("term.endswith", LanguageID)
            Case 5 ' excludes
                OptionType = Copient.PhraseLib.Lookup("term.excludes", LanguageID)
            Case 6 ' is
                OptionType = Copient.PhraseLib.Lookup("term.is", LanguageID)
            Case 7 ' is not
                OptionType = Copient.PhraseLib.Lookup("term.IsNot", LanguageID)
            Case Else ' default to contains
                OptionType = Copient.PhraseLib.Lookup("term.contains", LanguageID)
        End Select
        Return OptionType
    End Function
  
    Function GetPriorityOptionType(ByVal OptionIndex As Integer) As String
        Dim OptionType As String = "contains"
    
        Select Case OptionIndex
            Case 1
                OptionType = "="
            Case 2
                OptionType = "<="
            Case 3
                OptionType = ">="
        End Select
    
        Return OptionType
    End Function
  
    Function GetDateOption(ByRef MyCommon As Copient.CommonInc, ByVal OptionIndex As Integer, _
                           ByVal StartValue As String, ByVal EndValue As String, ByVal FieldName As String) As String
        Dim StartDate, EndDate As Date
        Dim FieldBuf As New StringBuilder()
    
        If (TryParseLocalizedDate(StartValue, StartDate, MyCommon) AndAlso OptionIndex <> 3) _
        OrElse (TryParseLocalizedDate(StartValue, StartDate, MyCommon) AndAlso TryParseLocalizedDate(EndValue, EndDate, MyCommon)) Then
            FieldBuf.Append(FieldName & " ")
            Select Case OptionIndex
                Case 0 ' on
                    FieldBuf.Append(" between '" & StartDate.ToString("yyyy-MM-ddT00:00:00") & "' and '" & StartDate.ToString("yyyy-MM-ddT23:59:59") & "' ")
                Case 1 ' before
                    FieldBuf.Append(" < '" & StartDate.ToShortDateString() & "' ")
                Case 2 ' after
                    FieldBuf.Append(" > '" & StartDate.ToShortDateString() & "' ")
                Case 3 ' between
                    FieldBuf.Append(" between '" & StartDate.ToShortDateString() & "' and '" & EndDate.ToShortDateString() & "' ")
                Case Else ' default to after
                    FieldBuf.Append(" > '" & StartDate.ToShortDateString() & "' ")
            End Select
        Else
            Throw New ApplicationException(Copient.PhraseLib.Lookup("term.invaliddateformat", LanguageID) & " (" & MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern & ").<br />")
        End If
    
        Return FieldBuf.ToString
    End Function
  
    Function TryParseLocalizedDate(ByVal DateStr As String, ByRef LocalizedDate As Date, ByRef MyCommon As Copient.CommonInc) As Boolean
        Return Date.TryParseExact(DateStr, MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern, _
                               MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, LocalizedDate)
    End Function
  
    Function GetDateOptionType(ByVal OptionIndex As Integer) As String
        Dim OptionType As String = "after"
        Select Case OptionIndex
            Case 0 ' on
                OptionType = Copient.PhraseLib.Lookup("term.on", LanguageID)
            Case 1 ' before
                OptionType = Copient.PhraseLib.Lookup("term.before", LanguageID)
            Case 2 ' after
                OptionType = Copient.PhraseLib.Lookup("term.after", LanguageID)
            Case 3 ' between
                OptionType = Copient.PhraseLib.Lookup("term.between", LanguageID)
            Case Else ' default to after
                OptionType = Copient.PhraseLib.Lookup("term.after", LanguageID)
        End Select
        Return OptionType
    End Function
  
    Function GetProductOfferList(ByRef MyCommon As Copient.CommonInc, ByVal ExtProductID As String, ByVal MatchType As Integer) As String
        Dim OfferListBuf As New StringBuilder("-1")
        Dim PadLen As Integer = 0
        Dim dt As DataTable
        Dim dt1 As DataTable
    
        If ExtProductID Is Nothing Then ExtProductID = ""
        ExtProductID = ExtProductID.Trim

        'get the padding length of the UPC code
        MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
        dt1 = MyCommon.LRT_Select
        If dt1 IsNot Nothing Then
            PadLen = Convert.ToInt32(dt1.Rows(0).Item("paddingLength"))
        End If
        ' pad if this is an exact match type (i.e. 2)
        If MatchType = 2 OrElse MatchType = 5 Then
            'Integer.TryParse(MyCommon.Fetch_SystemOption(52), PadLen)
            If PadLen > 0 Then
                ExtProductID = ExtProductID.PadLeft(PadLen, "0")
            End If
        End If
    
        ' get all the OfferID matching the Product
        MyCommon.QueryStr = "dbo.pa_GetOffersForProduct"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ExtProductID", SqlDbType.NVarChar, 120).Value = ExtProductID
        MyCommon.LRTsp.Parameters.Add("@MatchType", SqlDbType.Int).Value = MatchType
        dt = MyCommon.LRTsp_select()
        MyCommon.Close_LRTsp()

        For Each row As DataRow In dt.Rows
            OfferListBuf.Append(", " & MyCommon.NZ(row.Item("OfferID"), "-1"))
        Next
     
        Return OfferListBuf.ToString
    End Function
  
    Function GetExternalSourceOfferList(ByRef MyCommon As Copient.CommonInc) As String
        Dim ExternalSourceList As New StringBuilder("-1")
        Dim dt As DataTable
             
        MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "select ExtInterfaceID from ExtCRMInterfaces where ExtInterfaceID > 0 ;"
        dt = MyCommon.LRT_Select
       
        For Each row As DataRow In dt.Rows
            ExternalSourceList.Append(", " & MyCommon.NZ(row.Item("ExtInterfaceID"), "-1"))
        Next
     
        Return ExternalSourceList.ToString
    End Function

    Function GetMCLUOfferList(ByRef MyCommon As Copient.CommonInc, ByVal mclu As String, ByVal MatchType As Integer) As String
        Dim OfferListBuf As New StringBuilder("-1")
        Dim dt As DataTable
        Dim conditionQuery As String
    
        If mclu Is Nothing Then mclu = ""
        mclu = mclu.Trim

        MyCommon.Open_LogixRT()
        conditionQuery = GetOptionString(MyCommon, MatchType, mclu, "RXT.XmlText")

        MyCommon.QueryStr = "select distinct O.OfferID from RewardXmlTiers as RXT with (NoLock) " & _
                            "inner join OfferRewards as ORW with (NoLock) on ORW.RewardID=RXT.RewardID " & _
                            "inner join Offers as O with (NoLock) on O.OfferID=ORW.OfferID " & _
                            "where O.Deleted=0 and ORW.Deleted=0 and ORW.RewardTypeID=8 and " & conditionQuery & ";"
        dt = MyCommon.LRT_Select
       
        For Each row As DataRow In dt.Rows
            OfferListBuf.Append(", " & MyCommon.NZ(row.Item("OfferID"), "-1"))
        Next
     
        Return OfferListBuf.ToString
    End Function
  
    Function GetPriorityOfferList(ByRef MyCommon As Copient.CommonInc, ByVal Priority As String, ByVal PriorityOperator As Integer) As String
        Dim OfferListBuf As New StringBuilder("-1")
        Dim OperatorStr As String = "="
        Dim dt As DataTable
        
        Select Case PriorityOperator
            Case 2
                OperatorStr = "<="
            Case 3
                OperatorStr = ">="
        End Select
    
        MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where PriorityLevel " & OperatorStr & MyCommon.Extract_Val(Priority) & ";"
        dt = MyCommon.LRT_Select
    
        For Each row As DataRow In dt.Rows
            OfferListBuf.Append(", " & MyCommon.NZ(row.Item("OfferID"), "-1"))
        Next
     
        Return OfferListBuf.ToString
    End Function
  
    Sub AssignOfferToBanners(ByVal OfferID As Integer, ByRef MyCommon As Copient.CommonInc)
        Dim BannerIDs As String() = Nothing
        Dim BannerID As Integer
        Dim i As Integer
    
        If (Request.Form("banner") <> "") Then
            BannerIDs = Request.Form.GetValues("banner")
            If (BannerIDs IsNot Nothing AndAlso BannerIDs.Length > 0) Then
                For i = 0 To BannerIDs.GetUpperBound(0)
                    BannerID = MyCommon.Extract_Val(BannerIDs(i))
                    MyCommon.QueryStr = "insert into BannerOffers (BannerID, OfferID) values (" & BannerID & "," & OfferID & ");"
                    MyCommon.LRT_Execute()
                Next
            End If
        End If
        If (Request.Form("allbannersid") <> "") Then
            BannerIDs = Request.Form.GetValues("allbannersid")
            If (BannerIDs IsNot Nothing AndAlso BannerIDs.Length > 0) Then
                For i = 0 To BannerIDs.GetUpperBound(0)
                    BannerID = MyCommon.Extract_Val(BannerIDs(i))
                    MyCommon.QueryStr = "insert into BannerOffers (BannerID, OfferID) values (" & BannerID & "," & OfferID & ");"
                    MyCommon.LRT_Execute()
                Next
            End If
        End If

    End Sub
  
    Function GetBanners() As Integer()
        Dim Banners() As Integer = Nothing
        Dim BannerIDs() As String
        Dim BannerArray As New ArrayList(5)
        Dim BannerID As Integer
        Dim i As Integer
    
        If (Request.Form("banner") <> "") Then
            BannerIDs = Request.Form.GetValues("banner")
            If (BannerIDs IsNot Nothing AndAlso BannerIDs.Length > 0) Then
                For i = 0 To BannerIDs.GetUpperBound(0)
                    Integer.TryParse(BannerIDs(i), BannerID)
                    If (BannerID > 0) Then BannerArray.Add(BannerID)
                Next
            End If
        End If
        If (Request.Form("allbannersid") <> "") Then
            BannerIDs = Request.Form.GetValues("allbannersid")
            If (BannerIDs IsNot Nothing AndAlso BannerIDs.Length > 0) Then
                Integer.TryParse(BannerIDs(i), BannerID)
                If (BannerID > 0) Then BannerArray.Add(BannerID)
            End If
        End If
        If (BannerArray IsNot Nothing AndAlso BannerArray.Count > 0) Then
            ReDim Banners(BannerArray.Count - 1)
            BannerArray.CopyTo(Banners)
        End If
    
        Return Banners
    End Function
 
</script>
<form id="frmIter" name="frmIter" method="post" action="">
<input type="hidden" id="advSql" name="advSql" value="<% Sendb(Server.UrlEncode(AdvSearchSQL)) %>" />
<input type="hidden" id="advCrit" name="advCrit" value="<% Sendb(Server.UrlEncode(CriteriaMsg)) %>" />
<input type="hidden" id="advTokens" name="advTokens" value="<%Sendb(Server.UrlEncode(CriteriaTokens)) %>" />
</form>
<!-- overwrite the iteration links and post the form -->
<script type="text/javascript">
  var followLink = true;
  if (document.getElementById('firstPageLink') != null) { document.getElementById('firstPageLink').onclick = handleFirst; }
  if (document.getElementById('previousPageLink') != null) { document.getElementById('previousPageLink').onclick = handlePrev; }
  if (document.getElementById('nextPageLink') != null) { document.getElementById('nextPageLink').onclick = handleNext; }
  if (document.getElementById('lastPageLink') != null) { document.getElementById('lastPageLink').onclick = handleLast; }

// for closing the perform action div on click of Esc button and outside for all browsers
  map = {}
  keydown = function (e) {
      e = e || event
      map[e.keyCode] = true
      if (map[27]) {//Esc
	      //toggleDialogOfferDuplicate('DuplicateNoofOffer', false);
          toggleDialog("performactions", false);
          map = {}
          return false      
      }
  }
  keyup = function (e) {
      e = e || event
      map[e.keyCode] = false
  }
  onkeydown = keydown
  onkeyup = keyup//For Regular browsers
  try {//for IE
      document.attachEvent('onkeydown', keydown)
      document.attachEvent('onkeyup', keyup)      
  } catch (e) {
  
  }
  ///

  //Constatns

  var followLink = true;
  var selectedItems = new Array();
  var DUP_OFFERS = 12;
  var MASSDEPLOY_OFFERS = 13;
  var NAVIGATETO_REPORTS = 14;
  var SEND_OUTBOUND = 15;
  var WFSTAT_PREVALIDATE = 16;
  var WFSTAT_POSTVALIDATE = 17;
  var WFSTAT_READYTODEPLOY = 18;
  var TRANSFER_OFFERS = 20;  
	var SHOW_APPROVALACTIONITEMS = 21
  var MASSREQUESTAPPROVALFOROFFERS = 22;
  var DEPLOYANDAPPROVAL_REQUEST = 23;
  function handleFirst() {
    if (followLink) {
      followLink = false;
      handleIter('firstPageLink');
    } else {
      return false;
    }
  }

  function handlePrev() {
    if (followLink) {
      followLink = false;
      handleIter('previousPageLink');
    } else {
      return false;
    }
  }

  function handleNext() {
    if (followLink) {
      followLink = false;
      handleIter('nextPageLink');
    } else {
      return false;
    }
  }

  function handleLast() {
    if (followLink) {
      followLink = false;
      handleIter('lastPageLink');
    } else {
      return false;
    }
  }

  function handleIter(elemName) {
    var elem = document.getElementById(elemName);
    var frm = document.frmIter;

    if (elem != null && frm != null) {
      frm.action = elem.href;
      elem.href = "javascript:navTo();";
    }
  }

  function navTo() {
    document.frmIter.submit();
  }

  function handleFilterRegEx(newFilter) {
    var frm = document.frmIter;
    var elemAdv = frm.advSql;
    var currentURL = window.location.href;
    var newURL = "";

    if (elemAdv != null && elemAdv.value != "") {
      if (currentURL.indexOf('filterOffer=') > -1) {
        newURL = currentURL.replace(/filterOffer=[0-9]?/g, 'filterOffer=' + newFilter);
        newURL = newURL.replace(/pagenum=[0-9]+/g, '');
        if(document.getElementById('filterengine')!=null && currentURL.indexOf('filterengine=') > -1)
        {
         newURL = newURL.replace(/filterengine=[0-9]?/g, 'filterengine=' + document.getElementById('filterengine').value);
        }
      } else {
        if (currentURL.indexOf("&") > -1) {
          newURL = currentURL + "&amp;filterOffer=" + newFilter;
        } else {
          newURL = currentURL + "?filterOffer=" + newFilter;
        }
        if(document.getElementById('filterengine')!=null)
        {
         newURL = newURL + "&filterengine=" +  document.getElementById('filterengine').value;
        }
      }
      frm.action = newURL;
      frm.submit();
    } else {
      if (document.getElementById("searchform") != null) { document.getElementById("searchform").submit(); }
    }
  }

  function handleUserFilterRegEx(newFilter) {
    var frm = document.frmIter;
    var elemAdv = frm.advSql;
    var currentURL = window.location.href;
    var newURL = "";

    if (elemAdv != null && elemAdv.value != "") {
      if (currentURL.indexOf('filterUser=') > -1) {
        newURL = currentURL.replace(/filterUser=[0-9]?/g, 'filterUser=' + newFilter);
        newURL = newURL.replace(/pagenum=[0-9]+/g, '');
        if(document.getElementById('filterengine')!=null && currentURL.indexOf('filterengine=') > -1)
        {
         newURL = newURL.replace(/filterengine=[0-9]?/g, 'filterengine=' + document.getElementById('filterengine').value);
        }
      } else {
        if (currentURL.indexOf("&") > -1) {
          newURL = currentURL + "&amp;filterUser=" + newFilter;
        } else {
          newURL = currentURL + "?filterUser=" + newFilter;
        }
        if(document.getElementById('filterengine')!=null)
        {
         newURL = newURL + "&filterengine=" +  document.getElementById('filterengine').value;
        }
      }
      frm.action = newURL;
      frm.submit();
    } else {
      if (document.getElementById("searchform") != null) { document.getElementById("searchform").submit(); }
    }
  }

  function handleExcel() {
    var sUrl = document.getElementById("ExcelUrl");
    var form = document.forms['excelform'];

    form.action = sUrl.value;
    form.method = "Post";
    form.submit();

}

function submitToperformaction(itemid, bChecked) {
    var index = -1;
    var i = 0;
    var iID = 0;
    var elemAll = document.getElementById("allofferIDs");
    iID = parseInt(itemid);
    if (!bChecked) {
        index = search(selectedItems, iID, false)
        if (index > -1) {
            selectedItems.splice(index, 1);
        }

    } else {
        index = search(selectedItems, iID, true)
        selectedItems.splice(index, 0, iID);

    }
    elemAll.checked = iID.checked;
   
}

function handleAllItems() {
    var ids = new Array();
    var itemlist = document.getElementById("itemlist").value;
    var ids = itemlist.split(',')
    //alert(itemlist);     
    var elem = null;
    //var ID = 0;
    var elemAll = document.getElementById("allofferIDs");
    if (elemAll != null) {
        //alert('HI');
        selectedItems = new Array();
        //alert(ids);
        for (var i = 0; i < ids.length; i++) {
            elem = document.getElementById("linkID" + ids[i]);

            elem.checked = elemAll.checked;
            if (elem.checked) { updateidlist(ids[i], elem.checked); }
            //elem = document.getElementById("itemID" + ids[i]);

        }
        
    }
}

function updateidlist(itemid, bChecked) {
    var index = -1;
    var i = 0;
    var iID = 0;
    iID = parseInt(itemid);
    if (!bChecked) {
        index = search(selectedItems, iID, false)
        if (index > -1) {
            selectedItems.splice(index, 1);
        }

    } else {
        index = search(selectedItems, iID, true)
        selectedItems.splice(index, 0, iID);

    }
    
}

function search(o, v, i) {
    /*
    vector (o):  array that will be looked up
    value (v):   object that will be searched
    insert (i):  if true, the function will return the index where the value should be inserted
    to keep the array ordered, otherwise returns the index where the value was found
    or -1 if it wasn't found
    */
    var h = o.length, l = -1, m;
    while (h - l > 1)
        if (o[m = h + l >> 1] < v) l = m;
        else h = m;
    return o[h] != v ? i ? h : -1 : h;
};

   function getElementsByIdStartsWith(container, selectorTag, prefix) {
     var items = [];
     var myPosts = document.getElementById(container).getElementsByTagName(selectorTag);
     for (var i = 0; i < myPosts.length; i++) {
        //omitting undefined null check for brevity
       if (myPosts[i].id.lastIndexOf(prefix, 0) === 0) {
         items.push(myPosts[i]);
       }
     }
     return items;
   } 

  function showactions() {
    var offerrows = [];
    
    offerrows = getElementsByIdStartsWith("tb1","tr","errdesc"); 
        
    if (selectedItems.length > 0) {
	  
      deletecells(offerrows);
      ShowRequestApprovalActionItems(selectedItems);
      toggleDialog('performactions', true);

    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }    
  }

function clearErrorContents(){
  var performActionsElem = document.getElementById("performactionserror");
  var duplicateofferdiv = document.getElementById("execduplicateoffer");

  if (performActionsElem != null) {
    performActionsElem.style.display = 'none';
  }
  if (duplicateofferdiv != null) {
    duplicateofferdiv.style.visibility = 'hidden';
  }
}

function toggleDialog(elemName, shown) {
    var elem = document.getElementById(elemName);
    //var fadeElem = document.getElementById('fadeDiv');
   var tbelem = document.getElementById('ActionsTB');

    if (elem != null) {

        elem.style.display = (shown) ? 'block' : 'none';
           toggleDisabled(tbelem,shown);
    }
}
    

function toggleDisabled(el,shown) {
    var closebtnElem = document.getElementById('TB_closeWindowButton');
    var all = el.getElementsByTagName('input')
    //all = all + el.getElementsByTagName('a')
    var inp, i=0;
    while (inp = all[i++]) {        
        inp.disabled = (shown) ? true : false;
    }
    closebtnElem.style.display = (shown) ? 'none' : 'block';

}

function ExecAction() {
    var actionele = document.getElementById('Actionitems');
    var selectedaction = "";
    if (actionele != null) {
        selectedaction = actionele.options[actionele.selectedIndex].text;
    }
    if (selectedaction != "") {
        switch (selectedaction) {

          case '<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%>':
            handlemassDeploy();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%>':
            handleReports();
            break;      
          case '<%Sendb(Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID))%>':
            SendOutBound();
            break;         
          case '<%Sendb(Copient.PhraseLib.Lookup("term.prevalidate", LanguageID))%>':
            ChangeWFStatustoPreValidate();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.postvalidate", LanguageID))%>':
            ChangeWFStatustoPostValidate();
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID))%>':
            ChangeWFStatustoReadytoDeploy();
            break; 
          case '<%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%>':
                HandleMassRequestApproval_Async(13, '<%Sendb(Copient.PhraseLib.Lookup("folders.submit", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.submitapprovalconfirm", LanguageID))%>');
            break;	
          case '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%>':
                HandleMassRequestApproval_Async(14, '<%Sendb(Copient.PhraseLib.Lookup("folders.submit", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.approvewithdeployconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deployoffers", LanguageID))%>':
                HandleRequestForMixOffers_Async(1, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.deploycomfirm", LanguageID))%> ');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%>':
                HandleRequestForMixOffers_Async(2, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.submitapprovalconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%>':
                HandleRequestForMixOffers_Async(3, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.approvewithdeployconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%>':
                HandleRequestForMixOffers_Async(5, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deploy.submitapprovalconfirm", LanguageID))%>');
            break;
          case '<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.and", LanguageID))%> ' + '<%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%>':
                HandleRequestForMixOffers_Async(6, '<%Sendb(Copient.PhraseLib.Lookup("folders.performmassaction", LanguageID))%>' + ' ' + selectedItems.length + ' ' + '<%Sendb(Copient.PhraseLib.Lookup("deploy.approvewithdeployconfirm", LanguageID))%> ');
            break;
        }
    }
}

function handleAction() {

    var elemExecAction = document.getElementById('ExecAction');
    var elemdupoffer = document.getElementById('dupoffer');
    var actionele = document.getElementById('Actionitems');
    if (actionele.value == '1' || actionele.value == '7' || actionele.value == '8'){
        elemdupoffer.style.visibility = 'visible';
        elemExecAction.style.visibility = 'hidden';
    }
    else
    {
        elemExecAction.style.visibility = 'visible';
        elemdupoffer.style.visibility = 'hidden';
    }
}

function handlemassDeployConditional(offerswithoutcon) {
    //value of max offers should be fetched from a CM_System option
    
    if (selectedItems.length > 0) {
      
        if (confirm(offerswithoutcon + ' Offers do not have any condition.This operation will happen in background. Are you sure you want to continue?')) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true +'&OffersWithoutConditions=' + true;
        //ClosePopUp("");
        xmlhttpPost('folder-feeds.aspx?Action=DeployOffers', MASSDEPLOY_OFFERS, frmdata);
        }
     }
  }
function HandleMassRequestApproval_Async(approvalType, messageText) {
    //value of max offers should be fetched from a System option 152
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;   
    if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
         if (confirm(messageText)) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&ApprovalType=' + approvalType + '&FromOfferList=' + true + '&OffersWithoutConditions=' + false;
        xmlhttpPost('folder-feeds.aspx?Action=RequestApproval', MASSREQUESTAPPROVALFOROFFERS, frmdata);
        }
      } 
    else{
        alert('You are going to perform mass action on ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        } 
      }   
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 

  }
function HandleRequestForMixOffers_Async(deploymentType, messageText) {
    //value of max offers should be fetched from a System option 152
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;   
    if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
         if (confirm(messageText)) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&DeploymentType=' + deploymentType + '&OffersWithoutConditions=' + false + '&FromOfferList=' + true;
        xmlhttpPost('folder-feeds.aspx?Action=HandleDeployAndApprovalRequest', DEPLOYANDAPPROVAL_REQUEST, frmdata);
        }
      } 
        else{
        alert('You are going to perform mass action on ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        } 
      }   
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 

  }
function handlemassDeploy() {
    //value of max offers should be fetched from a CM_System option
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;
    if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.deployofferwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + ' This operation will happen in background .' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true +'&OffersWithoutConditions=' + false;
        //ClosePopUp("");
        xmlhttpPost('folder-feeds.aspx?Action=DeployOffers', MASSDEPLOY_OFFERS, frmdata);
        }
      }      
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.deployofferwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        } 
      }    
    else{
      alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    } 

  }

  function handleReports(){     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(103)) %>;
     if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.navigatetoreportswarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + ' This operation will happen in background . ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
           var performactionserrorElem = document.getElementById("performactionserror");
                performactionserrorElem.style.display = 'none';
                toggleDialog('performactions', false);
                //window.location = 'reports-custom.aspx';
          xmlhttpPost('folder-feeds.aspx?Action=NavigatetoReports', NAVIGATETO_REPORTS, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.navigatetoreportswarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }
}

  function SendOutBound(){     
       
     var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(154)) %>;
     if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.sendoutboundwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '. ' + ' This operation will happen in backgrouond' + ' . ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
          ClosePopUp("");
          xmlhttpPost('folder-feeds.aspx?Action=SendOutbound', SEND_OUTBOUND, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.sendoutboundwarning", LanguageID))%> ' + selectedItems.length + '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }
}

  function ChangeWFStatustoPreValidate(){     
     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(100)) %>;
     if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.prevalidatewarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + ' This operation will happen in background . ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
          //ClosePopUp("");
          xmlhttpPost('folder-feeds.aspx?Action=WFStatustoPreValidate', WFSTAT_PREVALIDATE, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.prevalidatewarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }
}

  function ChangeWFStatustoPostValidate(){     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(101)) %>;
     if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.postvalidatewarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + ' This operation will happen in background. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
          //ClosePopUp("");
          xmlhttpPost('folder-feeds.aspx?Action=WFStatustoPostValidate', WFSTAT_POSTVALIDATE, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.postvalidatewarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }
}

  function ChangeWFStatustoReadytoDeploy(){     
     var maxoffers = <% Sendb(MyCommon.Fetch_CM_SystemOption(102)) %>;
     if (selectedItems.length > 0) {
      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.readytodeploywarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + ' This operation will happen in background . '  + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
          //ClosePopUp("");
          xmlhttpPost('folder-feeds.aspx?Action=WFStatustoReadytoDeploy', WFSTAT_READYTODEPLOY, frmdata);
        }
      }
      else{
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.readytodeploywarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
        }
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }
}

function xmlhttpPost(strURL, action, frmdata) {
    var xmlHttpReq = false;
    var self = this;
    var tokens = new Array();

    if (window.XMLHttpRequest) { // Mozilla/Safari
        self.xmlHttpReq = new XMLHttpRequest();
    } else if (window.ActiveXObject) { // IE
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
    }
    self.xmlHttpReq.open('POST', strURL, true);
    self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    self.xmlHttpReq.onreadystatechange = function () {
        if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {

            switch (action) {
                case DUP_OFFERS:                    
                    cofirmsuccessdupoffer(self.xmlHttpReq.responseText);
                    break;
                case MASSDEPLOY_OFFERS:
                    ClosePopUp("");
                    break;
                case MASSREQUESTAPPROVALFOROFFERS:
                    ClosePopUp("");
                    break; 
                case DEPLOYANDAPPROVAL_REQUEST:
                    ClosePopUp("");
                    break;
                case NAVIGATETO_REPORTS:
                    handlenavtoreports(self.xmlHttpReq.responseText);
                    break;
                case SEND_OUTBOUND:
                    handlesendoutbound(self.xmlHttpReq.responseText);
                    break;
                case WFSTAT_PREVALIDATE:
                    handlewfstatprevalidate(self.xmlHttpReq.responseText);
                    break;
                case WFSTAT_POSTVALIDATE:
                    handlewfstatpostvalidate(self.xmlHttpReq.responseText);
                    break;
                case WFSTAT_READYTODEPLOY:
                    handlewfstatreadytodeploy(self.xmlHttpReq.responseText);
                    break;
                case TRANSFER_OFFERS:
                    confirmsuccesstransoffer(self.xmlHttpReq.responseText);
                    break;
                case SHOW_APPROVALACTIONITEMS:
                    OnSuccessResponse((self.xmlHttpReq.responseText).replace("\r\n",""));
                    break; 					
            }
        }
    }
    self.xmlHttpReq.send(frmdata);
}
function OnSuccessResponse(response) {
        if(response == "-1")
        {
            RemoveOptions('0, 9, 10, 11, 12, 13, 14, 15');
        }
        else
        {        
            if(response == "0")
            {
                AddOption('0', '<%= Copient.PhraseLib.Lookup("term.deploy", LanguageID) %>');

                RemoveOptions('9, 10, 11, 12, 13, 14, 15');
            }
            else if(response == "1")
            {
                RemoveOptions('0, 11, 12, 13, 14, 15');
            
                AddOption('10', '<%= Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('9', '<%= Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
            
            
            }
            else if(response == "2")
            {
                RemoveOptions('0, 9, 10');
                AddOption('11', '<%= Copient.PhraseLib.Lookup("term.deployoffers", LanguageID) %>');
                AddOption('15', '<%= Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('14', '<%= Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
                AddOption('13', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('12', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');
                
            
            }
            else if (response == "3") {
                RemoveOptions('0, 9, 10, 12, 13, 14, 15');
                
                AddOption('11', '<%= Copient.PhraseLib.Lookup("term.deployoffers", LanguageID) %>');

            }
            else if (response == "4") {
                RemoveOptions('0, 9, 10, 11, 14, 15');
                AddOption('13', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID) %>');
                AddOption('12', '<%= Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID) %>');

            }
        }
    }
    function isInArray(value, array) {
      return array.indexOf(value) > -1;
    }
    function RemoveOptions(optionValues){
        var values = optionValues.split(', ');
        var actionmenu = document.getElementById('Actionitems');
        var i;
        for(i=0; i<actionmenu.length; i++)
        {
            if(isInArray(actionmenu.options[i].value, values))
            {
                actionmenu.removeChild( actionmenu.options[i] ); 
                i--;
            }
        }
    }
    function AddOption(optionValue, optionText){
        var actionmenu = document.getElementById('Actionitems');
        var i, exists = 0;
        for(i=0; i<actionmenu.length; i++)
        {
            if(actionmenu.options[i].value == optionValue)
            {
                exists = 1;
                break;
            }
        }
        if(exists == 0)
        {
            actionmenu.options.add(new Option(optionText, optionValue), actionmenu.options[1])
        }
    }
function handlenavtoreports(responseText) {

//    if (responseText.substring(0, 2) == 'OK') {
        window.location = 'reports-custom.aspx';
//    }
}
function ShowRequestApprovalActionItems(itemIds) {
        xmlhttpPost('/logix/folder-feeds.aspx?Action=ShowRequestApprovalActionItem', SHOW_APPROVALACTIONITEMS, 'ItemIds=' + itemIds + '&FromOfferList=' + true)
    }
function handlesendoutbound(responseText){    
     var performactionserrorElem = document.getElementById("performactionserror");
//      var failedoffersdesc = responseText;
//      var failedoffers = responseText;
//      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
//      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));     

//      if (responseText.substring(0, 2) != 'OK') {
//        performactionserrorElem.style.display = 'block';
//        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.sendoutboundfail", LanguageID))%>';             
//        failedoffersdesc = failedoffersdesc.split('|');
//        failedoffers = failedoffers.split(',');
//        populaterow(failedoffers,failedoffersdesc);     
//      }
//      else if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'offer-list.aspx';
//      }
    }

  function handlewfstatprevalidate(responseText){     
     
      var performactionserrorElem = document.getElementById("performactionserror");
      var failedoffersdesc = responseText;
      var failedoffers = responseText;
      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));  

      if ((responseText.substring(0, 2) != 'OK') && (responseText.indexOf('||') >= 0)) {
        //performactionserrorElem.style.display = 'block';
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.requirerevalidaion", LanguageID))%>');       
        performactionserrorElem.style.display = 'none';
        toggleDialog('performactions', false);
        document.location = 'offer-list.aspx';         
      }
      else if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'offer-list.aspx';
      }
      else{
        performactionserrorElem.style.display = 'block';
        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.prevalidatefail", LanguageID))%>';
        failedoffersdesc = failedoffersdesc.split('|');
        failedoffers = failedoffers.split(',');
        populaterow(failedoffers,failedoffersdesc);          
      }
    }

function handlewfstatpostvalidate(responseText){     
      var performactionserrorElem = document.getElementById("performactionserror");
      var failedoffersdesc = responseText;
      var failedoffers = responseText;
      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));  

      if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'offer-list.aspx';
      }
      else{
        performactionserrorElem.style.display = 'block';
        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.postvalidatefail", LanguageID))%>';
        failedoffersdesc = failedoffersdesc.split('|');
        failedoffers = failedoffers.split(',');
        populaterow(failedoffers,failedoffersdesc);       
      }
    }

function handlewfstatreadytodeploy(responseText){     
      var performactionserrorElem = document.getElementById("performactionserror");
      var failedoffersdesc = responseText;
      var failedoffers = responseText;
      failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
      failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));  

      if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'offer-list.aspx';
      }
      else{
        performactionserrorElem.style.display = 'block';
        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.readytodeployfail", LanguageID))%>';
        failedoffersdesc = failedoffersdesc.split('|');
        failedoffers = failedoffers.split(',');
        populaterow(failedoffers,failedoffersdesc);         
      }
    }

function updateitemchecks(failedoffers) {
    
    for (var i = 0; i < failedoffers.length; i++) {
        elem = document.getElementById("linkID" + failedoffers[i]);
        var tr = elem.parentNode.parentNode;
        tr.style.backgroundColor = 'red';
    }
}

function resetitemscolor(offers) {

    for (var i = 0; i < (offers.length) - 1; i++) {
        elem = document.getElementById("linkID" + offers[i]);
        var tr = elem.parentNode.parentNode;
        tr.style.backgroundColor = 'transparent';
    }
}

function deletecells(rows){
     
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];        
      var mycell = row.cells[0];
		
      if (mycell != null){		           		  
        mycell.innerText = ""
      } 		
    }
  }

function populaterow(failedoffers, failedoffersdesc){
     
     for (var i = 0; i < failedoffers.length; i++) {
        var row = document.getElementById("errdesc" + failedoffers[i].trim());        
        createCell(row.insertCell(0), failedoffersdesc[i], 'row');
      }
    }

    // create DIV element and append to the table cell
    function createCell(cell, text, style) {      
      var div = document.createElement('div'), // create DIV element
      txt = document.createTextNode(text); // create text node
      div.appendChild(txt);                    // append text node to the DIV
      div.setAttribute('class', style);        // set DIV class attribute
      div.setAttribute('className', style);    // set DIV class attribute for IE (?!)      
      cell.appendChild(div);                   // append DIV to the table cell      
      cell.style.whiteSpace = "nowrap";
      cell.style.color = "#C82536";
      cell.colSpan = 6; 
                       
    }

  function assignNoofDuplicateOffers(shown) {
  
    //if (shown)  {
	//  toggleDialog('performactions', false);
	//} else {
	//  toggleDialog('performactions', true);
	//}
 
	 var elem = document.getElementById('DuplicateNoofOffer');
     var OfferfadeElem = document.getElementById('OfferfadeDiv');
	  var fadeElem = document.getElementById('ActionsTB');
	 
	 if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
      }

      if (OfferfadeElem != null) {
        OfferfadeElem.style.display = (shown) ? 'block' : 'none';
      }

	 // if (fadeElem != null) {
     //   fadeElem.style.display = (shown) ?  'none' : 'block';
     //}

	  if (shown)  {
	   document.getElementById('txtDuplicateOffersCnt').value='1';
	   document.getElementById('txtDuplicateOffersCnt').focus();
	   ClearNoOfDuplicateOfferserror();
	   return false;
	  } else {
	  return true;
	  }
   } 
   
   
  	function showNoOfDuplicateOfferserror(content){
     var duplicateofferElem = document.getElementById("DuplicateOffererror");
      
      duplicateofferElem.style.display = 'block';
      duplicateofferElem.innerHTML = content;
    }
	
	function ClearNoOfDuplicateOfferserror(){
     var duplicateofferElem = document.getElementById("DuplicateOffererror");
      if (duplicateofferElem != null) {
        duplicateofferElem.style.display = 'none';
       }  
    }
	
	function addDuplicateOfferscount() {
      var maxOffersperfolderduplicate = <% Sendb(MyCommon.NZ(MyCommon.Fetch_SystemOption(184),0)) %>;
	  if (maxOffersperfolderduplicate == 0 ) {
	    maxOffersperfolderduplicate = 99;
	  }
	  var dupOffersCntvalue =  document.getElementById("txtDuplicateOffersCnt").value;
	   if (dupOffersCntvalue != null && dupOffersCntvalue.trim() != "") {
	    if (!isNaN(dupOffersCntvalue)) {
		   ClearNoOfDuplicateOfferserror();
		   if (dupOffersCntvalue <= 0) {
		     showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.invalidDuplicateOfferCount", LanguageID))%>');
		   }
		   else if (dupOffersCntvalue > maxOffersperfolderduplicate) {
		     showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxOffersperfolderduplicate + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>');
		   }
		   else {
		       var actionele = document.getElementById('Actionitems');
               var folderlist = document.getElementById("folderList").value;
			   toggleDialogOfferDuplicate('DuplicateNoofOffer', false);
			   var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';
			   if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' ' + offerphrasetext + '. ' + ' <%Sendb(Copient.PhraseLib.Lookup("term.enteredDuplicateOffersCount", LanguageID))%> ' + dupOffersCntvalue + ' . ' +'This Operation will happen in background ' + '.  <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
               frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=' + dupOffersCntvalue + '&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
                //ClosePopUp("");
			   xmlhttpPost('folder-feeds.aspx?Action=DuplicateOffers', DUP_OFFERS, frmdata);
               selectedItems = new Array();
               selectedItemTypes = new Array(); 
			   }
            }
		}
		else {
		  showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.invalidDuplicateOfferCount", LanguageID))%>');
		}
	   }
	   else {
	     showNoOfDuplicateOfferserror('<%Sendb(Copient.PhraseLib.Lookup("term.enterDuplicateOfferCount", LanguageID))%>');
	   }

       }
     function ClosePopUp(responseText){
                var performactionserrorElem = document.getElementById("performactionserror");
                performactionserrorElem.style.display = 'none';
                toggleDialog('performactions', false);
                document.location = 'offer-list.aspx';
	}
  function toggleDialogOfferDuplicate(elemName, shown) {
	var elem = document.getElementById(elemName);
    var offerfadeElem = document.getElementById('OfferfadeDiv');
	
    if (elem != null) {
      elem.style.display = (shown) ? 'block' : 'none';
    }
    
    if (offerfadeElem != null) {
      offerfadeElem.style.display = (shown) ? 'block' : 'none';
    }	
	
    if (shown)  {
	  toggleDialog('performactions', false);
	} else {
	  toggleDialog('performactions', true);
	}
 }

function DuplicateOfferstofolder() {
    var actionele = document.getElementById('Actionitems');
     
    var folderlist = document.getElementById("folderList").value;
    
    if (selectedItems.length > 0) {
      if (actionele.value == 7){
        var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(157)) %>;
 
        if (maxoffers == 0 || selectedItems.length < maxoffers + 1) {
		  if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.AssignFolderWarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>.' + 'This operatiion will happen in background ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
            frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=1&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
            //ClosePopUp("");
            xmlhttpPost('folder-feeds.aspx?Action=DuplicateOffers', DUP_OFFERS, frmdata);
            selectedItems = new Array(); 
          }		  
        }
        else {
          alert('<%Sendb(Copient.PhraseLib.Lookup("folders.AssignFolderWarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>.' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>'); 
        }
      } 
	  else if (actionele.value == 8){
		TransferOffers();
	  }		  
      else {
        var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(153)) %>;

        if (maxoffers == 0 || selectedItems.length < maxoffers + 1) {
		 <%  if MyCommon.IsEngineInstalled(9) or MyCommon.IsEngineInstalled(0) then %>
		     assignNoofDuplicateOffers(true);
		 <% Else %>	 
             if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>.' + 'This operation will happen in background . ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
             frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=1&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
             //ClosePopUp("");
             xmlhttpPost('folder-feeds.aspx?Action=DuplicateOffers', DUP_OFFERS, frmdata);
             selectedItems = new Array();  
		  }	
		<% End If %>
        }
        else {
          alert('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>.' + '<%Sendb(Copient.PhraseLib.Lookup("folders.maxlimitreached", LanguageID))%>' + maxoffers + '. <%Sendb(Copient.PhraseLib.Lookup("foders.ActionCannotPerform", LanguageID))%>'); 
        }     
      }       
    }
    else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
    }   
}

function cofirmsuccessdupoffer(responseText) {
    //var performactionserrorElem = document.getElementById("performactionserror");
    //if (responseText.substring(0, 2) != 'OK') {
        //performactionserrorElem.style.display = 'block';
        //performactionserrorElem.innerHTML = responseText;
    //}
    //else if (responseText.substring(0, 2) == 'OK') {
      //  var CustomerInquiry = '<%Sendb(Request.QueryString("CustomerInquiry"))%>';
       // performactionserrorElem.style.display = 'none';
       // toggleDialog('performactions', false);
        //document.location = CustomerInquiry == "1" ? 'offer-list.aspx?CustomerInquiry=1' : 'offer-list.aspx';
    //}
	ClosePopUp("");
}

function confirmsuccesstransoffer(responseText) {
    //var performactionserrorElem = document.getElementById("performactionserror");
	//var failedoffers = responseText;
	//var failedoffersdesc = [];
    //if (responseText.trim() != 'OK') {
      //  performactionserrorElem.style.display = 'block';
      //  performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.transferoffersfail", LanguageID))%>';
      //  failedoffers = failedoffers.split(',');
//		for (var i = 0; i < failedoffers.length; i++) {
	//		  failedoffersdesc.push('<%Sendb(Copient.PhraseLib.Lookup("folders.transferofferserror", LanguageID))%>');
//		}
  //      populaterow(failedoffers,failedoffersdesc);
	//	document.getElementById('btnDupOffer').disabled = true; 
    //}
   // else if (responseText.trim() == 'OK') {
     //   performactionserrorElem.style.display = 'none';
      //  toggleDialog('performactions', false);
        //document.location = 'offer-list.aspx';
    //}
	ClosePopUp("");
}

  function TransferOffers(){    
    var actionele = document.getElementById('Actionitems');
    var destinationFolder = document.getElementById("folderList").value;
      if (selectedItems.length > 0) {
      frmdata = '&ItemIDs=' + selectedItems.join(',')+"&";
        frmdata += 'sFolder=0'+ '&dFolder=' + destinationFolder + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';	  
	    if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.transferofferswarning", LanguageID))%> ' + selectedItems.length + ' ' + offerphrasetext + '. '+ ' This operation will happen in background .' + ' <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
//          frmdata = 'sFolder=0&dFolder=' + destinationFolder + '&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
          //ClosePopUp("");
           xmlhttpPost('folder-feeds.aspx?Action=TransferOffers', TRANSFER_OFFERS, frmdata);
		  document.getElementById('btnDupOffer').disabled = false; 
        }
	}  
      else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
      }
  } 

     //Attach Folder status close function
      $(document).ready(function () {
        $('#statusClose').click(function() {
            var obj=this;
            var strURL="/logix/folder-feeds.aspx?Action=UpdateFolderStatus";
            $.post(strURL,{FromOfferList:true },function (data) { 
                        $(obj).parent().fadeTo(300,0,function(){
                             $(obj).remove();
                             $("#FolderStatus").empty();
                        });
                   },false);
          })  
      });

</script>
<script runat="server">
    Function IsValid_LongValue(ByVal str As String) As Boolean
        Dim temp As Long
        If (Long.TryParse(str, temp)) Then
            Return True
        Else
            Return False
        End If
    End Function
    
    Function Validate_Date_Range(ByVal startDate As String, ByVal endDate As String, ByVal date_Option As Integer, ByRef ErrorMsg As String) As Boolean
        Dim date1 As Date
        Dim date2 As Date
        Select Case date_Option
            Case 3
                If (startDate <> "") Then
                    If (Is_Valid_Date(startDate, ErrorMsg)) Then
                        date1 = startDate
                        If (endDate <> "") Then
                            If (Is_Valid_Date(endDate, ErrorMsg)) Then
                                date2 = endDate
                            End If
                        Else
                            ErrorMsg = Copient.PhraseLib.Lookup("term.invaliddaterange", LanguageID)
                        End If
                    End If
                ElseIf (endDate <> "") Then
                    If (Is_Valid_Date(endDate, ErrorMsg)) Then
                        date2 = endDate
                        ErrorMsg = Copient.PhraseLib.Lookup("term.invaliddaterange", LanguageID)
                    End If
                Else
                    Exit Select
                End If
                
                If (ErrorMsg = "") Then
                    Return Is_Valid_Date_Range(date1, date2, ErrorMsg)
                End If
            Case Else
                If (startDate <> "") Then
                    Return Is_Valid_Date(startDate, ErrorMsg)
                End If
        End Select
        Return False
    End Function
  
    Function Is_Valid_Date(ByVal dateString As String, ByRef ErrorMsg As String) As Boolean
        Dim tempDateTime As Date
        Dim DateFormats = {"MM/dd/yyyy", "M/d/yyyy", "MM/d/yyyy", "M/dd/yyyy", "yyyy-MM-dd", "yyyy-M-d", "yyyy-MM-d", "yyyy-M-dd"}

        ErrorMsg = ""
        If (Not DateTime.TryParseExact(dateString, DateFormats, New System.Globalization.CultureInfo("en-US"), System.Globalization.DateTimeStyles.None, tempDateTime)) Then
            ErrorMsg = Copient.PhraseLib.Lookup("term.invaliddateformat", LanguageID)
            Return False
        End If
        Return True
    End Function
        
    Function Is_Valid_Date_Range(ByVal startDate As Date, ByVal endDate As Date, ByRef ErrorMsg As String) As Boolean
        ErrorMsg = "Not valid date range"
        If (DateTime.Compare(startDate, endDate) < 0) Then
            ErrorMsg = ""
            Return True
        End If
        Return False
   End Function
   
   Function SortNoBanners(ByVal dst As DataTable, ByVal SortText As String, ByVal SortDir As String, ByRef MyCommon As Copient.CommonInc) As DataTable
      Dim dtConsolidated As DataTable = Nothing
      Dim dtSorted As DataTable = Nothing
      Dim row As DataRow
      Dim LastRowAdded As DataRow = Nothing
      Dim PrevOfferID As Integer = 0
      Dim OfferID As Integer = 0
      Dim sortedRows() As DataRow
    
      If (dst IsNot Nothing) Then
         dtConsolidated = dst.Clone()
         sortedRows = dst.Select("", "OfferID")
      
         For Each row In sortedRows
            OfferID = MyCommon.NZ(row.Item("OfferID"), 0)
            dtConsolidated.ImportRow(row)
            LastRowAdded = dtConsolidated.Rows(dtConsolidated.Rows.Count - 1)
            PrevOfferID = OfferID
         Next
      
         Select Case SortText
            Case "AOLV.OfferID"
               SortText = "OfferID"
            Case "AOLV.Name"
               SortText = "Name"
            Case "BAN.Name"
               SortText = "BannerName"
            Case "AOLV.StatusFlag"
               SortText = "StatusFlag"
            Case "AOLV.BuyerId"
               SortText = "BuyerId"
         End Select
      
         dtSorted = SelectIntoDataTable("", SortText & " " & SortDir, dtConsolidated)
      End If
    
      Return dtSorted
   End Function
   
    Function ConsolidateBanners(ByVal dst As DataTable, ByVal SortText As String, ByVal SortDir As String, ByRef MyCommon As Copient.CommonInc) As DataTable
        Dim dtConsolidated As DataTable = Nothing
        Dim dtSorted As DataTable = Nothing
        Dim row As DataRow
        Dim LastRowAdded As DataRow = Nothing
        Dim PrevOfferID As Integer = 0
        Dim OfferID As Integer = 0
        Dim sortedRows() As DataRow
    
        If (dst IsNot Nothing) Then
            dst.Columns.Add("Banners", System.Type.GetType("System.String"))
            dtConsolidated = dst.Clone()
            sortedRows = dst.Select("", "OfferID, BannerName")
      
            For Each row In sortedRows
                OfferID = MyCommon.NZ(row.Item("OfferID"), 0)
                If (OfferID = PrevOfferID AndAlso LastRowAdded IsNot Nothing) Then
                    LastRowAdded.Item("Banners") = LastRowAdded.Item("Banners") & "<br />" & MyCommon.NZ(row.Item("BannerName"), "")
                Else
                    row.Item("Banners") = row.Item("BannerName")
                    dtConsolidated.ImportRow(row)
                    LastRowAdded = dtConsolidated.Rows(dtConsolidated.Rows.Count - 1)
                End If
                PrevOfferID = OfferID
            Next
      
            Select Case SortText
                Case "AOLV.OfferID"
                    SortText = "OfferID"
                Case "AOLV.Name"
                    SortText = "Name"
                Case "BAN.Name"
                    SortText = "BannerName"
                Case "AOLV.StatusFlag"
                    SortText = "StatusFlag"
                Case "AOLV.BuyerId"
                    SortText = "BuyerId"
            End Select
      
            dtSorted = SelectIntoDataTable("", SortText & " " & SortDir, dtConsolidated)
        End If
    
        Return dtSorted
    End Function
  
    Private Function SelectIntoDataTable(ByVal selectFilter As String, ByVal order As String, _
                                         ByVal sourceDataTable As DataTable) As DataTable
        Dim newDataTable As DataTable = sourceDataTable.Clone
        Dim dataRows As DataRow() = sourceDataTable.Select(selectFilter, order)
        Dim typeDataRow As DataRow
    
        For Each typeDataRow In dataRows
            newDataTable.ImportRow(typeDataRow)
        Next
    
        Return newDataTable
    
    End Function
  
    Private Sub MarkOfferAsImported(ByVal OfferID As Long, ByRef MyCommon As Copient.CommonInc)
    
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    
        MyCommon.QueryStr = "update OfferIDs with (RowLock) set Imported=1 where OfferID=" & OfferID
        MyCommon.LRT_Execute()
    
    End Sub
  
    Private Function ExportListToExcel(ByRef dst As DataTable, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc) As String
        Dim bStatus As Boolean
        Dim sMsg As String = ""
        Dim CmExport As New Copient.ExportXml
        Dim sFileFullPath As String
        Dim sFullPathFileName As String
        Dim sFileName As String = "OfferList.xls"
        Dim dtExport As DataTable
        Dim dr As DataRow
        Dim drExport As DataRow
        Dim i64OfferId As Int64
        Dim sOfferStatus As String
        Dim oOfferStatus As Copient.LogixInc.STATUS_FLAGS
        Dim FolderList As String = ""


        If dst.Rows.Count > 0 Then
      
            dtExport = New DataTable()
            dtExport.Columns.Add("OfferID", Type.GetType("System.Int64"))
            dtExport.Columns.Add("XID", Type.GetType("System.String"))
            dtExport.Columns.Add("Engine", Type.GetType("System.String"))
            ''added
            dtExport.Columns.Add("ExternalBuyerId", Type.GetType("System.String"))
            dtExport.Columns.Add("Name", Type.GetType("System.String"))
            dtExport.Columns.Add("Description", Type.GetType("System.String"))
            dtExport.Columns.Add("Category", Type.GetType("System.String"))
            dtExport.Columns.Add("ProdStartDate", Type.GetType("System.String"))
            dtExport.Columns.Add("ProdEndDate", Type.GetType("System.String"))
            dtExport.Columns.Add("TestStartDate", Type.GetType("System.String"))
            dtExport.Columns.Add("TestEndDate", Type.GetType("System.String"))
            dtExport.Columns.Add("CreatedBy", Type.GetType("System.String"))
            dtExport.Columns.Add("LastUpdatedBy", Type.GetType("System.String"))
            dtExport.Columns.Add("Status", Type.GetType("System.String"))
            dtExport.Columns.Add("FolderName(s)", Type.GetType("System.String"))
      
            For Each dr In dst.Rows
                drExport = dtExport.NewRow()
                i64OfferId = MyCommon.NZ(dr.Item("OfferID"), 0)
                If i64OfferId > 0 Then
                    FolderList = ""
                    'Get the folder List for each offer
                    MyCommon.QueryStr = "dbo.pt_Get_OfferFolders"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = i64OfferId
                    MyCommon.LRTsp.Parameters.Add("@ReturnList", SqlDbType.NVarChar, -1).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    If (Not IsDBNull(MyCommon.LRTsp.Parameters("@ReturnList").Value)) Then
                        FolderList = MyCommon.LRTsp.Parameters("@ReturnList").Value
                    End If
                    MyCommon.Close_LRTsp()

                    drExport.Item("OfferID") = i64OfferId
                    drExport.Item("XID") = MyCommon.NZ(dr.Item("ExtOfferId"), "")
                    drExport.Item("Engine") = MyCommon.NZ(dr.Item("PromoEngine"), "")
                    'added
                    drExport.Item("ExternalBuyerId") = MyCommon.NZ(dr.Item("ExternalBuyerId"), 0)
                    drExport.Item("Name") = MyCommon.NZ(dr.Item("Name"), "")
                    drExport.Item("Description") = MyCommon.NZ(dr.Item("OfferDescription"), "")
                    drExport.Item("Category") = MyCommon.NZ(dr.Item("ODescription"), "")
                    drExport.Item("ProdStartDate") = Format(dr.Item("ProdStartDate"), "yyyy-MM-dd")
                    drExport.Item("ProdEndDate") = Format(dr.Item("ProdEndDate"), "yyyy-MM-dd")
                    drExport.Item("TestStartDate") = Format(dr.Item("TestStartDate"), "yyyy-MM-dd")
                    drExport.Item("TestEndDate") = Format(dr.Item("TestEndDate"), "yyyy-MM-dd")
                    drExport.Item("CreatedBy") = MyCommon.NZ(dr.Item("CreatedBy"), "Unknown")
                    drExport.Item("LastUpdatedBy") = MyCommon.NZ(dr.Item("LastUpdatedBy"), "Unknown")
                    sOfferStatus = Logix.GetOfferStatus(i64OfferId, LanguageID, oOfferStatus)
                    drExport.Item("Status") = Logix.GetOfferStatusText(oOfferStatus, LanguageID)
                    drExport.Item("FolderName(s)") = FolderList
                    dtExport.Rows.Add(drExport)
                End If
            Next

            sFileFullPath = MyCommon.Fetch_SystemOption(29)
            sFullPathFileName = sFileFullPath & "\" & sFileName

            bStatus = CmExport.ExportToExcel(sFullPathFileName, dtExport)
            If bStatus Then
                Dim oRead As System.IO.StreamReader
                Dim LineIn As String
                Dim Bom As String = ChrW(65279)
                oRead = System.IO.File.OpenText(sFullPathFileName)
                Response.Clear()
                Response.ContentEncoding = Encoding.Unicode
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment; filename=" & sFileName)
        
                'force little endian fffe bytes at front, why?  I don't know but is required.
                Sendb(Bom)
                While oRead.Peek <> -1
                    LineIn = oRead.ReadLine()
                    Send(LineIn)
                End While
                oRead.Close()
                Response.End()
                System.IO.File.Delete(sFullPathFileName)
            Else
                sMsg = CmExport.GetStatusMsg
            End If
        Else
            sMsg = Copient.PhraseLib.Lookup("offer-list.empty", LanguageID)
        End If
  
        Return sMsg
    End Function
  
</script>
<div id="performactions" class="folderdialog">
    <div class="foldertitlebar">
        <span class="dialogtitle">
            <%Sendb(Copient.PhraseLib.Lookup("folders.performaction", LanguageID))%></span>
        <span class="dialogclose" onclick="toggleDialog('performactions', false);">X</span>
    </div>
    <div id="performactionserror" style="display: none;">
    </div>
    <div class="dialogcontents">
        <br />
        <br class="half" />
        <label for="Theme">
            <%Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>:</label>&nbsp;&nbsp;
        <select name="Actionitems" id="Actionitems" onchange="javascript:handleAction();">
            <option value="-1">
                <%Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%></option>
            <%If (Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not bTestSystem Then%>
            <option value="0">
                <%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%></option>
            <%End If%>
            <%If Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers Then%>
                <option id="reqapproval" value="9"><%Sendb(Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%></option>
                <option id="reqapprovedeploy" value="10"><%Sendb(Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%></option>
                <option id="onlyreqapproval" value="12"><%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%></option>
                <option id="onlyreqapprovedeploy" value="13"><%Sendb(Copient.PhraseLib.Lookup("term.only", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%></option>
                <option id="deployreqapproval" value="14"><%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapproval", LanguageID))%></option>
                <option id="deployreqapprovedeploy" value="15"><%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID) & " " & Copient.PhraseLib.Lookup("term.and", LanguageID) & " " & Copient.PhraseLib.Lookup("term.reqapprovedeploy", LanguageID))%></option>
                <option id="deploymixoffers" value="11"><%Sendb(Copient.PhraseLib.Lookup("term.deployoffers", LanguageID)) %></option>
            <% End If %>
            <%If (Logix.UserRoles.CreateOfferFromBlank And Not bTestSystem) Then%>
            <option value="1">
                <%Sendb(Copient.PhraseLib.Lookup("folders.duplicateoffer", LanguageID))%></option>
            <%End If%>
            <%If (Logix.UserRoles.AssignPreValidate) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
            <option value="2">
                <%Sendb(Copient.PhraseLib.Lookup("term.prevalidate", LanguageID))%></option>
            <%End If%>
            <%If (Logix.UserRoles.AssignPostValidate) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
            <option value="3">
                <%Sendb(Copient.PhraseLib.Lookup("term.postvalidate", LanguageID))%></option>
            <%End If%>
            <%If (Logix.UserRoles.AssignReadyToDeploy) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
            <option value="4">
                <%Sendb(Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID))%></option>
            <%End If%>
            <%If (Logix.UserRoles.SendOffersToCRM) Then%>
            <option value="5">
                <%Sendb(Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID))%></option>
            <%End If%>
            <option value="8">
                <%Sendb(Copient.PhraseLib.Lookup("folders.transferoffers", LanguageID))%></option>
            <option value="6">
                <%Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%></option>
            <%If (Logix.UserRoles.EditFolders) Then%>
            <option value="7">
                <%Sendb(Copient.PhraseLib.Lookup("term.assignfolders", LanguageID))%></option>
            <%End If%>
              </select>
        <input type="button" name="ExecAction" id="ExecAction" value="Execute" onclick="javascript:ExecAction();"
            style="visibility: hidden;" />
        <div id="dupoffer" style="visibility: hidden;">
            <label for="lblselectfolder">
                <%Sendb(Copient.PhraseLib.Lookup("folders.PleaseSelect", LanguageID))%>:</label>&nbsp;&nbsp;
            <input type="hidden" id="folderList" name="folderList" value="" />&nbsp;
            <input type="button" name="folderbrowse" id="folderbrowse" value="Browse" onclick="javascript:openPopup('folder-browse.aspx');" />
            <br />
        </div>
        <div id="execduplicateoffer" style="visibility: hidden;">
            <table summary="">
                <tr>
                    <td valign="top" id="folderNames">
                        <%Sendb(Copient.PhraseLib.Lookup("term.noneselected", LanguageID))%>
                    </td>
                </tr>
                <tr>
                    <td>
                        <input type="button" name="btnDupOffer" id="btnDupOffer" value="Execute" onclick="javascript:DuplicateOfferstofolder();" />
                    </td>
                </tr>
            </table>
        </div>
    </div>
</div>
<div id="OfferfadeDiv">
</div>
<div id="DuplicateNoofOffer" class="folderdialog" style="position: relative; z-index: 1001;
    top: 100px; width: 400px; height: 150px">
    <div class="foldertitlebar">
        <span class="dialogtitle">
            <% Sendb(Copient.PhraseLib.Lookup("folders.copyofferstofolder", LanguageID))%></span>
        <span class="dialogclose" onclick="toggleDialogOfferDuplicate('DuplicateNoofOffer', false);">
            X</span>
    </div>
    <div class="dialogcontents">
        <div id="DuplicateOffererror" style="display: none; color: red;">
        </div>
        <table style="width: 90%">
            <tr>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    <label for="infoStart">
                        <% Sendb(Copient.PhraseLib.Lookup("term.duplicateOfferstoCreate", LanguageID).Replace("99", MyCommon.NZ(MyCommon.Fetch_SystemOption(184), 0).ToString()))%></label>
                    <input type="text" style="width: 20px" id="txtDuplicateOffersCnt" name="txtDuplicateOffersCnt"
                        maxlength="2" value="" />
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr align="right">
                <td>
                    <input type="button" name="btnOk" id="btnOk" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>"
                        onclick="addDuplicateOfferscount();" />
                    <input type="button" name="btnCancel" id="btnCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>"
                        onclick="toggleDialogOfferDuplicate('DuplicateNoofOffer', false);" />
                </td>
            </tr>
        </table>
    </div>
</div>
<script type="text/javascript">
<%  
If (Request.Form("mode") = "advancedsearch" AndAlso bEnableRestrictedAccessToUEOfferBuilder) Then %>
    if (document.getElementById("filterengine") != null) {
    document.getElementById("filterengine").value=<%=FilterEngine %>
    }
<% End If %>
 <% If (bEnableRestrictedAccessToUEOfferBuilder AndAlso CriteriaMsg <> "") Then %>
    if (document.getElementById("filterengine") != null) {document.getElementById("filterengine").disabled=true;}
 <% End If %>
</script>
<%
done:
    Send_BodyEnd("searchform", "searchterms")
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    Logix = Nothing
%>
