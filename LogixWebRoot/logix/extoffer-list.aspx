<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
    ' *****************************************************************************
    ' * FILENAME: extoffer-list.aspx 
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
    Dim rst As DataTable
    Dim dst As System.Data.DataTable
    Dim dstOffersUniqueUserHasRights As System.Data.DataTable
    Dim OffersUniqueUserHasRightsRows As DataRow
    Dim row As System.Data.DataRow
    Dim DT As System.Data.DataTable
    Dim DR As System.Data.DataRow
    Dim Shaded As String = "shaded"
    Dim idNumber As Integer
    Dim idSearch As String = ""
    Dim idSearchText As String = ""
    Dim PageNum As Integer = 0
    Dim MorePages As Boolean
    Dim linesPerPage As Integer = 20
    Dim sizeOfData As Integer
    Dim i As Integer = 0
    Dim File As HttpPostedFile
    Dim InstallPath As String
    Dim orderString As String
    Dim sSummaryPage As String
    Dim SearchMatchesROID As Boolean = False
    Dim RoidExtension As String = ""
    Dim WhereClause As String = ""
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
    Dim bStoreUser As Boolean = False
    Dim sValidLocIDs As String = ""
    Dim sValidSU As String = ""
    Dim wherestr As String = ""
    Dim sJoin As String = ""
    Dim iLen As Integer = 0
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "extoffer-list.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
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
  
    CustomerInquiry = IIf(Request.QueryString("CustomerInquiry") <> "", True, False)
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
%>
<style type="text/css">
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
</style>
<%
    Send_Scripts()
%>
<script type="text/javascript">
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
//   function chooseFile() {
//      document.getElementById("fileInput").click();
//   }
//   function fileonclick()
//   {
//   var filename=document.getElementById("fileInput").value;
//    document.getElementById("lblfileupload").innerText = filename.replace("C:\\fakepath\\", "");
//  }
</script>
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
        Send_Subtabs(Logix, 20, 3)
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
            And File.ContentType <> "application/x-gzip-compressed" And File.ContentType <> "application/gzip" Then
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
            'Dim DpImport As New Copient.ImportXmlDP
            CMS.AMS.CurrentRequest.Resolver.AppName = "extOffer-list.aspx"
            Dim UEImport As Copient.ImportXMLUE = CMS.AMS.CurrentRequest.Resolver.Resolve(Of Copient.ImportXMLUE)()
            Dim CpeFileName As String = ""
            Dim sMsg As String
      
            EngineId = MyImportXml.GetOfferEngineId(UploadFileName, sXml)
      
            If (EngineId = 2 OrElse EngineId = 3) Then
                If (MyCommon.Fetch_SystemOption(66) = "1") Then
                    CpeImport.SetBanners(GetBanners())
                End If
                CpeImport.ImportOffer(UploadFileName, sXml, AdminUserID, LanguageID, False)
                sMsg = CpeImport.GetErrorMsg()
                If (sMsg.Trim() = "") Then
                    MyCommon.Activity_Log(3, CpeImport.GetOfferId(), AdminUserID, Copient.PhraseLib.Lookup("offer.imported", LanguageID))
                    sSummaryPage = IIf(CpeImport.GetEngineId() = 2, "CPEoffer-sum.aspx", "web-offer-sum.aspx")
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", sSummaryPage & "?OfferID=" & CpeImport.GetOfferId())
                Else
                    infoMessage = sMsg
                End If
            ElseIf (EngineId = 9) Then
                If (MyCommon.Fetch_SystemOption(66) = "1") Then
                    UEImport.SetBanners(GetBanners())
                End If
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
            ElseIf (EngineId = -1) Then
                infoMessage = MyImportXml.GetStatusMsg
                If infoMessage = "" Then
                    infoMessage = Copient.PhraseLib.Lookup("offer-list.notoffer", LanguageID)
                End If
                'ElseIf (EngineId = 4) Then
                '  infoMessage = "1"
                '  bStatus = DpImport.ImportOfferLoad(UploadFileName, sXml, AdminUserID, LanguageID)
                '  infoMessage += " status: " & bStatus
                '  sMsg = DpImport.GetStatusMsg
                '  infoMessage += " " & sMsg.Length
                '  If sMsg.Length > 0 Then
                '    If bStatus Then
                '      ' display warning using sMsg
                '      sOfferId = DpImport.GetOfferId
                '      infoMessage = sMsg
                '      AssignOfferToBanners(CInt(sOfferId), MyCommon)
                '      Response.Status = "301 Moved Permanently"
                '      Response.AddHeader("Location", "DP-offer-gen.aspx?OfferID=" & sOfferId & "&amp;infoMessage=" & infoMessage)
                '    Else
                '      ' display error using sMsg
                '      infoMessage = sMsg
                '    End If
                '  Else
                '    If bStatus Then
                '      sOfferId = DpImport.GetOfferId
                '      AssignOfferToBanners(CInt(sOfferId), MyCommon)
                '      Response.Status = "301 Moved Permanently"
                '      Response.AddHeader("Location", "DP-offer-gen.aspx?OfferID=" & sOfferId)
                '    End If
                '  End If
                '  infoMessage += " before delete"
                '  If System.IO.File.Exists(UploadFileName) = True Then
                '    System.IO.File.Delete(UploadFileName)
                '  End If
            Else
                bStatus = MyImportXml.ImportOfferLoad(UploadFileName, sXml, AdminUserID, LanguageID)
                sMsg = MyImportXml.GetStatusMsg
                If sMsg.Length > 0 Then
                    If bStatus Then
                        ' display warning using sMsg
                        sOfferId = MyImportXml.GetOfferId
                        infoMessage = sMsg
                        AssignOfferToBanners(CInt(sOfferId), MyCommon)
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "offer-gen.aspx?OfferID=" & sOfferId & "&amp;infoMessage=" & infoMessage)
                    Else
                        ' display error using sMsg
                        infoMessage = sMsg
                    End If
                Else
                    If bStatus Then
                        sOfferId = MyImportXml.GetOfferId
                        AssignOfferToBanners(CInt(sOfferId), MyCommon)
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "offer-gen.aspx?OfferID=" & sOfferId)
                    End If
                End If
                If System.IO.File.Exists(UploadFileName) = True Then
                    System.IO.File.Delete(UploadFileName)
                End If
            End If
        
        End If
    End If
    
    ' handle an Advance Search Criteria
    If (Request.Form("mode") = "advancedsearch") Then
        Dim TempStr As String = ""
        Dim CritBuf As New StringBuilder()
        Dim CritTokenBuf As New StringBuilder()
    
        If (Request.Form("xid").Trim <> "" AndAlso Request.Form("xidOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("xidOption")), Request.Form("xid"), "AOLV.ExtOfferID"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.xid", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("xidOption"))) & " '" & Request.Form("xid").Trim & "'")
            CritTokenBuf.Append("XID," & Integer.Parse(Request.Form("xidOption")) & "," & Request.Form("xid").Trim & ",|")
        End If
    
        If (Request.Form("idSearch").Trim <> "" AndAlso Request.Form("idOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("idOption")), Request.Form("idSearch"), "AOLV.OfferID"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("idOption"))) & " '" & Request.Form("idSearch").Trim & "'")
            CritTokenBuf.Append("ID," & Integer.Parse(Request.Form("idOption")) & "," & Request.Form("idSearch").Trim & ",|")
        End If
    
        If (Request.Form("offerName").Trim <> "" AndAlso Request.Form("nameOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("nameOption")), Request.Form("offerName"), "AOLV.Name"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.name", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("nameOption"))) & " '" & Request.Form("offerName").Trim & "'")
            CritTokenBuf.Append("Name," & Integer.Parse(Request.Form("nameOption")) & "," & Request.Form("offerName").Trim & ",|")
        End If
    
        If (Request.Form("desc").Trim <> "" AndAlso Request.Form("descOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("descOption")), Request.Form("desc"), "Convert(nvarchar(1000),OfferDescription)"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.description", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("descOption"))) & " '" & Request.Form("desc").Trim & "'")
            CritTokenBuf.Append("Desc," & Integer.Parse(Request.Form("descOption")) & "," & Request.Form("desc").Trim & ",|")
        End If
    
        If (Request.Form("roid").Trim <> "" AndAlso Request.Form("roidOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("roidOption")), Request.Form("roid"), "RewardOptionID"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.roid", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("roidOption"))) & " '" & Request.Form("roid").Trim & "'")
            CritTokenBuf.Append("ROID," & Integer.Parse(Request.Form("roidOption")) & "," & Request.Form("roid").Trim & ",|")
        End If
    
        If (Request.Form("createdby").Trim <> "" AndAlso Request.Form("createdbyOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append("AOLV.CreatedByAdminID IN (select Distinct AdminUserID from AdminUsers where ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("createdbyOption")), Request.Form("createdby"), "UserName"))
            WhereBuf.Append(") ")
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.createdby", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("createdbyOption"))) & " '" & Request.Form("createdby").Trim & "'")
            CritTokenBuf.Append("CreatedBy," & Integer.Parse(Request.Form("createdbyOption")) & "," & Request.Form("createdby").Trim & ",|")
        End If
    
        If (Request.Form("lastupdatedby").Trim <> "" AndAlso Request.Form("lastupdatedbyOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append("AOLV.LastUpdatedByAdminID IN (select Distinct AdminUserID from AdminUsers where ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("lastupdatedbyOption")), Request.Form("lastupdatedby"), "UserName"))
            WhereBuf.Append(") ")
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.lastupdatedby", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("lastupdatedbyOption"))) & " '" & Request.Form("lastupdatedby").Trim & "'")
            CritTokenBuf.Append("LastUpdatedBy," & Integer.Parse(Request.Form("lastupdatedbyOption")) & "," & Request.Form("lastupdatedby").Trim & ",|")
        End If
    
        If (Request.Form("engine").Trim <> "" AndAlso Request.Form("engineOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("engineOption")), Request.Form("engine"), "PromoEngine"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.engine", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("engineOption"))) & " '" & Request.Form("engine").Trim & "'")
            CritTokenBuf.Append("Engine," & Integer.Parse(Request.Form("engineOption")) & "," & Request.Form("engine").Trim & ",|")
        End If
    
        If (BannersEnabled) Then
            If (Request.Form("banner").Trim <> "" AndAlso Request.Form("bannerOption") <> "") Then
                WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("bannerOption")), Request.Form("banner"), "BAN.Name"))
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                CritBuf.Append(Copient.PhraseLib.Lookup("term.banner", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("bannerOption"))) & " '" & Request.Form("banner").Trim & "'")
                CritTokenBuf.Append("BAN.Name," & Integer.Parse(Request.Form("bannerOption")) & "," & Request.Form("banner").Trim & ",|")
            End If
        End If
    
        If (Request.Form("category").Trim <> "" AndAlso Request.Form("categoryOption") <> "") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("categoryOption")), Request.Form("category"), "ODescription"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.category", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("categoryOption"))) & " '" & Request.Form("category").Trim & "'")
            CritTokenBuf.Append("Category," & Integer.Parse(Request.Form("categoryOption")) & "," & Request.Form("category").Trim & ",|")
        End If
    
        If (Request.Form("createdDate1").Trim <> "" AndAlso Request.Form("createdOption") <> "") Then
            TempStr = GetDateOption(MyCommon, Integer.Parse(Request.Form("createdOption")), Request.Form("createdDate1"), Request.Form("createdDate2"), "AOLV.CreatedDate")
            If (TempStr <> "") Then
                WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                WhereBuf.Append(TempStr)
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                If (TempStr.IndexOf("between") > -1) Then
                    CritBuf.Append(Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("createdOption"))) & " '" & Request.Form("createdDate1").Trim & "'")
                    If Request.Form("createdDate2").Trim <> "" Then
                        CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("createdDate2").Trim & "'")
                    End If
                Else
                    CritBuf.Append(Copient.PhraseLib.Lookup("term.created", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("createdOption"))) & " '" & Request.Form("createdDate1").Trim & "'")
                    CritTokenBuf.Append("Created," & Integer.Parse(Request.Form("createdOption")) & "," & Request.Form("createdDate1").Trim & ",|")
                End If
            End If
        End If
    
        If (Request.Form("startDate1").Trim <> "" AndAlso Request.Form("startOption") <> "") Then
            TempStr = GetDateOption(MyCommon, Integer.Parse(Request.Form("startOption")), Request.Form("startDate1"), Request.Form("startDate2"), "ProdStartDate")
            If (TempStr <> "") Then
                WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                WhereBuf.Append(TempStr)
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                If (TempStr.IndexOf("between") > -1) Then
                    CritBuf.Append(Copient.PhraseLib.Lookup("term.starts", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("startOption"))) & " '" & Request.Form("startDate1").Trim & "'")
                    If Request.Form("startDate2").Trim <> "" Then
                        CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("startDate2").Trim & "'")
                    End If
                    CritTokenBuf.Append("Starts," & Integer.Parse(Request.Form("startOption")) & "," & Request.Form("startDate1").Trim & "," & Request.Form("startDate2").Trim & "|")
                Else
                    CritBuf.Append(Copient.PhraseLib.Lookup("term.starts", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("startOption"))) & " '" & Request.Form("startDate1").Trim & "'")
                    CritTokenBuf.Append("Starts," & Integer.Parse(Request.Form("startOption")) & "," & Request.Form("startDate1").Trim & ",|")
                End If
            End If
        End If
    
        If (Request.Form("endDate1").Trim <> "" AndAlso Request.Form("endOption") <> "") Then
            TempStr = GetDateOption(MyCommon, Integer.Parse(Request.Form("endOption")), Request.Form("endDate1"), Request.Form("endDate2"), "ProdEndDate")
            If (TempStr <> "") Then
                WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                WhereBuf.Append(TempStr)
                If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                If (TempStr.IndexOf("between") > -1) Then
                    CritBuf.Append(Copient.PhraseLib.Lookup("term.ends", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("endOption"))) & " '" & Request.Form("endDate1").Trim & "'")
                    If Request.Form("endDate2").Trim <> "" Then
                        CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " '" & Request.Form("endDate2").Trim & "'")
                    End If
                    CritTokenBuf.Append("Ends," & Integer.Parse(Request.Form("endOption")) & "," & Request.Form("endDate1").Trim & "," & Request.Form("endDate2").Trim & "|")
                Else
                    CritBuf.Append(Copient.PhraseLib.Lookup("term.ends", LanguageID) & " " & GetDateOptionType(Integer.Parse(Request.Form("endOption"))) & " '" & Request.Form("endDate1").Trim & "'")
                    CritTokenBuf.Append("Ends," & Integer.Parse(Request.Form("endOption")) & "," & Request.Form("endDate1").Trim & ",|")
                End If
            End If
        End If
    
        If (Request.Form("favoriteOption") <> "0") Then
            WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("favoriteOption")), "1", "Favorite"))
            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
            CritBuf.Append(Copient.PhraseLib.Lookup("term.favorite", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("favoriteOption"))) & " on")
            CritTokenBuf.Append("Favorite," & Integer.Parse(Request.Form("favoriteOption")) & "," & "1" & ",|")
        End If
        
        If MyCommon.Fetch_SystemOption(156) = "1" Then
                        
            MyCommon.QueryStr = "select * from UserDefinedFields where AdvancedSearch = 1"
            dst = MyCommon.LRT_Select

            Dim RowCount As Integer = -1
            Dim udfrow As DataRow
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
                udfrow = udfdst.Rows(0)
                             
                Dim UDFPK = udfrow.Item("UDFPK")
                If (hasOption("udf-" & udfcount) And Request.Form("udf-" & udfcount) <> "") And udfrow.Item("DataType") <> 3 Then
			
                    'If udfAdvSearch = False Then	udfAdvSearch = True

                    Select Case udfrow.Item("DataType")
                        Case 0, 1, 4, 5, 6 'string, integer, listbox, likert
                            
                            If udfquery.Length > 0 Then udfquery.Append(" Intersect ")
                            
                            Dim testString As String
                            testString = "select OfferID from UserDefinedFieldsValues where " & GetOptionString(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), Request.Form("udf-" & udfcount), udfrow.Item("ColumnName"))
                            
                            
                            udfquery.Append("select OfferID from UserDefinedFieldsValues where " & GetOptionString(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), Request.Form("udf-" & udfcount), udfrow.Item("ColumnName")))
                            udfquery.Append(" and UDFPK = " & UDFPK & " and OfferID = AOLV.OfferID")
                            If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                            CritBuf.Append("UDF-" & UDFPK & " " & GetOptionType(Integer.Parse(Request.Form("udfOption-" & udfcount))) & " '" & Request.Form("udf-" & udfcount).Trim & "'")
                            CritTokenBuf.Append("UDFRow-" & udfcount & ",UDF-" & UDFPK & "," & Integer.Parse(Request.Form("udfOption-" & udfcount)) & "," & Request.Form("udf-" & udfcount).Trim & ",|")
                            
                           
                        
                        Case 2
                            Try
                                TempStr = "select OfferID from UserDefinedFieldsValues where " & GetDateOption(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), Request.Form("udf-" & udfcount), MyCommon.NZ(Request.Form("udfEnd-" & udfcount), ""), udfrow.Item("ColumnName"))
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
                ElseIf Request.Form("udfOption-" & udfcount) <> "0" And udfrow.Item("DataType") = "3" Then
                    'If udfAdvSearch = False Then	udfAdvSearch = True
                    If udfquery.Length > 0 Then udfquery.Append(" Intersect ")
                    udfquery.Append("select OfferID from UserDefinedFieldsValues where " & GetOptionString(MyCommon, Integer.Parse(Request.Form("udfOption-" & udfcount)), "1", udfrow.Item("ColumnName")))
                    udfquery.Append(" and UDFPK = " & UDFPK & " and OfferID = AOLV.OfferID")
                    If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                    CritBuf.Append("UDF-" & UDFPK & " " & GetOptionType(Integer.Parse(Request.Form("udfOption-" & udfcount))) & " on")
                    CritTokenBuf.Append("UDFRow-" & udfcount & ",UDF-" & UDFPK & "," & Integer.Parse(Request.Form("udfOption-" & udfcount)) & "," & "1" & ",|")
                Else
                    
                End If
            Next
            If udfquery.Length > 0 Then WhereBuf.Append(" and AOLV.OfferID = (" & udfquery.ToString & ")")

        End If 'ends UDF

    
        CriteriaMsg = CritBuf.ToString
        CriteriaTokens = CritTokenBuf.ToString
    End If
  
    Dim SortText As String = "AOLV.OfferID"
    Dim SortDirection As String = "DESC"
    Dim ShowExpired As String = ""
    Dim ShowActive As String = ""
    Dim PrctSignPos As Integer
    Dim FilterOffer As String
  
    FilterOffer = Request.QueryString("filterOffer")
    'If (FilterOffer = "") Then FilterOffer = "0"

    If BannersEnabled Then
        ShowExpired = " and AOLV.deleted=0 and isnull(AOLV.InboundCRMEngineID,0) > 0 "
    Else
        ShowExpired = " where AOLV.deleted=0 and isnull(AOLV.InboundCRMEngineID,0) > 0 "
    End If
    
    'If UE Engine is installed and User is assoiated with any Buyer and if user is not having View Offer Regardless of Buyer Permission, list User-Buyer specific Offers
    If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer)) Then
        ShowExpired = ShowExpired & " and ( AOLV.BuyerId in (select BuyerId from BuyerRoleUsers where AdminUserID=" & AdminUserID & "))"
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
        sJoin = "Full Outer Join OfferLocUpdate olu with (NoLock) on AOLV.ExtOfferID=olu.OfferID "
        wherestr = " and (LocationID in (" & sValidLocIDs & ") or CreatedByAdminID in (" & sValidSU & ")) "
    End If
  
    MyCommon.QueryStr = "from AllOffersListview AOLV with (NoLock) " & sJoin
  
    If (BannersEnabled) Then
        MyCommon.QueryStr &= "left join BannerOffers BO with (NoLock) " & _
                             "on BO.OfferID = AOLV.OfferID " & _
                             "left join Banners BAN with (NoLock) " & _
                             "on BAN.BannerID = BO.BannerID " & _
                             "where isnull(AOLV.InboundCRMEngineID,0) > 0 "
    End If
  
    If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
        If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
        If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
        If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))
    
        MyCommon.QueryStr += ShowExpired & " and IsNull(AOLV.isTemplate,0)=0 " & WhereBuf.ToString & ShowActive & wherestr
        MyCommon.QueryStr += " order by " & SortText & " " & SortDirection
        AdvSearchSQL = WhereBuf.ToString
    Else
        If (Request.QueryString("searchterms") <> "") Then
            If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
                idSearch = idNumber.ToString
            Else
                idSearch = "-1"
            End If
            idSearchText = MyCommon.Parse_Quotes(HttpUtility.UrlDecode(Request.QueryString("searchterms")))
            PrctSignPos = idSearchText.IndexOf("%")
            If (PrctSignPos > -1) Then
                idSearch = "-1"
                idSearchText = idSearchText.Replace("%", "[%]")
            End If
            If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
            MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired & " and(AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%' or AOLV.ExtOfferID like N'%" & idSearchText & "%') "
            MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 " & ShowActive & wherestr
            orderString = " order by " & SortText & " " & SortDirection
        Else
            MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired
            MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 " & ShowActive & wherestr
            orderString = " order by " & SortText & " " & SortDirection
        End If
    
        If (BannersEnabled) Then
            MyCommon.QueryStr &= " and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                                 " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                                 "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                                 "                     where AUB.AdminUserID = " & AdminUserID & ") ) "
        End If
    End If
  
    If Request.QueryString("filterOffer") <> "" Then
        If MyCommon.Extract_Val(Request.QueryString("filterOffer")) > 0 Then
            MyCommon.QueryStr &= " and AOLV.InboundCRMEngineID =" & MyCommon.Extract_Val(Request.QueryString("filterOffer"))
        End If
    End If
  
    'At this point, MyCommon.QueryStr contains the FROM and WHERE clauses of the query.  We need to build 2 versions of this query, one that will tell us the count of the total number of rows
    'and the second that will return the data for the sub (paganated) set of rows that we are going to return on the page
    'First we'll tack on what we need to query for the count of the total number of rows that meet the search & filter criteria  
  
    CountQuery = "select AOLV.OfferID " & MyCommon.QueryStr
    'Second we'll tack on what we need to query for the subset of data that needs to be displayed on this page
    'start by adding the names of the columns that we'll need for the page display.  This is not the completed SelectQuery, we'll add more to it later after we know the complete record count
    If BannersEnabled Then
        SelectQuery = "BAN.BannerID, BAN.Name as BannerName, AOLV.* " & MyCommon.QueryStr
    Else
        SelectQuery = "AOLV.* " & MyCommon.QueryStr
    End If
  
  
    'before we run the CountQuery or the SelectQuery, we need to see if we are doing an export to Excel
    If (Request.QueryString("excel") <> "") Then
        If BannersEnabled Then
            MyCommon.QueryStr = "select BAN.BannerID, BAN.Name as BannerName, AOLV.* " & MyCommon.QueryStr & " order by " & SortText & " " & SortDirection
        Else
            MyCommon.QueryStr = "select AOLV.* " & MyCommon.QueryStr & " order by " & SortText & " " & SortDirection
        End If
        dst = MyCommon.LRT_Select
        infoMessage = ExportListToExcel(dst, MyCommon, Logix, BannersEnabled)
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

    'now that we know the total record count and updated that value in 'sizeofdata', we can determine how we should slice up the selectquery. 
    If (Request.Form("favoriteoption") <> "0" And Request.Form("favoriteoption") <> "") Then
        If (Request.Form("favoriteoption") = "6") Then
            FavoriteOption6 = True
        Else
            Send("<!-- building normal query -->")
            SelectQuery1 = "select distinct " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
        End If
    Else
        Send("<!-- building normal query -->")
        ' query for all results for all pages
        SelectQuery1 = "select distinct " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
    End If

        ''add the SQL necessary to restrict the return set to only those rows that need to be displayed on this page
        'SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
  
  
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
    'If FavoriteOption6 Then
    '    ' Original Query:  SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by #temp.OfferID" & " " & SortDirection & ") as RowNumber, #temp.* from #temp" & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
    '    If (OffersInPage.Length > 0) Then
    '        SelectQuery = "select * from ( select #temp.* from #temp" & " ) as Table1 where #temp.OfferID in (" & OffersInPage & ")"
    '    Else
    '        SelectQuery = "select * from ( select #temp.* from #temp" & " ) as Table1"
    '    End If
    'Else
    '    'Original Query:  SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
    '    If (OffersInPage.Length > 0) Then
    '        SelectQuery = "select * from ( select " & SelectQuery & " ) as Table1 where OfferID in (" & OffersInPage & ")"
    '    Else
    '        SelectQuery = "select * from (select  " & SelectQuery & " ) as Table1"
    '    End If
    'End If
    If (sizeOfData > linesPerPage) And (((linesPerPage * PageNum) + 1) > (sizeOfData / 2)) Then
        Send("<!-- building reverse query -->")
        Send("<!-- SortText=" & SortText & "   SortDirection=" & SortDirection & " -->")
    
        'SelectQueryOrderBy = SortText
        'If (SortText.LastIndexOf(".") > 0) Then 'if the SortText (column name) is a dotted name (table.column), then grab just the column name off the end
        '  SelectQueryOrderBy = Right(SortText, (Len(SortText) - SortText.LastIndexOf(".") - 1))
        'End If
        If SortText = "AOLV.StatusFlag" Then
            SelectQueryOrderBy = "AOLV.StatusFlag"
        Else
            SelectQueryOrderBy = SortText
            If (SortText.LastIndexOf(".") > 0) Then 'if the SortText (column name) is a dotted name (table.column), then grab just the column name off the end
                SelectQueryOrderBy = Right(SortText, (Len(SortText) - SortText.LastIndexOf(".") - 1))
            End If
        End If
    
        If UCase(Trim(SortDirection)) = "DESC" Then
            SelectSortDirection = "asc"
            SelectQueryOrderBy = SelectQueryOrderBy & " desc"
        Else
            SelectSortDirection = "desc"
        End If

        If (sizeOfData - (linesPerPage + (linesPerPage * PageNum))) < 0 Then
            StartPoint = 1
        Else
            StartPoint = (sizeOfData - (linesPerPage + (linesPerPage * PageNum))) + 1
        End If
        EndPoint = (sizeOfData - (linesPerPage * PageNum))
        SelectQuery1 = "Select " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
        ' query for all results for all pages

        If SortText = "AOLV.StatusFlag" Then
            SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SelectQueryOrderBy & ") as RowNumber, " & SelectQuery & " ) as Table1 "
        Else
            SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SelectSortDirection & ", AOLV.CreatedDate " & SelectSortDirection & " ) as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & (StartPoint).ToString & " and " & (EndPoint).ToString
        End If
        'SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SelectSortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & (StartPoint).ToString & " and " & (EndPoint).ToString & " order by " & SelectQueryOrderBy
    Else
        
        Send("<!-- building normal query -->")
        If (Request.Form("favoriteOption") <> "0" And Request.Form("favoriteOption") <> "") Then
            If (Request.Form("favoriteOption") = "6") Then
                SelectQuery1 = SelectQuery.Replace("into #temp", String.Empty)
                SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by #temp.OfferID" & " " & SortDirection & ") as RowNumber, #temp.* from #temp" & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
            Else
                Send("<!-- building normal query -->")
                SelectQuery1 = "Select DISTINCT " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
                SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
            End If
        Else
            Send("<!-- building normal query -->")
            If SortText = "AOLV.StatusFlag" Then
                Dim defaultSortFieldOrder As String = "AOLV.StatusFlag  DESC"
                SelectQuery1 = "Select DISTINCT " & SelectQuery & " order by " & defaultSortFieldOrder & ""
                ' query for all results for all pages      
                'add the SQL necessary to restrict the return set to only those rows that need to be displayed on this page
                SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & defaultSortFieldOrder & ") as RowNumber, " & SelectQuery & " ) as Table1 "
            Else
                SelectQuery1 = "Select DISTINCT " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
                ' query for all results for all pages      
                'add the SQL necessary to restrict the return set to only those rows that need to be displayed on this page
                SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
                
                
            End If
        End If

        ''add the SQL necessary to restrict the return set to only those rows that need to be displayed on this page
        'SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
    End If
    MyCommon.QueryStr = SelectQuery
    'Send("<!-- Query=" & MyCommon.QueryStr & " -->")
    dst = MyCommon.LRT_Select
    If (BannersEnabled) Then
      dst = ConsolidateBanners(dst, SortText, SortDirection, MyCommon)
   Else
      dst = SortNoBanners(dst, SortText, SortDirection, MyCommon)
    End If
  
  
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
                    Sendb("<input type=""hidden"" id=""External"" name=""External"" value=""True"" />")
                    Send_New()
                    'Sendb("<input type=""submit"" accesskey=""n"" class=""regular"" id=""new"" name=""new"" value=""" & Copient.PhraseLib.Lookup("term.new", LanguageID) & """ />")
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
        If CustomerInquiry Then
            Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)
        Else
            Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)
        End If
        If (CriteriaMsg <> "") Then
            Send("<div id=""criteriabar"">" & CriteriaMsg & "<a href=""javascript:editSearchCriteria();"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "]</a><a href=""offer-list.aspx" & IIf(CustomerInquiry, "?CustomerInquiry=1", "") & """ class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.clear", LanguageID) & "]</a></div>")
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
                    <a id="xidLink" onclick="handleIter('xidLink');"href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ExtOfferID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
                    <a id="idLink" onclick="handleIter('idLink');" href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=AOLV.OfferID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
                <th align="left" class="th-category" scope="col">
                    <a id="engineLink" onclick="handleIter('engineLink');" href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=InboundCRMEngineID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.externalsource", LanguageID))%>
                    </a>
                    <%
                        If SortText = "InboundCRMEngineID" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-name" scope="col">
                    <a id="nameLink" onclick="handleIter('nameLink');" href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=AOLV.Name&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
                    <a id="createLink" onclick="handleIter('createLink');" href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=CreatedDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
                    <a id="startLink" onclick="handleIter('startLink');"href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ProdStartDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
                    <a id="endLink" onclick="handleIter('endLink');" href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ProdEndDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
                   <a id="statusLink" onclick="handleIter('statusLink');" href="extoffer-list.aspx?Sort=true&amp;pagenum=<%=PageNum%>&amp;searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&amp;filterOffer=<% Sendb(FilterOffer)%>&amp;SortText=AOLV.StatusFlag&amp;SortDirection=<% Sendb(SortDirection)%><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", ""))%>">
                    <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%>
                    <%
                        If SortText = "AOLV.StatusFlag" Then
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
        
                j = i
                'While (j < sizeOfData And j < linesPerPage + linesPerPage * PageNum)
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
                    StartIdx = (linesPerPage * PageNum)
                    Pagelength = linesPerPage + (linesPerPage * PageNum) - 1
                    If Pagelength > dst.Rows.Count Then
                        Pagelength = dst.Rows.Count - 1
                    End If
                End If
                'While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
                i = 0
                ' While (i < dst.Rows.Count)
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
                    'find the name of the external source
                    MyCommon.QueryStr = "select CASE WHEN P.Phrase IS NULL THEN ExtCRMInterfaces.Name ELSE P.Phrase END AS Name from ExtCRMInterfaces " & _
                                    "LEFT JOIN PhraseText P ON ExtCRMInterfaces.PhraseID = P.PhraseID where ExtInterfaceID=" & MyCommon.NZ(dst.Rows(i).Item("InboundCRMEngineID"), 0)
                    DT = MyCommon.LRT_Select()
                    If DT.Rows.Count > 0 Then
                        Send("  <td>" & MyCommon.NZ(DT.Rows(0).Item("Name"), "NONE") & "</td>")
                    End If
                    'Send("  <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0), LanguageID) & "</td>")
                    If (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CPE") Then
                        Sendb("  <td><a href=""CPEoffer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Website") Then
                        Sendb("  <td><a href=""web-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Email") Then
                        Sendb("  <td><a href=""email-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CAM") Then
                        Sendb("  <td><a href=""CAM/CAM-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "UE") Then
                        Sendb("  <td><a href=""UE/UEoffer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
                    Else
                        Sendb("  <td><a href=""offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
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
                    If (Not IsDBNull(dst.Rows(i).Item("ProdStartDate"))) And (dst.Rows(i).Item("ProdStartDate") > "1/1/1900") Then
                        Send("  <td>" & Logix.ToShortDateString(dst.Rows(i).Item("ProdStartDate"), MyCommon) & "</td>")
                    Else
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
                    End If
                    If (Not IsDBNull(dst.Rows(i).Item("ProdEndDate"))) And (dst.Rows(i).Item("ProdStartDate") > "1/1/1900") Then
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
                    'i = i + 1
                Next
                ' End While
            %>
        </tbody>
    </table>
</div>
<script runat="server">
    Private Sub MarkOfferAsImported(ByVal OfferID As Long, ByRef MyCommon As Copient.CommonInc)
    
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    
        MyCommon.QueryStr = "update OfferIDs with (RowLock) set Imported=1 where OfferID=" & OfferID
        MyCommon.LRT_Execute()
    
    End Sub
    Private Function hasOption(ByRef optionName As String) As Boolean
    
        Dim val As String = Request.Form(optionName)
        Return val IsNot Nothing AndAlso val.Trim().Length > 0
        
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
                OptionType = "contains"
        End Select
        Return OptionType
    End Function
  
    Function GetDateOption(ByRef MyCommon As Copient.CommonInc, ByVal OptionIndex As Integer, _
                           ByVal StartValue As String, ByVal EndValue As String, ByVal FieldName As String) As String
        Dim StartDate, EndDate As Date
        Dim FieldBuf As New StringBuilder()
        If ((Date.TryParse(StartValue, StartDate) AndAlso OptionIndex <> 3) _
           OrElse (Date.TryParse(StartValue, StartDate) AndAlso Date.TryParse(EndValue, EndDate))) Then
            FieldBuf.Append(FieldName & " ")
            Select Case OptionIndex
                Case 0 ' on
                    FieldBuf.Append(" between '" & StartDate.ToString("yyyy-MM-ddT00:00:00") & "' and '" & StartDate.ToString("yyyy-MM-ddT23:59:59") & "' ")
                Case 1 ' before
                    FieldBuf.Append(" < '" & StartDate.ToString("yyyy-MM-dd") & "' ")
                Case 2 ' after
                    FieldBuf.Append(" > '" & StartDate.ToString("yyyy-MM-dd") & "' ")
                Case 3 ' between
                    FieldBuf.Append(" between '" & StartDate.ToString("yyyy-MM-dd") & "' and '" & EndDate.ToString("yyyy-MM-dd") & "' ")
                Case Else ' default to after
                    FieldBuf.Append(" > '" & StartDate.ToString("yyyy-MM-dd") & "' ")
            End Select
        End If
        Return FieldBuf.ToString
    End Function
  
    Function GetDateOptionType(ByVal OptionIndex As Integer) As String
        Dim OptionType As String = "after"
        Select Case OptionIndex
            Case 0 ' on
                OptionType = "on"
            Case 1 ' before
                OptionType = "before"
            Case 2 ' after
                OptionType = "after"
            Case 3 ' between
                OptionType = "between"
            Case Else ' default to after
                OptionType = "after"
        End Select
        Return OptionType
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
            } else {
                if (currentURL.indexOf("&") > -1) {
                    newURL = currentURL + "&amp;filterOffer=" + newFilter;
                } else {
                    newURL = currentURL + "?filterOffer=" + newFilter;
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
</script>
<script runat="server">
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
        Dim Banners As String = ""
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

    Private Function ExportListToExcel(ByRef dst As DataTable, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc, BannersEnabled As Boolean) As String
        Dim bStatus As Boolean
        Dim sMsg As String = ""
        Dim CmExport As New Copient.ExportXml
        Dim sFileFullPath As String
        Dim sFullPathFileName As String
        Dim sFileName As String = "ExternalOfferList.xls"
        Dim dtExtSource As DataTable
        Dim dtExport As DataTable
        Dim dr As DataRow
        Dim drExport As DataRow
        Dim i64OfferId As Int64
        Dim sOfferStatus As String
        Dim oOfferStatus As Copient.LogixInc.STATUS_FLAGS
        Dim iPreviousExtSourceId As Integer = 0
        Dim iExtSourceId As Integer = 0
        Dim sExtSourceName As String = ""
        Dim sOfferDescription As String = ""


        If dst.Rows.Count > 0 Then
      
            dtExport = New DataTable()
            If BannersEnabled Then
                dtExport.Columns.Add("BannerId", Type.GetType("System.Int32"))
                dtExport.Columns.Add("BannerName", Type.GetType("System.Int32"))
            End If
            dtExport.Columns.Add("OfferID", Type.GetType("System.Int64"))
            dtExport.Columns.Add("XID", Type.GetType("System.String"))
            dtExport.Columns.Add("ExternalSource", Type.GetType("System.String"))
            dtExport.Columns.Add("Engine", Type.GetType("System.String"))
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
      
            For Each dr In dst.Rows
                drExport = dtExport.NewRow()
                i64OfferId = MyCommon.NZ(dr.Item("OfferID"), 0)
                If i64OfferId > 0 Then
                    iExtSourceId = MyCommon.NZ(dr.Item("InboundCRMEngineID"), 0)
                    If iExtSourceId <> iPreviousExtSourceId Then
                        iPreviousExtSourceId = iExtSourceId
                        MyCommon.QueryStr = "select CASE WHEN P.Phrase IS NULL THEN ExtCRMInterfaces.Name ELSE P.Phrase END AS Name from ExtCRMInterfaces " & _
                                "LEFT JOIN PhraseText P ON ExtCRMInterfaces.PhraseID = P.PhraseID where ExtInterfaceID=" & iExtSourceId
                        dtExtSource = MyCommon.LRT_Select()
                        If dtExtSource.Rows.Count > 0 Then
                            sExtSourceName = dtExtSource.Rows(0).Item(0)
                        Else
                            sExtSourceName = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
                        End If
                    End If
                    sOfferDescription = dr.Item("OfferDescription").ToString()
                    sOfferDescription = Replace(sOfferDescription, vbCrLf, vbNullString, 1, -1, vbTextCompare)
                    drExport.Item("OfferID") = i64OfferId
                    drExport.Item("XID") = MyCommon.NZ(dr.Item("ExtOfferId"), "")
                    drExport.Item("ExternalSource") = sExtSourceName
                    drExport.Item("Engine") = MyCommon.NZ(dr.Item("PromoEngine"), "")
                    drExport.Item("Name") = MyCommon.NZ(dr.Item("Name"), "")
                    drExport.Item("Description") = MyCommon.NZ(sOfferDescription, "")
                    drExport.Item("Category") = MyCommon.NZ(dr.Item("ODescription"), "")
                    drExport.Item("ProdStartDate") = Format(dr.Item("ProdStartDate"), "yyyy-MM-dd")
                    drExport.Item("ProdEndDate") = Format(dr.Item("ProdEndDate"), "yyyy-MM-dd")
                    drExport.Item("TestStartDate") = Format(dr.Item("TestStartDate"), "yyyy-MM-dd")
                    drExport.Item("TestEndDate") = Format(dr.Item("TestEndDate"), "yyyy-MM-dd")
                    drExport.Item("CreatedBy") = MyCommon.NZ(dr.Item("CreatedBy"), "Unknown")
                    drExport.Item("LastUpdatedBy") = MyCommon.NZ(dr.Item("LastUpdatedBy"), "Unknown")
                    sOfferStatus = Logix.GetOfferStatus(i64OfferId, LanguageID, oOfferStatus)
                    drExport.Item("Status") = Logix.GetOfferStatusText(oOfferStatus, LanguageID)
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
        
                'force little endian fffe bytes at front, why?  i dont know but is required.
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
<%
done:
    Send_BodyEnd("searchform", "searchterms")
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    Logix = Nothing
%>
