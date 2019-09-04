<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CAM-offer-list
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
  Dim Shaded As String = "shaded"
  Dim idNumber As Integer
  Dim idSearch As String = ""
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim File As HttpPostedFile
  Dim orderString As String
  Dim sSummaryPage As String
  Dim SearchMatchesROID As Boolean = False
  Dim RoidExtension As String = ""
  Dim WhereClause As String = ""
  Dim WhereBuf As New StringBuilder()
  Dim AdvSearchSQL As String = ""
  Dim CriteriaMsg As String = ""
  Dim CriteriaTokens As String = ""
  Dim BannersEnabled As Boolean = False
  Dim IE6ScrollFix As String = ""
  Dim infoMessage As String = ""
  Dim CustomerInquiry As Boolean = False
  Dim TotalUsers As Integer = 0
  Dim SourceName As String = ""
  Dim restrictLinks As Boolean = False
  Dim Handheld As Boolean = False
  
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CAM-offer-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  ' Check the logged in user to see if they're to be restricted to this page
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
  
  CustomerInquiry = IIf(Request.QueryString("CustomerInquiry") <> "", True, False)
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  If CustomerInquiry Then
    Send_HeadBegin("term.customerinquiry", "term.offers")
  Else
    Send_HeadBegin("term.camoffers")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
  #controls form {
    display: inline !important;
  }
  * html table {
   table-layout: auto !important;
  }
  * html #XIDcol {
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
  'If CustomerInquiry Then
  '  If (Not restrictLinks) Then
  '    Send_Tabs(Logix, 3)
  '    Send_Subtabs(Logix, 32, 2, , 0)
  '  Else
  '    Send_Subtabs(Logix, 91, 1, , 0)
  '  End If
  'Else
  Send_Tabs(Logix, 2)
  Send_Subtabs(Logix, 20, 4)
  'End If
  
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
    'check thats its not too big
    If (File.ContentLength = 0 AndAlso File.FileName <> "") Then
      infoMessage = Copient.PhraseLib.Lookup("term.upload-file-not-found", LanguageID) & " (" & File.FileName & ")"
    ElseIf (File.ContentLength = 0 AndAlso File.FileName = "") Then
      infoMessage = Copient.PhraseLib.Lookup("term.nofileselected", LanguageID)
    ElseIf File.ContentType <> "text/xml" And File.ContentType <> "application/octet-stream" And File.ContentType <> "application/x-gzip" _
        And File.ContentType <> "application/x-gzip-compressed" And File.ContentType <> "application/gzip" _
        And File.ContentType <> "application/x-tar" Then
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
      Dim CpeFileName As String = ""
      Dim sMsg As String
      
      EngineId = MyImportXml.GetOfferEngineId(UploadFileName, sXml)
      
      If (EngineId = 2 OrElse EngineId = 3 OrElse EngineId = 6) Then
        If (MyCommon.Fetch_SystemOption(66) = "1") Then
          CpeImport.SetBanners(GetBanners())
        End If
        CpeImport.ImportOffer(UploadFileName, sXml, AdminUserID, LanguageID, False)
        sMsg = CpeImport.GetErrorMsg()
        If (sMsg.Trim() = "") Then
          MyCommon.Activity_Log(3, CpeImport.GetOfferId(), AdminUserID, Copient.PhraseLib.Lookup("offer.imported", LanguageID))
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
          Response.AddHeader("Location", sSummaryPage & "?OfferID=" & CpeImport.GetOfferId())
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
    
    If (Request.Form("sourceOption") <> "0") Then
      MyCommon.QueryStr = "select Name from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=" & Request.Form("sourceOption") & ";"
      dst = MyCommon.LRT_Select
      If dst.Rows.Count > 0 Then
        SourceName = MyCommon.NZ(dst.Rows(0).Item("Name"), "")
      End If
      WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      WhereBuf.Append(GetOptionString(MyCommon, "6", Request.Form("sourceOption"), "InboundCRMEngineID"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.source", LanguageID) & " " & GetOptionType(6) & " '" & SourceName & "'")
      CritTokenBuf.Append("Source," & "6" & "," & Request.Form("sourceOption") & ",|")
    End If
    
    If (Request.Form("favoriteOption") <> "0") Then
      WhereBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("favoriteOption")), "1", "Favorite"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.favorite", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("favoriteOption"))) & " on")
      CritTokenBuf.Append("Favorite," & Integer.Parse(Request.Form("favoriteOption")) & "," & "1" & ",|")
    End If
    
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
  If (FilterOffer = "") Then FilterOffer = "1"
  If (FilterOffer = "0" OrElse FilterOffer = "3") Then
    ShowExpired = " where AOLV.deleted=0 and AOLV.EngineID = 6 and isnull(AOLV.InboundCRMEngineID,0) = 0 "
  ElseIf (FilterOffer = "1") Then
    ShowExpired = " where AOLV.deleted=0 and AOLV.EngineID = 6 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and isnull(AOLV.InboundCRMEngineID,0) = 0  "
  Else
    ShowExpired = " where AOLV.deleted=0 and AOLV.EngineID = 6 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and isnull(AOLV.InboundCRMEngineID,0) = 0  "
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
  
  MyCommon.QueryStr = "select OfferID, ExtOfferID, StatusFlag, Name, CreatedDate, isnull(ProdStartDate,0) as ProdStartDate, isnull(ProdEndDate,0) as ProdEndDate," & _
                      "C.Description as ODescription, PE.Description as PromoEngine, PE.PhraseID as EnginePhraseID, Favorite from Offers as O with (NoLock) " & _
                      "left join OfferCategories as C with (NoLock) on O.OfferCategoryID=C.OfferCategoryID " & _
                      "left join PromoEngines as PE on PE.EngineID=O.EngineID " & _
                      "where O.Deleted=0 and visible=1 and isnull(isTemplate,0)=0 "
  If (BannersEnabled) Then
    MyCommon.QueryStr = "select BAN.BannerID, BAN.Name as BannerName, AOLV.* from AllOffersListview AOLV with (NoLock) " & _
                        "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID " & _
                        "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID "
  Else
    MyCommon.QueryStr = "select AOLV.* from AllOffersListview AOLV with (NoLock) "
  End If
  
  If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
    If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
    If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
    If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))
    
    MyCommon.QueryStr += ShowExpired & " and IsNull(AOLV.isTemplate,0)=0 " & WhereBuf.ToString & ShowActive
    MyCommon.QueryStr += " order by " & SortText & " " & SortDirection
    AdvSearchSQL = WhereBuf.ToString
  Else
    If (Request.QueryString("searchterms") <> "") Then
      If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
        idSearch = idNumber.ToString
      Else
        idSearch = "-1"
      End If
      idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
      PrctSignPos = idSearchText.IndexOf("%")
      If (PrctSignPos > -1) Then
        idSearch = "-1"
        idSearchText = idSearchText.Replace("%", "[%]")
      End If
      If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
      MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired & " and(AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
      MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 " & ShowActive
      orderString = " order by " & SortText & " " & SortDirection
    Else
      MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired
      MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 " & ShowActive
      orderString = " order by " & SortText & " " & SortDirection
    End If
    
    ' check if banners are enabled
    If (BannersEnabled) Then
      MyCommon.QueryStr &= " and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                           " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                           "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                           "                     where AUB.AdminUserID = " & AdminUserID & ") ) "
    End If
    
    If (FilterOffer <> "3") Then
      MyCommon.QueryStr = MyCommon.QueryStr & orderString
    End If
  End If
  
  ShowExpired = IIf(FilterOffer = "0" OrElse FilterOffer = "3", "TRUE", "FALSE")
  
  If (FilterOffer = "2") Then
    If (BannersEnabled) Then
      MyCommon.QueryStr = "select BAN.BannerID, BAN.Name as BannerName, AOLV.* from AllActiveOffersListView AOLV with (NoLock) " & _
                          "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID " & _
                          "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                          " where AOLV.IsTemplate=0 and isnull(AOLV.InboundCRMEngineID,0) = 0 and AOLV.PromoEngine = 'CAM' and (AOLV.OfferID not in (select OfferID from BannerOffers BO with (NoLock)) " & _
                          " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                          "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                          "                     where AUB.AdminUserID = " & AdminUserID & " and isnull(AOLV.InboundCRMEngineID,0) = 0 ) ) " & AdvSearchSQL
      If (Request.QueryString("searchterms") <> "") Then
        MyCommon.QueryStr &= " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
      End If
    Else
      MyCommon.QueryStr = "select * from AllActiveOffersListView AOLV where AOLV.IsTemplate=0 and PromoEngine='CAM' "
      If (AdvSearchSQL <> "") Then
        MyCommon.QueryStr &= " and " & AdvSearchSQL
      Else
        If (Request.QueryString("searchterms") <> "") Then
          MyCommon.QueryStr &= " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
        End If
      End If
    End If
    
    MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection
  ElseIf (FilterOffer = "3") Then
    MyCommon.QueryStr &= " and AOLV.OfferID in (" & _
                         "select Distinct STI.IncentiveID as OfferID from CPE_ST_Incentives STI with (NoLock) " & _
                         "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = STI.IncentiveID " & _
                         "inner join OfferLocUpdate OLU with (NoLock) on OLU.OfferID = STI.IncentiveID " & _
                         "inner join Locations LOC with (NoLock) on LOC.LocationID = OLU.LocationID and LOC.TestingLocation = 0 " & _
                         "where STI.Deleted=0 and I.Deleted=0 and LOC.Deleted=0 " & _
                         "  and (getdate() between STI.StartDate and DateAdd(d, 1, STI.EndDate)) " & _
                         "  and (I.UpdateLevel <> STI.UpdateLevel or I.StatusFlag <> 0) " & _
                         "union " & _
                         "select Distinct STO.OfferID as OfferID from CM_ST_Offers STO with (NoLock) " & _
                         "inner join Offers O with (NoLock) on O.OfferID = STO.OfferID " & _
                         "inner join OfferLocUpdate OLU with (NoLock) on OLU.OfferID = STO.OfferID " & _
                         "inner join Locations LOC with (NoLock) on LOC.LocationID = OLU.LocationID and LOC.TestingLocation = 0 " & _
                         "where STO.Deleted=0 and O.Deleted=0 and LOC.Deleted=0 " & _
                         "  and (getdate() between STO.ProdStartDate and DateAdd(d, 1, STO.ProdEndDate)) " & _
                         "  and (O.UpdateLevel <> STO.UpdateLevel or O.StatusFlag <> 0)) and isnull(AOLV.InboundCRMEngineID,0) = 0 " & _
                         " order by " & SortText & " " & SortDirection
  ElseIf (FilterOffer = "4") Then
    MyCommon.QueryStr = "select OfferID, ExtOfferID, StatusFlag, Name, CreatedDate, isnull(ProdStartDate,0) as ProdStartDate, isnull(ProdEndDate,0) as ProdEndDate," & _
                        "C.Description as ODescription, PE.Description as PromoEngine, PE.PhraseID as EnginePhraseID, Favorite from Offers as O with (NoLock) " & _
                         "left join OfferCategories as C with (NoLock) on O.OfferCategoryID=C.OfferCategoryID " & _
                         "left join PromoEngines as PE on PE.EngineID=O.EngineID " & _
                         "where O.Deleted=0 and visible=1 and isnull(isTemplate,0)=0 " & _
                         "and OfferID in (" & _
                         " select IncentiveID as OfferID from CPE_Incentives where EngineID=6 and MutuallyExclusive=1 )"
    If (BannersEnabled) Then
      MyCommon.QueryStr = "select BAN.BannerID, BAN.Name as BannerName, AOLV.* from AllOffersListview AOLV with (NoLock) " & _
                          "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID " & _
                          "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                          "where AOLV.Deleted =0 and AOLV.IsTemplate=0 and AOLV.OfferID in ( " & _
                          "select IncentiveID as OfferID from CPE_Incentives where EngineID=6 and MutuallyExclusive=1)"
                          
      If (Request.QueryString("searchterms") <> "") Then
        MyCommon.QueryStr &= " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
      End If
    Else
      MyCommon.QueryStr = "select AOLV.* from AllOffersListview AOLV with (NoLock) " & _
                          "where AOLV.Deleted =0 and AOLV.IsTemplate=0 and AOLV.OfferID in ( " & _
                          "select IncentiveID as OfferID from CPE_Incentives where EngineID=6 and MutuallyExclusive=1)"
      If (Request.QueryString("searchterms") <> "") Then
        MyCommon.QueryStr &= " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
      End If
      MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection

    End If

  End If
  dst = MyCommon.LRT_Select

  If (BannersEnabled) Then
    dst = ConsolidateBanners(dst, SortText, SortDirection, MyCommon)
  End If
  
  sizeOfData = dst.Rows.Count
  
  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    If (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "CPE") Then
      sSummaryPage = "CPEoffer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
    ElseIf (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "Website") Then
      sSummaryPage = "web-offer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
    ElseIf (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "CAM") Then
      sSummaryPage = "CAM-offer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
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
        Sendb(Copient.PhraseLib.Lookup("term.camoffers", LanguageID))
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If CustomerInquiry Then
        Send("<form id=""controlsform"" name=""controlsform"" action=""#"">")
        Send("</form>")
      Else
        If (Logix.UserRoles.ImportOffer) Then
          Send_Import()
        End If
        Send("<form id=""controlsform"" name=""controlsform"" action=""../offer-new.aspx"">")
        If (Logix.UserRoles.CreateOfferFromBlank) Then
          Send("<input type=""hidden"" name=""NewCAM"" value=""Yes"" />")
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
<div id="main"<% Sendb(IE6ScrollFix) %>>
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    If CustomerInquiry Then
      Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, ShowExpired)
    Else
      Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, ShowExpired)
    End If
    If (CriteriaMsg <> "") Then
      Send("<div id=""criteriabar"">" & CriteriaMsg & "<a href=""javascript:editSearchCriteria();"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "]</a><a href=""CAMoffer-list" & IIf(CustomerInquiry, "?CustomerInquiry=1", "") & """ class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.clear", LanguageID) & "]</a></div>")
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
          <a id="xidLink" onclick="handleIter('xidLink');" href="CAM-offer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ExtOfferID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
          <a id="idLink" onclick="handleIter('idLink');" href="CAM-offer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=AOLV.OfferID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
          <a id="engineLink" onclick="handleIter('engineLink');" href="CAM-offer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=PromoEngine&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
        <th align="left" class="th-name" scope="col">
          <a id="nameLink" onclick="handleIter('nameLink');" href="CAM-offer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
          <a id="createLink" onclick="handleIter('createLink');" href="CAM-offer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=CreatedDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
          <a id="startLink" onclick="handleIter('startLink');" href="CAM-offer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ProdStartDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
          <a id="endLink" onclick="handleIter('endLink');" href="CAM-offer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ProdEndDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %>">
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
          <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%>
          <%
            If SortText = "StatusFlag" Then
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
        
        j = i
        While (j < sizeOfData And j < linesPerPage + linesPerPage * PageNum)
          arrList.Add(dst.Rows(j).Item("OfferID"))
          j += 1
        End While
        arrList.TrimToSize()
        ReDim OfferList(arrList.Count - 1)
        For j = 0 To arrList.Count - 1
          OfferList(j) = arrList.Item(j).ToString
        Next
        Statuses = Logix.GetStatusForOffers(OfferList, LanguageID)
        
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
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
          Send("  <td>" & Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0)), LanguageID, dst.Rows(i).Item("PromoEngine")) & "</td>")
          'Send("  <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0), LanguageID) & "</td>")
          If (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CPE") Then
            Sendb("  <td><a href=""CPEoffer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
          ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Website") Then
            Sendb("  <td><a href=""web-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
          ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Email") Then
            Sendb("  <td><a href=""email-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
          ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CAM") Then
            Sendb("  <td><a href=""CAM-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
          Else
            Sendb("  <td><a href=""offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
          End If
          If (MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0) > 0 AndAlso MyCommon.NZ(dst.Rows(i).Item("UpdateLevel"), 0) > 0 AndAlso CustomerInquiry = False) Then
            Send("<br /><span class=""red"" style=""font-size:10px;font-weight:bold;"">(" & Copient.PhraseLib.Lookup("alert.offermodified", LanguageID) & ")</span></td>")
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
          i = i + 1
          ' Next
        End While
      %>
    </tbody>
  </table>
</div>

<script runat="server">
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
  if (document.getElementById('firstPageLink') != null) { document.getElementById('firstPageLink').onclick = handleFirst; }
  if (document.getElementById('previousPageLink') != null) { document.getElementById('previousPageLink').onclick = handlePrev; }
  if (document.getElementById('nextPageLink') != null) { document.getElementById('nextPageLink').onclick = handleNext; }
  if (document.getElementById('lastPageLink') != null) { document.getElementById('lastPageLink').onclick = handleLast; }
  
  function handleFirst() {
    handleIter('firstPageLink');
  }
  
  function handlePrev() {
    handleIter('previousPageLink');
  }
  
  function handleNext() {
    handleIter('nextPageLink');
  }
  
  function handleLast() {
    handleIter('lastPageLink');
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
    
    if (elemAdv != null && elemAdv.value !="") {
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
</script>

<script runat="server">
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
</script>

<%
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
