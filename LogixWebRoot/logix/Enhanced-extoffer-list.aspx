<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: Enhanced-extoffer-list.aspx 
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
  ' * NOTES   : $Id: Enhanced-extoffer-list 55884 2012-09-18 20:34:28Z mark $
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
  Dim dst2 As System.Data.DataTable
  Dim row As System.Data.DataRow
  Dim DT As System.Data.DataTable
  Dim DR As System.Data.DataRow
  Dim Shaded As String = "shaded"
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim File As HttpPostedFile
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
  Dim SourceName As String = ""
  Dim restrictLinks As Boolean = False
  Dim Handheld As Boolean = False
  Dim CountQuery As String = ""
  Dim SelectQuery As String = ""
  Dim SelectQuery1 As String = ""
  Dim SelectQueryOrderBy As String = ""
  Dim SelectSortDirection As String = ""
  Dim StartPoint As Long
  Dim EndPoint As Long
  Dim bWorkflowActive As Boolean = False
  Dim bProductionSystem As Boolean = True
  Dim bTestSystem As Boolean = False
  Dim bArchiveSystem As Boolean = False
  Dim bCmInstalled As Boolean = False
  Dim bUeInstalled As Boolean = False
  dim bAdvancedExternalSearch As Boolean = False 
  Dim bSearchInternalAndExternal As Boolean = False
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
  Dim FilterEngine As Integer = 0 ' if System option #249 is enabled then CM is set as default option in the filterengine dropdown.

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "Enhanced-extoffer-list.aspx"
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
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
   Send_Scripts(New String() {"datePicker.js", "popup.js"})
%>
<style type="text/css">

  body {
    overflow: visible;
  }

  #controls form {
    display: inline !important;
  }
  * html table {
   table-layout: auto !important;
  }
  * html #XIDcol {
   width: auto !important;
  }

  #performactions  

 {
   position:absolute;
   z-index:999;
   top:10px; /* set top value */
   left:600px; /* set left value */
   width:400px;  /* set width value */
   height: 250px; 
 }
 
  #OfferfadeDiv {
  background-color: #e0e0e0;
  position: absolute;
  top: 0px;
  left: 0px;
  width:100%;
  height:100%;
  z-index: 1000;
  display:none;
  opacity: .4;
  filter: alpha(opacity=40);
  }
  
  #execduplicateoffer
 {
   overflow-y: auto;   
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
        Send("openPopup(""advanced-search.aspx?ExternalOfferList=1"");")
      End If
    %>
  }
  function editSearchCriteria() {
    var tokenStr = document.frmIter.advTokens.value;
    
    self.name = "OfferListWin";
    <%
      If CustomerInquiry Then
        Send("openPopup(""advanced-search.aspx?CustomerInquiry=1&tokens="" + tokenStr);")
      Else
        Send("openPopup(""advanced-search.aspx?ExternalOfferList=1&tokens="" + tokenStr);")
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
<script type="text/javascript" src="../javascript/jquery.js"></script>
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
    'check thats its not too big
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
      CMS.AMS.CurrentRequest.Resolver.AppName = "Enhanced-extoffer-list.aspx"
      Dim UEImport As Copient.ImportXMLUE = CMS.AMS.CurrentRequest.Resolver.Resolve(Of Copient.ImportXMLUE)()
      Dim CpeFileName As String = ""
      Dim sMsg As String
      
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
          If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not Logix.UserRoles.CreateUEOffers) Then
             infoMessage = Copient.PhraseLib.Lookup("term.unsupportedpromoengine", LanguageID)
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
        infoMessage = Copient.PhraseLib.Lookup("term.unsupportedpromoengine", LanguageID)
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
      WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("idOption")), Request.Form("idSearch"), "AOLV.OfferID"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.id", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("idOption"))) & " '" & Request.Form("idSearch").Trim & "'")
      CritTokenBuf.Append("ID," & Integer.Parse(Request.Form("idOption")) & "," & Request.Form("idSearch").Trim & ",|")
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
      WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("roidOption")), Request.Form("roid"), "RewardOptionID"))
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.roid", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("roidOption"))) & " '" & Request.Form("roid").Trim & "'")
      CritTokenBuf.Append("ROID," & Integer.Parse(Request.Form("roidOption")) & "," & Request.Form("roid").Trim & ",|")
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
      If(bEnableRestrictedAccessToUEOfferBuilder) Then
        FilterEngine = IIf(Request.Form("engine").Trim.ToLower.Equals("cm"),0,9)
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
    
    If (hasOption("createdDate1") AndAlso Request.Form("createdOption") <> "") Then
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
    End If
    
    If (hasOption("startDate1") AndAlso Request.Form("startOption") <> "") Then
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
    
    End If
    
    If (hasOption("endDate1") AndAlso Request.Form("endOption") <> "") Then
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
      bSearchInternalAndExternal = False
    End If
    
    If (Request.Form("favoriteOption") <> "0") Then
      WhereBuf.Append(" and " & GetOptionString(MyCommon, Integer.Parse(Request.Form("favoriteOption")), "1", "Favorite"))
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

        End If

    
    CriteriaMsg &= CritBuf.ToString
    CriteriaTokens = CritTokenBuf.ToString
  End If
  
  Dim SortText As String = "AOLV.OfferID"
  Dim SortDirection As String = "DESC"
  Dim ShowExpired As String = ""
  Dim ShowActive As String = ""
  Dim PrctSignPos As Integer
  Dim FilterOffer As String
  Dim FilterUser As String
  
  FilterOffer = Request.QueryString("filterOffer")
  FilterUser = Request.QueryString("filterUser")
  If ( bEnableRestrictedAccessToUEOfferBuilder AndAlso Request.Form("mode") <> "advancedsearch") Then
        If(Not String.IsNullOrEmpty(Request.QueryString("filterengine"))) Then 
            FilterEngine = Convert.ToInt16(Server.HtmlEncode(Request.QueryString("filterengine")))
            If(FilterEngine=9) Then
                If(FilterOffer="5" OrElse FilterOffer="6" OrElse FilterOffer="7" OrElse FilterOffer="8") Then FilterOffer=0
            End If 
        End If
  End If
  If (FilterOffer = "") Then FilterOffer = "1"
  If (FilterUser = "") Then FilterUser = AdminUserID.ToString
  If (FilterOffer = "0" OrElse FilterOffer = "3" OrElse FilterOffer = "4") Then
    ShowExpired = " AOLV.deleted=0 and isnull(AOLV.InboundCRMEngineID,0) > 0 "
  ElseIf (FilterOffer = "1") Then
    ShowExpired = " AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and isnull(AOLV.InboundCRMEngineID,0) > 0 "
  Else
    ShowExpired = " AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and isnull(AOLV.InboundCRMEngineID,0) > 0 "
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
  
  If BannersEnabled Then
    MyCommon.QueryStr = "from AllOffersListviewNoCAM AOLV with (NoLock) " & _
                        "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID " & _
                        "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                        "where "
  Else
    MyCommon.QueryStr = "from AllOffersListviewNoCAM AOLV with (NoLock) " & _
                        "where "
  End If
  
  If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
    If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
    If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
    If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))
    
    MyCommon.QueryStr += ShowExpired & " and IsNull(AOLV.isTemplate,0)=0 " & WhereBuf.ToString & ShowActive
    'MyCommon.QueryStr += " order by " & SortText & " " & SortDirection
    AdvSearchSQL = WhereBuf.ToString
    If (FilterOffer <> "3") Then
      'If this option is enabled and selected, return a dropdown of all users with at least one banner in common
      If (FilterOffer = "4") Then
        'MyCommon.QueryStr &= " and AOLV.CreatedByAdminID = " & AdminUserID & " "
        MyCommon.QueryStr &= " and AOLV.CreatedByAdminID = " & FilterUser & " "
      End If
    End If
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
      If bProductionSystem Then
        MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired & " and(AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & _
                            "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
      Else
        MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired & " and(AOLV.OfferID=" & idSearch & " or AOLV.ProductionID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & _
                            "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
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
  
  If (FilterOffer = "2") Then
    If (BannersEnabled) Then
      MyCommon.QueryStr = "from AllActiveOffersListView AOLV with (NoLock) " & _
                          "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID and AOLV.EngineID<>6 " & _
                          "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                          " where AOLV.IsTemplate=0 and isnull(AOLV.InboundCRMEngineID,0) > 0 and (AOLV.OfferID not in (select OfferID from BannerOffers BO with (NoLock)) " & _
                          " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                          "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                          "                     where AUB.AdminUserID = " & AdminUserID & " and isnull(AOLV.InboundCRMEngineID,0) > 0 ) ) "
    Else
      MyCommon.QueryStr = "from AllActiveOffersListView AOLV where AOLV.IsTemplate=0 and AOLV.PromoEngine<>'CAM' and InboundCRMEngineID>0 "
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
                         "where STO.Deleted=0 and O.Deleted=0 and LOC.Deleted=0 "
                          If Not bWorkflowActive Then
                              MyCommon.QueryStr &= "  and (getdate() between STO.ProdStartDate and DateAdd(d, 1, STO.ProdEndDate)) " 
                         End If
                         MyCommon.QueryStr &= "  and (O.UpdateLevel <> STO.UpdateLevel or O.StatusFlag <> 0)) and isnull(AOLV.InboundCRMEngineID,0) > 0 "
  ElseIf (FilterOffer = "5" Or FilterOffer = "6" Or FilterOffer = "7") Then
    MyCommon.QueryStr &= " and exists (select OfferId from Offers with (NoLock) where (isnull(WorkflowStatus,0) = " & (Integer.Parse(FilterOffer) - 4) & ") and OfferId=AOLV.OfferID) "

  ElseIf (FilterOffer = "8") Then
     'filtering to only display unexpired offers that have a status of testing and development. Includes scheduled offers which were not deployed
    If (BannersEnabled) Then
      MyCommon.QueryStr = "from AllOffersListviewNoCAM AOLV with (NoLock) " & _
                          "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID and AOLV.EngineID<>6 " & _
                          "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
						  "where AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and AOLV.DeploySuccessDate is null and isnull(AOLV.InboundCRMEngineID,0) > 0  " & _
                          " and isnull(AOLV.isTemplate,0)=0  and not exists ( Select OfferID from [AllActiveOffersListView] AAOFLV where AOLV.OfferID = AAOFLV.OfferID)  " & _
                          " and (AOLV.OfferID not in (select OfferID from BannerOffers BO with (NoLock))  or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                          "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                          "                     where AUB.AdminUserID = " & AdminUserID & " and isnull(AOLV.InboundCRMEngineID,0) > 0 ) ) "
    Else  
      MyCommon.QueryStr  = " from AllOffersListviewNoCAM AOLV with (NoLock) " & _
                         " where AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() and AOLV.DeploySuccessDate is null and " & _
                         " isnull(AOLV.InboundCRMEngineID,0) > 0  and isnull(AOLV.isTemplate,0)=0   " & _
                         " and not exists ( Select OfferID from [AllActiveOffersListView] AAOFLV where AOLV.OfferID = AAOFLV.OfferID) " 
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

  
  'At this point, MyCommon.QueryStr contains the FROM and WHERE clauses of the query.  We need to build 2 versions of this query, one that will tell us the count of the total number of rows
  'and the second that will return the data for the sub (paganated) set of rows that we are going to return on the page
  'First we'll tack on what we need to query for the count of the total number of rows that meet the search & filter criteria  
  CountQuery = "select count(*) as NumRows " & MyCommon.QueryStr
  'Second we'll tack on what we need to query for the subset of data that needs to be displayed on this page
  'start by adding the names of the columns that we'll need for the page display.  This is not the completed SelectQuery, we'll add more to it later after we know the complete record count
  If BannersEnabled Then
    SelectQuery = "BAN.BannerID, BAN.Name as BannerName, AOLV.* " & MyCommon.QueryStr
  Else
    SelectQuery = "AOLV.* " & MyCommon.QueryStr
  End If
  
  ' If advanced external search is on and searching for Offer ID
  ' Then search internal and external offers
  If bAdvancedExternalSearch And bSearchInternalAndExternal Then
    SelectQuery = SelectQuery.Replace("isnull(AOLV.InboundCRMEngineID,0) > 0 ", "isnull(AOLV.InboundCRMEngineID,0) > -1")
    CountQuery = CountQuery.Replace("isnull(AOLV.InboundCRMEngineID,0) > 0 ", "isnull(AOLV.InboundCRMEngineID,0) > -1")
  End If

    If(bEnableRestrictedAccessToUEOfferBuilder) Then
        If(Logix.UserRoles.CreateUEOffers AndAlso Logix.UserRoles.AccessTranslatedUEOffers) Then
              SelectQuery = SelectQuery & " and EngineID =" & FilterEngine & " "
              CountQuery = CountQuery & " and EngineID =" & FilterEngine & " "
        Else If (Logix.UserRoles.CreateUEOffers AndAlso Not Logix.UserRoles.AccessTranslatedUEOffers) Then
              SelectQuery = SelectQuery & " and EngineID =" & FilterEngine & "  and isnull(AOLV.InboundCRMEngineID,0) <> 10 "
              CountQuery = CountQuery & " and EngineID =" & FilterEngine & " and isnull(AOLV.InboundCRMEngineID,0) <> 10 "
         Else If (Not Logix.UserRoles.CreateUEOffers AndAlso  Logix.UserRoles.AccessTranslatedUEOffers) Then
            If(FilterEngine =0) Then
              SelectQuery = SelectQuery & " and EngineID = "& FilterEngine &" "
              CountQuery = CountQuery & " and EngineID = "& FilterEngine &" "
            Else If(FilterEngine=9) Then
              SelectQuery = SelectQuery & " and EngineID = "& FilterEngine &" and isnull(AOLV.InboundCRMEngineID,0) = 10 "
              CountQuery = CountQuery & " and EngineID = "& FilterEngine &" and isnull(AOLV.InboundCRMEngineID,0) = 10 "    
            End If
         Else If(Not Logix.UserRoles.CreateUEOffers AndAlso  Not Logix.UserRoles.AccessTranslatedUEOffers) Then
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
  dst = MyCommon.LRT_Select
  sizeOfData = 0
  If dst.Rows.Count > 0 Then
    sizeOfData = dst.Rows(0).Item("NumRows")
  End If
  dst = Nothing

  'Now that we know the total record count, we can determine how we should slice up the SelectQuery.  If we are wanting a subset of recrods that are past the middle of the complete record set,
  'that it is faster to switch the ordering of the records around, and grab our subset from the beginning of the list.
  
  'If the start position of the subset of records we are looking for is passed the mid-point of total size of the record set ... then ... 
  If (sizeOfData > linesPerPage) And (((linesPerPage * PageNum) + 1) > (sizeOfData / 2)) Then
    Send("<!-- building reverse query -->")
    Send("<!-- SortText=" & SortText & "   SortDirection=" & SortDirection & " -->")
    
    SelectQueryOrderBy = SortText
    If (SortText.LastIndexOf(".") > 0) Then 'if the SortText (column name) is a dotted name (table.column), then grab just the column name off the end
      SelectQueryOrderBy = Right(SortText, (Len(SortText) - SortText.LastIndexOf(".") - 1))
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
    ' query for all results for all pages
      SelectQuery1 = "Select " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
      ' query for all results for all pages
    SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SelectSortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & (StartPoint).ToString & " and " & (EndPoint).ToString & " order by " & SelectQueryOrderBy
  Else
    Send("<!-- building normal query -->")
     ' query for all results for all pages
      SelectQuery1 = "Select " & SelectQuery & " order by " & SortText & " " & SortDirection & ""
      ' query for all results for all pages      
    'add the SQL necessary to restrict the return set to only those rows that need to be displayed on this page
    SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
  End If

  'Run the query that returns the subset of data to be displayed on this page 
  MyCommon.QueryStr = SelectQuery
  Send("<!-- Query=" & MyCommon.QueryStr & " -->")
  'Response.End()
  dst = MyCommon.LRT_Select
  If (BannersEnabled) Then
    dst = ConsolidateBanners(dst, SortText, SortDirection, MyCommon)
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
<div id="performactions" class="folderdialog"> 
<div class="foldertitlebar">
<span class="dialogtitle"><%Sendb(Copient.PhraseLib.Lookup("folders.performaction", LanguageID))%></span>
<span class="dialogclose" onclick="toggleDialog('performactions', false);">X</span>
</div>
<div id="performactionserror" style="display:none;">
</div>
<div class="dialogcontents">
<br />
<br class="half"/>
<label for="Theme"><%Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>:</label>&nbsp;&nbsp;
<select name="Actionitems" id="Actionitems" onchange="javascript:handleAction();">
<option value="-1"><%Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%></option>
<%If (Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not bTestSystem Then%>
<option value="0"><%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%></option>
<%End If %>
<%If (Logix.UserRoles.CreateOfferFromBlank And Not bTestSystem) Then%>
<option value="1"><%Sendb(Copient.PhraseLib.Lookup("folders.duplicateoffer", LanguageID))%></option>
<%End If %>
<%If (Logix.UserRoles.AssignPreValidate) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
<option value="2"><%Sendb(Copient.PhraseLib.Lookup("term.prevalidate", LanguageID))%></option>
<%End If%>
<%If (Logix.UserRoles.AssignPostValidate) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
<option value="3"><%Sendb(Copient.PhraseLib.Lookup("term.postvalidate", LanguageID))%></option>
<%End If%>
<%If (Logix.UserRoles.AssignReadyToDeploy) AndAlso bWorkflowActive AndAlso bProductionSystem Then%>
<option value="4"><%Sendb(Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID))%></option>
<%End If%>
<%If (Logix.UserRoles.SendOffersToCRM) Then%>
<option value="5"><%Sendb(Copient.PhraseLib.Lookup("offer-sum.sendoutbound", LanguageID))%></option>
<%End If%>
<option value="8"><%Sendb(Copient.PhraseLib.Lookup("folders.transferoffers", LanguageID))%></option>
<option value="6"><%Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID))%></option>
<%If (Logix.UserRoles.EditFolders) Then %>
<option value="7"><%Sendb(Copient.PhraseLib.Lookup("term.assignfolders", LanguageID))%></option>
<%End If%>
</select>
<input type="button" name="ExecAction" id="ExecAction" value="Execute" onclick="javascript:ExecAction();" style="visibility:hidden;"/>
<div id="dupoffer"" style="visibility:hidden;">
<label for="lblselectfolder"><%Sendb(Copient.PhraseLib.Lookup("folders.PleaseSelect", LanguageID))%>:</label>&nbsp;&nbsp;
<input type="hidden" id="folderList" name="folderList" value="" />&nbsp;
<input type="button" name="folderbrowse" id="folderbrowse" value="Browse" onclick="javascript:openPopup('folder-browse.aspx');" />
<br />
</div>
<div id="execduplicateoffer" style="visibility:hidden;">
<table summary="">
<tr>
<td valign="top" id="folderNames"><%Sendb(Copient.PhraseLib.Lookup("term.noneselected", LanguageID))%></td>
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
<div id="OfferfadeDiv"></div>
<div id="DuplicateNoofOffer" class="folderdialog" style="position:relative; z-index:1001; top: 100px; WIDTH: 400px; HEIGHT: 150px">
  <div class="foldertitlebar">
    <span class="dialogtitle"><% Sendb(Copient.PhraseLib.Lookup("folders.copyofferstofolder", LanguageID)) %></span> 
	<span class="dialogclose" onclick="toggleDialogOfferDuplicate('DuplicateNoofOffer', false);">X</span>
  </div>
  <div class="dialogcontents">
    <div id="DuplicateOffererror" style="display: none; color: red;">
    </div>
    <table style="width:90%">
		<tr><td>&nbsp;</td></tr>
      <tr>
        <td>
		  <label for="infoStart"><% Sendb(Copient.PhraseLib.Lookup("term.duplicateOfferstoCreate", LanguageID).Replace("99", MyCommon.NZ(MyCommon.Fetch_SystemOption(184), 0).ToString()))%></label>
		  <input type="text" style="width:20px" id="txtDuplicateOffersCnt" name="txtDuplicateOffersCnt" maxlength="2" value="" /> 
        </td>
      </tr>
		<tr><td>&nbsp;</td></tr>
	  <tr align="right">
        <td>
          <input type="button" name="btnOk" id="btnOk" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>" onclick="addDuplicateOfferscount();" />
		  <input type="button" name="btnCancel" id="btnCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>" onclick="toggleDialogOfferDuplicate('DuplicateNoofOffer', false);" />
        </td>
       </tr>	  
    </table>
  </div>
</div>
<div id="main" <% Sendb(IE6ScrollFix) %>>
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    If CustomerInquiry Then
      Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, ShowExpired, , , AdminUserID)
    Else
      Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, ShowExpired, , , AdminUserID)
    End If
    If (CriteriaMsg <> "") Then
      Session.Add("AdvSearchquery", SelectQuery1.ToString())
      Dim AdvSearchq As String = "True"
	  
      Send("<div id=""criteriabar""" & IIf(CriteriaError, " style=""background-color:red;""", "") & ">" & CriteriaMsg & "<a href=""javascript:editSearchCriteria();"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "]</a><a href=""Enhanced-extoffer-list.aspx" & IIf(CustomerInquiry, "?CustomerInquiry=1", "") & """ class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.clear", LanguageID) & "]</a><a id= ""hrefAdvSearchq"" href= ""XMLFeeds.aspx?AdvSearchQuery=" & AdvSearchq & "&amp;height=400&amp;width=600 "" title=""Offers"" class=""thickbox"" style=""padding-left:15px;"" onclick=""javascript:reInitializeSelectedItems();"" >[Actions]</a></div>")
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
          <a id="xidLink" onclick="handleIter('xidLink');" href="Enhanced-extoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ExtOfferID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine="&FilterEngine, "")) %>">
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
          <a id="idLink" onclick="handleIter('idLink');" href="Enhanced-extoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=AOLV.OfferID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine="&FilterEngine, "")) %>">
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
          <a id="engineLink" onclick="handleIter('engineLink');" href="Enhanced-extoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=InboundCRMEngineID&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine="&FilterEngine, "")) %>">
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
          <a id="nameLink" onclick="handleIter('nameLink');" href="Enhanced-extoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=AOLV.Name&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine="&FilterEngine, "")) %>">
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
          <a id="createLink" onclick="handleIter('createLink');" href="Enhanced-extoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=CreatedDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine="&FilterEngine, "")) %>">
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
          <a id="startLink" onclick="handleIter('startLink');" href="Enhanced-extoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ProdStartDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine="&FilterEngine, "")) %>">
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
          <a id="endLink" onclick="handleIter('endLink');" href="Enhanced-extoffer-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ProdEndDate&amp;SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&amp;CustomerInquiry=1", "")) %><% Sendb(IIf(bEnableRestrictedAccessToUEOfferBuilder, "&amp;filterengine="&FilterEngine, "")) %>">
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
        
        For j = 0 To (dst.Rows.Count - 1)
          arrList.Add(dst.Rows(j).Item("OfferID"))
        Next
        arrList.TrimToSize()
        ReDim OfferList(arrList.Count - 1)
        For j = 0 To arrList.Count - 1
          OfferList(j) = arrList.Item(j).ToString
        Next
        Statuses = Logix.GetStatusForOffers(OfferList, LanguageID)
        
        'While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
        i = 0
        While (i < dst.Rows.Count)
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
          MyCommon.QueryStr = "select Name from ExtCRMInterfaces where ExtInterfaceID=" & MyCommon.NZ(dst.Rows(i).Item("InboundCRMEngineID"), 0)
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
          i = i + 1
          ' Next
        End While
      %>
    </tbody>
  </table>
</div>
<script runat="server">
  
  
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
          FieldBuf.Append(" < '" & StartDate.ToString("yyyy-MM-dd") & "' ")
        Case 2 ' after
          FieldBuf.Append(" > '" & StartDate.ToString("yyyy-MM-dd") & "' ")
        Case 3 ' between
          FieldBuf.Append(" between '" & StartDate.ToString("yyyy-MM-dd") & "' and '" & EndDate.ToString("yyyy-MM-dd") & "' ")
        Case Else ' default to after
          FieldBuf.Append(" > '" & StartDate.ToString("yyyy-MM-dd") & "' ")
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
            PadLen = Convert.ToInt32(dt1.Rows.Item("paddingLength"))
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
<form id="frmIter" name="frmIter" method="post" action="#">
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
	      toggleDialogOfferDuplicate('DuplicateNoofOffer', false);
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

  var selectedItems = new Array();
  var DUP_OFFERS = 12;
  var MASSDEPLOY_OFFERS = 13;
  var NAVIGATETO_REPORTS = 14;
  var SEND_OUTBOUND = 15;
  var WFSTAT_PREVALIDATE = 16;
  var WFSTAT_POSTVALIDATE = 17;
  var WFSTAT_READYTODEPLOY = 18;
  var TRANSFER_OFFERS = 20;
  
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

function ClosePopUp(responseText){
                var performactionserrorElem = document.getElementById("performactionserror");
                performactionserrorElem.style.display = 'none';
                toggleDialog('performactions', false);
                document.location = 'Enhanced-extoffer-list.aspx';
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


function toggleDisabled(el, shown) {
    var closebtnElem = document.getElementById('TB_closeWindowButton');
    var all = el.getElementsByTagName('input')
    //all = all + el.getElementsByTagName('a')
    var inp, i = 0;
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
        }
    }
}

function handleAction() {

    var elemExecAction = document.getElementById('ExecAction');
    var elemdupoffer = document.getElementById('dupoffer');
    var actionele = document.getElementById('Actionitems');
    if (actionele != null && (actionele.value == '0' || actionele.value == '2' || actionele.value == '3' || actionele.value == '4' || actionele.value == '5' || actionele.value == '6')) {
        elemExecAction.style.visibility = 'visible';
        elemdupoffer.style.visibility = 'hidden';
    }
    else if (actionele.value == '1' || actionele.value == '7' || actionele.value == '8'){
        elemdupoffer.style.visibility = 'visible';
        elemExecAction.style.visibility = 'hidden';
    }
}

function handlemassDeployConditional(offerswithoutcon) {
    //value of max offers should be fetched from a CM_System option
    
    if (selectedItems.length > 0) {
      
        if (confirm(offerswithoutcon + ' Offers do not have any condition. Are you sure you want to continue?')) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true +'&OffersWithoutConditions=' + true;
        xmlhttpPost('folder-feeds.aspx?Action=DeployOffers', MASSDEPLOY_OFFERS, frmdata);
        }
     }
  }

function handlemassDeploy() {
    //value of max offers should be fetched from a CM_System option
    var maxoffers = <% Sendb(MyCommon.Fetch_SystemOption(152)) %>;
    if (selectedItems.length > 0) {

      if (maxoffers == 0 || selectedItems.length < maxoffers + 1){
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.deployofferwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
        frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true +'&OffersWithoutConditions=' + false;
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
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.navigatetoreportswarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
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
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.sendoutboundwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
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
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.prevalidatewarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
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
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.postvalidatewarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
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
        if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.readytodeploywarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>. ' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true;
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
            }
        }
    }
    self.xmlHttpReq.send(frmdata);
}

function cofirmsuccessmassdeployoffer(responseText){
    var performactionserrorElem = document.getElementById("performactionserror");
    var failedoffersdesc = responseText;
    var failedoffers = responseText;
    failedoffersdesc = failedoffersdesc.substring(failedoffersdesc.indexOf(":") + 1);
    failedoffers = failedoffers.substring(0, failedoffers.indexOf(':'));
    
      if ((responseText.substring(0, 2) != 'OK') && (responseText.substring(0, 2) != 'NO')){
        performactionserrorElem.style.display = 'block';
        performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.deployfail", LanguageID))%>';
        failedoffersdesc = failedoffersdesc.split('|');
        failedoffers = failedoffers.split(',');
        populaterow(failedoffers,failedoffersdesc);                
      }

      else if (responseText.substring(0, 2) == 'NO') {
              
       var offerswithoutcon = responseText.substring(3, responseText.indexOf(','));        
       handlemassDeployConditional(offerswithoutcon);
      }
      else if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'Enhanced-extoffer-list.aspx';
      }
    }

function handlenavtoreports(responseText) {

//    if (responseText.substring(0, 2) == 'OK') {
        window.location = 'reports-custom.aspx';
//    }
}

function handlesendoutbound(responseText){    
//      var performactionserrorElem = document.getElementById("performactionserror");
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
      // performactionserrorElem.style.display = 'none';
      // toggleDialog('performactions', false);
       //document.location = 'Enhanced-extoffer-list.aspx';
//      }
ClosePopUp("");

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
        document.location = 'Enhanced-extoffer-list.aspx';         
      }
      else if (responseText.substring(0, 2) == 'OK') {
       performactionserrorElem.style.display = 'none';
       toggleDialog('performactions', false);
       document.location = 'Enhanced-extoffer-list.aspx';
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
       document.location = 'Enhanced-extoffer-list.aspx';
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
       document.location = 'Enhanced-extoffer-list.aspx';
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
			   if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' ' + offerphrasetext + '. ' + ' <%Sendb(Copient.PhraseLib.Lookup("term.enteredDuplicateOffersCount", LanguageID))%> ' + dupOffersCntvalue + '.  <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
               frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=' + dupOffersCntvalue + '&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
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
		  if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.AssignFolderWarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>.' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
            frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=1&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
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
		   if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.copyofferwarning", LanguageID))%> ' + selectedItems.length + ' <%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>.' + '<%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
            frmdata = 'FolderIDs=' + folderlist + '&DuplicateCnt=1&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
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
      //  performactionserrorElem.style.display = 'block';
      //  performactionserrorElem.innerHTML = responseText;
    //}
    //else if (responseText.substring(0, 2) == 'OK') {
      //  performactionserrorElem.style.display = 'none';
      //  toggleDialog('performactions', false);
      //  document.location = 'Enhanced-extoffer-list.aspx';
    //}
	ClosePopUp("");
}

  function confirmsuccesstransoffer(responseText) {
  
    ClosePopUp("");
    //var performactionserrorElem = document.getElementById("performactionserror");
	//var failedoffers = responseText;
	//var failedoffersdesc = [];
    //if (responseText.trim() != 'OK') {
      //  performactionserrorElem.style.display = 'block';
      //  performactionserrorElem.innerHTML = '<%Sendb(Copient.PhraseLib.Lookup("folders.transferoffersfail", LanguageID))%>';
      //  failedoffers = failedoffers.split(',');
	  //for (var i = 0; i < failedoffers.length; i++) {
		//	  failedoffersdesc.push('<%Sendb(Copient.PhraseLib.Lookup("folders.transferofferserror", LanguageID))%>');
		//}
        //populaterow(failedoffers,failedoffersdesc);
		//document.getElementById('btnDupOffer').disabled = true; 
    //}
    //else if (responseText.trim() == 'OK') {
      //  performactionserrorElem.style.display = 'none';
       // toggleDialog('performactions', false);
        //document.location = 'Enhanced-extoffer-list.aspx';
    //}
}

  function TransferOffers(){    
    var actionele = document.getElementById('Actionitems');
    var destinationFolder = document.getElementById("folderList").value;
      if (selectedItems.length > 0) {
        var offerphrasetext = (selectedItems.length == 1) ? '<%Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID))%>' : '<%Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>';	  
	    if (confirm('<%Sendb(Copient.PhraseLib.Lookup("folders.transferofferswarning", LanguageID))%> ' + selectedItems.length + ' ' + offerphrasetext + '. ' + ' <%Sendb(Copient.PhraseLib.Lookup("folders.confirmaction", LanguageID))%>')) {
          frmdata = 'sFolder=0&dFolder=' + destinationFolder + '&ItemIDs=' + selectedItems.join(',') + '&FromOfferList=' + true + '&ActionItem=' + actionele.value;
          xmlhttpPost('folder-feeds.aspx?Action=TransferOffers', TRANSFER_OFFERS, frmdata);
		  document.getElementById('btnDupOffer').disabled = false; 
        }
	}  
      else {
        alert('<%Sendb(Copient.PhraseLib.Lookup("folders.NothingSelectedtoPerformAction", LanguageID))%>');
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
    Dim sFileName As String = "ExternalOfferList.xls"
    Dim dtExtSource As DataTable
    Dim dtExport As DataTable
    Dim dr As DataRow
    Dim drExport As DataRow
    Dim i64OfferId As Int64
    Dim sOfferStatus As String
    Dim oOfferStatus As Copient.LogixInc.STATUS_FLAGS
    Dim FolderList As String = ""
    Dim iPreviousExtSourceId As Integer = 0
    Dim iExtSourceId As Integer = 0
    Dim sExtSourceName As String = ""

    If dst.Rows.Count > 0 Then
      
      dtExport = New DataTable()
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

          iExtSourceId = MyCommon.NZ(dr.Item("InboundCRMEngineID"), 0)
          If iExtSourceId <> iPreviousExtSourceId Then
            iPreviousExtSourceId = iExtSourceId
            MyCommon.QueryStr = "select Name from ExtCRMInterfaces where ExtInterfaceID=" & iExtSourceId
            dtExtSource = MyCommon.LRT_Select()
            If dtExtSource.Rows.Count > 0 Then
              sExtSourceName = dtExtSource.Rows(0).Item(0)
            Else
              sExtSourceName = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
            End If
          End If

          drExport.Item("OfferID") = i64OfferId
          drExport.Item("XID") = MyCommon.NZ(dr.Item("ExtOfferId"), "")
          drExport.Item("ExternalSource") = sExtSourceName
          drExport.Item("Engine") = MyCommon.NZ(dr.Item("PromoEngine"), "")
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
