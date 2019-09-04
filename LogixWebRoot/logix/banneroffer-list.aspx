<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: banneroffer-list.aspx 
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
  Dim idSearch As String
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
  Dim WhereBuf As New StringBuilder()
  Dim AdvSearchSQL As String = ""
  Dim CriteriaMsg As String = ""
  Dim CriteriaTokens As String = ""
  Dim BannersEnabled As Boolean = False
  Dim IE6ScrollFix As String = ""
  Dim infoMessage As String = ""
  Dim CustomerInquiry As Boolean = False
  Dim TotalUsers As Integer = 0
  Dim restrictLinks As Boolean = False
  Dim Handheld As Boolean = False
  Dim BannerTotal As Integer = 0
  Dim BannerList As String = 0
  Dim count As Integer = 0
  Dim CountQuery As String = ""
  Dim SelectQuery As String = ""
  Dim SelectQueryOrderBy As String = ""
  Dim SelectSortDirection As String = ""
  Dim StartPoint As Long
  Dim EndPoint As Long
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "banneroffer-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  'Get BannerIDs for all Banner this user is in
  MyCommon.QueryStr = "select BAN.BannerID as BannerID from Banners BAN with (NoLock) " & _
                      "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                      "WHERE(BAN.Deleted = 0 And AdminUserID = " & AdminUserID & ")"
  rst = MyCommon.LRT_Select()
  BannerTotal = rst.Rows.Count
  count = 0
  If BannerTotal > 0 Then
    While (count < BannerTotal)
      If count = 0 Then
        BannerList = MyCommon.NZ(rst.Rows(count).Item("BannerID"), 0).ToString()
        'ElseIf count = BannerTotal Then
        '  BannerList = BannerList & MyCommon.NZ(rst.Rows(i).Item("BannerID"), 0).ToString()
      Else
        BannerList = BannerList & "," & MyCommon.NZ(rst.Rows(count).Item("BannerID"), 0).ToString()
      End If
      count = count + 1
    End While
  End If
  
  'If rst.Rows.Count > 0 Then
  '  For Each row In rst.Rows
  '    BannerID = MyCommon.Extract_Val(row.Item("BannerID"))
  '    BannerHash.Add(BannerID, MyCommon.NZ(row.Item("BannerID"), 0))
  '  Next
  'End If
  
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
    Send_HeadBegin("term.offers")
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
        Send("openPopup(""advanced-search.aspx?CustomerInquiry=1&tokens="" + tokenStr);")
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
    Send_Subtabs(Logix, 20, 5)
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
    ElseIf File.ContentType <> "text/xml" AndAlso File.ContentType <> "application/octet-stream" AndAlso File.ContentType <> "application/x-gzip" _
        AndAlso File.ContentType <> "application/x-gzip-compressed" AndAlso File.ContentType <> "application/gzip" _
        AndAlso File.ContentType <> "application/x-tar" Then
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
      CMS.AMS.CurrentRequest.Resolver.AppName = "banneroffer-List.aspx"
      Dim UEImport As Copient.ImportXMLUE = CMS.AMS.CurrentRequest.Resolver.Resolve(Of Copient.ImportXMLUE)()
      'Dim DpImport As New Copient.ImportXmlDP
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
        '      Response.AddHeader("Location", "DP-offer-gen.aspx?OfferID=" & sOfferId & "&infoMessage=" & infoMessage)
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
            Response.AddHeader("Location", "offer-gen.aspx?OfferID=" & sOfferId & "&infoMessage=" & infoMessage)
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
    
  Dim SortText As String = "AOLV.OfferID"
  Dim SortDirection As String = "DESC"
  Dim ShowExpired As String = ""
  Dim ShowActive As String = ""
  Dim PrctSignPos As Integer
  Dim FilterOffer As String
  
  If Request.QueryString("filterOffer") <> "" Then
    FilterOffer = Request.QueryString("filterOffer")
  Else
    FilterOffer = -1
  End If

  'If (BannerTotal > 0) Then
  '  BannerQuery = " BAN.BannerID is NULL or BAN.BannerID in (" & BannerList & " ) "
  'ElseIf (BannerTotal = 0) Then
  '  BannerQuery = " BAN.BannerID is NULL "
  'End If
  
  'ShowExpired = " where " & IIf(FilterOffer = -1, BannerQuery, "BAN.BannerID=" & FilterOffer & " ") & "and AOLV.deleted=0 " 'and isnull(AOLV.InboundCRMEngineID,0) = 0 "
  ShowExpired = " where " & IIf(FilterOffer = -1, " AOLV.deleted=0 ", "BAN.BannerID=" & FilterOffer & " and AOLV.deleted=0 ")  'and isnull(AOLV.InboundCRMEngineID,0) = 0 "

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
  
  MyCommon.QueryStr = "from AllOffersListview AOLV with (NoLock) " & _
                      "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID " & _
                      "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
					   "inner join AdminUserBanners AUB with (NoLock) on BO.BannerID = AUB.BannerID"
  
  If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
    If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
    If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
    If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))
    
    MyCommon.QueryStr += ShowExpired & " and IsNull(AOLV.isTemplate,0)=0 and AUB.AdminUserID = " & AdminUserID & WhereBuf.ToString & ShowActive
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
      MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 and AUB.AdminUserID = " & AdminUserID & ShowActive
      orderString = " order by " & SortText & " " & SortDirection
    Else
      MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired
      MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=0 and AUB.AdminUserID = " & AdminUserID & ShowActive
      orderString = " order by " & SortText & " " & SortDirection
    End If

  End If
  
  
  'At this point, MyCommon.QueryStr contains the FROM and WHERE clauses of the query.  
  'We need to build 2 versions of this query, one that will tell us the count of the total number of rows
  'and the second that will return the data for the sub (paginated) set of rows that we are going to return on the page
  'First we'll tack on what we need to query for the count of the total number of rows that meet the search & filter criteria  
  CountQuery = "select count(*) as NumRows " & MyCommon.QueryStr
  'Second we'll tack on what we need to query for the subset of data that needs to be displayed on this page
  'start by adding the names of the columns that we'll need for the page display.  This is not the completed SelectQuery, we'll add more to it later after we know the complete record count
  SelectQuery = "BAN.BannerID, BAN.Name as BannerName, AOLV.* " & MyCommon.QueryStr
  
  'before we run the CountQuery or the SelectQuery, we need to see if we are doing an export to Excel
  If (Request.QueryString("excel") <> "") Then
    MyCommon.QueryStr = "select BAN.BannerID, BAN.Name as BannerName, AOLV.* " & MyCommon.QueryStr & " order by " & SortText & " " & SortDirection
    dst = MyCommon.LRT_Select
    infoMessage = ExportListToExcel(dst, MyCommon, Logix)
    If infoMessage = "" Then
      GoTo done
    End If
    dst = Nothing
  End If
  
  'Run the Count Query to determine the total number of rows that meet the search & filter criteria
  MyCommon.QueryStr = CountQuery
  dst = MyCommon.LRT_Select
  sizeOfData = 0
  If dst.Rows.Count > 0 Then
    sizeOfData = dst.Rows(0).Item("NumRows")
  End If
  dst = Nothing
  

  
  'Now that we know the total record count, we can determine how we should slice up the SelectQuery.  If we are wanting a subset of records that are past the middle of the complete record set,
  'that it is faster to switch the ordering of the records around, and grab our subset from the beginning of the list.
  
  'If the start position of the subset of records we are looking for is passed the mid-point of total size of the record set ... then ... 
  If (sizeOfData > linesPerPage) AndAlso (((linesPerPage * PageNum) + 1) > (sizeOfData / 2)) Then
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

    SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SelectSortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & (StartPoint).ToString & " and " & (EndPoint).ToString & " order by " & SelectQueryOrderBy
  Else
    Send("<!-- building normal query -->")
    'add the SQL necessary to restrict the return set to only those rows that need to be displayed on this page
    SelectQuery = "select * from ( select ROW_NUMBER() OVER (order by " & SortText & " " & SortDirection & ") as RowNumber, " & SelectQuery & " ) as Table1 where RowNumber between " & ((linesPerPage * PageNum) + 1).ToString & " and " & (linesPerPage + (linesPerPage * PageNum)).ToString
  End If

  
  
  

  
  'Run the query that returns the subset of data to be displayed on this page 
  MyCommon.QueryStr = SelectQuery
  Send("<!-- Query=" & MyCommon.QueryStr & " -->")
  dst = MyCommon.LRT_Select
  
 
  If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
    If (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "CPE") Then
      sSummaryPage = "CPEoffer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
    ElseIf (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "Website") Then
      sSummaryPage = "web-offer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
    ElseIf (MyCommon.NZ(dst.Rows(0).Item("PromoEngine"), "") = "CAM") Then
      sSummaryPage = "CAM/CAM-offer-sum.aspx?OfferID=" & dst.Rows(0).Item("OfferID")
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
<div id="main"<% Sendb(IE6ScrollFix) %>>
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    If CustomerInquiry Then
      Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection, ShowExpired, , , AdminUserID)
    Else
      Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection, ShowExpired, , , AdminUserID)
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
          <a id="xidLink" onclick="handleIter('xidLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=ExtOfferID&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
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
          <a id="idLink" onclick="handleIter('idLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=AOLV.OfferID&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
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
          <a id="engineLink" onclick="handleIter('engineLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=PromoEngine&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
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
          <a id="nameLink" onclick="handleIter('nameLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=AOLV.Name&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
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
        <% If (FilterOffer = "-1") Then%>
        <th align="left" class="th-category" scope="col">
          <a id="bannerLink" onclick="handleIter('bannerLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=BAN.Name&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.banner", LanguageID))%>
          </a>
          <%
            If SortText = "BAN.Name" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <% Else%>
        <th align="left" class="th-category" scope="col">
          <a id="categoryLink" onclick="handleIter('categoryLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=ODescription&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>
          </a>
          <%
            If SortText = "ODescription" Then
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
        <th align="left" class="th-date" scope="col" style="display: none;">
          <a id="createLink" onclick="handleIter('createLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=CreatedDate&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
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
          <a id="startLink" onclick="handleIter('startLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=ProdStartDate&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
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
          <a id="endLink" onclick="handleIter('endLink');" href="banneroffer-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&filterOffer=<% sendb(FilterOffer)%>&SortText=ProdEndDate&SortDirection=<% Sendb(SortDirection) %><% Sendb(IIf(CustomerInquiry, "&CustomerInquiry=1", "")) %>">
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
          ''Only show the banners that the user has access to
          'CheckBanner = MyCommon.NZ(dst.Rows(i).Item("BannerID"), 0)
          'If CheckBanner = 0 Or BannerHash.ContainsKey(CheckBanner) Then
          RoidExtension = IIf(SearchMatchesROID, "&nbsp;&nbsp;<span style=""color:green;"">(" & Copient.PhraseLib.Lookup("term.roid", LanguageID) & ": " & MyCommon.NZ(dst.Rows(i).Item("RewardOptionID") & ")</span>", ""), "")
          Send("<tr class=""" & Shaded & """>")
          If CustomerInquiry Then
            MyCommon.QueryStr = "select AdminUserID from AdminUserOffers where OfferID=" & dst.Rows(i).Item("OfferID") & ";"
            rst = MyCommon.LRT_Select
            Send("  <td style=""text-align:center;""><a href=""javascript:openPopup('offer-favorite.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & "&CustomerInquiry=1')"">" & rst.Rows.Count & "/" & TotalUsers & "</a></td>")
          Else
            Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("ExtOfferID"), "&nbsp;"), 9, "<br />") & "</td>")
          End If
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("OfferID"), 0) & "</td>")
          Send("  <td>" & dst.Rows(i).Item("PromoEngine") & "</td>")
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
          If (FilterOffer = "-1") Then
            Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("BannerName"), Copient.PhraseLib.Lookup("term.all", LanguageID)), 15) & "</td>")
          Else
            Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("ODescription"), "&nbsp;"), 15) & "</td>")
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
          ' Next
          'End If
          i = i + 1
        End While
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
        OptionType = "contains"
      Case 2 ' exact
        OptionType = "="
      Case 3 ' starts with
        OptionType = "starts with"
      Case 4 ' ends with
        OptionType = "ends with"
      Case 5 ' excludes
        OptionType = "excludes"
      Case 6 ' is
        OptionType = "is"
      Case 7 ' is not
        OptionType = "is not"
      Case Else ' default to contains
        OptionType = "contains"
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
    
    if (elemAdv != null && elemAdv.value !="") {
      if (currentURL.indexOf('filterOffer=') > -1) {
        newURL = currentURL.replace(/filterOffer=[0-9]?/g, 'filterOffer=' + newFilter);
        newURL = newURL.replace(/pagenum=[0-9]+/g, '');
      } else {
        if (currentURL.indexOf("&") > -1) {
          newURL = currentURL + "&filterOffer=" + newFilter;
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

  Private Function ExportListToExcel(ByRef dst As DataTable, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc) As String
    Dim bStatus As Boolean
    Dim sMsg As String = ""
    Dim CmExport As New Copient.ExportXml
    Dim sFileFullPath As String
    Dim sFullPathFileName As String
    Dim sFileName As String = "BannerOfferList.xls"
    Dim dtExport As DataTable
    Dim dr As DataRow
    Dim drExport As DataRow
    Dim i64OfferId As Int64
    Dim sOfferStatus As String
    Dim oOfferStatus As Copient.LogixInc.STATUS_FLAGS


    If dst.Rows.Count > 0 Then
      
      dtExport = New DataTable()
      dtExport.Columns.Add("OfferID", Type.GetType("System.Int64"))
      dtExport.Columns.Add("XID", Type.GetType("System.String"))
      dtExport.Columns.Add("Engine", Type.GetType("System.String"))
      dtExport.Columns.Add("Banner", Type.GetType("System.String"))
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
          drExport.Item("OfferID") = i64OfferId
          drExport.Item("XID") = MyCommon.NZ(dr.Item("ExtOfferId"), "")
          drExport.Item("Engine") = MyCommon.NZ(dr.Item("PromoEngine"), "")
          drExport.Item("Name") = MyCommon.NZ(dr.Item("Name"), "")
          drExport.Item("Banner") = MyCommon.NZ(dr.Item("BannerName"), "")
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
