<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: reports-list.aspx 
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
  Dim AdminUserID As Long
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim SortText As String = "OfferID"
  Dim SortDirection As String
  Dim ShowExpired As String = ""
  Dim ShowExpiredCPE As String = ""
  Dim ShowReportList As Boolean = False
  Dim dst As System.Data.DataTable
  Dim OCDate As New DateTime
  Dim PrctSignPos As Integer
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "reports-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    PageNum = Server.HtmlEncode(Request.QueryString("pagenum"))
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  Send_HeadBegin("term.reports")
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
  Send_Subtabs(Logix, 8, 8)
  
  
  If (Logix.UserRoles.AccessReports = False) Then
    Send_Denied(1, "perm.admin-reports")
    GoTo done
  End If
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

  ' Do not initially run the query and hide the results table
  'If (Request.QueryString.Count > 0) Then
    If (Server.HtmlEncode(Request.QueryString("ShowExpired")) = "FALSE") Then
    ShowExpired = " and dateadd (d, 1, ProdEndDate) > getdate() "
    ShowExpiredCPE = " and dateadd (d, 1, EndDate) > getdate() "
  End If
  
    If (Server.HtmlEncode(Request.QueryString("SortText")) <> "") Then
        SortText = Server.HtmlEncode(Request.QueryString("SortText"))
  End If
  
    If (Server.HtmlEncode(Request.QueryString("pagenum")) = "") Then
        If (Server.HtmlEncode(Request.QueryString("SortDirection")) = "ASC") Then
      SortDirection = "DESC"
        ElseIf (Server.HtmlEncode(Request.QueryString("SortDirection")) = "DESC") Then
      SortDirection = "ASC"
    Else
      SortDirection = "DESC"
    End If
  Else
        SortDirection = Server.HtmlEncode(Request.QueryString("SortDirection"))
  End If
  
    If (Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
        If (Integer.TryParse(Server.HtmlEncode(Request.QueryString("searchterms")), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
        idSearchText = MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("searchterms")))
    PrctSignPos = idSearchText.IndexOf("%")
    If (PrctSignPos > -1) Then
      idSearch = -1
      idSearchText = idSearchText.Replace("%", "[%]")
    End If
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    MyCommon.QueryStr = "select * from (select OfferID,ExtOfferID,StatusFlag,Name,CreatedDate,ProdStartDate,ProdEndDate,C.Description,NULL as BuyerID from Offers as O with (NoLock) Left Join OfferCategories as C with (NoLock) on O.OfferCategoryID=C.OfferCategoryID where O.Deleted=0 and visible=1 and isnull(isTemplate,0)=0 and(OfferID=" & idSearch & " or ExtOfferID like N'%" & idSearchText & "%' or Name like N'%" & idSearchText & "%') " & ShowExpired & _
                        " union " & _
                        "select IncentiveID as OfferID,ClientOfferId as ExtOfferID,StatusFlag,IncentiveName as Name,CreatedDate,StartDate as ProdStartDate,EndDate as ProdEndDate, C.Description,buy.ExternalBuyerId as BuyerID from CPE_Incentives I with (NoLock) Left Join OfferCategories as C with (NoLock) on I.PromoClassId=C.OfferCategoryID left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId  where I.Deleted=0 and isnull(isTemplate,0)=0 and(IncentiveID=" & idSearch & " or ClientOfferID like N'%" & idSearchText & "%' or IncentiveName like N'%" & idSearchText & "%') " & ShowExpiredCPE & ") as t1 "
  Else
    MyCommon.QueryStr = "select * from (select OfferID,ExtOfferID,StatusFlag,Name,CreatedDate,ProdStartDate,ProdEndDate,C.Description,NULL as BuyerID from Offers as O with (NoLock) Left Join OfferCategories as C with (NoLock) on O.OfferCategoryID=C.OfferCategoryID where O.Deleted=0 and visible=1 and isnull(isTemplate,0)=0 " & ShowExpired & _
                        " union " & _
                        "select IncentiveID as OfferID,ClientOfferId as ExtOfferID,StatusFlag,IncentiveName as Name,CreatedDate,StartDate as ProdStartDate,EndDate as ProdEndDate,C.Description,buy.ExternalBuyerId as BuyerID from CPE_Incentives I with (NoLock) Left Join OfferCategories as C with (NoLock) on I.PromoClassId=C.OfferCategoryID left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId  where I.Deleted=0 and isnull(isTemplate,0)=0 " & ShowExpiredCPE & ") as t1 "
  End If
  
  ' check if banners are enabled
  If (BannersEnabled) Then
    MyCommon.QueryStr &= " where (OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                         "   or OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                         "                  inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                         "                  where AUB.AdminUserID = " & AdminUserID & ") ) "
  End If
  MyCommon.QueryStr &= " order by " & SortText & " " & SortDirection

  dst = MyCommon.LRT_Select
  sizeOfData = dst.Rows.Count
  i = linesPerPage * PageNum
  
    If (sizeOfData = 1 AndAlso Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "reports-detail.aspx?OfferID=" & dst.Rows(0).Item("OfferID") & "&Start=" & MyCommon.NZ(dst.Rows(0).Item("ProdStartDate"), "") & "&End=" & MyCommon.NZ(dst.Rows(0).Item("ProdEndDate"), "") & "&Status=" & Copient.PhraseLib.Lookup("offer.status" & MyCommon.NZ(dst.Rows(0).Item("StatusFlag"), 0), LanguageID))
  End If
  
    If (Server.HtmlEncode(Request.QueryString("ShowExpired")) = "FALSE") Then
    ShowExpired = "FALSE"
  Else
    ShowExpired = "TRUE"
  End If
  ShowReportList = True
%>
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.reports", LanguageID))%>
  </h1>
  <div id="controls">
    <form action="#" id="controlsform" name="controlsform">
      <%
        'BZ2079: UE-feature-removal - only displying the instant win button if CPE is installed
        If (Logix.UserRoles.AccessInstantWinReports) And (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
          Send("<input type=""button"" name=""instantwin"" id=""instantwin"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.instantwin", LanguageID) & """ onclick=""location.href='instantwin-list.aspx'"" />")
        End If
        Send("<input type=""button"" id=""custom"" name=""custom"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.custom", LanguageID) & """ onclick=""location.href='reports-custom.aspx'"" />")
        
        Dim ERPath As String = MyCommon.Get_Install_Path()
        If (ERPath <> "") Then
          If Not (Right(ERPath, 1) = "\") Then ERPath &= "\"
          ERPath &= "LogixWebRoot\logix\reports-enhanced-list.aspx"
          Try
            Dim FileInfo As System.IO.FileInfo = New System.IO.FileInfo(ERPath)
            If FileInfo.Exists Then
              Send("<input type=""button"" name=""enhanced"" id=""enhanced"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.enhanced", LanguageID) & """ onclick=""location.href='reports-enhanced-list.aspx'"" />")
            End If
          Catch ex As Exception
            MyCommon.Error_Processor(ex.Message & vbCrLf & "Error locating file at the path:" & ERPath & vbCrLf, ex.StackTrace, MyCommon.AppName, MyCommon.InstallationName, )
          End Try
        End If
      %>
    </form>
  </div>
</div>
<div id="main">
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection, ShowExpired)%>
  <% If (ShowReportList) Then%>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.reports", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" class="th-xid" scope="col">
          <a href="reports-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=ExtOfferID&SortDirection=<% Sendb(SortDirection) %>&ShowExpired=<% Sendb(ShowExpired) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID))%>
          </a>
          <%If SortText = "ExtOfferID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-bigid" scope="col">
          <a href="reports-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=OfferID&SortDirection=<% Sendb(SortDirection) %>&ShowExpired=<% Sendb(ShowExpired)%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%If SortText = "OfferID" Then
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
          <a href="reports-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=Name&SortDirection=<% Sendb(SortDirection) %>&ShowExpired=<% Sendb(ShowExpired)%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%If SortText = "Name" Then
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
          <a href="reports-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=Description&SortDirection=<% Sendb(SortDirection) %>&ShowExpired=<% Sendb(ShowExpired)%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>
          </a>
          <%If SortText = "Description" Then
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
          <a href="reports-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=ProdStartDate&SortDirection=<% Sendb(SortDirection) %>&ShowExpired=<% Sendb(ShowExpired)%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.starts", LanguageID))%>
          </a>
          <%If SortText = "ProdStartDate" Then
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
          <a href="reports-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=ProdEndDate&SortDirection=<% Sendb(SortDirection) %>&ShowExpired=<% Sendb(ShowExpired)%>">
            <% Sendb(Copient.PhraseLib.Lookup("term.ends", LanguageID))%>
          </a>
          <%If SortText = "ProdEndDate" Then
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
          <%If SortText = "StatusFlag" Then
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
        Dim assocName As String=""
      While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          Send("<tr class=""" & Shaded & """>")
          Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("ExtOfferID"), "&nbsp;") & "</td>")
          Send("  <td>" & dst.Rows(i).Item("OfferID") & "</td>")
                  If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(dst.Rows(i).Item("BuyerID"), "") <> "") Then
                assocName = "Buyer " + dst.Rows(i).Item("BuyerID").ToString() + " - " + MyCommon.NZ(dst.Rows(i).Item("Name"), "").ToString()
                Else
                assocName = MyCommon.NZ(dst.Rows(i).Item("Name"),Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
          Send("  <td><a href=""reports-detail.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & "&Start=" & Logix.ToShortDateString(MyCommon.NZ(dst.Rows(i).Item("ProdStartDate"), New Date(1900, 1, 1)), MyCommon) & "&End=" & Logix.ToShortDateString(MyCommon.NZ(dst.Rows(i).Item("ProdEndDate"), New Date(1900, 1, 1)), MyCommon) & "&Status=" & Copient.PhraseLib.Lookup("offer.status" & MyCommon.NZ(dst.Rows(i).Item("StatusFlag"), 0), LanguageID) & """>" & MyCommon.SplitNonSpacedString(assocName, 25) & "</a></td>")
          Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("Description"), "&nbsp;"), 20) & "</td>")
          If (Not IsDBNull(dst.Rows(i).Item("ProdStartDate"))) Then
            Send("  <td>" & Logix.ToShortDateString(dst.Rows(i).Item("ProdStartDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
          End If
          If (Not IsDBNull(dst.Rows(i).Item("ProdEndDate"))) Then
            Send("  <td>" & Logix.ToShortDateString(dst.Rows(i).Item("ProdEndDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.never", LanguageID) & "</td>")
          End If
          Send("  <td>" & Logix.GetOfferStatus(dst.Rows(i).Item("OfferID"), LanguageID) & "</td>")
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
      %>
    </tbody>
  </table>
  <% End If%>
</div>
<%
done:
  Send_BodyEnd("searchform", "searchterms")
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
