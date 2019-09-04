<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
    ' *****************************************************************************
    ' * FILENAME: temp-list.aspx 
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
    Dim dst As System.Data.DataTable
    Dim row As System.Data.DataRow
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
    Dim BannersEnabled As Boolean = False
    Dim IE6ScrollFix As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim bStoreUser As Boolean = False
    Dim sValidLocIDs As String = ""
    Dim sValidSU As String = ""
    Dim wherestr As String = "" 
    Dim sJoin As String = "" 
    Dim iLen As Integer = 0
    Dim rst As DataTable
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "temp-list.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

  'Store User
  If(MyCommon.Fetch_CM_SystemOption(131) = "1") Then
    MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    iLen = rst.Rows.Count
    If iLen > 0 Then
      bStoreUser = True
      sValidSU = AdminUserID
      For i=0 to (iLen-1)
        If i=0 Then 
          sValidLocIDs = rst.Rows(0).Item("LocationID")
        Else 
          sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
        End If
      Next
    
      MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
      rst = MyCommon.LRT_Select
      iLen = rst.Rows.Count
      If iLen > 0 Then
        For i=0 to (iLen-1)
          sValidSU &= "," & rst.Rows(i).Item("UserID") 
        Next
      End If
    End If
  End If
    
    PageNum = Request.QueryString("pagenum")
    If PageNum < 0 Then PageNum = 0
    MorePages = False
  
    Send_HeadBegin("term.templates")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
%>
<style type="text/css">
    #controls form
    {
        display: inline !important;
    }
</style>
<%  
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(11)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 2)
    Send_Subtabs(Logix, 20, 2)
  
    If (Logix.UserRoles.AccessTemplates = False) Then
        Send_Denied(1, "perm.offers-access-templates")
        GoTo done
    End If
  
  
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
    Dim SortText As String = "AOLV.OfferID"
    Dim SortDirection As String = "DESC"
    Dim ShowExpired As String = ""
    Dim ShowActive As String = ""
    Dim PrctSignPos As Integer
    Dim FilterOffer As String
    Dim enableBuyerIdSearch As Boolean = IIf(MyCommon.Fetch_UE_SystemOption(169) = "1", True, False)
    
  
    'There's currently no filtering on templates, so I'm hard-coding this to zero:
    'FilterOffer = Request.QueryString("filterOffer")
    FilterOffer = "0"
    If (FilterOffer = "") Then FilterOffer = "1"
    If (FilterOffer = "0" OrElse FilterOffer = "3") Then
        ShowExpired = " where AOLV.deleted=0 "
    ElseIf (FilterOffer = "1") Then
        ShowExpired = " where AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() "
    Else
        ShowExpired = " where AOLV.deleted=0 and dateadd (d, 1, AOLV.ProdEndDate) > getdate() "
    End If

    
    'If UE Engine is installed and User is assoiated with any Buyer and if user is not having View Offer Regardless of Buyer Permission, list User-Buyer specific Offers
    If (MyCommon.IsEngineInstalled(9) AndAlso MyCommon.IsUserAssociatedWithAnyBuyer(AdminUserID) AndAlso Not (Logix.UserRoles.ViewOffersRegardlessBuyer)) Then
        ShowExpired = ShowExpired & " and ( AOLV.BuyerId in (select BuyerId from BuyerRoleUsers where AdminUserID=" & AdminUserID & "))"
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
                        "C.Description as ODescription, PE.Description as PromoEngine, PE.PhraseID as EnginePhraseID from Offers as O with (NoLock) " & _
                        "left join OfferCategories as C with (NoLock) on O.OfferCategoryID=C.OfferCategoryID " & _
                        "left join PromoEngines as PE on PE.EngineID=O.EngineID " & _
                        "where O.Deleted=0 and Visible=1 and isnull(isTemplate,0)=1 "
    If (BannersEnabled) Then
        MyCommon.QueryStr = "select BAN.BannerID, BAN.Name as BannerName, AOLV.* from AllOffersListview AOLV with (NoLock) " & _
                            "left join BannerOffers BO with (NoLock)on BO.OfferID = AOLV.OfferID " & _
                            "left join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID "
    Else
        MyCommon.QueryStr = "select AOLV.* from AllOffersListview AOLV with (NoLock) "
    End If
    
    If bStoreUser Then
      wherestr = " and CreatedByAdminID in (" & sValidSU & ") " 
    End If
    
    If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
        If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
        If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
        If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))
    
        MyCommon.QueryStr += ShowExpired & " and IsNull(AOLV.isTemplate,0)=1 " & WhereBuf.ToString & ShowActive & wherestr
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
            MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired & " and(AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%'"
           
            ''If buyerid sysopt is enabled and ue engine is installed append  condition to include the results for that buyer id
            If (enableBuyerIdSearch) AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                MyCommon.QueryStr = MyCommon.QueryStr & " or AOLV.ExternalbuyerId like N'%" & idSearchText & "%' )"
            Else
                MyCommon.QueryStr = MyCommon.QueryStr & ")"
            End If
             
            
           
            MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=1 " & ShowActive & wherestr
            orderString = " order by " & SortText & " " & SortDirection
            
            
            'MyCommon.QueryStr = MyCommon.QueryStr & " or( AOLV.BuyerId=" & idSearch & "and Deleted=0)"   ) "
            
            
        Else
            MyCommon.QueryStr = MyCommon.QueryStr & ShowExpired
            MyCommon.QueryStr = MyCommon.QueryStr & " and isnull(AOLV.isTemplate,0)=1 " & ShowActive & wherestr
            orderString = " order by " & SortText & " " & SortDirection
        End If
    
        ' check if banners are enabled
        If (BannersEnabled) Then
            MyCommon.QueryStr &= " and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                                 " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                                 "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                                 "                     where AUB.AdminUserID = " & AdminUserID & ") ) "
        End If
        If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not Logix.UserRoles.CreateUEOffers) Then
            MyCommon.QueryStr &=" and AOLV.EngineID =0 "
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
                                " where AOLV.IsTemplate=1 and (AOLV.OfferID not in (select OfferID from BannerOffers BO with (NoLock)) " & _
                                " or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                                "                     inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID" & _
                                "                     where AUB.AdminUserID = " & AdminUserID & ") ) " & AdvSearchSQL
            If (Request.QueryString("searchterms") <> "") Then
                MyCommon.QueryStr &= " and (AOLV.OfferID=" & idSearch & " or AOLV.ExtOfferID='" & idSearch & "' or (AOLV.RewardOptionID=" & idSearch & " and AOLV.RewardOptionID<>-1) or AOLV.Name like N'%" & idSearchText & "%') "
            End If
        Else
            MyCommon.QueryStr = "select * from AllActiveOffersListView AOLV where AOLV.IsTemplate=1 "
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
                             "union " & _
                             "select Distinct STO.OfferID as OfferID from CM_ST_Offers STO with (NoLock) " & _
                             "inner join Offers O with (NoLock) on O.OfferID = STO.OfferID " & _
                             "inner join OfferLocUpdate OLU with (NoLock) on OLU.OfferID = STO.OfferID " & _
                             "inner join Locations LOC with (NoLock) on LOC.LocationID = OLU.LocationID and LOC.TestingLocation = 0 " & _
                             "where STO.Deleted=0 and O.Deleted=0 and LOC.Deleted=0 " & _
                             " order by " & SortText & " " & SortDirection
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
        <% Sendb(Copient.PhraseLib.Lookup("term.offertemplates", LanguageID))%>
    </h1>
    <div id="controls">
        <form action="offer-new.aspx" id="controlsform" name="controlsform">
        <input type="hidden" name="NewTemplate" value="Yes" />
        <%
            If (Logix.UserRoles.CreateTemplate) Then
                Send_New()
            End If
        %>
        </form>
    </div>
</div>
<div id="main" <% Sendb(IE6ScrollFix) %>>
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection, ShowExpired)%>
    <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID)) %>">
        <thead>
            <tr>
                <th align="left" class="th-bigid" scope="col">
                    <a id="idLink" onclick="handleIter('idLink');" href="temp-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=AOLV.OfferID&amp;SortDirection=<% Sendb(SortDirection) %>">
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
                    <a id="engineLink" onclick="handleIter('engineLink');" href="temp-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=PromoEngine&amp;SortDirection=<% Sendb(SortDirection) %>">
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
                <%  If (MyCommon.Fetch_UE_SystemOption(169) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then%>
                <th align="left" class="th-name" scope="col">
                    <a id="A1" onclick="handleIter('idLink');" href="temp-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=AOLV.BuyerId&amp;SortDirection=<% Sendb(SortDirection) %>">
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
                    <a id="nameLink" onclick="handleIter('nameLink');" href="temp-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=Name&amp;SortDirection=<% Sendb(SortDirection) %>">
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
                <% If (BannersEnabled) Then%>
                <th align="left" class="th-category" scope="col">
                    <a id="bannerLink" onclick="handleIter('bannerLink');" href="temp-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=BAN.Name&amp;SortDirection=<% Sendb(SortDirection) %>">
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
                    <a id="categoryLink" onclick="handleIter('categoryLink');" href="temp-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=ODescription&amp;SortDirection=<% Sendb(SortDirection) %>">
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
                <th align="left" class="th-date" scope="col">
                    <a id="createLink" onclick="handleIter('createLink');" href="temp-list.aspx?searchterms=<%sendb(Server.HtmlEncode(Request.QueryString("searchterms"))) %>&amp;filterOffer=<% sendb(FilterOffer)%>&amp;SortText=CreatedDate&amp;SortDirection=<% Sendb(SortDirection) %>">
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
                    Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("OfferID"), 0) & "</td>")
                    Send("  <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(dst.Rows(i).Item("EnginePhraseID"), 0), LanguageID) & "</td>")
                    If (MyCommon.Fetch_UE_SystemOption(169) = "1") AndAlso (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE)) Then
                        Send(" <td>" & MyCommon.NZ(dst.Rows(i).Item("ExternalBuyerId"), "") & "</td>")
                    End If
                    
                    If (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CPE") Then
                        Sendb("  <td><a href=""CPEoffer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a></td>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Website") Then
                        Sendb("  <td><a href=""web-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a></td>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "Email") Then
                        Sendb("  <td><a href=""email-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a></td>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "CAM") Then
                        Sendb("  <td><a href=""CAM/CAM-offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a></td>")
                    ElseIf (MyCommon.NZ(dst.Rows(i).Item("PromoEngine"), "") = "UE") Then
                        Sendb("  <td><a href=""UE/UEoffer-gen.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a>")
                    Else
                        Sendb("  <td><a href=""offer-sum.aspx?OfferID=" & dst.Rows(i).Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 25) & RoidExtension & "</a></td>")
                    End If
                    If (BannersEnabled) Then
                        Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("Banners"), Copient.PhraseLib.Lookup("term.all", LanguageID)), 15) & "</td>")
                    Else
                        Send("  <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("ODescription"), "&nbsp;"), 15) & "</td>")
                    End If
                    If (Not IsDBNull(dst.Rows(i).Item("CreatedDate"))) Then
                        Send("  <td>" & Logix.ToShortDateString(dst.Rows(i).Item("CreatedDate"), MyCommon) & "</td>")
                    Else
                        Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                    End If
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
            Case Else ' default to contains
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
