<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: product-inquiry.aspx 
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
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim rst3 As DataTable
    Dim row3 As DataRow
    Dim amtMax As String
    Dim description As String = ""
    Dim productGroups() As String
    Dim productGroupsIDs() As String
    Dim productGroupList As String
    Dim prodGroups As String
    Dim promotionsList As String
    Dim promotions() As String
    Dim promotionsIDs() As String
    Dim result As Boolean = False
    Dim shaded As Boolean = True
    Dim PromoEngine As Integer = 0
    Dim Restricted As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim SearchTerms As String = ""
    Dim TempDate As Date
	Dim FilterOffer As String
    Dim SOfferStatus As String
    Dim oOfferStatus As Copient.LogixInc.STATUS_FLAGS
  
    Dim URLtrackBack As String = ""
    Dim inCardNumber As String = ""
    ' tack on the customercare remote links if needed
    Dim extraLink As String = ""
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "product-inquiry.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    ' lets check the logged in user and see if they are to be restricted to this page
    MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                        "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                        "where AU.AdminUserID=" & AdminUserID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        If (rst.Rows(0).Item("prestrict") = True) Then
            ' ok we got in here then we need to restrict the user from seeing any other pages
            Restricted = True
        End If
    End If
  
    If (Request.QueryString("searchterms") <> "") Then
        Dim idLength As Integer = 0
        MyCommon.QueryStr = "select paddingLength from producttypes where producttypeid=1"
        rst = MyCommon.LRT_Select
        If rst IsNot Nothing Then
            idLength = Convert.ToInt32(rst.Rows(0).Item("paddingLength"))
        End If
        SearchTerms = Request.QueryString("searchterms").PadLeft(idLength, "0")
        PromoEngine = Request.QueryString("EngineID")
		FilterOffer = Request.QueryString("FilterOffer")
        ' someone sent us a upc lets dig
        MyCommon.QueryStr = "select ProductID, ProductTypeID, Description from Products with (NoLock) " & _
                            "where ExtProductID='" & MyCommon.Parse_Quotes(SearchTerms) & "';"
        rst = MyCommon.LRT_Select()
        If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
                description = MyCommon.NZ(row.Item("Description"), "")
                ' ok in here we have the proudctid for what they typed in lets dig for groups its in
                MyCommon.QueryStr = "select distinct PG.ProductGroupID,PG.buyerid, PG.Name from ProdGroupItems as PGI with (NoLock) " & _
                            "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=PGI.ProductGroupID " & _
                            "where PGI.Deleted=0 and PG.Deleted=0 and ProductID=" & row.Item("ProductID") & ";"
                rst2 = MyCommon.LRT_Select
                If rst2.Rows.Count > 0 Then
                    Dim x As Integer
                    x = 0
                    For Each row2 In rst2.Rows
                        ' ok here we should have the productgroup list
                        If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(row2.Item("Buyerid"))) Then
                            Dim buyerid As Int32 = row2.Item("buyerid")
                            Dim externalbuyerid As String = MyCommon.GetExternalBuyerId(buyerid)
                            prodGroups = MyCommon.NZ("Buyer "&externalbuyerid & " - " & row2.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & Chr(30) & prodGroups
                        Else
                            prodGroups = MyCommon.NZ(row2.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & Chr(30) & prodGroups
                        End If
                        productGroupList = MyCommon.NZ(row2.Item("ProductGroupID"), 0) & Chr(30) & productGroupList
                    Next
                    If (prodGroups.Trim <> "") Then productGroups = Split(prodGroups, Chr(30))
                    If (productGroupList.Trim <> "") Then productGroupsIDs = Split(productGroupList, Chr(30))
                    result = True
                End If
            Next
            If result = False Then
                infoMessage = Copient.PhraseLib.Lookup("product-inquiry.notused", LanguageID)
            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("product-inquiry.notfound", LanguageID)
        End If
    End If
  
    ' lets check the logged in user and see if they are to be restricted to this page
    MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                        "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                        "where AU.AdminUserID=" & AdminUserID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        If (rst.Rows(0).Item("prestrict") = True) Then
            ' ok we got in here then we need to restrict the user from seeing any other pages
            Restricted = True
        End If
    End If
  
    If (Request.QueryString("mode") = "summary") Then
        URLtrackBack = Request.QueryString("exiturl")
        inCardNumber = Request.QueryString("cardnumber")
        extraLink = "&mode=summary&exiturl=" & URLtrackBack & "&cardnumber=" & inCardNumber
    End If
  
    Send_HeadBegin("term.productinquiry")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld, Restricted)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    If (Not Restricted) Then
        Send_Tabs(Logix, 4)
        Send_Subtabs(Logix, 40, 2)
    Else
        Send_Subtabs(Logix, 92, 2, LanguageID, ID, extraLink)
    End If
  
    If (Logix.UserRoles.AccessProductInquiry = False) Then
        Send_Denied(1, "perm.product-access")
        GoTo done
    End If
%>
<form id="mainform" name="mainform" action="">
<div id="intro">
    <h1 id="title">
        <% Sendb(Copient.PhraseLib.Lookup("term.productinquiry", LanguageID))%>
    </h1>
    <div id="controls">
        <%
            'If MyCommon.Fetch_SystemOption(75) Then
            '  If (Logix.UserRoles.AccessNotes) Then
            '    Send_NotesButton(6, 0, AdminUserID)
            '  End If
            'End If
        %>
    </div>
</div>
<div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% Sendb(Copient.PhraseLib.Lookup("product-inquiry.main", LanguageID))%>
    <br />
    <br class="half" />
    <input type="text" id="searchterms" name="searchterms" maxlength="100" value="" />
    <%
        If (Restricted) Then
            Send("<input type=""hidden"" id=""mode"" name=""mode"" value=""summary"" />")
            Send("<input type=""hidden"" id=""exiturl"" name=""exiturl"" value=" & URLtrackBack & " />")
            Send("<input type=""hidden"" id=""cardnumber"" name=""cardnumber"" value=" & inCardNumber & " />")
        End If
      
        MyCommon.QueryStr = "select EngineID,Description,PhraseID,DefaultEngine from PromoEngines with (NoLock) where Installed='True' and EngineID in (0,1,2,9);"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            Send("<select id=""EngineID"" name=""EngineID"">")
            For Each row In rst.Rows
                Sendb("  <option value=""" & row.Item("EngineID") & """" & IIf(row.Item("DefaultEngine") = 1, " selected=""selected""", "") & ">")
                If MyCommon.NZ(row.Item("PhraseID"), 0) > 0 Then
                    Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID))
                Else
                    Sendb(row.Item("Description"))
                End If
                Send("</option>")
            Next
            Send("</select>")
        End If
			Send("   <select id=""filterOffer"" name=""filterOffer"" onchange=""handleFilterRegEx(this.options[this.selectedIndex].value);"">")
            Send("    <option value=""2""" & IIf(FilterOffer = "2", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showonlyactive", LanguageID) & "</option>")
            Send("    <option value=""1""" & IIf(FilterOffer = "1", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall-noexpired", LanguageID) & "</option>")
            Send("    <option value=""0""" & IIf(FilterOffer = "0", " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & "</option>")
            Send("  </select>")
    %>
    <input type="submit" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>" />
    <div id="results">
        <%
            If (result) Then
                Send("<br />")
                Send("<table class=""list"" summary=""" & Copient.PhraseLib.Lookup("term.results", LanguageID) & """>")
                Send(" <thead>")
                Send("  <tr>")
                Send("   <th class=""th-longid"" scope=""col"">" & Copient.PhraseLib.Lookup("term.product", LanguageID) & "</th>")
                Send("   <th class=""th-groups"" scope=""col"">" & Copient.PhraseLib.Lookup("product-inquiry.groups", LanguageID) & "</th>")
                Send("   <th class=""th-offers"" scope=""col"">" & Copient.PhraseLib.Lookup("product-inquiry.offers", LanguageID) & "</th>")
                Send("  </tr>")
                Send(" </thead>")
                Send(" <tbody>")
                Dim z As Integer
                Dim x As Integer
                Send("  <tr>")
                Send("   <td class=""shadeddark"" valign=""top"" rowspan=""" & productGroups.GetUpperBound(0) & """ title=""" & description & """>" & SearchTerms & "</td>")
                For z = 0 To productGroups.GetUpperBound(0) - 1
                    If (shaded) Then
                        If (Restricted) Then
                            Send("   <td class=""shaded"" valign=""top"">" & MyCommon.SplitNonSpacedString(productGroups(z), 25) & "</td>")
                        Else
                            Send("   <td class=""shaded"" valign=""top""><a href=""pgroup-edit.aspx?ProductGroupID=" & IIf(productGroupsIDs(z).Trim <> "", productGroupsIDs(z), "-1") & """>" & MyCommon.SplitNonSpacedString(productGroups(z), 25) & "</a></td>")
                        End If
                        Send("   <td class=""shaded"" valign=""top"">")
                        shaded = False
                    Else
                        If (Restricted) Then
                            Send("   <td  valign=""top"">" & MyCommon.SplitNonSpacedString(productGroups(z), 25) & "</td>")
                       Else
                            Send("   <td  valign=""top""><a href=""pgroup-edit.aspx?ProductGroupID=" & IIf(productGroupsIDs(z).Trim <> "", productGroupsIDs(z), "-1") & """>" & MyCommon.SplitNonSpacedString(productGroups(z), 25) & "</a></td>")
                        End If
                        Send("   <td valign=""top"">")
                        shaded = True
                    End If
                    ' ok we have a list of products and ids, now we dig for the offers they are used in            
                    ' lets spit out the linked ones first
                    If (PromoEngine = 0) Or (PromoEngine = 1) Then
                        MyCommon.QueryStr = "select O.Name, O.OfferID, O.ProdStartDate, O.ProdEndDate from OfferConditions as OC with (NoLock) " & _
                                            "left join Offers as O with (NoLock) on O.OfferID=OC.OfferID " & _
                                            "left join OfferRewards as ord with (NoLock) on O.OfferID=ord.OfferID " & _
                                            "left join ProductGroups as PG with (NoLock) on ord.ProductGroupID=pg.ProductGroupID " & _
                                            "where ConditionTypeID=2 and O.Deleted=0 and ord.ProductGroupID=" & IIf(productGroupsIDs(z).Trim <> "", productGroupsIDs(z), "-1") & " " & _
                                            " union " & _
                                            "select O.Name, O.OfferID, O.ProdStartDate, O.ProdEndDate from OfferRewards as OC with (NoLock) " & _
                                            "left join Offers as O with (NoLock) on O.OfferID=OC.OfferID " & _
                                            "left join RewardTiers as RT with (NoLock) on OC.RewardID=RT.RewardID " & _
                                            "where O.EngineID=" & PromoEngine & " and OC.RewardTypeID=1 and O.Deleted=0 and OC.ProductGroupID=" & IIf(productGroupsIDs(z).Trim <> "", productGroupsIDs(z), "-1") & " " & _
                                            "group by O.Name,O.OfferID, O.ProdStartDate, O.ProdEndDate;"
                    ElseIf (PromoEngine = 2) Or (PromoEngine = 9) Then
                        ' AMSPS-2996 : When doing a product inquiry for items that are in the exclude group , the offer does not show in the inquiry
                        MyCommon.QueryStr = "select IncentiveName as Name, I.IncentiveID as OfferID, PG.ProductGroupID, i.EndDate as ProdEndDate," & _
                                            " I.StartDate as ProdStartDate,buy.ExternalBuyerId as BuyerID from CPE_Incentives as I with (NoLock) " & _
                                            " left join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID " & _
                                            " left join CPE_IncentiveProductGroups as PG with (NoLock) on RO.RewardOptionID=PG.RewardOptionID " & _
                                            " left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                            " where i.Deleted=0 and RO.Deleted=0 and PG.Deleted=0 " & _
                                            " and PG.ProductGroupID=" & IIf(productGroupsIDs(z).Trim <> "", productGroupsIDs(z), "-1") & _
                                            " and EngineID=" & PromoEngine & _
                                            " UNION " & _
                                            " select IncentiveName as Name, I.IncentiveID as OfferID, PCPG.ProductGroupID, i.EndDate as ProdEndDate," & _
                                            " I.StartDate as ProdStartDate,buy.ExternalBuyerId as BuyerID from CPE_Incentives as I with (NoLock)  " & _
                                            " left join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID  left join " & _
                                            " CPE_IncentiveProductGroups as IPG with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID  inner join" & _
                                            " ProductConditionProductGroups as PCPG with (NoLock) on PCPG.IncentiveProductGroupID=IPG.IncentiveProductGroupID  " & _
                                            " left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId  where i.Deleted=0 and " & _
                                            " RO.Deleted = 0 And IPG.Deleted = 0 And PCPG.ProductGroupID = " & IIf(productGroupsIDs(z).Trim <> "", productGroupsIDs(z), "-1") & _
                                            " and EngineID=" & PromoEngine & ";"
                    End If
                    rst = MyCommon.LRT_Select
                    Send("    <table style=""width:100%;"" summary=""" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & """>")
                    'If (rst.Rows.Count > 0) Then
                    '    Send("     <tr>")
                    '    Send("      <td>" & Copient.PhraseLib.Lookup("term.usedby", LanguageID) & ":</td>")
                    '    Send("      <td align=""right"">" & Copient.PhraseLib.Lookup("term.maxreward", LanguageID) & "</td>")
                    '    Send("     </tr>")
                    'End If
                    If rst.Rows.Count > 0 Then
                        For Each row In rst.Rows
								Dim bDisplayoffer As Boolean = False
                            
                                SOfferStatus = Logix.GetOfferStatus(row.Item("OfferID"), LanguageID, oOfferStatus)
                                                            
                                If (FilterOffer = "2") And (oOfferStatus = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
                                    bDisplayoffer = True
                                ElseIf (FilterOffer = "1") And (oOfferStatus <> Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED) Then
                                    bDisplayoffer = True
                                ElseIf (FilterOffer = "0") Then
                                    bDisplayoffer = True
                                End If
                                If bDisplayoffer = True Then
                            'We know the offer ID; let's go digging for the max reward amount for the offer
                            If (PromoEngine = 0) Or (PromoEngine = 1) Then
                                MyCommon.QueryStr = "select top 1 RewardAmount as Amt, OfferID, RewardAmountTypeID as ATID, D.SVLinkID as SVProgramID " & _
                                                    "from OfferRewards as OFR with (NoLock) " & _
                                                    "left join RewardTiers as RT with (NoLock) on OFR.RewardID=RT.RewardID " & _
                                                    "left join Discounts as D with (NoLock) on D.DiscountID=OFR.LinkID " & _
                                                    "where (RewardTypeID=1 and Deleted=0 and OfferID=" & row.Item("OfferID") & ") " & _
                                                    "and OFR.ProductGroupID=" & IIf(productGroupsIDs(z).Trim <> "", productGroupsIDs(z), "-1") & " " & _
                                                    "order by ATID, Amt DESC;"
                            ElseIf (PromoEngine = 2) Or (PromoEngine = 9) Then
                                MyCommon.QueryStr = "select D.DiscountAmount as Amt, I.IncentiveID as OfferID, D.AmountTypeID as ATID, D.SVProgramID " & _
                                                    " from CPE_Incentives as I with (NoLock) " & _
                                                    " left join CPE_Rewardoptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID " & _
                                                    " left join CPE_Deliverables as Del with (NoLock) on Del.RewardOptionID=RO.RewardOptionID " & _
                                                    " left join CPE_Discounts as D with (NoLock) on D.DiscountID=Del.OutputID " & _
                                                    " where (I.Deleted=0 and RO.Deleted=0) and DiscountTypeID=1 and DeliverableTypeID=2 and I.IncentiveID=" & row.Item("OfferID") & _
                                                    ";"
                            End If
                            rst3 = MyCommon.LRT_Select
                            amtMax = ""
                            If (rst3.Rows.Count > 0) Then
                                amtMax = MyCommon.NZ(rst3.Rows(0).Item("Amt"), "")
                            End If
                            Send("     <tr>")
                            If (PromoEngine = 0) Or (PromoEngine = 1) Then
                                If (Restricted) Then
                                    Send("      <td>" & row.Item("OfferID") & ": " & MyCommon.SplitNonSpacedString(row.Item("Name"), 25) & "")
                                Else
                                    Send("      <td>" & row.Item("OfferID") & ": <a href=""offer-sum.aspx?OfferID=" & row.Item("OfferID") & """>" & MyCommon.SplitNonSpacedString(row.Item("Name"), 25) & "</a>")
                                End If
                            ElseIf (PromoEngine = 2) Or (PromoEngine = 9) Then
                                Dim Name As String=""
                                 If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                                    Name = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.SplitNonSpacedString(row.Item("Name"), 25).ToString()
                                Else
                                    Name = MyCommon.NZ(MyCommon.SplitNonSpacedString(row.Item("Name"), 25), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                                End If
                                If (Restricted) Then
                                    Send("      <td>" & row.Item("OfferID") & ": " & Name & "")
                                Else
                                    Send("      <td>" & row.Item("OfferID") & ": <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & Name & "</a>")
                                End If
                            End If
                            If (Not IsDBNull(row.Item("ProdStartDate")) AndAlso Not IsDBNull(row.Item("ProdEndDate"))) Then
                                Sendb("        <br />&nbsp;&nbsp;" & Logix.ToShortDateString(row.Item("ProdStartDate"), MyCommon) & " - " & Logix.ToShortDateString(row.Item("ProdEndDate"), MyCommon))
                                If (Date.TryParse(row.Item("ProdEndDate").ToString, TempDate)) Then
                                    TempDate = TempDate.AddDays(1D)
                                    If (TempDate < Now) Then
                                        Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                                    End If
                                End If
                            End If
                            Send("      </td>")
                            Sendb("      <td align=""right"">")
                            If (amtMax = "") Then
                                Sendb("&nbsp;")
                            Else
                                If (PromoEngine = 0) Then
                                    Select Case (MyCommon.NZ(rst3.Rows(0).Item("ATID"), 0))
                                        Case 1, 3, 4
                                            Send("$" & Format(MyCommon.NZ(rst3.Rows(0).Item("Amt"), ""), "#####0.00#") & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                        Case 2
                                            Send(MyCommon.NZ(rst3.Rows(0).Item("Amt"), "") & "% " & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                        Case 7
                                            Send(Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
                                        Case 5, 6
                                            Send("$" & Format(MyCommon.NZ(rst3.Rows(0).Item("Amt"), ""), "#####0.00#") & "&nbsp;")
                                        Case 10
                                            MyCommon.QueryStr = "select SVProgramID, Name from StoredValuePrograms with (NoLock) where SVProgramID=" & MyCommon.NZ(rst3.Rows(0).Item("SVProgramID"), 0) & ";"
                                            rst3 = MyCommon.LRT_Select
                                            If (rst3.Rows.Count > 0) Then
                                                Send(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & " (<a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst3.Rows(0).Item("SVProgramID"), 0), 25) & """>" & MyCommon.NZ(rst3.Rows(0).Item("Name"), "") & "</a>) " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase) & " ")
                                            End If
                                        Case Else
                                            Send(MyCommon.NZ(rst3.Rows(0).Item("Amt"), "") & "&nbsp;")
                                    End Select
                                Else
                                    Select Case (MyCommon.NZ(rst3.Rows(0).Item("ATID"), 0))
                                        Case 1, 5
                                            Send("$" & Format(MyCommon.NZ(rst3.Rows(0).Item("Amt"), ""), "#####0.00#") & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                        Case 3
                                            Send(MyCommon.NZ(rst3.Rows(0).Item("Amt"), "") & "% " & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                                        Case 4
                                            Send(Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
                                        Case 2, 6
                                            Send("$" & Format(MyCommon.NZ(rst3.Rows(0).Item("Amt"), ""), "#####0.00#") & "&nbsp;")
                                        Case 7
                                            MyCommon.QueryStr = "select SVProgramID, Name from StoredValuePrograms with (NoLock) where SVProgramID=" & MyCommon.NZ(rst3.Rows(0).Item("SVProgramID"), 0) & ";"
                                            rst3 = MyCommon.LRT_Select
                                            If (rst3.Rows.Count > 0) Then
                                                Send(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & " (<a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst3.Rows(0).Item("SVProgramID"), 0), 25) & """>" & MyCommon.NZ(rst3.Rows(0).Item("Name"), "") & "</a>) " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase) & " ")
                                            End If
                                        Case Else
                                            Send(MyCommon.NZ(rst3.Rows(0).Item("Amt"), "") & "&nbsp;")
                                    End Select
                                End If
                            End If
                            Send("</td>")
                            Send("     </tr>")
							End If
                        Next
                    Else
                        Send("     <tr>")
                        Send("      <td>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</td>")
                        Send("      <td align=""right"">&nbsp;</td>")
                        Send("     </tr>")
                    End If
                    '''''''''''''''''''''''''''''''''''''''
                    '' now lets do excluded
                    '' we'll spit out the linked ones first
                    '' Send("Checking group: " & productGroupsIDs(z) & "<br />")
                    'MyCommon.QueryStr = "select O.Name,O.OfferID from OfferConditions as OC left join Offers as O on " & _
                    '" O.OfferID=OC.OfferID where conditionTypeID=2 and O.deleted=0 and ExcludedID=" & productGroupsIDs(z) & " union " & _
                    '"select distinct O.Name,O.OfferID from OfferRewards as OC " & _
                    '"left join Offers as O on O.OfferID=OC.OfferID left join RewardTiers as RT on " & _
                    '"OC.RewardID=RT.RewardID where  (O.ProdEndDate >= DATEADD(day, - 1, GETDATE())) and O.deleted=0  " & _
                    '"and ExcludedProdGroupID=" & productGroupsIDs(z) & " group by O.Name,O.OfferID"
                    'rst = MyCommon.LRT_Select
                    'If (rst.Rows.Count > 0) Then
                    '    Send("     <tr>")
                    '    Send("      <td style=""border-top: 1px solid #ffffff;"">" & Copient.PhraseLib.Lookup("term.excludedby", LanguageID) & ":</td>")
                    '    Send("      <td style=""border-top: 1px solid #ffffff;"" align=""right""></td>")
                    '    Send("     </tr>")
                    'End If
                    'For Each row In rst.Rows
                    '    ' well we know the offer id lets go digging for the max reward amount for the offer
                    '    MyCommon.QueryStr = "select max(RewardAmount) as Amt from OfferRewards as OFR left join RewardTiers as RT on " & _
                    '    "OFR.RewardID=RT.RewardID where RewardTypeID=1 and Deleted=0 and OfferID=" & row.Item("OfferID")
                    '    rst3 = MyCommon.LRT_Select
                    '    amtMax = ""
                    '    If (rst3.Rows.Count > 0) Then
                    '        amtMax = MyCommon.NZ(rst3.Rows(0).Item("Amt"), "")
                    '    End If
                    '    Send("     <tr>")
                    '    Send("      <td>&nbsp;<a href=""offer-sum.aspx?OfferID=" & row.Item("OfferID") & "&amp;Name=" & row.Item("Name") & """>" & row.Item("Name") & "</a></td>")
                    '    Send("      <td align=""right"">" & amtMax & "</td>")
                    'Next
                    '''''''''''''''''''''''''''''''''''''''
                    Send("    </table>")
                    Send("   </td>")
                    Send("  </tr>")
                    If z < (productGroups.GetUpperBound(0) - 1) Then
                        Send("  <tr>")
                    End If
                Next
                Send("</table>")
            Else
            End If
        %>
    </div>
</div>
</form>
<%
    'If MyCommon.Fetch_SystemOption(75) Then
    '  If (Logix.UserRoles.AccessNotes) Then
    '    Send_Notes(6, 0, AdminUserID)
    '  End If
    'End If
done:
    Send_BodyEnd("mainform", "searchterms")
    MyCommon.Close_LogixRT()
    Logix = Nothing
    MyCommon = Nothing
%>
