<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-hist.aspx 
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
    Dim MyCryptLib As New Copient.CryptLib
  Dim dt As DataTable
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim row As System.Data.DataRow
  Dim CustomerPK As Long
  Dim CardPK As Long
  Dim ExtCardID As String = ""
  Dim FirstName As String = ""
  Dim MiddleName As String = ""
  Dim LastName As String = ""
  Dim FullName As String = ""
  Dim IsHouseholdID As Boolean = False
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim x As Integer = 0
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim Shaded As String = "shaded"
  Dim restrictLinks As Boolean = False
  Dim extraLink As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim CAMCustomer As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-hist.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
  
  Dim SortText As String = "ActivityDate"
  Dim SortDirection As String = ""
  
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
  
  ' lets check the logged in user and see if they are to be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    If (MyCommon.NZ(rst.Rows(0).Item("prestrict"), False) = True) Then
      ' ok we got in here then we need to restrict the user from seeing any other pages
      restrictLinks = True
    End If
  End If
  
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustPK"))
  ' Check in case it was a POST instead of get
  If (CustomerPK = 0) Then
    CustomerPK = Request.Form("CustomerPK")
  End If
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  If CardPK > 0 Then
        MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CardPK=" & CardPK & ";"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      ExtCardID = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("ExtCardID").ToString())
    End If
  End If
  CAMCustomer = IIf(Request.QueryString("CAM") = "1", True, False)
  
  If (CustomerPK = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "customer-inquiry.aspx")
  End If
  
  MyCommon.QueryStr = "select FirstName, MiddleName, LastName, CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
  dt = MyCommon.LXS_Select
  If (dt.Rows.Count > 0) Then
    IsHouseholdID = (MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0) = 1)
  End If
  
  If CardPK > 0 Then
    Send_HeadBegin("term.customer", "term.history", MyCommon.Extract_Val(ExtCardID))
  Else
    Send_HeadBegin("term.customer", "term.history")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If (Not restrictLinks) Then
    Send_Tabs(Logix, 3)
    If (CAMCustomer) Then
      If CardPK > 0 Then
        Send_Subtabs(Logix, 33, 8, LanguageID, CustomerPK, , CardPK)
      Else
        Send_Subtabs(Logix, 33, 8, LanguageID, CustomerPK)
      End If
    Else
      If CardPK > 0 Then
        Send_Subtabs(Logix, 32, 8, LanguageID, CustomerPK, , CardPK)
      Else
        Send_Subtabs(Logix, 32, 8, LanguageID, CustomerPK)
      End If
    End If
  Else
    If CardPK > 0 Then
      Send_Subtabs(Logix, 91, 4, LanguageID, CustomerPK, extraLink, CardPK)
    Else
      Send_Subtabs(Logix, 91, 4, LanguageID, CustomerPK, extraLink)
    End If
  End If
  
  If (Logix.UserRoles.AccessCustomerInquiry = False) Then
    Send_Denied(1, "perm.customer-access")
    GoTo done
  End If
  If (Logix.UserRoles.ViewHistory = False) Then
    Send_Denied(1, "perm.admin-history")
    GoTo done
  End If
  
  If (Request.QueryString("searchterms") <> "") Then
    If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
      idSearch = idNumber.ToString
    Else
      idSearch = "-1"
    End If
    idSearchText = MyCommon.Parse_Quotes(Request.QueryString("searchterms"))
    If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
    If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1") Then
      MyCommon.QueryStr = "select AU.FirstName, AU.LastName, AL.ActivityDate, AL.Description, AL.LinkID2, AL.LinkID3, AL.ActivitySubTypeID from ActivityLog as AL with (NoLock) " & _
                          "left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID " & _
                          "where ActivityTypeID='25' and LinkID='" & CustomerPK & "' and LinkID3 is null " & _
                          "and (FirstName like N'%" & idSearchText & "%' or LastName Like N '%" & idSearchText & "%') " & _
                          "union " & _
                          "select '" & Copient.PhraseLib.Lookup("term.cashier", LanguageID) & "' as FirstName, AL.LinkID3 as LastName, AL.ActivityDate, AL.Description, AL.LinkID2, AL.ActivitySubTypeID from ActivityLog as AL with (NoLock) " & _
                          "where ActivityTypeID='25' and LinkID='" & CustomerPK & "' and LinkID3 is not null " & _
                          "and ('" & Copient.PhraseLib.Lookup("term.cashier", LanguageID) & "' like N'%" & idSearchText & "%' or LinkID3 Like N '%" & idSearchText & "%') " & _
                          "order by " & SortText & " " & SortDirection & ";"
    Else
      MyCommon.QueryStr = "select case when AL.AdminID is null then '" & Copient.PhraseLib.Lookup("term.systemuser", LanguageID) & "' else AU.FirstName end AS FirstName, case when AL.AdminID is null then AL.LinkID3 else AU.LastName end AS LastName, AL.ActivityDate, AL.Description, AL.LinkID2, AL.LinkID3, AL.ActivitySubTypeID from ActivityLog as AL with (NoLock) " & _
                          "left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID " & _
                          "where ActivityTypeID in (25, 100010) and LinkID='" & CustomerPK & "' " & _
                          "and (FirstName like N'%" & idSearchText & "%' or LastName Like N '%" & idSearchText & "%') " & _
                          "order by " & SortText & " " & SortDirection & ";"
    End If
  Else
    If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) AndAlso (MyCommon.Fetch_CM_SystemOption(35) = "1") Then
      MyCommon.QueryStr = "select AU.FirstName, AU.LastName, AL.ActivityDate, AL.Description, AL.LinkID2, AL.ActivitySubTypeID from ActivityLog as AL with (NoLock) " & _
                          "left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID " & _
                          "where ActivityTypeID='25' and LinkID='" & CustomerPK & "' and LinkID3 is null " & _
                          "union " & _
                          "select '" & Copient.PhraseLib.Lookup("term.cashier", LanguageID) & "' as FirstName, AL.LinkID3 as LastName, AL.ActivityDate, AL.Description, AL.LinkID2, AL.ActivitySubTypeID from ActivityLog as AL with (NoLock) " & _
                          "where ActivityTypeID='25' and LinkID='" & CustomerPK & "' and LinkID3 is not null " & _
                          "order by " & SortText & " " & SortDirection & ";"
    Else
      MyCommon.QueryStr = "select case when AL.AdminID is null then '" & Copient.PhraseLib.Lookup("term.systemuser", LanguageID) & "' else AU.FirstName end AS FirstName, case when AL.AdminID is null then AL.LinkID3 else AU.LastName end AS LastName, AL.ActivityDate,AL.ActivityID, AL.Description, AL.LinkID2, AL.ActivitySubTypeID " & _
                            " from ActivityLog as AL with (NoLock) left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID" & _
                            " where ActivityTypeID in (25, 100010) and LinkID='" & CustomerPK & "' order by " & SortText & " " & SortDirection & ";"
    End If
  End If
    dt = MyCommon.LRT_Select
    sizeOfData = dt.Rows.Count
    i = linesPerPage * PageNum
%>
<div id="intro">
  <h1 id="title">
    <%
      If CardPK = 0 Then
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID))
        End If
      Else
        If (IsHouseholdID) Then
          Sendb(Copient.PhraseLib.Lookup("term.householdcard", LanguageID) & " #" & ExtCardID)
        Else
          Sendb(Copient.PhraseLib.Lookup("term.customercard", LanguageID) & " #" & ExtCardID)
        End If
      End If
      MyCommon.QueryStr = "select Prefix, FirstName, MiddleName, LastName, Suffix from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
      rst2 = MyCommon.LXS_Select
      If rst2.Rows.Count > 0 Then
        FullName = IIf(MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Prefix, MyCommon.NZ(rst2.Rows(0).Item("Prefix"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_FirstName, MyCommon.NZ(rst2.Rows(0).Item("FirstName"), "") & " ", "")
        FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_MiddleName, Left(MyCommon.NZ(rst2.Rows(0).Item("MiddleName"), ""), 1) & ". ", "")
        FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("LastName"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_LastName, MyCommon.NZ(rst2.Rows(0).Item("LastName"), ""), "")
        FullName &= IIf(MyCommon.NZ(rst2.Rows(0).Item("Suffix"), "") <> "" AndAlso Logix.UserRoles.AccessCustomerIdData_Suffix, " " & MyCommon.NZ(rst2.Rows(0).Item("Suffix"), ""), "")
      End If
      If FullName <> "" Then
        Sendb(": " & MyCommon.TruncateString(FullName, 30))
      End If
    %>
  </h1>
  <div id="controls">
    <%
      If (CustomerPK > 0 And Logix.UserRoles.ViewCustomerNotes) Then
        Send_CustomerNotes(CustomerPK, CardPK)
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    If (InfoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")
    End If
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection & IIf(CAMCustomer, "&amp;CAM=1", ""), , "CustPK=" & CustomerPK & IIf(CardPK > 0, "&amp;CardPK=" & CardPK, ""), True)
  %>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.history", LanguageID)) %>">
    <thead>
      <tr>
        <th align="left" scope="col" class="th-timedate">
          <a href="customer-hist.aspx?CustPK=<% Sendb(CustomerPK) %><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ActivityDate&amp;SortDirection=<% Sendb(SortDirection) %><%Sendb(IIf(CAMCustomer, "&amp;CAM=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.timedate", LanguageID))%>
          </a>
          <%
            If SortText = "ActivityDate" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" class="th-user">
          <a href="customer-hist.aspx?CustPK=<% Sendb(CustomerPK) %><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=FirstName&amp;SortDirection=<% Sendb(SortDirection) %><%Sendb(IIf(CAMCustomer, "&amp;CAM=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.user", LanguageID))%>
          </a>
          <%
            If SortText = "FirstName" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" scope="col" class="th-action">
          <a href="customer-hist.aspx?CustPK=<% Sendb(CustomerPK) %><% Sendb(IIf(CardPK > 0, "&amp;CardPK=" & CardPK, "")) %>&amp;searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Description&amp;SortDirection=<% Sendb(SortDirection) %><%Sendb(IIf(CAMCustomer, "&amp;CAM=1", "")) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID))%>
          </a>
          <%
            If SortText = "Description" Then
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
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          Send("<tr class=""" & Shaded & """>")
          If (Not IsDBNull(dt.Rows(i).Item("ActivityDate"))) Then
            Send("  <td>" & Logix.ToShortDateTimeString(dt.Rows(i).Item("ActivityDate"), MyCommon) & "</td>")
          Else
            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          FullName = MyCommon.NZ(dt.Rows(i).Item("FirstName"), "") & " " & MyCommon.NZ(dt.Rows(i).Item("LastName"), "")
          Send("  <td>" & MyCommon.SplitNonSpacedString(FullName, 25) & "</td>")
              Sendb("  <td style='white-space: -moz-pre-wrap white-space: -pre-wrap; white-space: -o-pre-wrap;white-space: pre-wrap;word-wrap: break-word;word-break: break-all;white-space: normal;'>")
          Sendb(MyCommon.NZ(dt.Rows(i).Item("Description"), ""))
          If MyCommon.NZ(dt.Rows(i).Item("ActivitySubTypeID"), 0) = 12 Then
            MyCommon.QueryStr = "select distinct 1 as EngineID, O.Name, O.OfferID from PointsPrograms as PP with (NoLock) " & _
                                "inner join OfferConditions as OC with (NoLock) on PP.ProgramID=OC.LinkID and PP.Deleted=0 and OC.ConditionTypeID=3 " & _
                                "inner join Offers as O with (NoLock) on OC.OfferID=O.OfferID and O.Deleted=0 and O.IsTemplate=0 and OC.Deleted=0 " & _
                                "where PP.ProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & " " & _
                                " union " & _
                                "select distinct 1 as EngineID, O.Name, O.OfferID from RewardPoints as RP with (NoLock) " & _
                                "inner join OfferRewards AS OFFR WITH (NoLock) on RP.RewardPointsID=OFFR.LinkID and OFFR.RewardTypeID=2 " & _
                                "inner join Offers as O with (NoLock) on OFFR.OfferID=O.OfferID and O.Deleted=0 and O.IsTemplate=0 and OFFR.Deleted=0 and O.ProdEndDate>=getdate() " & _
                                "where RP.ProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & " " & _
                                " union " & _
                                "select distinct 2 as EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_IncentivePointsGroups IPG with (NoLock) " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                                "inner join CPE_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                "inner join PointsPrograms PP with (NoLock) on IPG.ProgramID=PP.ProgramID " & _
                                "where IPG.ProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & " and IPG.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and PP.Deleted=0 and I.IsTemplate=0 " & _
                                " union " & _
                                "select distinct 2 as EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_DeliverablePoints DP " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DP.RewardOptionID " & _
                                "inner join CPE_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                "where DP.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and ProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & _
                                "order by OfferID desc;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              x = 1
              Sendb("<br /><div style=""font-size:11px;margin-left:6px;"">(" & Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID) & ": ")
              For Each row In rst.Rows
                Sendb("<a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), 0) & """>" & MyCommon.NZ(row.Item("OfferID"), 0) & "</a>")
                If x < rst.Rows.Count Then
                  Sendb(", ")
                End If
                x += 1
              Next
              Sendb(")</div>")
            End If
          ElseIf MyCommon.NZ(dt.Rows(i).Item("ActivitySubTypeID"), 0) = 13 Then
            MyCommon.QueryStr = "select distinct 1 as EngineID, O.Name, O.OfferID from StoredValuePrograms as SV with (NoLock) " & _
                                "inner join OfferConditions as OC with (NoLock) on SV.SVProgramID=OC.LinkID and SV.Deleted=0 and OC.ConditionTypeID=6 " & _
                                "inner join Offers as O with (NoLock) on OC.OfferID=O.OfferID and O.Deleted=0 and O.IsTemplate=0 and OC.Deleted=0 " & _
                                "where SV.SVProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & " " & _
                                " union " & _
                                "select distinct 1 as EngineID, O.Name, O.OfferID from CM_RewardStoredValues as RSV with (NoLock) " & _
                                "inner join OfferRewards AS OFFR WITH (NoLock) on RSV.RewardStoredValuesID=OFFR.LinkID and OFFR.RewardTypeID=10 " & _
                                "inner join Offers as O with (NoLock) on OFFR.OfferID=O.OfferID and O.Deleted=0 and O.IsTemplate=0 and OFFR.Deleted=0 and O.ProdEndDate>=getdate() " & _
                                "where RSV.ProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & " " & _
                                " union " & _
                                "select distinct 2 as EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_IncentiveStoredValuePrograms ISVP with (NoLock) " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ISVP.RewardOptionID " & _
                                "inner join CPE_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                "inner join StoredValuePrograms SV with (NoLock) on ISVP.SVProgramID=SV.SVProgramID " & _
                                "where ISVP.SVProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & " and ISVP.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and SV.Deleted=0 and I.IsTemplate=0 " & _
                                " union " & _
                                "select distinct 2 as EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_DeliverableStoredValue DSV " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DSV.RewardOptionID " & _
                                "inner join CPE_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                "where DSV.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and I.IsTemplate=0 and SVProgramID=" & MyCommon.NZ(dt.Rows(i).Item("LinkID2"), 0) & " " & _
                                "order by OfferID desc;"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              x = 1
              Sendb("<br /><div style=""font-size:11px;margin-left:6px;"">(" & Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID) & ": ")
              For Each row In rst.Rows
                Sendb("<a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), 0) & """>" & MyCommon.NZ(row.Item("OfferID"), 0) & "</a>")
                If x < rst.Rows.Count Then
                  Sendb(", ")
                End If
                x += 1
              Next
              Sendb(")</div>")
            End If
          End If
          Send("</td>")
          Send("</tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
        If (sizeOfData = 0) Then
          Send("<tr>")
          Send("  <td colspan=""3""></td>")
          Send("</tr>")
        End If
      %>
    </tbody>
  </table>
</div>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
