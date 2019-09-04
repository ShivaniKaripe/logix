<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: status.aspx 
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
    Dim MyCryptLib As New Copient.CryptLib
  Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
  Dim Logix As New Copient.LogixInc
  Dim dst As System.Data.DataTable
  Dim dst1 As System.Data.DataTable
  Dim dst2 As System.Data.DataTable
  Dim DT As DataTable
  Dim rst As DataTable
  Dim row As System.Data.DataRow
  Dim row2 As System.Data.DataRow
  Dim offerCount As Integer
  Dim folderCount As Integer
  Dim templateCount As Integer
  Dim offerDeployed As Integer
  Dim deployDate As String
  Dim deployStatus As String
  Dim cgroupCount As Integer
  Dim cgroupDeployed As Integer
  Dim pgroupCount As Integer
  Dim pgroupDeployed As Integer
  Dim pointCount As Integer
  Dim svCount As Integer
  Dim graphicCount As Integer
  Dim layoutCount As Integer
  Dim terminalCount As Integer
  Dim storeCount As Integer
  Dim sgroupCount As Integer
  Dim userCount As Integer
  Dim searchterms As String
  Dim sizeOfData As Integer
  Dim i As Integer = 0
  Dim maxEntries As Integer = 8
  Dim Shaded As String = " class=""shaded"""
  Dim ShadedOn As Boolean = True
  Dim hDate As New DateTime
  Dim hDateString As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim WarningTypeId As Integer
  Dim IsWarning As Boolean
  Dim WarningClass As String = ""
  Dim dtHealth As DataTable
  Dim WarningTextDisplayed As Boolean = False
  Dim LowWarning, HealthErrCount As Long
  Dim AlteredSinceDeploy As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim EPMHostURI As String = ""
  
  Dim bStoreUser As Boolean = False
  Dim sValidLocIDs As String = ""
  Dim sValidSU As String = ""
  Dim sValidLocGroups As String = ""
  Dim wherestr As String = ""
  Dim sJoin As String = ""
  Dim iLen As Integer = 0
  Dim BrokerEnabled As Boolean = False
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
  Dim conditionalQuery = String.Empty
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "status.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  'Store User
  If (MyCommon.Fetch_CM_SystemOption(131) = "1") Then
    'Figure out to which locations the store user has access
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
      
      'Figure out the location groups containing the respective locations
      MyCommon.QueryStr = "select LocationGroupID from locgroupitems where LocationID in (" & sValidLocIDs & ");"
      rst = MyCommon.LRT_Select
      iLen = rst.Rows.Count
      If iLen > 0 Then
        For i = 0 To (iLen - 1)
          If i = 0 Then
            sValidLocGroups = rst.Rows(0).Item("LocationGroupID")
          Else
            sValidLocGroups &= "," & rst.Rows(i).Item("LocationGroupID")
          End If
        Next
      End If
    
      'Figure out the other store users that have access to the same locations
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
  
  Send_HeadBegin("term.status")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  
  ' Write the style cookie
  MyCommon.QueryStr = "select StyleID from AdminUsers where AdminUserID=" & AdminUserID
  dst = MyCommon.LRT_Select
  If dst.Rows.Count > 0 Then
    Write_StyleCookie(MyCommon.NZ(dst.Rows(0).Item("StyleID"), 1))
  End If
  
  Send_Scripts()
  Send("<script type=""text/javascript"" src=""/javascript/jquery.min.js""></script>")
%>
<script type="text/javascript">
  function launchHierarchy(url) {
    var popW = 700;
    var popH = 570;
    lhierWindow = window.open(url, "hierTree", "width=" + popW + ", height=" + popH + ", top=" + calcTop(popH) + ", left=" + calcLeft(popW) + ", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
    lhierWindow.focus();
  }

  function updateSearchArea() {
    var elemSearchArea = document.getElementById("searcharea");
    if (elemSearchArea.value == 1) {
      document.getElementById("offers").style.display = 'block'
      document.getElementById("offertemplates").style.display = 'none'
      document.getElementById("customers").style.display = 'none'
      document.getElementById("customergroups").style.display = 'none'
      document.getElementById("products").style.display = 'none'
      document.getElementById("productgroups").style.display = 'none'
    } else if (elemSearchArea.value == 2) {
      document.getElementById("offers").style.display = 'none'
      document.getElementById("offertemplates").style.display = 'block'
      document.getElementById("customers").style.display = 'none'
      document.getElementById("customergroups").style.display = 'none'
      document.getElementById("products").style.display = 'none'
      document.getElementById("productgroups").style.display = 'none'
    } else if (elemSearchArea.value == 3) {
      document.getElementById("offers").style.display = 'none'
      document.getElementById("offertemplates").style.display = 'none'
      document.getElementById("customers").style.display = 'block'
      document.getElementById("customergroups").style.display = 'none'
      document.getElementById("products").style.display = 'none'
      document.getElementById("productgroups").style.display = 'none'
    } else if (elemSearchArea.value == 4) {
      document.getElementById("offers").style.display = 'none'
      document.getElementById("offertemplates").style.display = 'none'
      document.getElementById("customers").style.display = 'none'
      document.getElementById("customergroups").style.display = 'block'
      document.getElementById("products").style.display = 'none'
      document.getElementById("productgroups").style.display = 'none'
    } else if (elemSearchArea.value == 5) {
      document.getElementById("offers").style.display = 'none'
      document.getElementById("offertemplates").style.display = 'none'
      document.getElementById("customers").style.display = 'none'
      document.getElementById("customergroups").style.display = 'none'
      document.getElementById("products").style.display = 'block'
      document.getElementById("productgroups").style.display = 'none'
    } else if (elemSearchArea.value == 6) {
      document.getElementById("offers").style.display = 'none'
      document.getElementById("offertemplates").style.display = 'none'
      document.getElementById("customers").style.display = 'none'
      document.getElementById("customergroups").style.display = 'none'
      document.getElementById("products").style.display = 'none'
      document.getElementById("productgroups").style.display = 'block'
    }
  }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 1)
  Send_Subtabs(Logix, 1, 1)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  If (Request.QueryString("searcharea") = "customers") Then
    ' Someone's looking for a customer...
    searchterms = MyCommon.NZ(Request.QueryString("searchterms"), "")
    If IsNumeric(searchterms) Then
      ' The search term is entirely numeric, so assume it's an ID.
      MyCommon.QueryStr = "select CustomerPK, CardTypeID from CardIDs where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(searchterms) & "';"
        
      dst = MyCommon.LXS_Select
      If dst.Rows.Count > 0 Then
        If dst.Rows(0).Item("CardTypeID") = 0 Then
          ' It's an individual:
          Response.Redirect("customer-inquiry.aspx?search=Search&CustomerPK=&searchby=1&hhID=&cardID=" & searchterms)
        ElseIf dst.Rows(0).Item("CardTypeID") = 1 Then
          ' It's a household:
          Response.Redirect("customer-inquiry.aspx?search=Search&CustomerPK=&searchby=2&hhID=" & searchterms & "&cardID=")
        End If
      Else
        ' No match, so just pass it through as an individual card:
        Response.Redirect("customer-inquiry.aspx?search=Search&CustomerPK=&searchby=1&hhID=&cardID=" & searchterms)
      End If
    Else
      ' The search is not entirely numeric, so assume it's a name.
      Response.Redirect("customer-inquiry.aspx?search=Search&CustomerPK=&searchby=4&lastname=" & searchterms)
    End If
  End If
%>
<div id="intro">
  <div id="version">
    <%
      MyCommon.QueryStr = "select top 1 VersionID, MajorVersion, MinorVersion, Build, Revision, InstallDate from InstalledVersions with (NoLock) order by InstallDate Desc;"
      dst = MyCommon.LRT_Select
      If dst.Rows.Count > 0 Then
        hDate = dst.Rows(0).Item("InstallDate")
        hDateString = Logix.ToShortDateString(hDate, MyCommon)
        Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID) & " " & dst.Rows(0).Item("MajorVersion") & "." & dst.Rows(0).Item("MinorVersion") & " ")
        Sendb(StrConv(Copient.PhraseLib.Lookup("term.build", LanguageID), VbStrConv.Lowercase) & " " & dst.Rows(0).Item("Build") & " ")
        Sendb(Left(StrConv(Copient.PhraseLib.Lookup("term.revision", LanguageID), VbStrConv.Lowercase), 3) & " " & dst.Rows(0).Item("Revision") & ", ")
        Sendb(StrConv(Copient.PhraseLib.Lookup("term.installed", LanguageID), VbStrConv.Lowercase) & " " & hDateString)
      End If
    %>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="column1">
    <%
      ' Offer Count
        
      If (bEnableRestrictedAccessToUEOfferBuilder) Then
        conditionalQuery = GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "AOLV")
      End If
        
      If (BannersEnabled) Then
        MyCommon.QueryStr = "select count(AOLV.OfferID) as OfferCount from AllOffersListview AOLV with (NoLock) " & _
                            "where AOLV.deleted=0 and isnull(AOLV.isTemplate,0)=0 " & _
                            "and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                            "or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID " & _
                            "where AUB.AdminUserID = " & AdminUserID & ") )"
      Else
        MyCommon.QueryStr = "select count(AOLV.OfferID) as OfferCount from AllOffersListview AOLV with (NoLock) where AOLV.deleted=0 and AOLV.IsTemplate=0"
      End If
        
      If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
      
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        offerCount = MyCommon.NZ(dst1.Rows(0).Item("OfferCount"), 0)
      End If
      
      ' Folder count
      MyCommon.QueryStr = "SELECT count(FolderID) as FolderCount FROM Folders WITH (NoLock);"
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        folderCount = MyCommon.NZ(dst1.Rows(0).Item("FolderCount"), 0)
      End If
      
      ' Template Count
      If (BannersEnabled) Then
        MyCommon.QueryStr = "select count(AOLV.OfferID) as TemplateCount from AllOffersListview AOLV with (NoLock) " & _
                            "where AOLV.deleted=0 and isnull(AOLV.isTemplate,0)=1 " & _
                            "and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                            "or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID " & _
                            "where AUB.AdminUserID = " & AdminUserID & ") )"
      Else
        MyCommon.QueryStr = "select count(AOLV.OfferID) as TemplateCount from AllOffersListview AOLV with (NoLock) where AOLV.deleted=0 and AOLV.IsTemplate=1"
      End If
        
      If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
        
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        templateCount = MyCommon.NZ(dst1.Rows(0).Item("TemplateCount"), 0)
      End If
      
      ' Offers Deployed
      If (BannersEnabled) Then
        MyCommon.QueryStr = "select count(AOLV.OfferID) as DeployedCount from AllOffersListview AOLV with (NoLock) " & _
                            "where(AOLV.deleted = 0 And isnull(AOLV.isTemplate, 0) = 0 And AOLV.StatusFlag = 0) " & _
                            "and (AOLV.OfferID not in ( select OfferID from BannerOffers BO with (NoLock)) " & _
                            "or AOLV.OfferID in (select BO.OfferID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join BannerOffers BO with (NoLock) on BO.BannerID = AUB.BannerID " & _
                            "where AUB.AdminUserID = " & AdminUserID & ") )"
      Else
        MyCommon.QueryStr = "select count(AOLV.OfferID) as DeployedCount from AllOffersListview AOLV with (NoLock) " & _
                            "where AOLV.deleted=0 and isnull(AOLV.isTemplate,0)=0 and AOLV.StatusFlag=0"
      End If
        
      If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
        
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        offerDeployed = MyCommon.NZ(dst1.Rows(0).Item("DeployedCount"), 0)
      End If
      
      ' Customer Group count
      MyCommon.QueryStr = "select count(*) as CgCount from CustomerGroups with (NoLock) " & _
                          "where Deleted=0 and CustomerGroupID <> 1 and CustomerGroupID <> 2 and BannerID is null and NewCardholders=0;"
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        cgroupCount = MyCommon.NZ(dst1.Rows(0).Item("CgCount"), 0)
      End If
      
      ' Customer Group Deployed
      If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
        MyCommon.QueryStr = "SELECT Count(CustomerGroupID) as DeployedCgCount FROM CustomerGroups with (NoLock) " & _
                            "WHERE CustomerGroups.Deleted=0 and (CMOADeployStatus=1 or CPEStatusFlag=0 or UEStatusFlag=0) and CustomerGroupID <> 1 and CustomerGroupID <> 2 and NewCardholders=0;"
        dst1 = MyCommon.LRT_Select
        If (dst1.Rows.Count > 0) Then
          cgroupDeployed = MyCommon.NZ(dst1.Rows(0).Item("DeployedCgCount"), 0)
        End If
      End If
      
      ' Product Group 
      MyCommon.QueryStr = "select count(ProductGroupID) as PgCount from ProductGroups with (NoLock) " & _
                          "where Deleted=0 and ProductGroupID <> 1;"
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        pgroupCount = MyCommon.NZ(dst1.Rows(0).Item("PgCount"), 0)
      End If
      
      ' Product Group deployed
      MyCommon.QueryStr = "SELECT count(ProductGroupID) as DeployedPgCount FROM ProductGroups with (NoLock) " & _
                          "WHERE Deleted=0 and (CMOADeployStatus=1 or CPEStatusFlag=0 or UEStatusFlag=0) and ProductGroupID <> 1;"
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        pgroupDeployed = MyCommon.NZ(dst1.Rows(0).Item("DeployedPgCount"), 0)
      End If
      
      ' Points Programs
      If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
        MyCommon.QueryStr = "SELECT count(PP.ProgramID) as PointsCount FROM PointsPrograms AS PP WITH (NoLock) WHERE PP.Deleted=0;"
        dst1 = MyCommon.LRT_Select
        If (dst1.Rows.Count > 0) Then
          pointCount = MyCommon.NZ(dst1.Rows(0).Item("PointsCount"), 0)
        End If
      End If
      
      'Stored Value Programs
      If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
        MyCommon.QueryStr = "SELECT count(SVProgramID) as StoredValueCount FROM StoredValuePrograms WITH (NoLock) WHERE Deleted=0;"
        dst1 = MyCommon.LRT_Select
        If (dst1.Rows.Count > 0) Then
          svCount = MyCommon.NZ(dst1.Rows(0).Item("StoredValueCount"), 0)
        End If
      End If
      
      ' Graphics
      If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
        MyCommon.QueryStr = "select Count(OnScreenAdID) as GraphicsCount from OnScreenAds with (nolock) where Deleted=0;"
        dst1 = MyCommon.LRT_Select
        If (dst1.Rows.Count > 0) Then
          graphicCount = MyCommon.NZ(dst1.Rows(0).Item("GraphicsCount"), 0)
        End If
      End If
      
      ' Layouts
      If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
        MyCommon.QueryStr = "select count(LayoutID) as LayoutCount from ScreenLayouts with (nolock) where Deleted=0;"
        dst1 = MyCommon.LRT_Select
        If (dst1.Rows.Count > 0) Then
          layoutCount = MyCommon.NZ(dst1.Rows(0).Item("LayoutCount"), 0)
        End If
      End If
      
      ' Terminals
      MyCommon.QueryStr = "select Count(TerminalTypeId) as TerminalCount  from TerminalTypes TT with (NoLock) " & _
                          "inner join PromoEngines PE with (NoLock) on PE.EngineId = TT.EngineId " & _
                          "left join Banners BAN with (NoLock) on TT.BannerID = BAN.BannerID and BAN.Deleted=0 " & _
                          "where AnyTerminal=0 and TT.Deleted=0;"
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        terminalCount = MyCommon.NZ(dst1.Rows(0).Item("TerminalCount"), 0)
      End If
      
      ' Stores
      If (BannersEnabled) Then
        MyCommon.QueryStr = "select count(L.LocationID) as StoreCount from Locations as L with (NoLock) " & _
                            "inner join PromoEngines as PE with (NoLock) on L.EngineID=PE.EngineID " & _
                            "where Deleted = 0 and (BannerID is Null or BannerID =0 or BannerID in (select BannerID from AdminUserBanners where AdminUserID=" & AdminUserID & "))"
      Else
        MyCommon.QueryStr = "select count(L.LocationID) as StoreCount from Locations as L with (NoLock) " & _
                            "inner join PromoEngines as PE with (NoLock) on L.EngineID=PE.EngineID " & _
                            "where Deleted = 0"
      End If
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        storeCount = MyCommon.NZ(dst1.Rows(0).Item("StoreCount"), 0)
      End If
      
      ' Store Groups
      If (BannersEnabled) Then
        MyCommon.QueryStr = "select count(*) as SgCount from (select BannerID, a.LocationGroupID,a.Name,(select count(*) from Locations with (NoLock) where Deleted = 0) as Locations,a.CreatedDate,a.LastUpdate,a.AllLocations, pe.description as EngineName from LocationGroups a with (NoLock) inner join PromoEngines pe on a.engineid = pe.engineid where a.AllLocations = 1 and a.Deleted = 0" & _
                            "  union select BannerID, a.LocationGroupID,a.Name,(select count(*) from LocGroupItems b with (NoLock) where b.Deleted = 0 and b.LocationGroupId = a.LocationGroupId) as Locations,a.CreatedDate,a.LastUpdate,a.AllLocations, pe.description as EngineName from LocationGroups a with (NoLock) inner join PromoEngines pe with (NoLock) on a.engineid = pe.engineid where a.AllLocations = 0 and a.Deleted = 0 " & _
                            "  ) LocGroups " & _
                            "where (BannerID is Null or BannerID =0 or BannerID in (select BannerID from AdminUserBanners where AdminUserID=" & AdminUserID & "))"
      Else
        MyCommon.QueryStr = "select count(*) as SgCount from (select BannerID, a.LocationGroupID,a.Name,(select count(*) from Locations with (NoLock) where Deleted = 0) as Locations,a.CreatedDate,a.LastUpdate,a.AllLocations, pe.description as EngineName from LocationGroups a with (NoLock) inner join PromoEngines pe on a.engineid = pe.engineid where a.AllLocations = 1 and a.Deleted = 0" & _
                            "  union select BannerID, a.LocationGroupID,a.Name,(select count(*) from LocGroupItems b with (NoLock) where b.Deleted = 0 and b.LocationGroupId = a.LocationGroupId) as Locations,a.CreatedDate,a.LastUpdate,a.AllLocations, pe.description as EngineName from LocationGroups a with (NoLock) inner join PromoEngines pe with (NoLock) on a.engineid = pe.engineid where a.AllLocations = 0 and a.Deleted = 0 " & _
                            "  ) LocGroups "
      End If
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        sgroupCount = MyCommon.NZ(dst1.Rows(0).Item("SgCount"), 0)
      End If
      
      ' Users
      MyCommon.QueryStr = "SELECT count (AU.AdminUserID) as UserCount FROM AdminUsers AS AU WITH (NoLock) "
      dst1 = MyCommon.LRT_Select
      If (dst1.Rows.Count > 0) Then
        userCount = MyCommon.NZ(dst1.Rows(0).Item("UserCount"), 0)
      End If
    %>
    <div class="box" id="systemstatus">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.systemstatus", LanguageID))%>
        </span>
      </h2>
      <table cellpadding="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.systemstatus", LanguageID))%>">
        <thead>
          <tr>
            <th align="left" class="th-name" scope="col" colspan="2">
              <% Sendb(Copient.PhraseLib.Lookup("term.item", LanguageID))%>
            </th>
            <th align="left" class="th-deployed" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.deployed", LanguageID))%>
            </th>
            <th align="left" class="th-total" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.total", LanguageID))%>
            </th>
          </tr>
        </thead>
        <tbody>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="offer-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateOfferFromBlank) Then
                  Send("<small class=""noprint""><a href=""offer-new.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>
              <% Response.Write(offerDeployed)%>
            </td>
            <td>
              <% Response.Write(offerCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="folders.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.folders", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateTemplate) Then
                  Send("<small class=""noprint""><a href=""folders.aspx?New=Yes"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td></td>
            <td>
              <% Response.Write(folderCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="temp-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.templates", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateTemplate) Then
                  Send("<small class=""noprint""><a href=""offer-new.aspx?NewTemplate=Yes"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(templateCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="cgroup-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.customergroups", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateCustomerGroups) Then
                  Send("<small class=""noprint""><a href=""cgroup-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>
              <% Response.Write(cgroupDeployed)%>
            </td>
            <td>
              <% Response.Write(cgroupCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <% End If%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="pgroup-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.productgroups", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateProductGroups) Then
                  Send("<small class=""noprint""><a href=""pgroup-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>
              <% Response.Write(pgroupDeployed)%>
            </td>
            <td>
              <% Response.Write(pgroupCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="point-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.pointsprograms", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreatePointsPrograms) Then
                  Send("<small class=""noprint""><a href=""point-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(pointCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <% End If%>
          <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="sv-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprograms", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateStoredValuePrograms) Then
                  Send("<small class=""noprint""><a href=""sv-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(svCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <% End If%>
          <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="graphic-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.graphics", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateGraphics) Then
                  Send("<small class=""noprint""><a href=""graphic-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(graphicCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <% End If%>
          <% If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="layout-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.layouts", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateLayouts) Then
                  Send("<small class=""noprint""><a href=""layout-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(layoutCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <% End If%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="store-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.stores", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateStores) Then
                  Send("<small class=""noprint""><a href=""store-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(storeCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="lgroup-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateStoreGroups) Then
                  Send("<small class=""noprint""><a href=""lgroup-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(sgroupCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="terminal-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.EditTerminals) Then
                  Send("<small class=""noprint""><a href=""terminal-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(terminalCount)%>
            </td>
          </tr>
          <% If (ShadedOn) Then ShadedOn = False Else ShadedOn = True%>
          <tr <% If ShadedOn Then Sendb(Shaded)%>>
            <td>
              <a href="user-list.aspx">
                <% Sendb(Copient.PhraseLib.Lookup("term.users", LanguageID))%>
              </a>
            </td>
            <td>
              <%
                If (Logix.UserRoles.CreateAdminUsers) Then
                  Send("<small class=""noprint""><a href=""user-edit.aspx"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a></small>")
                End If
              %>
            </td>
            <td>&nbsp;
            </td>
            <td>
              <% Response.Write(userCount)%>
            </td>
          </tr>
        </tbody>
      </table>
      <hr class="hidden" />
    </div>
    <div class="box" id="userstatus">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.userstatus", LanguageID))%>
        </span>
      </h2>
      <%
        Send("<b>" & AdminName & "</b><br />")
        MyCommon.QueryStr = "select R.RoleID,R.RoleName,R.PhraseID from AdminRoles as R with (NoLock) where R.RoleID in(select RoleID from AdminUserRoles where AdminUserID=" & AdminUserID & ") ORDER BY RoleName"
        dst = MyCommon.LRT_Select
        For Each row In dst.Rows
          If IsDBNull(row.Item("PhraseID")) Then
            Send("      " & row.Item("RoleName") & "<br />")
          Else
            If (row.Item("PhraseID") = 0) Then
              Send("      " & row.Item("RoleName") & "<br />")
            Else
              Send("      " & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "<br />")
            End If
          End If
        Next
      %>
      <br class="half" />
      <h3>
        <a href="user-edit.aspx?UserID=<% Sendb(AdminUserID)%>">
          <% Sendb(Copient.PhraseLib.Lookup("status.editprefs", LanguageID))%>
        </a>
      </h3>
      <br class="half" />
      <table cellpadding="0" summary="<% Sendb(Copient.PhraseLib.Lookup("status.yra", LanguageID))%>">
        <thead>
          <tr>
            <th align="left" class="th-activity" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("status.yra", LanguageID))%>
            </th>
            <th align="left" class="th-date" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
            </th>
          </tr>
        </thead>
        <tbody>
          <%
              If (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                  If Not (MyCommon.PMRTadoConn.State = ConnectionState.Open) Then
                      MyCommon.Open_PrefManRT()
                  End If

                  EPMHostURI = IntegrationVals.HTTP_RootURI
                  If Not (Right(EPMHostURI, 1) = "/") Then
                      EPMHostURI = EPMHostURI & "/"
                  End If
                  EPMHostURI = EPMHostURI & "UI/"
              End If

              MyCommon.QueryStr = "select top " & maxEntries & " AU.FirstName, AU.LastName, AL.ActivityDate, AL.Description, ActT.Name as ActivityTypeName, AL.LinkID, isnull(AL.ActivityTypeID, 0) as ActivityTypeID " & _
                                  "from ActivityLog as AL with (NoLock) " & _
                                  "left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID " & _
                                  "left join ActivityTypes as ActT with (NoLock) on ActT.ActivityTypeID=AL.ActivityTypeID " & _
                                    "where AdminUserID=" & AdminUserID & "  order by ActivityDate desc;"
              dst = MyCommon.LRT_Select
              sizeOfData = dst.Rows.Count
              While (i < sizeOfData And i < maxEntries)
                  Send("<tr" & Shaded & ">")

                  If (dst.Rows(i).Item("ActivityTypeID") = 100003) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                      'EPM Connector Activity (GUID change)
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "connector-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "connector-detail.aspx?ConnectorID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If

                  ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100005) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                      'EPM Agent Activity 
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "agent-detail.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "agent-detail.aspx?appid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      End If

                  ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100006) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                      'EPM SystemOptions Activity 
                      Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")

                  ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100007) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                      'EPM Role Activity 
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "role-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "role-edit.aspx?RoleID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      End If

                  ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100008) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                      'EPM Theme Activity 
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefstheme-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefstheme-edit.aspx?themeid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      End If

                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Offer") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""offer-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""offer-redirect.aspx?OfferID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Logged on") Then
                      Send("    <td>" & Copient.PhraseLib.Lookup("term.loggedin", LanguageID) & "</td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Logged off") Then
                      Send("    <td>" & Copient.PhraseLib.Lookup("term.loggedout", LanguageID) & "</td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Customer Group") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""cgroup-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""cgroup-edit.aspx?CustomerGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Product Group") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""pgroup-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Points Program") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""point-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""point-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Location") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""store-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""store-edit.aspx?LocationID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Location Group") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""lgroup-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""lgroup-edit.aspx?LocationGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Graphic") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""graphic-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""graphic-edit.aspx?OnScreenAdID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Screen Layout") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""layout-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""layout-edit.aspx?LayoutID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Categories") Then
                      Send("    <td><a href=""categories.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Departments") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""department-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""department-edit.aspx?DeptID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Product Hierarchy") Then
                      Send("   <td><a href=""javascript:launchHierarchy('phierarchytree.aspx');"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Store Hierarchy") Then
                      Send("   <td><a href=""javascript:launchHierarchy('lhierarchytree.aspx');"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Roles") Then
                      Send("    <td><a href=""role-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Terminals") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""terminal-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""terminal-edit.aspx?TerminalID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Tenders") Then
                      If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then
                          Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                      ElseIf MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then
                          Send("    <td><a href=""tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      ElseIf MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) Then
                          Send("    <td><a href=""tender.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Admin Users") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""user-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""user-edit.aspx?UserID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "settings") Then
                      Send("    <td><a href=""settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Customer Inquiry") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""customer-inquiry.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""customer-inquiry.aspx?CustPK=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Stored Value") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""sv-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""sv-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Promotion Variables") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""promovar-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""promovar-edit.aspx?PromoVarID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Banner") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""banner-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""banner-edit.aspx?BannerID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Reports") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""reports-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""reports-detail.aspx?OfferID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Agents") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""agent-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""agent-detail.aspx?appid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Issuance") Then
                      Send("    <td><a href=""issuance.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CM Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CM settings") Then
                      Send("    <td><a href=""CM-settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CPE Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CPE settings") Then
                      Send("    <td><a href=""CPEsettings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "UE Settings") Then
                      Send("    <td><a href=""UE\UESettings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Web Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Web settings") Then
                      Send("    <td><a href=""websettings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      'ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "DP Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "DP settings") Then
                      '  Send("    <td><a href=""DP-settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Attribute Product Group Builder Configuration") Then
                      Send("    <td><a href=""Attribute-PGBuilderConfig.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "External Sources") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""sources-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""sources-edit.aspx?SourceID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Issuance") Then
                      Send("    <td><a href=""issuance.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Vendor") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""vendor-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""vendor-edit.aspx?VendorID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Campaign") Then
                      Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Event") Then
                      Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Scorecard") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""scorecard-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""scorecard-edit.aspx?ScorecardID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "TerminalLockingGroup") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""terminal-lockgroup-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""terminal-lockgroup-edit.aspx?TerminalLockingGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Connectors") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""connector-detail.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""connector-detail.aspx?ConnectorID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Attributes") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""attribute-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""attribute-edit.aspx?AttributeID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Advanced Limits") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""CM-advlimit-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""CM-advlimit-edit.aspx?LimitID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Folders") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""folders.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""folders.aspx?FolderID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Health Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Health settings") Then
                      Send("    <td><a href=""health-settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Customer supplemental field") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""customer-supplemental-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""customer-supplemental-edit.aspx?FieldID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Terminal Sets") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""terminal-sets-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""terminal-sets-edit.aspx?TerminalSetID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Mutual exclusion groups") Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          Send("    <td><a href=""MEG-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          Send("    <td><a href=""MEG-edit.aspx?MutualExclusionGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If

                  ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Preferences") And MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals) Then
                      If dst.Rows(i).Item("LinkID") = 0 Then
                          'we don't have a link to the preference referenced in the activity log ... just send a link to the prefernece folders page
                          Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "folders.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                      Else
                          'see if this is a custom or system preference so we know what page to send the user to
                          MyCommon.QueryStr = "select UserCreated from Preferences where PreferenceID=" & dst.Rows(i).Item("LinkID") & ";"
                          DT = MyCommon.PMRT_Select
                          If DT.Rows.Count > 0 Then
                              If DT.Rows(0).Item("UserCreated") = True Then
                                  Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefscustom-edit.aspx?prefid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                              Else
                                  Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefsstd-edit.aspx?prefid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                              End If
                          Else
                              'we weren't able to find the preference referenced in the activity log ... just send a link to the prefernece folders page
                              Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "folders.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                          End If
                          DT = Nothing
                      End If  'dst.Rows(i).Item("LinkID") = 0

                  Else
                      Send("    <td>&nbsp;</td>")
                  End If
                  hDate = dst.Rows(i).Item("ActivityDate")
                  hDateString = Logix.ToShortDateString(hDate, MyCommon)
                  Send("    <td>" & hDateString & "</td>")
                  Send("</tr>")
                  If Shaded = " class=""shaded""" Then
                      Shaded = ""
                  Else
                      Shaded = " class=""shaded"""
                  End If
                  i = i + 1
              End While
          %>
        </tbody>
      </table>
      <hr class="hidden" />
    </div>
    <% BrokerEnabled = (MyCommon.Fetch_SystemOption(80) = "3")
      If BrokerEnabled Then
    %>
    <div class="box" id="brokerstatus">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.extpointsconnector", LanguageID))%>
        </span>
      </h2>
      <table cellpadding="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.comingup", LanguageID))%>">
        <thead>
          <tr>
            <th align="left" class="th-activity" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.server", LanguageID))%>
            </th>
            <th align="left" class="th-activity" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.messages", LanguageID))%>
            </th>
            <th align="left" class="th-activity" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.time", LanguageID))%>
            </th>
          </tr>
        </thead>
        <tbody>
          <%
            MyCommon.QueryStr = "select ExtLocationCode, ExtHostPartnerCode, ExtHostSuccess, Messages, LastUpdate, " & _
                                "DateDiff(minute, CURRENT_TIMESTAMP, LastUpdate) AS MinutesAgo FROM ExtHostStatus"
            Dim brokerHostSuccess As Boolean = False
            Dim brokerSuccessStr As String = ""
            Dim minutesAgo As Integer = 0
            Dim statusColor As String = ""
            Dim numMessages As Integer = 0
            WarningClass = "<span class=""" & IIf(IsWarning, "redlight", "greenlight") & """>" & _
                              IIf(IsWarning, "&#9679;", "&#9679;") & "</span>"

            dst = MyCommon.LRT_Select
            If (dst.Rows.Count = 0) Then
              Send("          <tr colspan=4>")
              Send("            <td>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</td>")
              Send("          </tr>")
            Else
              For Each row In dst.Rows
                hDate = row.Item("LastUpdate")
                hDateString = Logix.ToShortDateTimeString(hDate, MyCommon)
                minutesAgo = row.Item("MinutesAgo")
                numMessages = MyCommon.NZ(row.Item("Messages"), 0)
                brokerHostSuccess = row.Item("ExtHostSuccess")
                If (Not brokerHostSuccess Or minutesAgo < -3) Then
                  statusColor = "redlight"
                ElseIf (numMessages > 0) Then
                  statusColor = "blacklight"
                Else
                  statusColor = "greenlight"
                End If
                Send("<tr>")
                Send("<td><span class=""" & statusColor & """>&#9679;</span>&nbsp;" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & _
                     "&nbsp;-&nbsp;" & MyCommon.NZ(row.Item("ExtHostPartnerCode"), "") & "</td>")
                Send("<td>" & numMessages & "</td>")
                Send("<td>" & hDateString & "</td>")
                Send("</tr>")
              Next
            End If
          %>
        </tbody>
      </table>
    </div>
    <% End If%>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="warnings">
      <h2>
        <span class="white">
          <% Sendb(Copient.PhraseLib.Lookup("term.systemnotices", LanguageID))%>
        </span>
      </h2>
      <div class="boxscroll">
        <%
          If (bEnableRestrictedAccessToUEOfferBuilder) Then
            If (Not String.IsNullOrEmpty(conditionalQuery)) Then conditionalQuery = String.Empty
            conditionalQuery = GetRestrictedAccessToUEBuilderQuery(MyCommon, Logix, "")
          End If
          ' Admin System Warning
          MyCommon.QueryStr = "dbo.pa_GetSystemWarnings"
          MyCommon.Open_LRTsp()
          rst = MyCommon.LRTsp_select()
          For Each row In rst.Rows
            WarningTypeId = MyCommon.NZ(row.Item("WarningTypeID"), 0)
            IsWarning = (MyCommon.NZ(row.Item("Warning"), 0) = 1)
            WarningClass = "<span class=""" & IIf(IsWarning, "redlight", "greenlight") & """>" & _
                            IIf(IsWarning, "&#9679;", "&#9679;") & "</span>"
            
            Dim serverHealthEnabled As Boolean = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso MyCommon.Fetch_UE_SystemOption(91) = 1

            Select Case WarningTypeId
              Case 1
                Send("  " & WarningClass & "<a href=""agent-list.aspx"">" & Copient.PhraseLib.Lookup("term.agents", LanguageID) & "</a><br />")
              Case 2
                'Disabling this for now: we don't have a filter on the store health page that's specifically for failed sanity checks
                'If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso Not serverHealthEnabled) Then
                '  Send("  " & WarningClass & "<a href=""store-health-cpe.aspx?searchterms=&amp;filterhealth=2"">" & Copient.PhraseLib.Lookup("term.sanitycheck", LanguageID) & "</a><br />")
                'End If
              Case 3
                If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso Logix.UserRoles.AccessStoreHealth) Then
                  Send("  " & WarningClass & "<a href=""store-health-cpe.aspx?searchterms=&amp;filterhealth=7"">" & Copient.PhraseLib.Lookup("term.cpe", LanguageID) & " " & Copient.PhraseLib.Lookup("term.iplneeded", LanguageID) & "</a><br />")
                End If
              Case 4
                'Disabling this for now: we don't have a filter on the store health page that's specifically for no incentive fetches
                'If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso !serverHealthEnabled) Then
                '  Send("  " & WarningClass & "<a href=""store-health-cpe.aspx?searchterms=&amp;filterhealth=6"">" & Copient.PhraseLib.Lookup("term.incentivefetch", LanguageID) & "</a><br />")
                'End If
              Case 5
                If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso Logix.UserRoles.AccessStoreHealth Then
                  If (serverHealthEnabled) Then
                    Send("  " & WarningClass & "<a href=""UE/UEServerHealthSummary.aspx"">" & Copient.PhraseLib.Lookup("term.ue", LanguageID) & " " & Copient.PhraseLib.Lookup("term.iplneeded", LanguageID) & "</a><br />")
                  Else
                    Send("  " & WarningClass & "<a href=""UE/store-health-ue.aspx?searchterms=&amp;filterhealth=7"">" & Copient.PhraseLib.Lookup("term.ue", LanguageID) & " " & Copient.PhraseLib.Lookup("term.iplneeded", LanguageID) & "</a><br />")
                  End If
                End If
            End Select
          Next
          
          wherestr = ""
          If bStoreUser Then
            wherestr = " and OL.LocationGroupID in (" & sValidLocGroups & ") "
          End If
            
          ' Offer Validation Section
          If (MyCommon.Fetch_SystemOption(77) = 1) Then
            ' method GetOfferHealthTable is found in include file offer-health-cb.aspx
            If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) OrElse MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then
              dtHealth = GetOfferHealthTable(MyCommon, wherestr, True)
              WarningClass = IIf((dtHealth Is Nothing OrElse dtHealth.Rows.Count = 0), "greenlight", "redlight")
              Send("<span class=""" & WarningClass & """>&#9679;</span><a href=""offer-health.aspx?searchterms=&amp;filterhealth=2"">" & Copient.PhraseLib.Lookup("term.offervalidation", LanguageID) & "</a><br />")
            End If
          ElseIf (MyCommon.Fetch_SystemOption(77) = 2) Then
            Send("<span class=""blacklight"">&#9679;</span><a href=""offer-health.aspx?searchterms=&amp;filterhealth=2"">" & Copient.PhraseLib.Lookup("term.offervalidation", LanguageID) & "</a><br />")
          Else
          End If
          
          ' Store Health Section for CPE
          LowWarning = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(58))
          LowWarning = IIf(LowWarning <= 0, 90, LowWarning)
          
          If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso Logix.UserRoles.AccessStoreHealth Then
            MyCommon.QueryStr = "dbo.pa_StoreHealth_GetCentralErrorCount"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@Low", SqlDbType.BigInt).Value = LowWarning
            MyCommon.LRTsp.Parameters.Add("@ErrCount", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            HealthErrCount = MyCommon.LRTsp.Parameters("@ErrCount").Value
            MyCommon.Close_LRTsp()
            If (HealthErrCount = 0) Then
              MyCommon.QueryStr = "dbo.pa_StoreHealth_GetLocalErrorCount"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@ErrCount", SqlDbType.BigInt).Direction = ParameterDirection.Output
              MyCommon.LRTsp.ExecuteNonQuery()
              HealthErrCount += MyCommon.LRTsp.Parameters("@ErrCount").Value
              MyCommon.Close_LRTsp()
            End If

            WarningClass = "<span class=""" & IIf(HealthErrCount > 0, "redlight", "greenlight") & """>" & IIf(HealthErrCount > 0, "&#9679;", "&#9679;") & "</span>"
            
            Send("  " & WarningClass & "<a href=""store-health-cpe.aspx?filterhealth=2"">" & Copient.PhraseLib.Lookup("term.cpe", LanguageID) & " " & Copient.PhraseLib.Lookup("term.storehealth", LanguageID) & "</a><br />")
          End If
          
          ' Store Health Section for UE
          If MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso Logix.UserRoles.AccessStoreHealth Then
            If (MyCommon.Fetch_UE_SystemOption(91) = 1) Then
              Send("  " & WarningClass & "<a href=""UE/UEServerHealthSummary.aspx"">" & Copient.PhraseLib.Lookup("term.ue", LanguageID) & " " & Copient.PhraseLib.Lookup("term.serverhealth", LanguageID) & "</a><br />")
            Else
              MyCommon.QueryStr = "dbo.pa_StoreHealth_UE_GetCentralErrorCount"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@Low", SqlDbType.BigInt).Value = LowWarning
              MyCommon.LRTsp.Parameters.Add("@ErrCount", SqlDbType.BigInt).Direction = ParameterDirection.Output
              MyCommon.LRTsp.ExecuteNonQuery()
              HealthErrCount = MyCommon.LRTsp.Parameters("@ErrCount").Value
              MyCommon.Close_LRTsp()

              WarningClass = "<span class=""" & IIf(HealthErrCount > 0, "redlight", "greenlight") & """>" & IIf(HealthErrCount > 0, "&#9679;", "&#9679;") & "</span>"
              Send("  " & WarningClass & "<a href=""UE/store-health-ue.aspx?filterhealth=2"">" & Copient.PhraseLib.Lookup("term.ue", LanguageID) & " " & Copient.PhraseLib.Lookup("term.storehealth", LanguageID) & "</a><br />")
            End If
          End If
          
          MyCommon.QueryStr = "dbo.pa_Status_GetMismatchedOfferCount"
          MyCommon.Open_LRTsp()
          rst = MyCommon.LRTsp_select
          AlteredSinceDeploy = rst.Rows.Count
          MyCommon.Close_LRTsp()
          
          Sendb("<span class=""" & IIf(AlteredSinceDeploy > 0, "redlight", "greenlight") & """>&#9679;</span><a href=""offer-list.aspx?filterOffer=3"">")
          Sendb(IIf(AlteredSinceDeploy > 0, AlteredSinceDeploy, "No") & " ")
          Send(Copient.PhraseLib.Lookup("status.modified-since-deploy", LanguageID) & "</a><br />")
          
          Send("<br class=""half"" />")
          
          sJoin = ""
          wherestr = ""
          If bStoreUser Then
            sJoin = "Full Outer Join OfferLocUpdate olu with (NoLock) on O.OfferID=olu.OfferID "
            wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) "
          End If
          
          MyCommon.QueryStr = "select O.OfferID,Name,CMOADeployStatus as Status,CMOADeployRpt as DeployRpt,CMOARptDate as RptDate,CMOADeploySuccessDate " & _
                              "from Offers as O with (nolock)" & _
                              sJoin & _
                              " where Deleted=0 and isnull(CMOADeployRpt, '') not in ('', 'Export Okay') " & _
                              wherestr & _
                              "union " & _
                      "select IncentiveID as OfferID, IncentiveName as Name, CPEOADeployStatus as Status, CPEOADeployRpt as DeployRpt, CPEOARptDate as RptDate, CPEOADeploySuccessDate " & _
                "from CPE_Incentives with (NoLock) where Deleted=0 and isnull(CPEOADeployRpt, '') not in ('', 'Export Okay')"
          If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
          
          rst = MyCommon.LRT_Select()
          For Each row In rst.Rows
            If (Not WarningTextDisplayed) Then
              Send(Copient.PhraseLib.Lookup("status.warning", LanguageID) & "<br />")
              WarningTextDisplayed = True
            End If
            Send("<a href=""offer-redirect.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), "") & """>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & MyCommon.NZ(row.Item("OfferID"), "") & ": " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 20) & "</a><br />")
          Next
          
          If Not bStoreUser Then
            MyCommon.QueryStr = "select CustomerGroupID,Name,CMOADeployStatus,CMOADeployRpt,CMOARptDate,CMOADeploySuccessDate from CustomerGroups with (nolock) " & _
                                "where Deleted=0 and CustomerGroupID not in (1,2) and isnull(CMOADeployRpt, '') not in ('', 'Export Okay');"
            rst = MyCommon.LRT_Select()
            For Each row In rst.Rows
              If (Not WarningTextDisplayed) Then
                Send(Copient.PhraseLib.Lookup("status.warning", LanguageID) & "<br />")
                WarningTextDisplayed = True
              End If
              Send("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & """>" & Copient.PhraseLib.Lookup("term.customergroup", LanguageID) & " " & MyCommon.NZ(row.Item("CustomerGroupID"), "") & ": " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 20) & "</a><br />")
            Next
          End If
          
          sJoin = ""
          wherestr = ""
          If bStoreUser Then
            sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on pg.ProductGroupID = pglu.ProductGroupID "
            wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) "
          End If
          
          MyCommon.QueryStr = "select pg.ProductGroupID,Name,CMOADeployStatus,CMOADeployRpt,CMOARptDate,CMOADeploySuccessDate from ProductGroups as pg with (nolock) " & _
                              sJoin & _
                              "where Deleted=0 and pg.ProductGroupID <> 1 and isnull(CMOADeployRpt, '') not in ('', 'Export Okay')" & _
                              wherestr & ";"
          rst = MyCommon.LRT_Select()
          For Each row In rst.Rows
            If (Not WarningTextDisplayed) Then
              Send(Copient.PhraseLib.Lookup("status.warning", LanguageID) & "<br />")
              WarningTextDisplayed = True
            End If
            Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), "") & """>" & Copient.PhraseLib.Lookup("term.productgroup", LanguageID) & " " & MyCommon.NZ(row.Item("ProductGroupID"), "") & ": " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 20) & "</a><br />")
          Next
        %>
      </div>
      <hr class="hidden" />
    </div>
    <div class="box" id="search">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>
        </span>
      </h2>
      <select id="searcharea" name="searcharea" class="large" onchange="updateSearchArea();">
        <option value="1">
          <% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>
        </option>
        <option value="2">
          <% Sendb(Copient.PhraseLib.Lookup("term.offertemplates", LanguageID))%>
        </option>
        <%
          If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Or (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE)) Then
            Send("<option value=""3"">" & Copient.PhraseLib.Lookup("term.customers", LanguageID) & "</option>")
            Send("<option value=""4"">" & Copient.PhraseLib.Lookup("term.customergroups", LanguageID) & "</option>")
          End If
        %>
        <option value="5">
          <% Sendb(Copient.PhraseLib.Lookup("term.products", LanguageID))%>
        </option>
        <option value="6">
          <% Sendb(Copient.PhraseLib.Lookup("term.productgroups", LanguageID))%>
        </option>
      </select>
      <div id="offers" style="display: block;">
        <form action="offer-list.aspx" name="searchform1" style="display: inline;">
          <input type="text" name="searchterms" class="long" maxlength="100" value="" />
          <input type="submit" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
          <input type="hidden" name="searcharea" value="offers" />
        </form>
      </div>
      <div id="offertemplates" style="display: none;">
        <form action="temp-list.aspx" name="searchform2" style="display: inline;">
          <input type="text" name="searchterms" class="long" maxlength="100" value="" />
          <input type="submit" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
          <input type="hidden" name="searcharea" value="templates" />
        </form>
      </div>
      <div id="customers" style="display: none;">
        <form action="#" name="searchform3" style="display: inline;">
          <input type="text" name="searchterms" class="long" maxlength="100" value="" />
          <input type="submit" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
          <input type="hidden" name="searcharea" value="customers" />
        </form>
      </div>
      <div id="customergroups" style="display: none;">
        <form action="cgroup-list.aspx" name="searchform4" style="display: inline;">
          <input type="text" name="searchterms" class="long" maxlength="100" value="" />
          <input type="submit" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
          <input type="hidden" name="searcharea" value="customergroups" />
        </form>
      </div>
      <div id="products" style="display: none;">
        <form action="product-inquiry.aspx" name="searchform5" style="display: inline;">
          <input type="text" name="searchterms" class="long" maxlength="100" value="" />
          <input type="submit" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
          <input type="hidden" name="searcharea" value="products" />
        </form>
      </div>
      <div id="productgroups" style="display: none;">
        <form action="pgroup-list.aspx" name="searchform6" style="display: inline;">
          <input type="text" name="searchterms" class="long" maxlength="100" value="" />
          <input type="submit" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID))%>" />
          <input type="hidden" name="searcharea" value="productgroups" />
        </form>
      </div>
    </div>
    <div class="box" id="events">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.events", LanguageID))%>
        </span>
      </h2>
      <table cellpadding="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.today", LanguageID))%>" style="height: 5px; overflow-y: scroll;">
        <thead>
          <tr>
            <th align="left" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.today", LanguageID))%>
            </th>
          </tr>
        </thead>
      </table>
      <div class="notificationDiv" style="position: relative; text-align: left; display: block; width: 99%; max-height: 150px; overflow-y: auto; margin: 0px; padding-left: 2px; padding-right: 0;">
        <%
          Dim aPF As AjaxProcessingFunctions = New AjaxProcessingFunctions()
          Dim notificationData As String = aPF.GetNotificationList(bEnableRestrictedAccessToUEOfferBuilder, conditionalQuery, bStoreUser, sValidLocIDs, sValidSU, LanguageID, 1, 100)
          Dim today As New DateTime
          Dim oDate As New DateTime
          today = DateTime.Today
          Shaded = " class=""shaded"""
          If (String.IsNullOrWhiteSpace(notificationData)) Then
            Send("<p style=""padding-left: 2px;"">" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</p>")
          Else
            Send(notificationData)
          End If
        %>
      </div>
      <br class="half" />
      <table cellpadding="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.comingup", LanguageID))%>">
        <thead>
          <tr>
            <th align="left" class="th-activity" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.comingup", LanguageID))%>
            </th>
            <th align="left" class="th-date" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
            </th>
          </tr>
        </thead>
        <tbody>
          <%
            Shaded = " class=""shaded"""
            
            If bStoreUser Then
              sJoin = "Inner Join OfferLocUpdate olu with (NoLock) on Table1.OfferID=olu.OfferID "
              wherestr = " where (LocationID in (" & sValidLocIDs & ") or CreatedByAdminID in (" & sValidSU & ")) "
            End If
            
            MyCommon.QueryStr = "select top 8 Table1.OfferID, Date, Event, CreatedByAdminID from ( " & _
                                " select OfferID,ProdStartDate as Date, 'starts' as Event, CreatedByAdminID from Offers as O with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and ProdStartDate>getdate() " & _
                                " union " & _
                                " select OfferID, ProdEndDate  as Date, 'ends' as Event, CreatedByAdminID from Offers as O with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and ProdEndDate>getdate() " & _
                                " union " & _
                                " select IncentiveID as OfferID,StartDate as Date, 'starts' as Event, CreatedByAdminID from CPE_Incentives as I with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and StartDate>getdate() "
            If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
            MyCommon.QueryStr &= " union " & _
                                " select IncentiveID as OfferID,EndDate as Date, 'ends' as Event, CreatedByAdminID from CPE_Incentives as I with (NoLock) where Deleted=0 and isnull(isTemplate,0)=0 and EndDate>getdate() "
            If (bEnableRestrictedAccessToUEOfferBuilder AndAlso Not String.IsNullOrEmpty(conditionalQuery)) Then MyCommon.QueryStr &= conditionalQuery & " "
            MyCommon.QueryStr &= ") as Table1 " & sJoin & wherestr & " order by Date, Event Desc;"
            dst = MyCommon.LRT_Select
            If (dst.Rows.Count = 0) Then
              Send("          <tr>")
              Send("            <td colspan=""2"">" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</td>")
              Send("          </tr>")
            Else
              For Each row In dst.Rows
                oDate = row.Item("Date")
                If DateTime.Compare(today, oDate) < 0 Then
                  Send("          <tr" & Shaded & ">")
                  Sendb("            <td><a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & row.Item("OfferID") & "</a>")
                  If (row.Item("Event") = "starts") Then
                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.starts", LanguageID), VbStrConv.Lowercase))
                  ElseIf (row.Item("Event") = "ends") Then
                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.ends", LanguageID), VbStrConv.Lowercase))
                  End If
                  Send("</td>")
                  Send("            <td>" & Logix.ToShortDateString(oDate, MyCommon) & "</td>")
                  Send("          </tr>")
                  If Shaded = " class=""shaded""" Then
                    Shaded = ""
                  Else
                    Shaded = " class=""shaded"""
                  End If
                End If
              Next
            End If
          %>
        </tbody>
      </table>
      <br class="half" />
      <table cellpadding="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.recentactivity", LanguageID))%>">
        <thead>
          <tr>
            <th align="left" class="th-activity" scope="col" colspan="2">
              <% Sendb(Copient.PhraseLib.Lookup("term.recentactivity", LanguageID))%>
            </th>
            <th align="left" class="th-date" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.date", LanguageID))%>
            </th>
          </tr>
        </thead>
        <tbody>
          <%
            If (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
              If Not (MyCommon.PMRTadoConn.State = ConnectionState.Open) Then
                MyCommon.Open_PrefManRT()
              End If

              EPMHostURI = IntegrationVals.HTTP_RootURI
              If Not (Right(EPMHostURI, 1) = "/") Then
                EPMHostURI = EPMHostURI & "/"
              End If
              EPMHostURI = EPMHostURI & "UI/"
            End If

            wherestr = ""
            If bStoreUser Then
              wherestr = " AdminID in (" & sValidSU & ") and "
            End If
            
            Shaded = " class=""shaded"""
            MyCommon.QueryStr = "select top " & maxEntries & " AU.FirstName as Actor, AU.LastName, AL.ActivityDate, AL.Description, ActT.Name as ActivityTypeName, AL.LinkID, isnull(AL.ActivityTypeID, 0) as ActivityTypeID " & _
                                "from ActivityLog as AL with (Nolock) left join AdminUsers as AU with (NoLock) on AL.AdminID=AU.AdminUserID " & _
                                "left Join ActivityTypes as ActT with (NoLock) on ActT.ActivityTypeID=AL.ActivityTypeID " & _
                                "where " & wherestr & " AL.ActivityTypeID > 2 order by ActivityDate desc;"
            dst = MyCommon.LRT_Select
            sizeOfData = dst.Rows.Count
            i = 0
            If (sizeOfData = 0) Then
              Send("<tr>")
              Send("  <td colspan=""2"">" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</td>")
              Send("</tr>")
            Else
              While (i < sizeOfData And i < maxEntries)
                Send("<tr" & Shaded & ">")
                Send("  <td>" & dst.Rows(i).Item("Actor") & "</td>")
                
                If (dst.Rows(i).Item("ActivityTypeID") = 100003) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                  'EPM Connector Activity (GUID change)
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "connector-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "connector-detail.aspx?ConnectorID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                
                ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100005) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                  'EPM Agent Activity 
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "agent-detail.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "agent-detail.aspx?appid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  End If
                
                ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100006) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                  'EPM SystemOptions Activity 
                  Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")

                ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100007) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                  'EPM Role Activity 
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "role-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "role-edit.aspx?RoleID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  End If
                
                ElseIf (dst.Rows(i).Item("ActivityTypeID") = 100008) And (MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
                  'EPM Theme Activity 
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefstheme-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefstheme-edit.aspx?themeid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  End If
                
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Offer") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""offer-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""offer-redirect.aspx?OfferID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Logged on") Then
                  Send("    <td>" & Copient.PhraseLib.Lookup("term.loggedin", LanguageID) & "</td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Logged off") Then
                  Send("    <td>" & Copient.PhraseLib.Lookup("term.loggedout", LanguageID) & "</td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Customer Group") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""cgroup-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""cgroup-edit.aspx?CustomerGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Product Group") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""pgroup-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Points Program") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""point-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""point-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Location") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""store-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""store-edit.aspx?LocationID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Location Group") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""lgroup-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""lgroup-edit.aspx?LocationGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Graphic") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""graphic-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""graphic-edit.aspx?OnScreenAdID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Screen Layout") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""layout-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""layout-edit.aspx?LayoutID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Categories") Then
                  Send("    <td><a href=""categories.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Departments") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""department-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""department-edit.aspx?DeptID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Product Hierarchy") Then
                  Send("   <td><a href=""javascript:launchHierarchy('phierarchytree.aspx');"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Store Hierarchy") Then
                  Send("   <td><a href=""javascript:launchHierarchy('lhierarchytree.aspx');"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Roles") Then
                  Send("    <td><a href=""role-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Terminals") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""terminal-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""terminal-edit.aspx?TerminalID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Tenders") Then
                  If (MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)) Then
                    Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                  ElseIf MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) Then
                    Send("    <td><a href=""tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  ElseIf MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM) Then
                    Send("    <td><a href=""tender.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Admin Users") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""user-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""user-edit.aspx?UserID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "settings") Then
                  Send("    <td><a href=""settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Attribute Product Group Builder Configuration") Then
                  Send("    <td><a href=""Attribute-PGBuilderConfig.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Customer Inquiry") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""customer-inquiry.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""customer-inquiry.aspx?CustPK=" & dst.Rows(i).Item("LinkID") & """ onClick=""return targetOpener(this)"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Stored Value") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""sv-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""sv-edit.aspx?ProgramGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Promotion Variables") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""promovar-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""promovar-edit.aspx?PromoVarID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Banner") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""banner-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""banner-edit.aspx?BannerID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Reports") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""reports-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""reports-detail.aspx?OfferID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Agents") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""agent-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""agent-detail.aspx?appid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Issuance") Then
                  Send("    <td><a href=""issuance.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CM Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CM settings") Then
                  Send("    <td><a href=""CM-settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CPE Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "CPE settings") Then
                  Send("    <td><a href=""CPEsettings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "UE Settings") Then
                      Send("    <td><a href=""UE\UESettings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                 ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Web Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Web settings") Then
                  Send("    <td><a href=""websettings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  'ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "DP Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "DP settings") Then
                  '  Send("    <td><a href=""DP-settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "External Sources") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""sources-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""sources-edit.aspx?SourceID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Issuance") Then
                  Send("    <td><a href=""issuance.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Vendor") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""vendor-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""vendor-edit.aspx?VendorID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Campaign") Then
                  Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Event") Then
                  Send("    <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Scorecard") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""scorecard-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""scorecard-edit.aspx?ScorecardID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "TerminalLockingGroup") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""terminal-lockgroup-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""terminal-lockgroup-edit.aspx?TerminalLockingGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Connectors") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""connector-detail.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""connector-detail.aspx?ConnectorID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Attributes") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""attribute-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""attribute-edit.aspx?AttributeID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Advanced Limits") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""CM-advlimit-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""CM-advlimit-edit.aspx?LimitID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Folders") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""folders.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""folders.aspx?FolderID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Health Settings") OrElse (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Health settings") Then
                  Send("    <td><a href=""health-settings.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Customer supplemental field") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""customer-supplemental-edit.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""customer-supplemental-edit.aspx?FieldID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Terminal Sets") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""terminal-sets-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""terminal-sets-edit.aspx?TerminalSetID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Mutual exclusion groups") Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    Send("    <td><a href=""MEG-list.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    Send("    <td><a href=""MEG-edit.aspx?MutualExclusionGroupID=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                  End If
                  
                ElseIf (MyCommon.NZ(dst.Rows(i).Item("ActivityTypeName"), "") = "Preferences") And MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals) Then
                  If dst.Rows(i).Item("LinkID") = 0 Then
                    'we don't have a link to the preference referenced in the activity log ... just send a link to the prefernece folders page
                    Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "folders.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & "</a></td>")
                  Else
                    'see if this is a custom or system preference so we know what page to send the user to
                    MyCommon.QueryStr = "select UserCreated from Preferences where PreferenceID=" & dst.Rows(i).Item("LinkID") & ";"
                    DT = MyCommon.PMRT_Select
                    If DT.Rows.Count > 0 Then
                      If DT.Rows(0).Item("UserCreated") = True Then
                        Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefscustom-edit.aspx?prefid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      Else
                        Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "prefsstd-edit.aspx?prefid=" & dst.Rows(i).Item("LinkID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                      End If
                    Else
                      'we weren't able to find the preference referenced in the activity log ... just send a link to the prefernece folders page
                      Send("    <td><a href=""authtransfer.aspx?SendToURI=" & EPMHostURI & "folders.aspx"">" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Description"), 40) & ": #" & dst.Rows(i).Item("LinkID") & "</a></td>")
                    End If
                    DT = Nothing
                  End If  'dst.Rows(i).Item("LinkID") = 0
                  
                Else
                  Send("    <td>&nbsp;</td>")
                End If
                hDate = dst.Rows(i).Item("ActivityDate")
                hDateString = Logix.ToShortDateString(hDate, MyCommon)
                Send("  <td>" & hDateString & "</td>")
                Send("</tr>")
                If Shaded = " class=""shaded""" Then
                  Shaded = ""
                Else
                  Shaded = " class=""shaded"""
                End If
                i = i + 1
              End While
            End If
          %>
        </tbody>
      </table>
    </div>
  </div>
  <br clear="all" />
</div>
<script type="text/javascript">
    $(document).ready(function () {
        var pageNum = 1;

        $('.notificationDiv').bind('scroll', function () {
            if ($(this).scrollTop() + $(this).innerHeight() >= this.scrollHeight) {
                getNextNotificationList(++pageNum);
            }
        });

        function getNextNotificationList(pageNum) {
            <% 
          Dim urlStr As String = "/Connectors/AjaxProcessingFunctions.asmx/GetNotificationList?"
          urlStr = urlStr & "enableRestrictedAccessToUEOB=" & bEnableRestrictedAccessToUEOfferBuilder
          If Not String.IsNullOrWhiteSpace(conditionalQuery) Then
            urlStr = urlStr & "&conditionalQuery=" & conditionalQuery
          Else
            urlStr = urlStr & "&conditionalQuery=''"
          End If
          urlStr = urlStr & "&IsStoreUser=" & bStoreUser
  
          If Not String.IsNullOrWhiteSpace(sValidLocIDs)
            urlStr = urlStr & "&ValidLocIDs=" & sValidLocIDs
          Else
            urlStr = urlStr & "&ValidLocIDs=''"
          End If
          If Not String.IsNullOrWhiteSpace(sValidSU)
            urlStr = urlStr & "&ValidSU=" & sValidSU
          Else
            urlStr = urlStr & "&ValidSU=''"
          End If
          urlStr = urlStr & "&LanguageId=" & LanguageID
      %>
            var urlStr = "<%=urlStr%>";
            urlStr += "&PageNum=" + pageNum + "&MaxRecords=100";
            $.ajax({
                url: urlStr,
                cache: false,
                dataType: "xml"
            })
              .done(function (html) {
                  $('.notificationDiv').append($(html).text());
              });
        }
    });
</script>
<!-- #Include virtual="/include/offer-health-cb.inc" -->
<%
done:
  Send_BodyEnd("searchform1", "searchterms")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  If MyCommon.PMRTadoConn.State = ConnectionState.Open Then MyCommon.PMRTadoConn.Close()
  MyCommon = Nothing
  Logix = Nothing
%>
