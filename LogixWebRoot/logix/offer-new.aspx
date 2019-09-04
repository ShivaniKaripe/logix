<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>
<%
    ' *****************************************************************************
    ' * FILENAME: offer-new.aspx
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
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim rst2 As System.Data.DataTable
    Dim OfferID As Integer
    Dim SchemeID As Integer
    Dim IsNewTemplate As Boolean = False
    Dim IntroID As String = ""
    Dim ShowSave As Boolean = False
    Dim ActiveTab As Integer = 1
    Dim AllInfoEntered As Boolean = False
    Dim EngineID As Integer = -1
    Dim EngineSubTypePKID As Integer = -1
    Dim EngineSubTypeID As Integer = -1
    Dim HeaderTitle As String = ""
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannerCt As Integer = 0
    Dim EngineCt As Integer = 0
    Dim BannerID As Integer = 0
    Dim BannerIdList As String() = Nothing
    Dim i As Integer = 0
    Dim AllowMultipleBanners As Boolean = False
    Dim BannersEnabled As Boolean = False
    Dim AllBannersCheckBox As String = ""
    Dim DefaultCAMScorecardExists As Boolean = False
    Dim External As Boolean = False
    Dim Save As Boolean = True
    Dim EngineOptionName As String = ""
    Dim OutboundCRM As Integer = 0
    Dim CategoryID As Integer = 0
    Dim CAMCategoryID As Integer = 0
    Const CAMEngineID As Integer = 6
    Const CPEEngineID As Integer = 2
    Dim TempEngineID As Integer = 0
    Dim WrongFolderSelected As Boolean = False
    Dim AllowSpecialCharacters As String
    Dim DefaultEngine As String = ""
    Dim BuyerSelected As Boolean = False
    Dim tempbuyerId As Int32 = -2
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)
    Dim m_OAWService As IOfferApprovalWorkflowService
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
    CurrentRequest.Resolver.AppName = "offer-new.aspx"
    Response.Expires = 0
    MyCommon.AppName = "offer-new.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    m_OAWService = CurrentRequest.Resolver.Resolve(Of IOfferApprovalWorkflowService)()
    AllowSpecialCharacters = MyCommon.Fetch_SystemOption(171)
    IsNewTemplate = (Request.QueryString("NewTemplate") = "Yes" Or Request.QueryString("NewTemplate") = "True")
    If (IsNewTemplate) Then
        IntroID = "intro"
        HeaderTitle = "term.newtemplate"
        ShowSave = Logix.UserRoles.CreateTemplate
        ActiveTab = 2
    Else
        IntroID = "intro"
        HeaderTitle = "term.newoffer"
        ShowSave = Logix.UserRoles.CreateOfferFromBlank
        ActiveTab = 1
    End If

    If Request.QueryString("External") <> "" Then
        If Request.QueryString("External") = "True" Then
            External = True
            ActiveTab = 3
        End If
    End If

    Send_HeadBegin("term.offers")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript" language="javascript">
  function handleKeyDown(e) {
    var keycode;
    var submitThing;

    if (window.event) keycode = window.event.keyCode;
    else if (e) keycode = e.which;
    else return true;

    if (keycode == 13) {
      submitThing = document.getElementById("searchPressed");
      submitThing.value = "Search";
    <% If Request.Browser.Browser <> "IE" Then %>
      document.mainform.submit();
    <% End If %>
      return false;
    } else {
      return true;
    }
    return true;
  }

  function submitForm(val) {
    if (val != "-1") {
      document.mainform.submit();
    }
  }

  function toggleSchemeSelector() {
    var elemSchemeSelector = document.getElementById("schemeselector");

    if (elemSchemeSelector.style.display == "none") {
      elemSchemeSelector.style.display = "block";
    } else {
      elemSchemeSelector.style.display = "none";
    }
  }

  function updateSchemeID() {
    var elemSchemeID = document.getElementById("schemeID");
    var elemSchemeSelector = document.getElementById("schemeselector");

    elemSchemeID.value = elemSchemeSelector.value;
    elemSchemeSelector.style.display = "none";
  }

    var tempfolderList;
    var tempfodlerName="None Selected";
    var updated =false;
  function UpdateBuyer(val){
    var selectedEngine = $('#EngineSubTypePKID :selected').val();
    //If selected Engine is UE display Buyer section.
    if(val == 12){
        xmlhttpPost('UE//UEfolderFeeds.aspx');
        updated=false;
		if (document.getElementById("buyers") != null)
        buyers.style.display = "block";
      }
    else{
          if(!updated){
          tempfodlerName=document.getElementById("folderNames").innerHTML.toString() ;
          tempfolderList=document.getElementById("folderList").value.toString();
          updated=true;
          }
          document.getElementById("folderNames").innerHTML = "None Selected";
          document.getElementById("folderList").value="";
		  if (document.getElementById("buyers") != null)
          buyers.style.display = "none";
        }
  }

  function xmlhttpPost(strURL) {
    document.getElementById("folderNames").innerHTML = "<div class=\"loading\"><img id=\"clock\" src=\"../images/clock22.png\" \ /><br \ />" +'<\/div>';

    var _buyerId = document.getElementById("buyerID")!=null? $("#buyerID").val():-1;
    strURL += "?buyerID="+ _buyerId;

    $.post(strURL,{buyerID : _buyerId},function (data) { updatepage(data); });
  }

  function updatepage(str) {
    document.getElementById("folderNames").innerHTML = str;
    document.getElementById("folderList").value="";
    document.getElementById("folderList").value=getfoldervalue(str);
  }

  function getfoldervalue(str) {
    var strtIndex= str.indexOf("hidden");
    var elem = str.substr(strtIndex);
    var endIndex= elem.indexOf("value");
    var val ;
    elem= elem.substr(endIndex);
    endIndex = elem.indexOf("/>");
    if(endIndex == -1)
      endIndex = elem.indexOf(" />");
    elem= elem.substr(0,endIndex);
    endIndex = elem.indexOf("=");
    elem = elem.substr(endIndex+1);

    if(elem !="\"\"")
        val= parseInt(elem.match(/[0-9]+/)[0], 10);
    else
        val ="";
    return val.toString();
  }
</script>
<%
    Send_HeadEnd()
    If (IsNewTemplate) Then
        Send_BodyBegin(11)
    Else
        Send_BodyBegin(1)
    End If
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 2)
    Send_Subtabs(Logix, 20, ActiveTab)

    If (Logix.UserRoles.AccessOffers = False AndAlso Not IsNewTemplate) Then
        Send_Denied(1, "perm.offers-access")
        GoTo done
    End If
    If (Logix.UserRoles.AccessTemplates = False AndAlso IsNewTemplate) Then
        Send_Denied(1, "perm.offers-access-templates")
        GoTo done
    End If

    ' store banner options
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    AllowMultipleBanners = (MyCommon.Fetch_SystemOption(67) = "1")

    If (Request.QueryString("NewCAM") = "Yes") Then
        EngineID = CAMEngineID
        EngineSubTypePKID = CAMEngineID
    End If

    If (Request.QueryString("EngineSubTypePKID") <> "") Then
        EngineSubTypePKID = MyCommon.Extract_Val(Request.QueryString("EngineSubTypePKID"))
        MyCommon.QueryStr = "select PromoEngineID, SubTypeID from PromoEngineSubTypes with (NoLock) where PKID=" & EngineSubTypePKID
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("PromoEngineID"), -1)
            EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("SubTypeID"), -1)
        End If
    End If

    If (Request.QueryString("SchemeID") <> "") Then
        SchemeID = MyCommon.Extract_Val(Request.QueryString("SchemeID"))
    End If
    'AL-5888 System option 132 split into 132 (Offers must be associated with at least one folder) and 192 (Offer dates should be within folder dates)
    If MyCommon.Fetch_SystemOption(132) = "1" AndAlso Not IsNewTemplate Then
        AllInfoEntered = Logix.TrimAll(Request.QueryString("Name")) <> "" AndAlso EngineSubTypePKID >= 0 _
                           AndAlso (MyCommon.Extract_Val(Request.QueryString("banner")) > 0 OrElse Request.QueryString("allbannersid") <> "" OrElse Not BannersEnabled) _
                           AndAlso Request.QueryString("folderList") <> ""
    Else
        AllInfoEntered = Logix.TrimAll(Request.QueryString("Name")) <> "" AndAlso EngineSubTypePKID >= 0 _
                           AndAlso (MyCommon.Extract_Val(Request.QueryString("banner")) > 0 OrElse Request.QueryString("allbannersid") <> "" OrElse Not BannersEnabled)
    End If
    'For UE engine, if Buyer is present in system, Buyer ID should be selected.
    If (Request.QueryString("EngineSubTypePKID") = 12 AndAlso AllInfoEntered) Then
        AllInfoEntered = Not (Request.QueryString("buyerID") = "-1")
    End If

    If ((Request.QueryString("save") <> "" OrElse Request.QueryString("searchPressed") <> "") AndAlso AllInfoEntered) Then
        ' User wants to save something, so run stored procedure for saving group
        ' OK its not the CPE so lets make ourselves an offer and redirect
        ' Before making an offer check for Lockout Days validation

        If External = True Then
            If MyCommon.Extract_Val(Request.QueryString("crmEngineID")) > 0 Then
                Save = True
            Else
                Save = False
            End If
        End If

        If Save Then

            ' Check if the selected folder has a corresponding theme with the banner
            If MyCommon.Fetch_SystemOption(131) = "1" AndAlso Not IsNewTemplate Then
                If BannersEnabled AndAlso Not AllowMultipleBanners Then
                    MyCommon.QueryStr = "SELECT bt.Lockoutdays FROM BannerThemes bt INNER JOIN  FolderThemes ft ON ft.ThemeID = bt.ThemeID " & _
                                        " AND bt.BannerID = " & MyCommon.Extract_Val(Request.QueryString("Banner")) & " INNER JOIN Folders fo ON ft.FolderID=fo.FolderID " & _
                                        " WHERE fo.FolderID = " & MyCommon.Extract_Val(Request.QueryString("folderList")) & ""
                    rst2 = MyCommon.LRT_Select
                    If rst2.Rows.Count = 0 Then
                        WrongFolderSelected = True
                    End If
                End If
            End If

            Dim OffName As String = Logix.TrimAll(Request.QueryString("Name"))
            Dim buyerID As Integer = MyCommon.Extract_Val(Request.QueryString("buyerID"))
            'If buyer option doesn't exist in offer creation page set it to -1
            If (EngineID = 9) Then
                buyerID = If(buyerID = 0, -1, buyerID)
            Else
                buyerID = -1
            End If
            MyCommon.QueryStr = "dbo.pt_Offers_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = OffName
            MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
            MyCommon.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = EngineSubTypeID
            MyCommon.LRTsp.Parameters.Add("@IsTemplate", SqlDbType.Bit).Value = IIf(IsNewTemplate, 1, 0)
            MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@BuyerID", SqlDbType.Int).Value = If(buyerID = -1, DBNull.Value, buyerID)

            If EngineID = 6 Then
                MyCommon.QueryStr = "select ScorecardID, EngineID, DefaultForEngine from Scorecards with (NoLock) " & _
                                    "where EngineID=6 and DefaultForEngine=1 and ScorecardTypeID=1 and Deleted=0;"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    DefaultCAMScorecardExists = True
                End If
            End If

            MyCommon.QueryStr = "SELECT Name FROM Offers with (NoLock) WHERE Name = @IncentiveName AND Deleted=0 " & _
                                "UNION " & _
                                "SELECT IncentiveName FROM CPE_Incentives with (NoLock) WHERE IncentiveName = @IncentiveName AND Deleted=0"
            MyCommon.DBParameters.Add("@IncentiveName", SqlDbType.NVarChar).Value = OffName
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (rst.Rows.Count > 0) Then
                If (IsNewTemplate) Then
                    infoMessage = Copient.PhraseLib.Lookup("template-gen.nameused", LanguageID)
                Else
                    infoMessage = Copient.PhraseLib.Lookup("offer-gen.nameused", LanguageID)
                End If
            ElseIf (EngineID = 6) AndAlso (DefaultCAMScorecardExists = False) Then
                If (IsNewTemplate) Then
                    infoMessage = Copient.PhraseLib.Lookup("template-new.MissingDefaultCAMScorecard", LanguageID)
                Else
                    infoMessage = Copient.PhraseLib.Lookup("offer-new.MissingDefaultCAMScorecard", LanguageID)
                End If
            ElseIf WrongFolderSelected Then
                infoMessage = Copient.PhraseLib.Lookup("folder.BadTheme", LanguageID)
            Else
                MyCommon.QueryStr = "dbo.pt_Offers_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = OffName
                MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                MyCommon.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = EngineSubTypeID
                MyCommon.LRTsp.Parameters.Add("@IsTemplate", SqlDbType.Bit).Value = IIf(IsNewTemplate, 1, 0)
                MyCommon.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.Parameters.Add("@BuyerID", SqlDbType.Int).Value = If(buyerID = -1, DBNull.Value, buyerID)
                MyCommon.LRTsp.ExecuteNonQuery()
                OfferID = MyCommon.LRTsp.Parameters("@OfferID").Value
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-create", LanguageID), buyerID)
                MyCommon.Close_LRTsp()

                'Save Offer Folder selection,
                Dim cookie As HttpCookie
                If Request.Cookies("DefaultBuyer") Is Nothing Then
                    cookie = New HttpCookie("DefaultBuyer")
                Else
                    cookie = HttpContext.Current.Request.Cookies("DefaultBuyer")
                End If

                cookie.Value = buyerID
                cookie.Expires = DateTime.MaxValue
                Response.Cookies.Add(cookie)

                ' add offer to the selected banner
                If (BannersEnabled) Then
                    If (Request.QueryString("allbannersid") <> "") Then
                        For i = 0 To Request.QueryString.GetValues("allbannersid").GetUpperBound(0)
                            MyCommon.QueryStr = "insert into BannerOffers (BannerID, OfferID) values (" & MyCommon.Extract_Val(Request.QueryString.GetValues("allbannersid")(i)) & ", " & OfferID & ");"
                            MyCommon.LRT_Execute()
                        Next
                    ElseIf (Request.QueryString("banner") <> "") Then
                        BannerIdList = Request.QueryString.GetValues("banner")
                        For i = 0 To BannerIdList.GetUpperBound(0)
                            MyCommon.QueryStr = "insert into BannerOffers (BannerID, OfferID) values (" & MyCommon.Extract_Val(BannerIdList(i)) & ", " & OfferID & ");"
                            MyCommon.LRT_Execute()
                        Next
                    End If
                End If
                'insert the record into Offer Approvals table if approval is enabled
                Dim m_isOAWEnabled As Boolean = False
                If (BannersEnabled) Then
                    Dim BannerIds As Integer() = Logix.GetBannersForOffer(OfferID)
                    m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabledForBanners(BannerIds).Result
                Else
                    m_isOAWEnabled = m_OAWService.CheckIfOfferApprovalIsEnabled().Result
                End If

                If OfferID > 0 AndAlso m_isOAWEnabled Then
                    m_OAWService.InsertUpdateOfferApprovalRecord(OfferID, AdminUserID)
                End If

                'set outbound engine
                OutboundCRM = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(39))
                If EngineID = 2 AndAlso OfferID > 0 Then
                    MyCommon.QueryStr = "Update CPE_Incentives with (RowLock) set CRMEngineID=" & OutboundCRM & " " & _
                                         "where IncentiveID=" & OfferID & " and Deleted=0;"
                    MyCommon.LRT_Execute()
                ElseIf EngineID = 0 AndAlso OfferID > 0 Then
                    MyCommon.QueryStr = "Update Offers with (RowLock) set CRMEngineID=" & OutboundCRM & " " & _
                                        "where OfferID=" & OfferID & " and Deleted=0;"
                    MyCommon.LRT_Execute()
                End If

                ' save the external source if one is selected
                If EngineID = 0 Then
                    If (OfferID > 0 AndAlso MyCommon.Extract_Val(Request.QueryString("crmEngineID")) > 0) Then
                        MyCommon.QueryStr = "update Offers with (RowLock) set InboundCRMEngineID = " & MyCommon.Extract_Val(Request.QueryString("crmEngineID")) & " " & _
                                            "where OfferID = " & OfferID & " and Deleted=0;"
                        MyCommon.LRT_Execute()
                    End If
                Else
                    If (OfferID > 0 AndAlso MyCommon.Extract_Val(Request.QueryString("crmEngineID")) > 0) Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set InboundCRMEngineID = " & MyCommon.Extract_Val(Request.QueryString("crmEngineID")) & " " & _
                                            "where IncentiveID = " & OfferID & " and Deleted=0;"
                        MyCommon.LRT_Execute()
                    End If
                End If

                ' save the folder selections
                If (OfferID > 0 AndAlso Request.QueryString("folderList") <> "") Then
                    MyCommon.QueryStr = "dbo.pa_FolderOffers_Insert"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@OfferIDs", SqlDbType.NVarChar).Value = OfferID
                    MyCommon.LRTsp.Parameters.Add("@FolderIDs", SqlDbType.NVarChar).Value = Request.QueryString("folderList")
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Activity_Log2(3, 22, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.folder-additem", LanguageID), , Request.QueryString("folderList"))
                    MyCommon.Close_LRTsp()
                End If
                If MyCommon.Fetch_SystemOption(192) = "1" Then  '192 - Offer dates should be within folder dates
                    MyCommon.QueryStr = "dbo.pt_OfferDates_Default"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                    MyCommon.LRTsp.ExecuteNonQuery()
                    MyCommon.Close_LRTsp()
                End If

                ' set the category
                CategoryID = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(4))
                CAMCategoryID = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(160))
                If OfferID > 0 Then
                    If (EngineID = 2 Or EngineID = 3 Or EngineID = 9) Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set PromoClassID=" & CategoryID & " where IncentiveID=" & OfferID & " and Deleted=0;"
                    ElseIf (EngineID = 6) Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set PromoClassID=" & CAMCategoryID & " where IncentiveID=" & OfferID & " and Deleted=0;"
                    ElseIf (EngineID = 0) Then
                        MyCommon.QueryStr = "update Offers with (RowLock) set OfferCategoryID=" & CategoryID & " where OfferID=" & OfferID & " and Deleted=0;"
                    End If
                End If
                MyCommon.LRT_Execute()

                Select Case EngineID
                    Case 2
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "CPEoffer-gen.aspx?new=New&OfferID=" & OfferID)
                    Case 3
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "web-offer-gen.aspx?new=New&OfferID=" & OfferID)
                        'Case 4
                        '  Response.Status = "301 Moved Permanently"
                        '  Response.AddHeader("Location", "DP-offer-gen.aspx?new=New&OfferID=" & OfferID)
                    Case 5
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "email-offer-gen.aspx?new=New&OfferID=" & OfferID)
                    Case 6
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "CAM/CAM-offer-gen.aspx?new=New&OfferID=" & OfferID)
                        'Case 7
                        '  Response.Status = "301 Moved Permanently"
                        '  Response.AddHeader("Location", "PDEoffer-gen.aspx?OfferID=" & OfferID & "&SchemeID=" & SchemeID)
                    Case 9
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "UE/UEoffer-gen.aspx?new=New&OfferID=" & OfferID)
                    Case Else
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "offer-gen.aspx?new=New&OfferID=" & OfferID)
                End Select
            End If
        Else
            infoMessage = Copient.PhraseLib.Lookup("newoffer.externalsource", LanguageID)
        End If

    ElseIf ((Request.QueryString("save") <> "" OrElse Request.QueryString("searchPressed") <> "") AndAlso Not AllInfoEntered) Then
        If (Logix.TrimAll(Request.QueryString("Name")) = "") Then
            If (IsNewTemplate) Then
                infoMessage = Copient.PhraseLib.Lookup("template-gen.noname", LanguageID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("offer-gen.noname", LanguageID)
            End If
        ElseIf (Request.QueryString("EngineSubTypePKID") = "12" AndAlso Request.QueryString("buyerID") = "-1") Then
            If (IsNewTemplate) Then
                infoMessage = Copient.PhraseLib.Lookup("template.noBuyerID", LanguageID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("offer.noBuyerID", LanguageID)
            End If
        ElseIf (EngineSubTypePKID = -1 OrElse EngineID = -1) Then
            If (IsNewTemplate) Then
                infoMessage = Copient.PhraseLib.Lookup("template.noengine", LanguageID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("offer.noengine", LanguageID)
            End If
        ElseIf (BannersEnabled AndAlso MyCommon.Extract_Val(Request.QueryString("banner")) <= 0) Then
            infoMessage = Copient.PhraseLib.Lookup("offer.nobanner", LanguageID)
        ElseIf ((Request.QueryString("folderList") = "" OrElse Request.QueryString("folderList") = String.Empty) AndAlso MyCommon.Fetch_SystemOption(132) = "1") Then
            If (IsNewTemplate) Then
                infoMessage = Copient.PhraseLib.Lookup("template-new.SelectFolder", LanguageID)
            Else
                infoMessage = Copient.PhraseLib.Lookup("offer-new.SelectFolder", LanguageID)
            End If
        End If
    End If
%>
<form id="mainform" name="mainform" action="offer-new.aspx">
<input type="hidden" name="searchPressed" id="searchPressed" value="" />
<div id="<% Sendb(IntroID)%>">
    <h1 id="title">
        <%Sendb(Copient.PhraseLib.Lookup(HeaderTitle, LanguageID))%>
    </h1>
    <input type="hidden" name="NewTemplate" id="NewTemplate" value="<%Sendb(IsNewTemplate)%>" />
    <input type="hidden" name="External" id="External" value="<%Sendb(External)%>" />
    <div id="controls">
        <% If (ShowSave) Then Send_Save()%>
    </div>
</div>
<div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
        <table summary="">
            <%
                Send("<tr>")
                Send("  <td valign=""top"">")
                Send("    <label for=""name"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label>")
                Send("  </td>")
                Send("  <td valign=""top"">")
                If (AllowSpecialCharacters <> "") Then
                    If Request.QueryString("name") <> "" Then
                        Send("    <input type=""text"" class=""longest"" id=""name"" name=""name"" maxlength=""100"" onkeydown=""handleKeyDown(event);"" value=""" & Request.QueryString("name").Replace(Chr(34), "&quot;") & """ />")
                    Else
                        Send("    <input type=""text"" class=""longest"" id=""name"" name=""name"" maxlength=""100"" onkeydown=""handleKeyDown(event);"" value=""" & Request.QueryString("name") & """ />")
                    End If
                Else
                    Send("    <input type=""text"" class=""longest"" id=""name"" name=""name"" maxlength=""100"" onkeydown=""handleKeyDown(event);"" value=""" & Request.QueryString("name") & """ />")
                End If
                Send("  </td>")
                Send("</tr>")

                If (BannersEnabled) Then
                    ' find the installed engines for the user-permitted banners
                    MyCommon.QueryStr = "select distinct PE.EngineID, PE.PhraseID, PE.Description, PE.DefaultEngine, " & _
                                        "  PEST.PKID as SubTypePKID, PEST.SubTypeID, PEST.SubTypeName, PEST.PhraseID as SubPhraseID " & _
                                        "from BannerEngines BE with (NoLock)   " & _
                                        "inner join PromoEngines PE with (NoLock) on PE.EngineID = BE.EngineID  " & _
                                        "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BE.BannerID  " & _
                                        "left join PromoEngineSubTypes as PEST with (NoLock)  " & _
                                        "  on PEST.PromoEngineID = PE.EngineID and PEST.Installed=1 " & _
                                        "where PE.Installed=1 and PE.OfferBuilderSupported=1 and AUB.AdminUserID = " & AdminUserID & " " & _
                                        "union  " & _
                                        "select distinct PE.EngineID, PE.PhraseID, PE.Description, PE.DefaultEngine, " & _
                                        "  PEST.PKID as SubTypePKID, PEST.SubTypeID, PEST.SubTypeName, PEST.PhraseID as SubPhraseID " & _
                                        "from PromoEngines PE with (NoLock)  " & _
                                        "left join PromoEngineSubTypes as PEST with (NoLock)  " & _
                                        "  on PEST.PromoEngineID = PE.EngineID and PEST.Installed=1 " & _
                                        "where BannerSupported=0 and PE.Installed=1 and OfferBuilderSupported=1 "
                Else
                    ' find all installed engines
                    MyCommon.QueryStr = "select PE.EngineID, PE.PhraseID, PE.Description, PE.DefaultEngine, " & _
                                        "  PEST.PKID as SubTypePKID, PEST.SubTypeID, PEST.SubTypeName, PEST.PhraseID as SubPhraseID " & _
                                        "from PromoEngines PE with (NoLock)  " & _
                                        "left join PromoEngineSubTypes as PEST with (NoLock)  " & _
                                        "  on PEST.PromoEngineID = PE.EngineID and PEST.Installed=1 " & _
                                        "where PE.Installed=1 and PE.OfferBuilderSupported=1 "
                End If
                If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Not Logix.UserRoles.CreateUEOffers) Then
                      MyCommon.QueryStr &= " and PE.EngineID <> 9 "
                End If
                If (IsNewTemplate) Then
                    MyCommon.QueryStr &= " and PE.TemplateSupported = 1;"
                End If
                rst = MyCommon.LRT_Select
                EngineCt = rst.Rows.Count

                If (EngineCt = 1) Then
                    EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
                    EngineSubTypePKID = MyCommon.NZ(rst.Rows(0).Item("SubTypePKID"), -1)
                    Send("<input type=""hidden"" id=""EngineSubTypePKID"" name=""EngineSubTypePKID"" value=""" & EngineSubTypePKID & """ />")
                    'If Only UE engine is installed, then set Default Engine to UE
                    If (EngineSubTypePKID = 12) Then
                        DefaultEngine = "UE"
                    End If
                ElseIf (EngineCt > 1) Then
                    Send("<tr id=""engines"">")
                    Send("  <td valign=""top"">")
                    Send("    <label for=""EngineSubTypePKID"">" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</label>")
                    Send("  </td>")
                    Send("  <td valign=""top"">")

                    Sendb("    <select id=""EngineSubTypePKID"" name=""EngineSubTypePKID"" class=""medium"" onchange=""javascript:UpdateBuyer(this.value);"& IIf(BannersEnabled,"submitForm(this.value);","")&"""  ")
                    'If (BannersEnabled) Then
                    '    Sendb(" onchange=""submitForm(this.value);""")
                    'End If
                    Send(">")
                    Send("      <option value=""-1"">— " & Copient.PhraseLib.Lookup("offer.select-engine", LanguageID) & " —</option>")
                    For Each row In rst.Rows
                        EngineOptionName = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), -1), LanguageID, MyCommon.NZ(row.Item("Description"), "")) & " " & _
                                           Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("SubPhraseID"), -1), LanguageID, MyCommon.NZ(row.Item("SubTypeName"), ""))

                        If (EngineID = -1 AndAlso MyCommon.NZ(row.Item("DefaultEngine"), -1) = 1) OrElse (EngineID >= 0 And EngineSubTypePKID = MyCommon.NZ(row.Item("SubTypePKID"), -1)) Then
                            Send("      <option value=""" & MyCommon.NZ(row.Item("SubTypePKID"), -1) & """ selected=""selected"">" & EngineOptionName & "</option>")
                            EngineID = MyCommon.NZ(row.Item("EngineID"), -1)
                            EngineSubTypePKID = MyCommon.NZ(row.Item("SubTypePKID"), -1)
                            If (EngineOptionName.Trim() = "UE") Then
                                DefaultEngine = "UE"
                            End If
                        Else
                            Send("      <option value=""" & MyCommon.NZ(row.Item("SubTypePKID"), -1) & """>" & EngineOptionName & "</option>")
                        End If
                    Next
                    Send("    </select>")
                    Send("  </td>")
                    Send("</tr>")
                End If

                '*******************************************************************'
                ' find all Buyers Created
                MyCommon.QueryStr = "select B.BuyerId,ExternalBuyerId from Buyers B inner join buyerroleusers BU on B.BuyerId = BU.BuyerId where BU.AdminUserID=" & AdminUserID & ";"
                rst = MyCommon.LRT_Select
                'if We have any buyers in Logix, then show Buyers section
                If rst.Rows.Count > 0 Then
                    'preselect the BuyerID, which is used before
                    If Not Request.Cookies("DefaultBuyer") Is Nothing AndAlso DefaultEngine = "UE" Then
                        tempbuyerId = Convert.ToInt32(Request.Cookies("DefaultBuyer").Value)
                    End If
                    Send("<tr id=""buyers""")
                    If (DefaultEngine <> "UE") Then
                        Send("style=""display:none""")
                    End If
                    Send(">")

                    Send("  <td valign=""top"">")
                    Send("    <label for=""buyerID"">Buyer ID:</label>")
                    Send("  </td>")
                    Send("  <td valign=""top"">")
                    Sendb("    <select id=""buyerID"" name=""buyerID"" onchange=""xmlhttpPost('UE//UEfolderFeeds.aspx');"" class=""medium"" ")
                    Send(">")
                    Send("      <option value=""-1"">— Select a Buyer —</option>")
                    For Each row In rst.Rows
                        Send("      <option value=""" & MyCommon.NZ(row.Item("BuyerId"), -1) & """")
                        If (tempbuyerId = MyCommon.NZ(row.Item("BuyerId"), -1)) Then
                            Send(" selected=""selected""")
                            BuyerSelected = True
                        End If
                        Send(">" & MyCommon.NZ(row.Item("ExternalBuyerId"), "") & "</option>")
                    Next
                    Send("    </select>")
                    Send("  </td>")
                    Send("</tr>")
                End If
                '*******************************************************************'

                ' show banners dropdown (if necessary)
                TempEngineID = EngineID
                'The CAM engine should be treated as a sub-engine of CPE. If the CPE engine does not support banners then CAM will not as well and vice versa.
                If EngineID = CAMEngineID Then TempEngineID = CPEEngineID 'If the engine selected is CAM then show the CPE banners
                MyCommon.QueryStr = "select EngineID from PromoEngines PE with (NoLock) " & _
                                     "where BannerSupported=1 and Installed=1 and EngineID=" & TempEngineID
                rst = MyCommon.LRT_Select
                If (BannersEnabled AndAlso EngineID > -1 AndAlso rst.Rows.Count > 0) Then
                    'MyCommon.QueryStr = "select EngineID from PromoEngines PE with (NoLock) " & _
                    '                    "where BannerSupported=0 and Installed=1 and EngineID=" & EngineID
                    'rst = MyCommon.LRT_Select

                    'If (rst.Rows.Count > 0) AndAlso EngineID <> CAMEngineID Then
                    '  ' a non-banner supported engine was selected, so show all the banners
                    '  MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name, BAN.AllBanners as AllBanners from Banners BAN with (NoLock) " & _
                    '                      "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                    '                      "where BAN.Deleted=0 and AdminUserID = " & AdminUserID & " order by BAN.Name;"
                    'Else
                    ' a banner supported engine was selected, so show only those banners for which the user is permitted.
                    MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name, BAN.AllBanners from Banners BAN with (NoLock) " & _
                                        "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                        "inner join BannerEngines BE with (NoLock) on BE.BannerID = AUB.BannerID " & _
                                        "where BAN.Deleted=0 and AdminUserID = " & AdminUserID & " and BE.EngineID=" & TempEngineID & " " & _
                                        "order by BAN.Name;"
                    'End If
                    rst = MyCommon.LRT_Select
                    BannerCt = rst.Rows.Count
                    If (BannerCt > 0) Then
                        BannerID = MyCommon.Extract_Val(Request.QueryString("banner"))
                        Send("<tr id=""banners"">")
                        Send("  <td valign=""top"">")
                        Send("    <label for=""banner"">" & Copient.PhraseLib.Lookup("term.banners", LanguageID) & ":</label>")
                        Send("  </td>")
                        Send("  <td valign=""top"">")
                        Send("    <select class=""longest"" name=""banner"" id=""banner""" & If(AllowMultipleBanners, " size=""5"" multiple=""multiple""", "") & ">")
                        i = 0
                        For Each row In rst.Rows
                            If (AllowMultipleBanners AndAlso MyCommon.NZ(row.Item("AllBanners"), False)) Then
                                ' exclude this all banners from the list box and store the option for later display
                                i += 1
                                AllBannersCheckBox &= "<input type=""checkbox"" name=""allbannersid"" id=""allbannersid1" & i & """ value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """ onclick=""updateBanners(this.checked);"" />"
                                AllBannersCheckBox &= "<label for=""allbannersid1" & i & """>" & MyCommon.NZ(row.Item("Name"), "") & "</label><br />"
                                AllBannersCheckBox &= "<br class=""half"" />"
                            Else
                                Send("      <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """" & IIf(BannerID = MyCommon.NZ(row.Item("BannerID"), -1), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                            End If
                        Next
                        Send("    </select><br />")
                        Send("    <br class=""half"" />")
                        If (AllBannersCheckBox <> "") Then
                            Send(AllBannersCheckBox)
                        End If
                        Send("  </td>")
                        Send("</tr>")
                    End If
                End If

                If (BannerCt < 1 AndAlso EngineCt = 1 And Request.QueryString("save") = "" And Not Logix.UserRoles.CreateExternalOffers) Then
                    ' there is only one engine installed lets just make the offer
                    ' Response.Redirect("offer-new.aspx?save=Save&NewTemplate=" & IsNewTemplate & "&name=" & Server.HtmlEncode("Offer-" & Now.ToString) & "&EngineID=" & EngineID & "&banner=" & BannerID)
                End If

                ' create a new offer as a proxy for an external source
                If Logix.UserRoles.CreateExternalOffers Then
                    MyCommon.QueryStr = "select ExtInterfaceID, ExtCode, ExtCodePhraseID , NamePhraseID, Name from ExtCRMInterfaces where Active=1 and Deleted=0 and ExtInterfaceID>0 and ExtInterfaceTypeID<2;"
                    rst = MyCommon.LRT_Select
                    Dim ExtCodePhraseId As Integer
                    Dim NamePhraseID As Integer
                    If External = True Then
                        If rst.Rows.Count > 0 Then
                            Send("<tr>")
                            Send("  <td valign=""top"" colspan=""2"" >")
                            Send("    <label for=""crmEngineID"">" & Copient.PhraseLib.Lookup("offer-new.create-external", LanguageID) & ":</label>")
                            Send("  </td>")
                            Send("  <tr>")
                            Send("  <td></td>")
                            Send("  <td valign=""top"">")
                            Send("    <select name=""crmEngineID"" id=""crmEngineID"" >")
                            Send("      <option value=""-1"">[" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "]</option>")
                            For Each row In rst.Rows
                                ExtCodePhraseId = IIf(IsDBNull(row.Item("ExtCodePhraseID")), -1, row.Item("ExtCodePhraseID"))
                                NamePhraseID = IIf(IsDBNull(row.Item("NamePhraseID")), -1, row.Item("NamePhraseID"))
                                Send("      <option value=""" & MyCommon.NZ(row.Item("ExtInterfaceID"), -1) & """>" & IIf(ExtCodePhraseId > 0, Copient.PhraseLib.Lookup(ExtCodePhraseId, LanguageID, "PhraseNotFound"), row.Item("ExtCode")) & " - " & IIf(NamePhraseID > 0, Copient.PhraseLib.Lookup(NamePhraseID, LanguageID, "PhraseNotFound"), row.Item("Name")) & "</option>")
                            Next
                            Send("    </select>")
                            Send("  </td>")
                            Send("</tr>")
                        End If
                    End If
                End If
                ' folder assignments for offer
                If (Logix.UserRoles.EditFolders OrElse Logix.UserRoles.AssignFolders) Then
                    Send("<tr>")
                    Send("  <td valign=""top"">")
                    Send("   " & Copient.PhraseLib.Lookup("term.folders", LanguageID) & ":")
                    Send("  </td>")
                    Send("  <td valign=""top"" id=""folderNames"">")
                    Dim folderList As String
                    'If Buyer is selected by default, get default folder assigned for Buyer
                    If (BuyerSelected Or DefaultEngine = "UE") Then
                        MyCommon.QueryStr = "select FI.FolderID,F.FolderName from FolderItems FI " & _
                                 "inner join Folders F on FI.FolderID=F.FolderID where(LinkID = " & tempbuyerId & "And LinkTypeID = 2)"
                        rst = MyCommon.LRT_Select
                        Dim folderNames As String = ""
                        folderNames += "<ul>"
                        For Each row In rst.Rows
                            folderList = MyCommon.NZ(row.Item("FolderID"), "").ToString()
                            folderNames += "<li>"
                            folderNames += MyCommon.NZ(row.Item("FolderName"), "")
                            folderNames += "</li>"
                        Next
                        'If Buyer is not assigned any folder, get the default folder set for UE
                        If DefaultEngine = "UE" AndAlso folderNames = "<ul>" Then
                            MyCommon.QueryStr = "select FolderID,FolderName from Folders where DefaultUEFolder=1"
                            rst = MyCommon.LRT_Select
                            For Each row In rst.Rows
                                folderList = MyCommon.NZ(row.Item("FolderID"), "").ToString()
                                folderNames += "<li>"
                                folderNames += MyCommon.NZ(row.Item("FolderName"), "")
                                folderNames += "</li>"
                            Next
                        End If
                        folderNames += "</ul>"
                        Send(folderNames)
                        'If no folder set, then show none selected
                    Else
                        'If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(132)) = 1) Then
                        Send(Copient.PhraseLib.Lookup("term.noneselected", LanguageID))
                        'End If

                    End If
                    Send("</td>")
                    Send("</tr>")
                    Send("<tr>")
                    Send("  <td>&nbsp;</td>")
                    Send("  <td>")
                    Send("    <input type=""hidden"" id=""folderList"" name=""folderList"" value=""" & folderList & """ />")
                    Send("    <input type=""button"" class=""regular"" name=""btnBrowse"" id=""btnBrowse"" value=""" & Copient.PhraseLib.Lookup("term.browse", LanguageID) & """ onclick=""javascript:openPopup('folder-browse.aspx');"" />")
                    Send("  </td>")
                    Send("</tr>")
                End If

            %>
        </table>
    </div>
</div>
</form>
<%
done:
    Send_BodyEnd("mainform", "name")
    MyCommon.Close_LogixRT()
    MyCommon = Nothing
    Logix = Nothing
%>