<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%
' *****************************************************************************
' * FILENAME: offer-channels.aspx 
' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' * Copyright © 2002 - 2013.  All rights reserved by:
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
%>
<script runat="server">

    Public Structure OfferData
        Dim OfferID As Long
        Dim RewardOptionID As Long
        Dim Name As String
        Dim IsTemplate As Boolean
        Dim FromTemplate As Boolean
        Dim EngineID As Integer
        Dim EngineSubTypeID As Integer
        Dim HasAnyCustomer As Boolean
        Dim TemplatePermissions As TemplatePermissionTypes
    End Structure

    Public Structure TemplatePermissionTypes
        Dim DisAllow_Channels As Boolean
    End Structure

    Public Structure Channel
        Dim ID As Integer
        Dim Name As String
        Dim Selected As Boolean
        Dim UsesStartDate As Boolean
        Dim UsesEndDate As Boolean
        Dim UsesLimits As Boolean
        Dim StartDate As Date
        Dim EndDate As Date
        Dim Limit As EligibilityLimit
        Dim PrintedMessage As PrintedMessageRec
        Dim PosImgURL As PosImgUrlRec
        Dim CashierMessage As CashierMessageRec
        Dim AccumMessage As AccumMessageRec
        Dim Graphic As List(Of GraphicDeliverable)
        Dim Assets As List(Of ChannelOfferAssestRec)
        Dim PosNotificationCheck As Integer
    End Structure

    Public Structure EligibilityLimit
        Dim P1DistQtyLimit As Integer
        Dim P1DistPeriod As Integer
        Dim P1DistTimeType As Integer
    End Structure

    Public Structure PrintedMessageRec
        Dim MessageID As Integer
        Dim DeliverableID As Integer
        Dim BodyText As String
    End Structure

    Public Structure PosImgUrlRec
        Dim PTPKID As Integer
        Dim DeliverableID As Integer
        Dim Data As String
    End Structure

    Public Structure CashierMessageRec
        Dim DeliverableID As Integer
        Dim MessageID As Integer
        Dim Line1 As String
        Dim Line2 As String
    End Structure

    Public Structure GraphicDeliverable
        Dim DeliverableID As Integer
        Dim OnScreenAdID As Integer
        Dim Name As String
        Dim CellSelectID As Integer
        Dim ImageType As Integer
        Dim URI As String
    End Structure

    Public Structure AccumMessageRec
        Dim MessageID As Integer
        Dim DeliverableID As Integer
        Dim BodyText As String
    End Structure

    Public Structure ChannelOfferAssestRec
        Dim MediaTypeID As Integer
        Dim LanguageID As Integer
        Dim MediaData As String
        Dim Name As String
        Dim UIFormField As DisplayTypes
        Dim MediaParamTypes As List(Of MediaParamTypeRec)
    End Structure

    Public Structure MediaParamTypeRec
        Dim ParamTypeID As Integer
        Dim Name As String
        Dim DecimalPlaces As Integer
        Dim MinValue As Decimal
        Dim MaxValue As Decimal
    End Structure

    Public Structure TimeType
        Dim ID As Integer
        Dim Name As String
    End Structure

    Public Structure Language
        Dim ID As Integer
        Dim Name As String
    End Structure

    Public Structure LegacyStatusRec
        Dim DisabledOnCFW As Boolean
        Dim DisplayOnWebKiosk As Boolean
        Dim UseLegacy As Boolean
    End Structure

    Public Enum LimitTypes As Integer
        CUSTOM = 0
        NO_LIMIT = 1
        ONCE_PER_TRANSACTION = 2
        ONCE_PER_DAY = 3
        ONCE_PER_WEEK = 4
        ONCE_PER_OFFER = 5
        DAYS_ROLLING = 6
        PER_TRANSACTION = 7
    End Enum

    Public Enum DisplayTypes As Integer
        UNKNOWN = 0
        INPUT_FILE = 3
        INPUT_TEXT = 10
        SELECT_MULTIPLE = 12
        POS_PRINTED_MESSAGE = 14
        POS__CASHIER_MESSAGE = 15
        POS_GRAPHIC = 16
        POS_ACCUM_MESSAGE = 17
        POS_IMAGE_URL = 18
    End Enum

    Public Enum MediaTypes As Integer  'these correspond to the metadata in the ChannelMedaiTypes table
        POS_Receipt_Message = 1
        Website_Offer_Description = 2
        Website_Offer_Graphic = 3
        Broker_Offer_Description = 4
        Broker_Offer_Graphic = 5
        POS_Cashier_Message = 6
        Kiosk_Offer_Description = 7
        Kiosk_Offer_Graphic = 8
        POS_Graphic = 9
        Broker_Offer_Name = 10
        POS_Accumulation_Message = 11
        POS_Image_Url = 17
    End Enum

    Private Const POS_CHANNEL_ID As Integer = 1

    Dim Common As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim BrokerChannel As Copient.BrokerChannel
    Dim ChannelCommon As Copient.ChannelCommon
    Dim Handheld As Boolean = False
    Dim BannersEnabled As Boolean
    Dim Offer As New OfferData
    Dim Channels As New List(Of Channel)
    Dim TimeTypes As New List(Of TimeType)
    Dim AvailableLanguages As New List(Of Language)
    Dim MyCommon As New Copient.CommonInc

    Property m_EditOfferRegardlessOfBuyer As Boolean
        Get
            MyCommon.Open_LogixRT()
            Return Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, Offer.OfferID)
        End Get
        Set(value As Boolean)

        End Set
    End Property
    ' Dim m_EditOfferRegardlessOfBuyer As Boolean   = Logix.UserRoles.EditOffersRegardlessBuyer Or MyCommon.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, Offer.OfferID) 
    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_Page()
        Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
        Dim CopientFileVersion As String = "7.3.1.138972"
        Dim CopientProject As String = "Copient Logix"
        Dim CopientNotes As String = ""
        Dim ShowPage As Boolean = False

        'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
        CheckIfValidOffer(MyCommon, Offer.OfferID)

        Send_HeadBegin("term.offer", "term.channels", Offer.OfferID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts(New String() {"datePicker.js", "popup.js", "ajaxSubmit.js"})
        Send_Page_Script()
        Send_HeadEnd()

        ' send page chrome
        Send_BodyBegin(IIf(Offer.IsTemplate, 11, 1))
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 2)

        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)

        Select Case Offer.EngineID
            Case CommonInc.InstalledEngines.CM
                Send_Subtabs(Logix, IIf(Offer.IsTemplate, 22, 21), 9, , Offer.OfferID)
            Case CommonInc.InstalledEngines.CPE
                Send_Subtabs(Logix, IIf(Offer.IsTemplate, 25, 24), 7, , Offer.OfferID)
            Case CommonInc.InstalledEngines.CAM
                Send_Subtabs(Logix, IIf(Offer.IsTemplate, 206, 205), 7, , Offer.OfferID)
            Case CommonInc.InstalledEngines.UE
                Send_Subtabs(Logix, IIf(Offer.IsTemplate, 209, 208), 7, , Offer.OfferID)
            Case Else 'CPE
                Send_Subtabs(Logix, IIf(Offer.IsTemplate, 25, 24), 7, , Offer.OfferID)
        End Select

        ' check permissions
        If (Logix.UserRoles.AccessOffers = False AndAlso Not Offer.IsTemplate) Then
            Send_Denied(1, "perm.offers-access")
        ElseIf (Logix.UserRoles.AccessTemplates = False AndAlso Offer.IsTemplate) Then
            Send_Denied(1, "perm.offers-access-templates")
        ElseIf (Logix.UserRoles.AccessInstantWinOffers = False AndAlso Offer.EngineID = CommonInc.InstalledEngines.CPE AndAlso Offer.EngineSubTypeID = 1) Then
            Send_Denied(1, "perm.offers-accessinstantwin")
        ElseIf (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, Offer.OfferID)) Then
            Send("<script type=""text/javascript"" language=""javascript"">")
            Send("  function updateCookie() { return true; } ")
            Sendb("</")
            Send("script>")
            Send_Denied(1, "banners.access-denied-offer")
        Else
            ShowPage = True
        End If

        If ShowPage Then
            Send_Intro()
            Send_Main()
            Send_Graphic_Selector()
        End If

        If Common.Fetch_SystemOption(75) Then
            If (Offer.OfferID > 0 And Logix.UserRoles.AccessNotes AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                Send_Notes(3, Offer.OfferID, AdminUserID)
            End If
        End If

        Send_BodyEnd()



    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Page_Script()
        Send("<script type=""text/javascript"">")
        Send("")
        Send("  window.name = ""offerChannels"";")
        Send("  var datePickerDivID = ""datepicker"";")
        Send("")
        Send("  if (window.captureEvents) {")
        Send("    window.captureEvents(Event.CLICK);")
        Send("    window.onclick=handlePageClick;")
        Send("  } else {")
        Send("    document.onclick=handlePageClick;")
        Send("  }")
        Send("")
        Send_Calendar_Overrides(MyCommon)
        Send("")
        Send("  function setCheckboxValue(obj) { ")
        Send("  document.getElementById('SaveChLockState').src='../images/save.png';")
        Send("  if(obj.checked)")
        Send("    document.getElementById(""IsChannelsLocked"").value=1;")
        Send("  else")
        Send("    document.getElementById(""IsChannelsLocked"").value=0;")
        Send("  } ")
        Send("  function saveChannelLockState() {")
        Send("      document.getElementById('mainform').submit();")
        Send("  }")
        Send("")
        Send("  function SetReadonlyChannel(channel) {")
        Send("    var channelobj = document.getElementById('tblchannel' + channel);")
        Send("    var controls = channelobj.getElementsByTagName('a');")
        Send("    for (i = 0; i < controls.length; i++) {")
        Send("      controls[i].onclick = function () {")
        Send("        return false;")
        Send("      }")
        Send("    }")
        Send("    controls = channelobj.getElementsByTagName('input');")
        Send("    for (i = 0; i < controls.length; i++) {")
        Send("      control = controls[i];")
        Send("      control.disabled = true;")
        Send("    }")
        Send("    controls = channelobj.getElementsByTagName('select');")
        Send("    for (i = 0; i < controls.length; i++) {")
        Send("      control = controls[i];")
        Send("      control.disabled = true;")
        Send("    }")
        Send("    controls = channelobj.getElementsByTagName('img');")
        Send("    for (i = 0; i < controls.length; i++) {")
        Send("      if (controls[i].id != 'chExpand' + channel) {")
        Send("        controls[i].onclick = function () {")
        Send("          return false;")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  } ")
        Send("  function handlePageClick(e) {")
        Send("    var el=(typeof event!=='undefined')? event.srcElement : e.target        ")
        Send("  ")
        Send("    if (el != null) {")
        Send("      var pickerDiv = document.getElementById(datePickerDivID);")
        Send("      if (pickerDiv != null && pickerDiv.style.visibility == ""visible"") {")
        Send("        if (el.id.indexOf(""Start-picker"") != 0 && el.id.indexOf(""End-picker"") != 0) { ")
        Send("          if (!isDatePickerControl(el.className)) {")
        Send("            pickerDiv.style.visibility = ""hidden"";")
        Send("            pickerDiv.style.display = ""none""; ")
        Send("          }")
        Send("        } else  {")
        Send("          pickerDiv.style.visibility = ""visible"";            ")
        Send("          pickerDiv.style.display = ""block"";     ")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function isDatePickerControl(ctrlClass) {")
        Send("    var retVal = false;")
        Send("    ")
        Send("    if (ctrlClass != null && ctrlClass.length >= 2) {")
        Send("      if (ctrlClass.substring(0,2) == ""dp"") {")
        Send("        retVal = true;")
        Send("      }")
        Send("    }")
        Send("    return retVal;")
        Send("  }")
        Send("")
        Send("  function handleChannelDetail(channelId) {")
        Send("    var dtlElem = document.getElementById('chDetail' + channelId);")
        Send("    var imgElem = document.getElementById('chExpand' + channelId);")
        Send("    var chElem  = document.getElementById('channel' + channelId);")
        Send("    var expand = (imgElem != null && imgElem.src.indexOf('plus.png') > -1) ? true : false;")
        Send("")
        Send("    if (dtlElem != null) { ")
        Send("      dtlElem.style.display = (expand) ? """" : ""none"";")
        Send("    }")
        Send("")
        Send("    if (imgElem != null) { ")
        Send("      imgElem.src = (expand) ? ""/images/minus.png"" : ""/images/plus.png"";")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function selectChannel(channelId) {")
        Send("    var elemSave = document.getElementById('btnSaveCh' + channelId);")
        Send("    var elemFrm = document.getElementById('frmChannel' + channelId);")
        Send("    var frmIsChanged = false;")
        Send("    ")
        Send("    if (elemFrm != null && elemSave != null) {")
        Send("      frmIsChanged = IsFormChanged(elemFrm);")
        Send("      updateChannelSave(channelId, frmIsChanged);")
        Send("      updateChannelStatus(channelId, (frmIsChanged) ? '" & Copient.PhraseLib.Lookup("offer-channels.unsavedChanges", LanguageID) & "' : '" & Copient.PhraseLib.Lookup("term.ready", LanguageID) & "');")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateP1limit(channelId) {")
        Send("    if(document.getElementById('p1typeCh' + channelId).value == 7   ) {")
        Send("    document.getElementById('lbllimit1periodCh' + channelId).style.display='none';")
        Send("    document.getElementById('limit1periodCh' + channelId).value='-1';")
        Send("    document.getElementById('limit1periodCh' + channelId).style.display='none';")
        Send("  }else {")
        Send("    document.getElementById('lbllimit1periodCh' + channelId).style.display='';")
        Send("    document.getElementById('limit1periodCh' + channelId).value='0';")
        Send("    document.getElementById('limit1periodCh' + channelId).style.display='';")
        Send("  }")
        Send("    if(document.getElementById('p1typeCh' + channelId).value == 0   ) {")
        Send("      document.getElementById('customRowCh' + channelId).style.display ='';")
        If Offer.HasAnyCustomer Then
            Send("    document.getElementById('limit1periodCh' + channelId).value='1';")
            Send("    document.getElementById('P1DistTimeTypeCh' + channelId).value='2';")
        End If
        Send("  } else if(document.getElementById('p1typeCh' + channelId).value == 6   ) {")
        Send("      document.getElementById('customRowCh' + channelId).style.display ='';")
        Send("    document.getElementById('P1DistTimeTypeCh' + channelId).value='1';")
        If Offer.HasAnyCustomer Then
            Send("    document.getElementById('limit1periodCh' + channelId).value='1';")
        End If
        Send("  } else if(document.getElementById('p1typeCh' + channelId).value == 7   ) {")
        Send("      document.getElementById('customRowCh' + channelId).style.display ='';")
        Send("      document.getElementById('P1DistTimeTypeCh' + channelId).value='2';")
        Send("    } else {")
        Send("      document.getElementById('customRowCh' + channelId).style.display ='none';")
        Send("      if(document.getElementById('p1typeCh' + channelId).value == '1'){")
        Send("      // no limit 0 0 2")
        Send("        document.getElementById('limit1Ch' + channelId).value = '0';")
        Send("        document.getElementById('limit1periodCh' + channelId).value = '0';")
        Send("        document.getElementById('P1DistTimeTypeCh' + channelId).value = '2';")
        Send("      }")
        Send("      else if(document.getElementById('p1typeCh' + channelId).value == '2'){")
        Send("      // no limit 0 0 2")
        Send("        document.getElementById('limit1Ch' + channelId).value = '1';")
        Send("        document.getElementById('limit1periodCh' + channelId).value = '1';")
        Send("        document.getElementById('P1DistTimeTypeCh' + channelId).value = '2';")
        Send("      }")
        Send("      else if(document.getElementById('p1typeCh' + channelId).value == '3'){")
        Send("      // no limit 0 0 2")
        Send("        document.getElementById('limit1Ch' + channelId).value = '1';")
        Send("        document.getElementById('limit1periodCh' + channelId).value = '1';")
        Send("        document.getElementById('P1DistTimeTypeCh' + channelId).value = '1';")
        Send("      }")
        Send("      else if(document.getElementById('p1typeCh' + channelId).value == '4'){")
        Send("      // no limit 0 0 2")
        Send("        document.getElementById('limit1Ch' + channelId).value = '1';")
        Send("        document.getElementById('limit1periodCh' + channelId).value = '7';")
        Send("        document.getElementById('P1DistTimeTypeCh' + channelId).value = '1';")
        Send("      }")
        Send("      else if(document.getElementById('p1typeCh' + channelId).value == '5'){")
        Send("      // no limit 0 0 2")
        Send("        document.getElementById('limit1Ch' + channelId).value = '1';")
        Send("        document.getElementById('limit1periodCh' + channelId).value = '3650';")
        Send("        document.getElementById('P1DistTimeTypeCh' + channelId).value = '1';")
        Send("      }")
        Send("    }")
        Send("    selectChannel(channelId);")
        Send("  }")
        Send("")
        Send("  function openPMsgPopup(deliverableID, messageID) {")
        If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
            Send("    var pageName = 'CPEoffer-rew-pmsg.aspx';")
        ElseIf Offer.EngineID = CommonInc.InstalledEngines.UE Then
            Send("    var pageName = 'UE/UEoffer-rew-pmsg.aspx';")
        Else
            Send("    var pageName = 'offer-rew-pmsg.aspx';")
        End If
        Send("    var qryStr = '?DeliverableID=' + deliverableID + '&OfferID=" & Offer.OfferID & "&Phase=1&RewardID=" & Offer.RewardOptionID & "&MessageID=' + messageID;")
        Send("")
        Send("    openPopup(pageName + qryStr);")
        Send("  }")
        Send("")

        Send("  function openImgUrlPopup(deliverableID, PTPKID) {")
        Send("    var pageName = 'UE/pos-channel-imageurl.aspx';")
        Send("    var qryStr = '?DeliverableID=' + deliverableID + '&OfferID=" & Offer.OfferID & "&PTPKID=' + PTPKID;")
        Send("")
        Send("    openPopup(pageName + qryStr);")
        Send("  }")
        Send("")

        Send("  function openCMsgPopup(deliverableID, messageID) {")
        If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
            Send("    var pageName = 'CPEoffer-rew-cmsg.aspx';")
        ElseIf Offer.EngineID = CommonInc.InstalledEngines.UE Then
            Send("    var pageName = 'UE/UEoffer-rew-cmsg.aspx';")
        Else
            Send("    var pageName = 'offer-rew-cmsg.aspx';")
        End If
        Send("    var qryStr = '?DeliverableID=' + deliverableID + '&OfferID=" & Offer.OfferID & "&Phase=1&RewardID=" & Offer.RewardOptionID & "&MessageID=' + messageID;")
        Send("")
        Send("    openPopup(pageName + qryStr);")
        Send("  }")
        Send("")
        Send("  function openAddGraphicsPopup() {")
        Send("    var pageName = 'CPEoffer-rew-graphic.aspx';")
        Send("    var qryStr = '?DeliverableID=0&OfferID=" & Offer.OfferID & "&Phase=1&RewardID=" & Offer.RewardOptionID & "';")
        Send("")
        Send("    openPopup(pageName + qryStr);")
        Send("  }")
        Send("")
        Send("  function openPreviewGraphicsPopup(ad, cell, imgType) {")
        Send("    var pageName = 'CPEoffer-rew-graphic.aspx';")
        Send("    var qryStr = '?OfferID=" & Offer.OfferID & "&ad=' + ad + '&cellselect=' + cell + '&imagetype=' + imgType +'&preview=1&Phase=1';")
        Send("")
        Send("    openPopup(pageName + qryStr);")
        Send("  }")
        Send("")
        Send("  function openAccumMsgPop(deliverableID, messageID) {")
        If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
            Send("    var pageName = 'CPEoffer-rew-pmsg.aspx';")
        ElseIf Offer.EngineID = CommonInc.InstalledEngines.UE Then
            Send("    var pageName = 'UE/UEoffer-rew-pmsg.aspx';")
        Else
            Send("    var pageName = 'offer-rew-pmsg.aspx';")
        End If
        Send("    var qryStr = '?DeliverableID=' + deliverableID + '&OfferID=" & Offer.OfferID & "&Phase=2&RewardID=" & Offer.RewardOptionID & "&MessageID=' + messageID;")
        Send("")
        Send("    openPopup(pageName + qryStr);")
        Send("  }")
        Send("")
        Send("  function saveChannel(channelId) {")
        Send("    var elemSave = document.getElementById('btnSaveCh' + channelId);")
        Send("    var elemMode = document.getElementById('modeChannel' + channelId);")
        Send("")
        Send("    if (elemSave != null && elemSave.src.indexOf('save.png') > 0 && elemMode != null) {")
        Send("      elemMode.value = 'SaveChannel';")
        Send("      xmlhttpPostForm('offer-channels.aspx', 'frmChannel' + channelId, updateFromSave);")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateFromSave(jsonResponse) {")
        Send("    try {")
        Send("      // convert into an object ")
        Send("      if (jsonResponse != null) {")
        Send("        var obj = eval ('(' + jsonResponse + ')');")
        Send("        if (obj != null) {")
        Send("          updateChannelStatus(obj.ChannelID, obj.Message);")
        Send("          updateChannelSave(obj.ChannelID, (obj.Status == 0));")
        Send("          updateChannelDeployMsg(obj.ChannelID, obj.DeployMsg, obj.DeployMsgColor);")
        Send("          DisplayOfferModified(); ")
        Send("          if (obj.Status == 1) { ")
        Send("            updateActiveIcon(obj.ChannelID, true);")
        Send("            updateDefaultValues(obj.ChannelID);")
        Send("          } else { ")
        Send("            alert(obj.Message);")
        Send("          }")
        Send("        }")
        Send("      }")
        Send("    } catch(err)  {")
        'Send("      alert('Error Encountered! \r\n\r\n' + err + '\r\n' + jsonResponse);") '- Use this for debug
        Send("      alert('Error encountered while processing your request!, Please try again');")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function DisplayOfferModified() { ")
        If Not (IsOfferExpired(Offer.OfferID, Offer.EngineID)) Then
            Send("    if (document.getElementById('statusbar')) { ")
            Send("      document.getElementById('statusbar').style.display='block'; ")
            If HasOfferBeenDeployed(Offer.OfferID, Offer.EngineID) Then
                Send("      document.getElementById('statusbar').className='red-background'; ")
                Send("      document.getElementById('statusbar').innerHTML='" & Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID) & "'; ")
            Else
                Send("      document.getElementById('statusbar').className='orange-background'; ")
                Send("      document.getElementById('statusbar').innerHTML='" & Copient.PhraseLib.Lookup("offer.status1msg", LanguageID) & "'; ")
            End If
            Send("    } ")
        End If
        Send("  } ")
        Send("")
        Send("  function updateChannelStatus(channelID, msg) {")
        Send("    var elemStatus = document.getElementById('channelMsg' + channelID);")
        Send("    if (elemStatus != null) {")
        Send("      elemStatus.innerHTML = msg;")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateChannelDeployMsg(channelID, deployMsg, deployMsgColor) {")
        Send("    var elemDeploy = document.getElementById('ChannelDeployMsg' + channelID);")
        Send("    if (elemDeploy != null) {")
        Send("      elemDeploy.innerHTML = deployMsg;")
        Send("      elemDeploy.style.color = deployMsgColor;")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateChannelSave(channelID, enabled) {")
        Send("    var elemSave = document.getElementById('btnSaveCh' + channelID);")
        Send("    if (elemSave != null) {")
        Send("      //elemSave.src = (enabled) ? '../images/save.png' : '../images/save-off.png';")
        Send("      elemSave.title = (enabled) ? '" & Copient.PhraseLib.Lookup("offer-channels.unsavedChanges", LanguageID) & "' : '" & Copient.PhraseLib.Lookup("offer-channels.noUnsavedChanges", LanguageID) & "';")
        Send("      elemSave.alt = (enabled) ? '" & Copient.PhraseLib.Lookup("offer-channels.unsavedChanges", LanguageID) & "' : '" & Copient.PhraseLib.Lookup("offer-channels.noUnsavedChanges", LanguageID) & "';")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function channelChangesMade(channelID) {")
        Send("    updateChannelSave(channelID, true);")
        Send("    updateChannelStatus(channelID, '" & Copient.PhraseLib.Lookup("offer-channels.unsavedChanges", LanguageID) & "');")
        Send("  }")
        Send("")
        Send("  function removeChannel(channelId, modeElemName) {")
        Send("    var elemMode = document.getElementById('modeChannel' + channelId);")
        Send("")
        Send("    if (elemMode != null) {")
        Send("      elemMode.value = 'RemoveChannel';")
        Send("      xmlhttpPostForm('offer-channels.aspx', 'frmChannel' + channelId, updateFromRemove);")
        Send("      DisplayOfferModified(); ")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateFromRemove(jsonResponse) {")
        Send("    if (jsonResponse != null) {")
        Send("      // convert into an object ")
        Send("      var obj = eval ('(' + jsonResponse + ')');")
        Send("      if (obj != null) {")
        Send("        updateChannelStatus(obj.ChannelID, obj.Message);")
        Send("        updateChannelDeployMsg(obj.ChannelID, obj.DeployMsg, obj.DeployMsgColor);")
        Send("        updateChannelSave(obj.ChannelID, false);")
        Send("        updateActiveIcon(obj.ChannelID, (parseInt(obj.Status)==0));")
        Send("        if (obj.ChannelID==1){  // reset the buttons for the POS channel assets ")
        Send("          if(document.getElementById('btnAddCmsgCh1') != null) { // reset the POS cashier message add button ")
        Send("            document.getElementById('btnAddCmsgCh1').disabled = false; ")
        Send("          } ")
        Send("          if(document.getElementById('btnAddPmsgCh1') != null) { // reset the POS printed message add button ")
        Send("            document.getElementById('btnAddPmsgCh1').disabled = false; ")
        Send("          } ")
        Send("          if(document.getElementById('btnAddImgUrlCh1') != null) { // reset the POS printed message add button ")
        Send("            document.getElementById('btnAddImgUrlCh1').disabled = false; ")
        Send("          } ")
        Send("          if(document.getElementById('btnAddGraphicCh1') != null) { // reset the POS graphic add button ")
        Send("            document.getElementById('btnAddGraphicCh1').disabled = false; ")
        Send("          } ")
        Send("          if(document.getElementById('btnDelCmsgCh1') != null) { // reset the POS cashier message delete button ")
        Send("            document.getElementById('btnDelCmsgCh1').disabled = true; ")
        Send("          } ")
        Send("          if(document.getElementById('btnDelPmsgCh1') != null) { // reset the POS printed message delete button ")
        Send("            document.getElementById('btnDelPmsgCh1').disabled = true; ")
        Send("          } ")
        Send("          if(document.getElementById('btnDelImgUrlCh1') != null) { // reset the POS printed message delete button ")
        Send("            document.getElementById('btnDelImgUrlCh1').disabled = true; ")
        Send("          } ")
        Send("          if(document.getElementById('btnAddGraphicCh1') != null) { // reset the POS graphic delete button ")
        Send("            document.getElementById('btnDelGraphicCh1').disabled = true; ")
        Send("          }")
        Send("        }")
        Send("        var elemFrm = document.getElementById('frmChannel' + obj.ChannelID);")
        Send("        if (elemFrm != null) {")
        Send("          clearForm(elemFrm);")
        Send("          clearAssets(obj.DelAssets);")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function clearAssets(DelAssets) { ")
        Send("    for(var i in DelAssets) { ")
        Send("      var obj=DelAssets[i]; ")
        Send("      var elem = document.getElementById(obj.assetID); ")
        Send("      if (elem != null) { ")
        Send("        elem.style.display = 'none'; ")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateActiveIcon(channelId, enabled) {")
        Send("    elem = document.getElementById('activeImgCh' + channelId);")
        Send("  ")
        Send("    if (elem != null) {")
        Send("      elem.src = '../images/star-' + ((enabled) ? 'on' : 'off') + '.png';")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function clearForm(elemFrm) {")
        Send("    if (elemFrm != null) {")
        Send("      var frm_elements = elemFrm.elements;")
        Send("")
        Send("      for(var i=0; i<frm_elements.length; i++) {")
        Send("        field_type = frm_elements[i].type.toLowerCase();")
        Send("        switch(field_type) {")
        Send("          case ""text"":")
        Send("          case ""password"":")
        Send("          case ""textarea"":")
        Send("            frm_elements[i].value = '';")
        Send("            frm_elements[i].defaultValue = '';")
        Send("            break;")
        Send("          case ""radio"":")
        Send("          case ""checkbox"":")
        Send("            if (frm_elements[i].checked) {")
        Send("              frm_elements[i].checked = false;")
        Send("              frm_elements[i].defaultChecked = false;")
        Send("            }")
        Send("            break;")
        Send("          case ""select-one"":")
        Send("          case ""select-multi"":")
        Send("            frm_elements[i].selectedIndex = -1;")
        Send("            frm_elements[i].defaultSelected = null;")
        Send("            break;")
        Send("          default:")
        Send("            break;")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateDefaultValues(channelId) {")
        Send("    var elemFrm = document.getElementById('frmChannel' + channelId);")
        Send("")
        Send("    if (elemFrm != null) {")
        Send("      var frm_elements = elemFrm.elements;")
        Send("      for(var i=0; i<frm_elements.length; i++) {")
        Send("        field_type = frm_elements[i].type.toLowerCase();")
        Send("        switch(field_type) {")
        Send("          case ""text"":")
        Send("          case ""password"":")
        Send("          case ""textarea"":")
        Send("            frm_elements[i].defaultValue = frm_elements[i].value;")
        Send("            break;")
        Send("          case ""radio"":")
        Send("          case ""checkbox"":")
        Send("            frm_elements[i].defaultChecked = frm_elements[i].checked;")
        Send("            break;")
        Send("          case ""select-one"":")
        Send("          case ""select-multi"":")
        Send("            frm_elements[i].options[frm_elements[i].selectedIndex].defaultSelected = true;")
        Send("            break;")
        Send("          default:")
        Send("            break;")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function removeOfferAsset(channelId, mediaTypeId, langId) {")
        Send("    xmlhttpPostDataCallback('offer-channels.aspx','mode=RemoveAsset&OfferID=" & Offer.OfferID & "&ChannelID=' + channelId + '&MediaTypeID=' + mediaTypeId + '&LanguageID=' + langId, updateRemovedAsset);")
        Send("  }")
        Send("")
        Send("  function removeCashierMsg(channelId, deliverableId, messageId) {")
        Send("    xmlhttpPostDataCallback('offer-channels.aspx','mode=RemoveCMsg&OfferID=" & Offer.OfferID & "&ChannelID=' + channelId + '&DeliverableID=' + deliverableId + '&MessageID=' + messageId, updateRemovedPosAsset);")
        Send("  }")
        Send("")
        Send("  function removePrintedMsg(channelId, deliverableId, messageId) {")
        Send("    xmlhttpPostDataCallback('offer-channels.aspx','mode=RemovePMsg&OfferID=" & Offer.OfferID & "&ChannelID=' + channelId + '&DeliverableID=' + deliverableId + '&MessageID=' + messageId, updateRemovedPosAsset);")
        Send("  }")
        Send("")
        Send("  function removeImageUrl(channelId, deliverableId, PTPKID) {")
        Send("    xmlhttpPostDataCallback('offer-channels.aspx','mode=RemoveImgUrl&OfferID=" & Offer.OfferID & "&ChannelID=' + channelId + '&DeliverableID=' + deliverableId + '&PTPKID=' + PTPKID, updateRemovedPosAsset);")
        Send("  }")
        Send("")
        Send("  function removeGraphic(channelId, deliverableId, onScreenAdId) {")
        Send("    xmlhttpPostDataCallback('offer-channels.aspx','mode=RemoveGraphic&OfferID=" & Offer.OfferID & "&ChannelID=' + channelId + '&DeliverableID=' + deliverableId + '&OnScreenAdID=' + onScreenAdId, updateRemovedPosAsset);")
        Send("  }")
        Send("")
        Send("  function removeAccumMsg(channelId, deliverableId, messageId) {")
        Send("    xmlhttpPostDataCallback('offer-channels.aspx','mode=RemoveAMsg&OfferID=" & Offer.OfferID & "&ChannelID=' + channelId + '&DeliverableID=' + deliverableId + '&MessageID=' + messageId, updateRemovedPosAsset);")
        Send("  }")
        Send("")
        Send("  function updateRemovedAsset(jsonResponse) {")
        Send("    if (jsonResponse != null) {")
        Send("      // convert into an object ")
        Send("      var obj = eval ('(' + jsonResponse + ')');")
        Send("      if (obj != null) {")
        Send("        updateChannelStatus(obj.ChannelID, obj.Message);")
        Send("        updateChannelDeployMsg(obj.ChannelID, obj.DeployMsg, obj.DeployMsgColor);")
        Send("        var elemBtn = document.getElementById('btnassetCh' + obj.ChannelID + 'Mt' + obj.MediaTypeID + 'L' + obj.LanguageID);")
        Send("        var elem = document.getElementById('assetCh' + obj.ChannelID + 'Mt' + obj.MediaTypeID + 'L' + obj.LanguageID);")
        Send("        if (elem != null) {")
        Send("          switch(elem.nodeName.toUpperCase()) {")
        Send("            case 'IMG':")
        Send("              elem.style.display = 'none';")
        Send("              if (elemBtn != null) elemBtn.style.display = '';")
        Send("              break;")
        Send("            default:")
        Send("              elem.value='';")
        Send("              elem.defaultValue = '';")
        Send("              break;")
        Send("          }")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateRemovedPosAsset(jsonResponse) {")
        Send("    if (jsonResponse != null) {")
        Send("      // convert into an object ")
        Send("      var obj = eval ('(' + jsonResponse + ')');")
        Send("      if (obj != null) {")
        Send("        updateChannelStatus(obj.ChannelID, obj.Message);")
        Send("        var elemStr ='assetCh' + obj.ChannelID + 'Mt' + obj.MediaTypeID;")
        Send("        if (obj.OnScreenAdID) elemStr += 'Ad' + obj.OnScreenAdID;")
        Send("        var elem = document.getElementById(elemStr);")
        Send("        var elemDelBtn = document.getElementById(getButtonName(obj.MediaTypeID, obj.ChannelID, obj.OnScreenAdID, 0));")
        Send("        var elemAddBtn = document.getElementById(getButtonName(obj.MediaTypeID, obj.ChannelID, obj.OnScreenAdID, 1));")
        Send("        if (elem != null) elem.innerHTML = '';")
        Send("        if (elemDelBtn != null) { ")
        Send("          elemDelBtn.disabled = true;")
        Send("        }")
        Send("        if (elemAddBtn != null) {")
        Send("          elemAddBtn.disabled = false;")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function getButtonName(mediaTypeID, channelId, adId, buttonType) {")
        Send("    var btnName = 'btn' + ((buttonType==1) ? 'Add' : 'Del');")
        Send("    switch(mediaTypeID) {")
        Send("      case " & MediaTypes.POS_Cashier_Message & ":")
        Send("        btnName += 'CmsgCh' + channelId;")
        Send("        break;")
        Send("      case " & MediaTypes.POS_Receipt_Message & ":")
        Send("        btnName += 'PmsgCh' + channelId;")
        Send("        break;")
        Send("      case " & MediaTypes.POS_Image_Url & ":")
        Send("        btnName += 'ImgUrlCh' + channelId;")
        Send("        break;")
        Send("      case " & MediaTypes.POS_Graphic & ":")
        Send("        btnName += 'GraphicCh' + channelId;")
        Send("        break;")
        Send("      case " & MediaTypes.POS_Accumulation_Message & ":")
        Send("        btnName += 'AmsgCh' + channelId;")
        Send("        break;")
        Send("      default:")
        Send("        break;")
        Send("    }")
        Send("    return btnName;")
        Send("  }")
        Send("")
        Send("  function showGraphicSelector(channelId, mediaTypeId, languageId) {")
        Send("    var elemIfrm = document.getElementById('ifrmGraphic');")
        Send("")
        Send("    if (elemIfrm != null) {")
        Send("      elemIfrm.src = 'offer-channels-graphic.aspx?OfferID=" & Offer.OfferID & "&ChannelID=' + channelId + '&MediaTypeID=' + mediaTypeId + '&LanguageID=' + languageId;")
        Send("      elemIfrm.style.display = '';")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function hideGraphicSelector() {")
        Send("    var elemIfrm = document.getElementById('ifrmGraphic');")
        Send("")
        Send("    if (elemIfrm != null) {")
        Send("      elemIfrm.style.display = 'none';")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function reloadGraphic(id, channelId) {")
        Send("    var obj = document.getElementById(id);")
        Send("    var objBtn = document.getElementById('btn' + id);")
        Send("")
        Send("    if (obj != null && objBtn != null) { ")
        Send("      var src = obj.src;")
        Send("      var pos = src.indexOf('&t=');")
        Send("      if (pos >= 0) {")
        Send("        src = src.substr(0, pos);")
        Send("      }")
        Send("      var date = new Date();")
        Send("      obj.src = src + '&t=' + date.getTime();")
        Send("      obj.style.display = '';")
        Send("      objBtn.style.display = 'none';")
        Send("")
        Send("      updateChannelStatus(channelId, '" & Copient.PhraseLib.Lookup("offer-assets.graphicSaved", LanguageID) & "');")
        Send("    }")
        Send("    return false;")
        Send("  }")
        Send("")
        Send("  function deployChannel(channelId) {")
        'Send("    if (isDeployableChannel(channelId)) { ")
        Send("      xmlhttpPostDataCallback('offer-channels.aspx','mode=DeployChannel&OfferID=" & Offer.OfferID & "&ChannelID=' + channelId, updateDeployChannel);")
        'Send("    }")
        Send("  }")
        Send("")
        Send("  function isDeployableChannel(channelId) {")
        Send("    var deployable = false;")
        Send("    var elemSave = document.getElementById('btnSaveCh' + channelId);")
        Send("")
        Send("    deployable = (elemSave != null && elemSave.src.indexOf('save-off.png') > 0);")
        Send("")
        Send("    if (!deployable) { ")
        Send("      updateChannelStatus(channelId, '" & Copient.PhraseLib.Lookup("offer-channels.noDeploywithChanges", LanguageID) & "');")
        Send("      alert('" & Copient.PhraseLib.Lookup("offer-channels.noDeploywithChanges", LanguageID) & "');")
        Send("    }")
        Send("    return deployable;")
        Send("  }")
        Send("")
        Send("  function updateDeployChannel(jsonResponse) {")
        Send("    if (jsonResponse != null) {")
        Send("      // convert into an object ")
        Send("      var obj = eval ('(' + jsonResponse + ')');")
        Send("      if (obj != null) {")
        Send("        updateChannelDeployMsg(obj.ChannelID, obj.DeployMsg, obj.DeployMsgColor);")
        Send("        updateChannelStatus(obj.ChannelID, obj.Message);")
        Send("        if (obj.Status==0) {")
        Send("          alert(obj.Message);")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")

        Send("</" & "script>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_FormBegin(ByVal c As Channel)
        Send("<form action=""offer-channels.aspx"" id=""frmChannel" & c.ID & """ name=""frmChannel" & c.ID & """ method=""post"" enctype=""multipart/form-data"">")
        Send("  <input type=""hidden"" name=""OfferID"" value=""" & Offer.OfferID & """ />")
        Send("  <input type=""hidden"" name=""ChannelID"" value=""" & c.ID & """ />")
        Send("  <input type=""hidden"" name=""mode"" id=""modeChannel" & c.ID & """ value=""SaveChannel"" />")
        Send("  <input type=""hidden"" name=""IsTemplate"" value=""" & IIf(Offer.IsTemplate, "IsTemplate", "Not") & """ />")
        Send_Channel_MediaTypeIDs(c)
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Channel_MediaTypeIDs(ByVal c As Channel)
        Dim MediaTypeIDs As New List(Of Integer)

        If c.Assets IsNot Nothing Then
            For Each rec As ChannelOfferAssestRec In c.Assets
                If Not MediaTypeIDs.Contains(rec.MediaTypeID) Then
                    MediaTypeIDs.Add(rec.MediaTypeID)
                    Send("<input type=""hidden"" name=""assetCh" & c.ID & "MediaTypeID" & """ value=""" & rec.MediaTypeID & """ />")
                End If
            Next
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_FormEnd()
        Send("</form>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Intro()
        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)

        Send("  <div id=""intro"">")
        If (Offer.IsTemplate) Then
            Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & Offer.OfferID & ": " & Common.TruncateString(Offer.Name, 50) & "</h1>")
        Else
            Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & Offer.OfferID & ": " & Common.TruncateString(Offer.Name, 50) & "</h1>")
        End If
        Send("    <div id=""controls"">")
        If Common.Fetch_SystemOption(75) Then
            If (Offer.OfferID > 0 And Logix.UserRoles.AccessNotes AndAlso (Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                Send_NotesButton(3, Offer.OfferID, AdminUserID)
            End If
        End If
        Send("    </div>")
        Send("  </div>")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Main()
        Send("  <div id=""main"">")
        Sent_Status_Bar()
        Send_Legacy_Selectors()
        Send("<div id=""column"">")
        Send_Channels()
        Send("</div>")

        If Get_Legacy_Status(Offer.OfferID, Offer.EngineID).UseLegacy Then
            Send("<script type=""text/javascript"">")
            Send("  var elemColumn= document.getElementById('column');")
            Send("  if (elemColumn!=null) { elemColumn.style.visibility=""hidden""; }")
            Send("</scr" & "ipt>")
            Send("  </div>")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Sent_Status_Bar()

        Dim dt As DataTable

        If (Not Offer.IsTemplate) Then
            Select Case Offer.EngineID
                Case 2, 3, 6, 9
                    Common.QueryStr = "select incentiveId from CPE_Incentives with (NoLock) where CreatedDate = LastUpdate and IncentiveID=" & Offer.OfferID
                    dt = Common.LRT_Select
                    If (dt.Rows.Count = 0) Then
                        Send_Status(Offer.OfferID, Offer.EngineID)
                    End If
                Case 0, 1
                    Common.QueryStr = "select OfferID from Offers with (NoLock) where CreatedDate = LastUpdate and OfferID=" & Offer.OfferID
                    dt = Common.LRT_Select
                    If (dt.Rows.Count = 0) Then
                        Send_Status(Offer.OfferID)
                    End If
            End Select
        End If
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Function Get_Legacy_Status(ByRef OfferID As Long, ByVal EngineId As Integer) As LegacyStatusRec

        Dim TempStatus As LegacyStatusRec
        Dim DT As DataTable

        TempStatus.DisabledOnCFW = False
        TempStatus.DisplayOnWebKiosk = False
        TempStatus.UseLegacy = False
        If EngineId = CommonInc.InstalledEngines.CPE OrElse EngineId = CommonInc.InstalledEngines.UE OrElse EngineId = CommonInc.InstalledEngines.CAM Then
            Common.QueryStr = "select isnull(DisabledOnCFW, 0) as DisabledOnCFW, DisplayOnWebKiosk, UseLegacyWebKiosk from cpe_incentives where IncentiveID=" & OfferID & ";"
        ElseIf EngineId = CommonInc.InstalledEngines.CM Then
            Common.QueryStr = "select isnull(DisabledOnCFW, 0) as DisabledOnCFW, DisplayOnWebKiosk, UseLegacyWebKiosk from Offers where OfferId=" & OfferID & ";"
        End If
        DT = Common.LRT_Select
        If DT.Rows.Count > 0 Then
            TempStatus.DisabledOnCFW = DT.Rows(0).Item("DisabledOnCFW")
            TempStatus.DisplayOnWebKiosk = DT.Rows(0).Item("DisplayOnWebKiosk")
            TempStatus.UseLegacy = DT.Rows(0).Item("UseLegacyWebKiosk")
        End If
        DT = Nothing
        If (TempStatus.DisabledOnCFW OrElse TempStatus.DisplayOnWebKiosk) AndAlso Not (TempStatus.UseLegacy) Then
            'The UseLegacyWebKiosk bit is out of sync with DisabledOnCFW and DisplayOnWebKiosk
            If EngineId = CommonInc.InstalledEngines.CPE OrElse EngineId = CommonInc.InstalledEngines.UE OrElse EngineId = CommonInc.InstalledEngines.CAM Then
                Common.QueryStr = "Update CPE_Incentives set UseLegacyWebKiosk=1 where IncentiveID=@OfferID;"
            ElseIf EngineId = CommonInc.InstalledEngines.CM Then
                Common.QueryStr = "Update Offers set UseLegacyWebKiosk=1 where OfferId=@OfferID;"
            End If
            Common.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            Common.ExecuteNonQuery(Copient.DataBases.LogixRT)
            TempStatus.UseLegacy = True
        End If
        Return TempStatus

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Legacy_Selectors()

        ' this procedure is only partially developed - development stopped for 5.18 release
        ' phrases need to be created for code in this procedure
        ' AJAX call back needs to be added to the page to handle clicking on the 'save' icon
        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)

        Dim LegacyStatus As LegacyStatusRec
        Dim DT As DataTable

        LegacyStatus = Get_Legacy_Status(Offer.OfferID, Offer.EngineID)

        Send("<script type=""text/javascript"">")
        Send("  function EnableLegacySelectors() {")
        Send("    document.getElementById('DisabledOnCFW').disabled = false;")
        Send("    document.getElementById('DisplayOnWebKiosk').disabled = false;")
        Send("  }")
        Send("")
        Send("  function DisableLegacySelectors() {")
        Send("    document.getElementById('DisabledOnCFW').disabled = true;")
        Send("    document.getElementById('DisplayOnWebKiosk').disabled = true;")
        Send("  }")
        Send("  function UncheckLegacySelectors() {")
        Send("    document.getElementById('DisabledOnCFW').checked = false;")
        Send("    document.getElementById('DisplayOnWebKiosk').checked = false;")
        Send("  }")
        Send("  function EnableLegacySave() { ")
        Send("    document.getElementById('btnSaveLegacy').src='../images/save.png';")
        Send("    document.getElementById('btnSaveLegacy').title='" & Copient.PhraseLib.Lookup("term.clicktosave", LanguageID) & "';") 'Click to save changes
        Send("    document.getElementById('btnSaveLegacy').alt='" & Copient.PhraseLib.Lookup("term.clicktosave", LanguageID) & "';") 'Click to save changes
        Send("  }")
        Send("  function DisableLegacySave() { ")
        Send("    //document.getElementById('btnSaveLegacy').src='../images/save-off.png';")
        Send("    document.getElementById('btnSaveLegacy').title='" & Copient.PhraseLib.Lookup("term.nounsavedchanges", LanguageID) & "';")  'There are no unsaved changes
        Send("    document.getElementById('btnSaveLegacy').alt='" & Copient.PhraseLib.Lookup("term.nounsavedchanges", LanguageID) & "';")  'There are no unsaved changes
        Send("  }")
        Send("")
        Send("  function saveLegacy() {")
        Send("    var elemSave = document.getElementById('btnSaveLegacy');")
        Send("    if (elemSave != null && elemSave.src.indexOf('save.png') > 0) {")
        Send("      xmlhttpPostForm('offer-channels.aspx', 'frmLegacy', updateFromLegacySave);")
        If Not HasChannelContent() Then
            Send("      DisplayOfferModified();")
        End If
        Send("    }")
        Send("  }")
        Send("")
        Send("  function updateFromLegacySave(jsonResponse) {")
        Send("    if (jsonResponse != null) {")
        Send("      var obj = eval ('(' + jsonResponse + ')');")
        Send("      if (obj != null) {")
        Send("        var elemLegacySaveBtn = document.getElementById('btnSaveLegacy'); ")
        Send("        if (elemLegacySaveBtn != null) { DisableLegacySave(); } ")
        Send("        var elemStatus = document.getElementById('LegacyMsg');")
        Send("        if (elemStatus != null) { elemStatus.innerHTML = obj.Message; }")
        Send("        var elemColumn= document.getElementById('column');")
        Send("        if (obj.hidechannels==""1"") { ")
        Send("          if (elemColumn!=null) { elemColumn.style.visibility=""hidden""; }")
        Send("        } ")
        Send("        if (obj.hidechannels==""0"") { ")
        Send("          if (elemColumn!=null) { elemColumn.style.visibility=""visible""; }")
        Send("        } ")
        Send("        if (obj.forcechannelrule==""1"") { ")
        Send("          var elemLegacyOff=document.getElementById('legacyoff'); ")
        Send("          if (elemLegacyOff !=null) { elemLegacyOff.checked=true; } ")
        Send("          UncheckLegacySelectors(); ")
        Send("          DisableLegacySelectors(); ")
        Send("        } ")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("")
        Send("</" & "script>")
        Send("<div class=""box"" id=""legacy"" style=""width: 728px;"">")
        Send("<h2>" & Copient.PhraseLib.Lookup("offer-channels.deliveryselection", LanguageID) & "</h2>")  'Delivery Selection
        Send("<form action=""offer-channels.aspx"" id=""frmLegacy"" name=""frmLegacy"" method=""post"" enctype=""multipart/form-data"">")
        Send("<input type=""hidden"" name=""mode"" id=""mode"" value=""SaveLegacy"">")
        Send("<input type=""hidden"" name=""OfferID"" id=""OfferID"" value=""" & Offer.OfferID & """>")
        Send("<table border=0 cellpadding=0 cellspacing=0 width=""100%""><TR><TD valign=top width=""*"">")
        Send("<input type=""radio"" id=""legacyon"" name=""enablechannels"" value=""false"" onclick=""EnableLegacySelectors(); EnableLegacySave();"" " & IIf(LegacyStatus.UseLegacy = True, "checked", "") & "> " & Copient.PhraseLib.Lookup("offer-channels.uselegacydelivery", LanguageID)) 'Use Legacy web/kiosk delivery
        Send("<BR>&nbsp;&nbsp;&nbsp;&nbsp;<input type=""checkbox"" id=""DisabledOnCFW"" name=""DisabledOnCFW"" value=""true"" " & IIf(LegacyStatus.DisabledOnCFW = True, "checked", "") & " onclick=""EnableLegacySave();"">")
        Send("<label for=""DisabledOnCFW"">" & Copient.PhraseLib.Lookup("offer-gen.showonwebsite", LanguageID) & "</label>")
        Send("<BR>&nbsp;&nbsp;&nbsp;&nbsp;<input type=""checkbox"" id=""DisplayOnWebKiosk"" name=""DisplayOnWebKiosk"" value=""true"" " & IIf(LegacyStatus.DisplayOnWebKiosk = True, "checked", "") & " onclick=""EnableLegacySave();"">")
        Send("<label for=""DisplayOnWebKiosk"">" & Copient.PhraseLib.Lookup("offer-gen.displayonwebkiosk", LanguageID) & "</label>")
        Send("<BR>&nbsp;")
        Send("<BR><input type=""radio"" id=""legacyoff"" name=""enablechannels"" value=""true"" onclick=""UncheckLegacySelectors(); DisableLegacySelectors(); EnableLegacySave();"" " & IIf(LegacyStatus.UseLegacy = False, "checked", "") & "> " & Copient.PhraseLib.Lookup("offer-channels.usechannels", LanguageID)) ' Use Channels
        Send("</TD><TD valign=top width=""300"" align=right>")
        Send(" <span id=""LegacyMsg"" style=""font: 10px arial,sans-serif; color: #606060;""></span>")
        Send("</TD><TD valign=top>")
        If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
            Send("<a href=""#"" onclick=""javascript:saveLegacy();""")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send(">")
            'Send("<img src=""../images/save-off.png"" name=""btnSaveLegacy"" id=""btnSaveLegacy"" alt=""" & Copient.PhraseLib.Lookup("term.nounsavedchanges", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.nounsavedchanges", LanguageID) & """ />") 'There are no unsaved changes
            Send("<img src=""../images/save.png"" name=""btnSaveLegacy"" id=""btnSaveLegacy"" alt=""" & Copient.PhraseLib.Lookup("term.nounsavedchanges", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.nounsavedchanges", LanguageID) & """ />") 'There are no unsaved changes
            Send("</a>")
        End If
        Send("</TD></TR></table>")
        Send("</form>")
        If LegacyStatus.UseLegacy = False Then
            Send("<script type=""text/javascript"">")
            Send("DisableLegacySelectors();")
            Send("")
            Send("</" & "script>")
        End If
        Send("</div>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Channels()
        Send("      <div class=""box"" id=""channels"">")
        Send("        <h2>")
        Send("          <span>" & Copient.PhraseLib.Lookup("term.channels", LanguageID) & "</span>")
        Send("        </h2>")
        If Offer.IsTemplate AndAlso Logix.UserRoles.EditTemplates Then
            Send("<form action='#' id='mainform' name='mainform' method='post'>")
            Send("<input type=""hidden"" id=""OfferID"" name=""OfferID"" value=" & Offer.OfferID & " />")
            Send("<input type=""hidden"" id=""IsTemplate"" name=""IsTemplate"" value=" & Offer.IsTemplate & " />")
            Send("<input type=""hidden"" id=""IsChannelsLocked"" name=""IsChannelsLocked"" />")
            Send("<input type=""hidden"" name=""mode"" id=""modeChannel"" "" value=""SaveChannelLockedState"" />")
            Send("<span class=""temp"">")
            Send("  <input type=""checkbox"" class=""tempcheck"" id=""Disallow_Channels"" name=""Disallow_Channels"" " & IIf(Offer.TemplatePermissions.DisAllow_Channels, " checked=""checked""", "") & "onclick=""setCheckboxValue(this)"" />")
            Send("  <label for=""Disallow_Channels"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("  &nbsp; <a href=""#"" onclick=""javascript:saveChannelLockState();""")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send(">")
            Send("  <img src='../images/save-off.png' name='SaveChLockState' id='SaveChLockState' alt='Save Lock State for Channels' title='Save Lock State for Channels' />")
            Send("  </a>")
            Send("</span>")
            Send_FormEnd()
        End If
        For Each c As Channel In Channels
            If (AllowToViewChannel(c)) Then
                Send_Channel(c)
                If (Offer.FromTemplate AndAlso Offer.TemplatePermissions.DisAllow_Channels) OrElse AllowToEditChannel(c) = False Then
                    SetReadonlyChannel(c.ID)
                End If
            End If
        Next
        Send("      </div>")
        Send("      <div id=""datepicker"" class=""dpDiv""></div>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Function AllowToViewChannel(ByVal c As Channel) As Boolean
        AllowToViewChannel = False
        Select Case (c.ID)
            Case "1"
                AllowToViewChannel = Logix.UserRoles.AccessChannelPos
            Case "2"
                AllowToViewChannel = Logix.UserRoles.AccessChannelWebSite
            Case "3"
                AllowToViewChannel = Logix.UserRoles.AccessChannelOfferBroker
            Case "4"
                AllowToViewChannel = Logix.UserRoles.AccessChannelKiosk
            Case "5"
                AllowToViewChannel = Logix.UserRoles.AccessChannelFlyer ' add new role for flyer
        End Select
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Function AllowToEditChannel(ByVal c As Channel) As Boolean
        AllowToEditChannel = False
        Select Case (c.ID)
            Case "1"
                AllowToEditChannel = Logix.UserRoles.EditChannelPos
            Case "2"
                AllowToEditChannel = Logix.UserRoles.EditChannelWebsite
            Case "3"
                AllowToEditChannel = Logix.UserRoles.EditChannelOfferBroker
            Case "4"
                AllowToEditChannel = Logix.UserRoles.EditChannelKiosk
            Case "5"
                AllowToEditChannel = Logix.UserRoles.EditChannelFlyer ' 	add new role for flyer
        End Select
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub SetReadonlyChannel(ByVal ChannelID As Integer)
        Send("<script type='text/javascript' language='javascript'>")
        Send("  window.onload = SetReadonlyChannel(" & ChannelID & ");")
        Sendb("</")
        Send("script>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Channel(ByVal c As Channel)
        Send_FormBegin(c)
        Send("        <table id=""tblchannel" & c.ID & """ style=""table-layout:fixed; width: 635px;"">")
        Send_Channel_Summary(c)
        Send_Channel_Detail(c)
        Send("        </table>")
        Send_FormEnd()
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Channel_Summary(ByVal c As Channel)
        Dim ExpanderImage As String
        Dim ChannelActiveText As String

        ExpanderImage = IIf(c.Selected, "minus.png", "plus.png")
        ChannelActiveText = IIf(c.Selected, Copient.PhraseLib.Lookup("offer-channels.Configured", LanguageID), Copient.PhraseLib.Lookup("offer-channels.NotConfigured", LanguageID))

        ' write the summary row
        Send("          <tr>")
        Send("            <td style=""width:12px;""><img src=""/images/" & ExpanderImage & """ id=""chExpand" & c.ID & """ onclick=""handleChannelDetail(" & c.ID & ");"" /></td>")
        Send("            <td style=""width:23px;""><img src=""/images/star-" & IIf(c.Selected, "on", "off") & ".png"" id=""activeImgCh" & c.ID & """  alt=""" & ChannelActiveText & """ title=""" & ChannelActiveText & """ /></td>")
        Send("            <td style=""width:600px;"">" & c.Name & "</td>")
        Send("          </tr>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Channel_Detail(ByVal c As Channel)
        Dim StyleStr As String

        StyleStr = IIf(c.Selected, "", "display: none;")

        ' write the detail row
        Send("          <tr id=""chDetail" & c.ID & """ style=""" & StyleStr & """ >")
        Send("            <td colspan=""2""/>")
        Send("            <td style=""background-color:#f0f0f0;"">")
        Send("              <table style=""width: 100%; border: 2px solid #707070;"">")
        Send_Channel_Toolbar(c)
        Send_SubChannels(c)
        Send_Dates_Row(c)
        Send_Limits_Row(c)
        If c.ID = POS_CHANNEL_ID Then
            Send_PosNotificationProduction(c)
            Send_POS_Assets(c)
        Else
            Send_Assets(c)
        End If
        Send_Channel_StatusBar(c)
        Send("              </table>")
        Send("            </td>")
        Send("          </tr>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_SubChannels(ByVal c As Channel)
        'This procedure sends the sub channel selector (if there are any related SubChannels)
        Dim SubChannelList As List(Of Copient.ChannelCommon.SubChannelOfferRec)
        Dim SubChannelRec As Copient.ChannelCommon.SubChannelOfferRec
        Dim SelectBoxSize As Integer

        SubChannelList = ChannelCommon.GetOfferSubChannels(c.ID, Offer.OfferID)
        If SubChannelList.Count > 0 Then
            SelectBoxSize = SubChannelList.Count
            If SelectBoxSize > 4 Then SelectBoxSize = 4
            Send("<TR>")
            Send("<TD colspan=""4"">")
            Send("<table border=0 cellpadding=0 cellspacing=1>")
            Send("<TR><TD valign=""top"" width=""100"" nowrap>Make content available to<BR>the following sub-channels: </TD><TD valign=""top"">")
            'send the multiselect list box
            Send("<select multiple=""multiple"" id=""Ch" & c.ID & "subchannels"" name=""Ch" & c.ID & "subchannels"" size=""" & SelectBoxSize & """ onchange=""selectChannel(" & c.ID & ");"">")
            For Each SubChannelRec In SubChannelList
                'send the listbox item
                Send("<option value=""" & SubChannelRec.SubChannelID & """" & IIf(SubChannelRec.Associated, " selected=""selected""", "") & ">" & SubChannelRec.SubChannelName & "</option>")
            Next
            'close the multiselect list box
            Send("</select></TD></TR></table>")
            Send("</TD>")
            Send("</TR>")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Channel_Toolbar(ByVal c As Channel)
        Dim SaveText As String = Copient.PhraseLib.Lookup("offer-channels.noUnsavedChanges", LanguageID)
        Dim RemoveText As String = Copient.PhraseLib.Lookup("offer-channels.removeChannel", LanguageID)
        Dim DeployText As String = Copient.PhraseLib.Lookup("offer-channels.deployChannel", LanguageID)

        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)

        Send("                <tr style=""background-color:#808080;"">")
        Send("                  <td colspan=""4"">")
        If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
            Send("                    <a href=""#"" onclick=""javascript:saveChannel(" & c.ID & ");""")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send(">")
            '
            Send("                      <img src=""../images/save.png"" name=""btnSaveCh" & c.ID & """ id=""btnSaveCh" & c.ID & """ alt=""" & SaveText & """ title=""" & SaveText & """ />")
            Send("                    </a>")
            Send("                    &nbsp;&nbsp;")
            Send("                    <a href=""#"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) javascript:removeChannel(" & c.ID & ");""")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send(">")
            Send("                      <img src=""../images/delete.png"" name=""btnRemoveCh" & c.ID & """ id=""btnRemoveCh" & c.ID & """ alt=""" & RemoveText & """ title=""" & RemoveText & """ />")
            Send("                    </a>")
            If c.ID <> POS_CHANNEL_ID And Not (Offer.IsTemplate) Then
                Send("                    &nbsp;&nbsp;")
                Send("                    <a href=""#"" onclick=""javascript:deployChannel(" & c.ID & ");""")
                If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                    Send("style=""visibility: hidden""")
                End If
                Send(">")
                Send("                      <img src=""../images/deploy.png"" name=""btnDeployCh" & c.ID & """ id=""btnDeployCh" & c.ID & """ alt=""" & DeployText & """ title=""" & DeployText & """ />")
                Send("                    </a>")
                Send("<span id=""ChannelDeployMsg" & c.ID.ToString & """></span>")
                If ChannelCommon.HasChannelChanged(c.ID, Offer.OfferID) Then
                    Send("<script type=""text/javascript"">")
                    Send("document.getElementById('ChannelDeployMsg" & c.ID.ToString & "').style.color='#FF3333';")
                    Send("document.getElementById('ChannelDeployMsg" & c.ID.ToString & "').innerHTML="" " & Copient.PhraseLib.Lookup("term.changesnotdeployed", LanguageID) & """;")
                    Send("</scr" & "ipt>")
                ElseIf ChannelCommon.IsChannelWaitingDeployment(c.ID, Offer.OfferID) Then
                    Send("<script type=""text/javascript"">")
                    Send("document.getElementById('ChannelDeployMsg" & c.ID.ToString & "').style.color='#00FF00';")
                    Send("document.getElementById('ChannelDeployMsg" & c.ID.ToString & "').innerHTML="" " & Copient.PhraseLib.Lookup("term.awaitingdeployment", LanguageID) & """;")
                    Send("</scr" & "ipt>")
                End If
            End If
        End If
        Send("                  </td>")
        Send("                </tr>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Channel_StatusBar(ByVal c As Channel)
        Send("                <tr style=""background-color:#e8e8e8;"">")
        Send("                  <td colspan=""4"">")
        Send("                    <div id=""channelMsg" & c.ID & """ style=""font: 10px arial,sans-serif; color: #606060;border-top: 1px solid #d0d0d0;"">" & Copient.PhraseLib.Lookup("term.ready", LanguageID) & "</div>")
        Send("                  </td>")
        Send("                </tr>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Dates_Row(ByVal c As Channel)
        Dim DateInput, DateImg, StartDateStr, EndDateStr As String

        ' AL-1392 - Logix 5.19.0 SRD, section 4.1.1, Start and end date shall default to the date the offer is first created
        If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.UE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
            StartDateStr = CheckChannelDate(c.StartDate, "StartDate")
            EndDateStr = CheckChannelDate(c.EndDate, "EndDate")
        ElseIf Offer.EngineID = CommonInc.InstalledEngines.CM Then
            StartDateStr = CheckChannelDate(c.StartDate, "ProdStartDate")
            EndDateStr = CheckChannelDate(c.EndDate, "ProdEndDate")
        Else
            StartDateStr = CheckChannelDate(c.StartDate, "ProdStartDate")
            EndDateStr = CheckChannelDate(c.EndDate, "ProdEndDate")
        End If
        If c.UsesStartDate OrElse c.UsesEndDate Then
            DateInput = "<input type=""text"" id=""ch{0}{1}"" name=""ch{0}{1}"" style=""width:70px;"" maxlength=""10"" onkeyup=""selectChannel({1});"" onmousedown=""selectChannel({1});"" onchange=""selectChannel({1}); if(isNaN(Date.parse(ConvertToISODate(this.value,'" & Common.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern & "' )))){{ alert('{3}'); this.value='{2}';}}"" value=""{2}"" />"
            DateImg = "<img src=""../images/calendar.png"" class=""calendar"" id=""{0}-picker" & c.ID & """ alt=""Date picker"" title=""Date picker"" onclick=""displayDatePicker('ch{0}" & c.ID & "', event); selectChannel({1});""/>"

            Send("                <tr>")
            Send("                  <td colspan=""4"">")
            If c.UsesStartDate Then
                Send(Copient.PhraseLib.Lookup("term.startdate", LanguageID) & ": ")
                Send("                    " & String.Format(DateInput, "Start", c.ID, StartDateStr, Copient.PhraseLib.Lookup("term.InvalidStartDate", LanguageID)))
                Send("                    " & String.Format(DateImg, "Start", c.ID))
            End If
            If c.UsesEndDate Then
                Send(IIf(c.UsesStartDate, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", ""))
                Send(Copient.PhraseLib.Lookup("term.enddate", LanguageID) & ": ")
                Send("                    " & String.Format(DateInput, "End", c.ID, EndDateStr, Copient.PhraseLib.Lookup("term.InvalidEndDate", LanguageID)))
                Send("                    " & String.Format(DateImg, "End", c.ID))
            End If
            Send("                  </td>")
            Send("                </tr>")
        End If

    End Sub

    '------------------------------------------------------------------------------------------------------------- 

    Private Sub Send_Limits_Row(ByVal c As Channel)
        Dim LimitType As LimitTypes
        Dim TextBoxHandlers As String = ""

        If c.UsesLimits Then
            LimitType = GetSelectedLimitType(c.Limit)
            TextBoxHandlers = "onkeyup=""selectChannel(" & c.ID & ");"" onmousedown=""selectChannel(" & c.ID & ");"" onchange=""selectChannel(" & c.ID & ");"""

            Send("<tr><td colspan=""4""><b><u>" & Copient.PhraseLib.Lookup("term.limits", LanguageID) & "</u></b></td></tr>")
            If Offer.EngineID = CommonInc.InstalledEngines.UE Then
                Send("<tr>")
                Send("  <td colspan=""2"">")
                Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.eligibility", LanguageID) & """>")
                Send("    <tr>")
                Send("      <td colspan=""6"">")
                Send("        <label for=""p1typeCh" & c.ID & """>" & Copient.PhraseLib.Lookup("term.frequency", LanguageID) & ":</label>&nbsp;")
                Send("        <select name=""p1typeCh" & c.ID & """ id=""p1typeCh" & c.ID & """ onchange=""updateP1limit(" & c.ID & ");""" & ">")
                If Not (Offer.HasAnyCustomer) Then
                    Send("            <option value=""6""" & IIf(LimitType = LimitTypes.DAYS_ROLLING, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("UE_DistributionTimeTypes.Days", LanguageID) & "</option>")
                End If
                Send("          <option value=""2""" & IIf(LimitType = LimitTypes.ONCE_PER_TRANSACTION, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID) & "</option>")
                Send("          <option value=""1""" & IIf(LimitType = LimitTypes.NO_LIMIT, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.nolimit", LanguageID) & "</option>")
                If (IsCustomerAssigned()) AndAlso (Not Offer.HasAnyCustomer) Then
                    Send("          <option value=""5""" & IIf(LimitType = LimitTypes.ONCE_PER_OFFER, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID) & "</option>")
                End If
                Send("        </select>")
                Send("      </td>")
                Send("    </tr>")
                Send("    <tr id=""customRowCh" & c.ID & """ style=""" & IIf((LimitType = LimitTypes.DAYS_ROLLING Or LimitType = LimitTypes.PER_TRANSACTION), "", "display: none;") & """>")
                Send("      <td>")
                Send("        <label for=""limit1Ch" & c.ID & """>" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & ":</label>&nbsp;")
                'Send("      </td>")
                'Send("      <td>")
                Send("        <input type=""text"" class=""shorter"" id=""limit1Ch" & c.ID & """ name=""limit1Ch" & c.ID & """ maxlength=""6"" " & TextBoxHandlers & " value=""" & c.Limit.P1DistQtyLimit & """" & " />")
                Send("      </td>")
                Send("      <td >")
                Send("        <label id=""lbllimit1periodCh" & c.ID & """ for=""limit1periodCh" & c.ID & """  style=""" & IIf(LimitType <> LimitTypes.PER_TRANSACTION, "", "display: none;") & """  >" & Copient.PhraseLib.Lookup("term.period", LanguageID) & ":</label>&nbsp;")
                'Send("      </td>")
                'Send("      <td colspan=2>")
                Sendb("      <input type=""text"" class=""shorter"" id=""limit1periodCh" & c.ID & """  style=""" & IIf(LimitType <> LimitTypes.PER_TRANSACTION, "", "display: none;") & """ name=""limit1periodCh" & c.ID & """ maxlength=""6"" " & TextBoxHandlers & " value=""" & c.Limit.P1DistPeriod)
                Send("""" & IIf(Offer.HasAnyCustomer, " disabled=""disabled""", "") & " />")
                Send("        <label style='display:none' for=""P1DistTimeTypeCh" & c.ID & """>" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label>")
                Send("        <input type=""hidden"" id=""P1DistTimeTypeCh" & c.ID & """ name=""P1DistTimeTypeCh" & c.ID & """ value=""" & c.Limit.P1DistTimeType & """ />")
                Send("      </td>")
                Send("    </tr>")
                Send("  </table>")
                Send("  </td>")
                'Send("</tr>")


            Else
                Send("<tr>")
                Send("  <td colspan=""4"">")
                Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.eligibility", LanguageID) & """>")
                Send("    <tr>")
                Send("      <td colspan=""6"">")
                Send("        <label for=""p1typeCh" & c.ID & """>" & Copient.PhraseLib.Lookup("term.frequency", LanguageID) & ":</label>&nbsp;")
                Send("        <select name=""p1typeCh" & c.ID & """ id=""p1typeCh" & c.ID & """ onchange=""updateP1limit(" & c.ID & ");""" & ">")
                Send("          <option value=""1""" & IIf(LimitType = LimitTypes.NO_LIMIT, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.nolimit", LanguageID) & "</option>")
                Send("          <option value=""2""" & IIf(LimitType = LimitTypes.ONCE_PER_TRANSACTION, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID) & "</option>")
                If Not (Offer.HasAnyCustomer) Then
                    Send("          <option value=""3""" & IIf(LimitType = LimitTypes.ONCE_PER_DAY, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.onceperday", LanguageID) & "</option>")
                    Send("          <option value=""4""" & IIf(LimitType = LimitTypes.ONCE_PER_WEEK, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.onceperweek", LanguageID) & "</option>")
                    Send("          <option value=""5""" & IIf(LimitType = LimitTypes.ONCE_PER_OFFER, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID) & "</option>")
                End If
                Send("            <option value=""0""" & IIf(LimitType = LimitTypes.CUSTOM, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.custom", LanguageID) & "</option>")
                Send("        </select>")
                Send("      </td>")
                Send("    </tr>")
                Send("    <tr id=""customRowCh" & c.ID & """ style=""" & IIf(LimitType = LimitTypes.CUSTOM, "", "display: none;") & """>")
                Send("      <td>")
                Send("        <label for=""limit1Ch" & c.ID & """>" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & ":</label>")
                Send("      </td>")
                Send("      <td>")
                Send("        <input type=""text"" class=""shorter"" id=""limit1Ch" & c.ID & """ name=""limit1Ch" & c.ID & """ maxlength=""6"" " & TextBoxHandlers & " value=""" & c.Limit.P1DistQtyLimit & """" & " />")
                Send("      </td>")
                Send("      <td>")
                Send("        <label id=""lbllimit1periodCh" & c.ID & """ for=""limit1periodCh" & c.ID & """>" & Copient.PhraseLib.Lookup("term.period", LanguageID) & ":</label>")
                Send("      </td>")
                Send("      <td>")
                Sendb("      <input type=""text"" class=""shorter"" id=""limit1periodCh" & c.ID & """ name=""limit1periodCh" & c.ID & """ maxlength=""6"" " & TextBoxHandlers & " value=""" & c.Limit.P1DistPeriod)
                Send("""" & IIf(Offer.HasAnyCustomer, " disabled=""disabled""", "") & " />")
                Send("      </td>")
                Send("      <td>")
                Send("        <label for=""P1DistTimeTypeCh" & c.ID & """>" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label>")
                Send("      </td>")
                Send("      <td>")
                Send("        <select id=""P1DistTimeTypeCh" & c.ID & """ name=""P1DistTimeTypeCh" & c.ID & """ onchange=""selectChannel(" & c.ID & ")"" " & ">")

                For Each t As TimeType In TimeTypes
                    Send("<option value=""" & t.ID & """" & IIf(c.Limit.P1DistTimeType = t.ID, " selected=""selected""", "") & ">" & t.Name & "</option>")
                Next

                Send("        </select>")
                Send("        <input type=""hidden"" id=""BeginP1TimeTypeIDCh" & c.ID & """ name=""BeginP1TimeTypeIDCh" & c.ID & """ value=""" & c.Limit.P1DistTimeType & """ />")
                Send("      </td>")
                Send("    </tr>")
                Send("  </table>")
                Send("  </td>")
                'Send("</tr>")
            End If
        End If
    End Sub

    '------------------------------------------------------------------------------------------------------------- 

    Private Sub Send_PosNotificationProduction(ByVal channel As Channel)
        Dim m_Offer As IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of IOffer)()
        Dim CheckQCust As String = IIf(channel.PosNotificationCheck = 0, " checked=""checked""", "")
        Dim CheckProdPurchase As String = IIf(channel.PosNotificationCheck = 1, " checked=""checked""", "")
        Dim listProdCondition As AMSResult(Of List(Of RegularProductCondition)) = m_Offer.GetRegularProductConditionsByOfferId(Offer.OfferID)
        Dim DisableProdPurchase As String = IIf(listProdCondition.Result.Count = 0, " disabled ", "")

        Send("  <td colspan=""2"">")
        Send("  <input type=""radio"" name=""PosProductionConditionCheck"" value=""0"" " & CheckQCust & "> " & Copient.PhraseLib.Lookup("term.PosNotificationCheckBox1", LanguageID) & "<br>" &
         "  <input type=""radio"" name=""PosProductionConditionCheck"" value=""1"" " & CheckProdPurchase & "" & DisableProdPurchase & "> " & Copient.PhraseLib.Lookup("term.PosNotificationCheckBox2", LanguageID) & "")
        Send("  </td>")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_PrintedMessage(ByVal c As Channel)

        Send("<tr>")
        Send("  <td style=""width: 33%;"">")
        Send("    <input type=""button"" class=""ex"" id=""btnDelPmsgCh" & c.ID & """ value=""X"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) removePrintedMsg(" & c.ID & "," & c.PrintedMessage.DeliverableID & "," & c.PrintedMessage.MessageID & ");"" " & IIf(c.PrintedMessage.DeliverableID = 0, "disabled=""disabled"" ", "") & "")
        If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
            Send("style=""visibility: hidden""")
        End If
        Send("/>")
        Send("    <input type=""button"" class=""adjust"" id=""btnAddPmsgCh" & c.ID & """ value=""+""" & IIf(c.PrintedMessage.DeliverableID > 0, " disabled=""disabled""", "") & " onclick=""openPMsgPopup(0, 0);""")
        If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
            Send("style=""visibility: hidden""")
        End If
        Send("/>")
        Send(Copient.PhraseLib.Lookup("term.printedmessage", LanguageID) & ": ")
        Send("  </td>")
        Send("  <td colspan=""3"">")
        Send("    <span id=""assetCh" & c.ID & "Mt" & MediaTypes.POS_Receipt_Message & """>")
        Send("      <a href=""javascript:openPMsgPopup(" & c.PrintedMessage.DeliverableID & "," & c.PrintedMessage.MessageID & ");"" ")
        Send("         alt=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """>")
        Send("         " & HttpUtility.HtmlEncode(Common.TruncateString(c.PrintedMessage.BodyText, 40)))
        Send("      </a>")
        Send("    </span>")
        Send("  </td>")
        Send("</tr>")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_ImageURL(ByVal c As Channel)

        Send("<tr>")
        Send("  <td style=""width: 33%;"">")
        Send("    <input type=""button"" class=""ex"" id=""btnDelImgUrlCh" & c.ID & """ value=""X"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) removeImageUrl(" & c.ID & "," & c.PosImgURL.DeliverableID & "," & c.PosImgURL.PTPKID & ");"" " & IIf(c.PosImgURL.DeliverableID = 0, "disabled=""disabled"" ", "") & "")
        If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
            Send("style=""visibility: hidden""")
        End If
        Send("/>")
        Send("    <input type=""button"" class=""adjust"" id=""btnAddImgUrlCh" & c.ID & """ value=""+""" & IIf(c.PosImgURL.DeliverableID > 0, " disabled=""disabled""", "") & " onclick=""openImgUrlPopup(0, 0);""")
        If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
            Send("style=""visibility: hidden""")
        End If
        Send("/>")
        Send(Copient.PhraseLib.Lookup("mediatype.imgurl", LanguageID) & ": ")
        Send("  </td>")
        Send("  <td colspan=""3"">")
        Send("    <span id=""assetCh" & c.ID & "Mt" & MediaTypes.POS_Image_Url & """>")
        Send("      <a href=""javascript:openImgUrlPopup(" & c.PosImgURL.DeliverableID & "," & c.PosImgURL.PTPKID & ");"" ")
        Send("         alt=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """>")
        Send("         " & HttpUtility.HtmlEncode(Common.TruncateString(c.PosImgURL.Data, 40)))
        Send("      </a>")
        Send("    </span>")
        Send("  </td>")
        Send("</tr>")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_CashierMessage(ByVal c As Channel)
        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)
        Send("<tr>")
        Send("  <td style=""width: 33%;"">")
        If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
            Send("    <input type=""button"" class=""ex"" id=""btnDelCmsgCh" & c.ID & """ value=""X"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) removeCashierMsg(" & c.ID & "," & c.CashierMessage.DeliverableID & "," & c.CashierMessage.MessageID & ");"" " & IIf(c.CashierMessage.DeliverableID = 0, "disabled=""disabled"" ", "") & "")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send("/>")
            Send("    <input type=""button"" class=""adjust"" id=""btnAddCmsgCh" & c.ID & """ value=""+""" & IIf(c.CashierMessage.DeliverableID > 0, " disabled=""disabled""", "") & " onclick=""openCMsgPopup(0, 0);""")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send("/>")
            Send(Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID) & ": ")
        End If
        Send("  </td>")
        Send("  <td colspan=""3"">")
        Send("    <span id=""assetCh" & c.ID & "Mt" & MediaTypes.POS_Cashier_Message & """>")
        Send("      <a href=""javascript:openCMsgPopup(" & c.CashierMessage.DeliverableID & "," & c.CashierMessage.MessageID & ");"" ")
        Send("         alt=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """>")
        Send("         " & Common.TruncateString(c.CashierMessage.Line1, 40) & "<br />" & Common.TruncateString(c.CashierMessage.Line2, 40))
        Send("      </a>")
        Send("    </span>")
        Send("  </td>")
        Send("</tr>")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Graphic(ByVal c As Channel)
        Dim GraphicPath As String = ""

        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)

        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)

        GraphicPath = Common.Fetch_SystemOption(47)
        If Not (Right(GraphicPath, 1) = "\") Then
            GraphicPath = GraphicPath & "\"
        End If

        For Each gr As GraphicDeliverable In c.Graphic
            Send("<tr>")
            Send("  <td style=""width: 33%;"">")
            If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                Send("    <input type=""button"" class=""ex"" id=""btnDelGraphicCh" & c.ID & """ value=""X"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) removeGraphic(" & c.ID & "," & gr.DeliverableID & "," & gr.OnScreenAdID & ");"" " & IIf(gr.DeliverableID = 0, "disabled=""disabled"" ", "") & "")
                If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                    Send("style=""visibility: hidden""")
                End If
                Send("/>")
                Send("    <input type=""button"" class=""adjust"" id=""btnAddGraphicCh" & c.ID & """ value=""+"" onclick=""openAddGraphicsPopup();"" " & IIf(gr.DeliverableID <> 0, "disabled=""disabled"" ", "") & "")
                If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                    Send("style=""visibility: hidden""")
                End If
                Send("/>")
                Send(Copient.PhraseLib.Lookup("term.graphic", LanguageID) & ":")
            End If
            Send("  </td>")
            Send("  <td colspan=""3"">")
            Send("    <span id=""assetCh" & c.ID & "Mt" & MediaTypes.POS_Graphic & "Ad" & gr.OnScreenAdID & """>")
            Send("      <a href=""javascript:openPreviewGraphicsPopup(" & gr.OnScreenAdID & ", " & gr.CellSelectID & ", " & gr.ImageType & ");""")
            Send("         alt=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """>")
            If GraphicPath <> "" AndAlso gr.OnScreenAdID > 0 Then
                Send("        <img src=""graphic-display-img.aspx?path=" & GraphicPath & gr.OnScreenAdID & "img_tn.jpg"" /><br />")
            End If
            Send("         " & Common.TruncateString(gr.Name, 40))
            Send("      </a>")
            Send("      <br />")
            Send("    </span>")
            Send("  </td>")
            Send("</tr>")
        Next

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_AccumulationMessage(ByVal c As Channel)
        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)
        Send("<tr>")
        Send("  <td style=""width: 33%;"">")
        If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
            Send("    <input type=""button"" class=""ex"" id=""btnDelAmsgCh" & c.ID & """ value=""X"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) removeAccumMsg(" & c.ID & "," & c.AccumMessage.DeliverableID & "," & c.AccumMessage.MessageID & ");"" " & IIf(c.AccumMessage.DeliverableID = 0, "disabled=""disabled"" ", "") & "")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send("/>")
            Send("    <input type=""button"" class=""adjust"" id=""btnAddAmsgCh" & c.ID & """ value=""+""" & IIf(c.AccumMessage.DeliverableID > 0, " disabled=""disabled""", "") & " onclick=""openAccumMsgPop(0, 0);""")
            If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                Send("style=""visibility: hidden""")
            End If
            Send("/>")
            Send(Copient.PhraseLib.Lookup("term.accumulationmessage", LanguageID) & ": ")
        End If
        Send("  </td>")
        Send("  <td colspan=""3"">")
        Send("    <span id=""assetCh" & c.ID & "Mt" & MediaTypes.POS_Accumulation_Message & """>")
        Send("      <a href=""javascript:openAccumMsgPop(" & c.AccumMessage.DeliverableID & "," & c.AccumMessage.MessageID & ");"" ")
        Send("         alt=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """>")
        Send("         " & Common.TruncateString(c.AccumMessage.BodyText, 40))
        Send("      </a>")
        Send("    </span>")
        Send("  </td>")
        Send("</tr>")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_POS_Assets(ByVal c As Channel)
        Dim WroteCMsg, WrotePMsg, WroteImgUrl, WroteGraphic, WroteAccumMsg As Boolean

        Send("<tr><td colspan=""4""><b><u>" & Copient.PhraseLib.Lookup("term.notifications", LanguageID) & "</u></b></td></tr>")

        For Each rec As ChannelOfferAssestRec In c.Assets
            ' send this as the appropriate UI send function for this asset
            Select Case rec.UIFormField
                Case DisplayTypes.POS__CASHIER_MESSAGE
                    If Not WroteCMsg Then Send_CashierMessage(c)
                    WroteCMsg = True

                Case DisplayTypes.POS_GRAPHIC
                    If Not WroteGraphic Then Send_Graphic(c)
                    WroteGraphic = True

                Case DisplayTypes.POS_PRINTED_MESSAGE
                    If Not WrotePMsg Then Send_PrintedMessage(c)
                    WrotePMsg = True

                Case DisplayTypes.POS_IMAGE_URL
                    If Not WroteImgUrl Then Send_ImageURL(c)
                    WroteImgUrl = True

                Case DisplayTypes.POS_ACCUM_MESSAGE
                    If ((Not WroteAccumMsg) AndAlso IsAccumulationEnabled()) Then
                        Send_AccumulationMessage(c)
                        WroteAccumMsg = True
                    End If
            End Select
        Next

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Assets(ByVal c As Channel)
        Dim LangAssets As List(Of ChannelOfferAssestRec)
        Dim L As Language
        Dim dtAdFieldDetails As DataTable
        Dim rowAdFieldDetails As DataRow
        Dim MediaData As String = ""
        Dim DefaultMediaData As String = ""
        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)

        If c.Assets IsNot Nothing AndAlso c.Assets.Count > 0 Then

            For Each L In AvailableLanguages
                Send("<tr><td colspan=""4""><b><u>" & L.Name & "</b></u></td></tr>")
                LangAssets = c.Assets.FindAll(Function(asset) FindAssetByLanguageID(asset, L.ID))

                For Each rec As ChannelOfferAssestRec In LangAssets
                    If (Not rec.UIFormField = DisplayTypes.INPUT_FILE) AndAlso (rec.MediaData IsNot Nothing) Then
                        MediaData = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(rec.MediaData))
                    End If
                    ' send this as the appropriate UI send function for this asset
                    Select Case rec.UIFormField
                        Case DisplayTypes.INPUT_FILE, DisplayTypes.INPUT_TEXT
                            Send("<tr>")
                            Send("  <td style=""width: 30%;"">")
                            If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                Send("    <input type=""button"" class=""ex"" value=""X"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) removeOfferAsset(" & c.ID & "," & rec.MediaTypeID & "," & rec.LanguageID & ");"" ")
                                If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                                    Send("style=""visibility: hidden""")
                                End If
                                Send("/>")
                                Send(rec.Name & " : ")
                            End If
                            Send("</td>")
                            Send("  <td colspan=""3"">" & GetAssetFormField(rec, c.ID) & "</td>")
                            Send("</tr>")
                        Case DisplayTypes.SELECT_MULTIPLE

                            Send("<tr>")
                            Send("  <td style=""width: 30%;"">")
                            If ((Not bEnableRestrictedAccessToUEOfferBuilder OrElse Not isTranslatedOffer) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                                Send("    <input type=""button"" class=""ex"" value=""X"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.delete", LanguageID) & "')) removeOfferAsset(" & c.ID & "," & rec.MediaTypeID & "," & rec.LanguageID & ");""")
                                If (Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                                    Send("style=""visibility: hidden""")
                                End If
                                Send("/>")
                                Send(rec.Name & " : ")
                            End If
                            Send("</td>")
                            Send("  <td colspan=""3"">" & GetAssetFormField(rec, c.ID) & "")
                            Send("<option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
                            Common.QueryStr = "Select MS.MediaData, MS.DefaultValue from MediaAssetValues MS with (NoLock) inner join " &
                                              "ChannelMedia as CM with (NoLock) on CM.ChannelMediaID=MS.ChannelMediaID " &
                                              "inner join ChannelMediaTypes as CMT with (NoLock) on CMT.MediaTypeID = cm.MediaTypeID " &
                                              "where ms.ChannelMediaID = " & rec.MediaTypeID & ""
                            dtAdFieldDetails = Common.LRT_Select()
                            For Each rowAdFieldDetails In dtAdFieldDetails.Rows
                                If Common.NZ(rowAdFieldDetails.Item("DefaultValue"), 0) = True AndAlso (MediaData = "0" OrElse MediaData = "") Then
                                    MediaData = rowAdFieldDetails.Item("MediaData")
                                End If
                                If MediaData = rowAdFieldDetails.Item("MediaData") Then
                                    Send("<option value=""" & rowAdFieldDetails.Item("MediaData") & """ selected=""selected"">" & rowAdFieldDetails.Item("MediaData") & " </option>")
                                Else
                                    Send("<option value=""" & rowAdFieldDetails.Item("MediaData") & """>" & rowAdFieldDetails.Item("MediaData") & " </option>")
                                End If
                            Next
                            Send("</Select>")
                            Send("</td>")
                            Send("</tr>")
                    End Select
                Next
            Next

        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Send_Graphic_Selector()
        Send("<iframe id=""ifrmGraphic"" name=""ifrmGraphic"" src=""blank.html"" style=""position:absolute; top: 30%; height: 40%; left: 25%; width: 50%; z-index: 999; display: none; background-color: #e0e0e0;"">")
        Send("</iframe>")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Function GetSelectedLimitType(ByVal Limit As EligibilityLimit) As LimitTypes

        Dim LimitType As LimitTypes = LimitTypes.CUSTOM
        If Offer.EngineID = CommonInc.InstalledEngines.UE Then
            LimitType = LimitTypes.DAYS_ROLLING
            If Limit.P1DistTimeType = 2 AndAlso Limit.P1DistPeriod = 1 AndAlso Limit.P1DistQtyLimit = 1 Then
                LimitType = LimitTypes.ONCE_PER_TRANSACTION
            End If
            If Limit.P1DistQtyLimit = 0 AndAlso Limit.P1DistPeriod = 0 AndAlso Limit.P1DistTimeType = 2 Then
                LimitType = LimitTypes.NO_LIMIT
            End If
            If Limit.P1DistQtyLimit = 1 AndAlso Limit.P1DistPeriod = 3650 AndAlso Limit.P1DistTimeType = 1 Then
                LimitType = LimitTypes.ONCE_PER_OFFER
            End If
            Return LimitType
        End If
        If Limit.P1DistQtyLimit = 0 AndAlso Limit.P1DistPeriod = 0 AndAlso (Limit.P1DistTimeType = 2 OrElse Limit.P1DistTimeType = -1) Then
            LimitType = LimitTypes.NO_LIMIT
        ElseIf Limit.P1DistQtyLimit = 1 Then
            If Limit.P1DistPeriod = 1 Then
                If Limit.P1DistTimeType = 2 Then
                    LimitType = LimitTypes.ONCE_PER_TRANSACTION
                ElseIf Limit.P1DistTimeType = 1 Then
                    LimitType = LimitTypes.ONCE_PER_DAY
                End If
            ElseIf Limit.P1DistPeriod = 7 AndAlso Limit.P1DistTimeType = 1 Then
                LimitType = LimitTypes.ONCE_PER_WEEK
            ElseIf Limit.P1DistPeriod = 3650 AndAlso Limit.P1DistTimeType = 1 Then
                LimitType = LimitTypes.ONCE_PER_OFFER
            End If
        End If

        Return LimitType
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Function GetAssetFormField(ByVal rec As ChannelOfferAssestRec, ByVal ChannelID As Integer) As String
        Dim FieldStr As String = ""
        Dim MediaData As String = ""
        Dim AssetId As String = ""
        Dim isTranslatedOffer As Boolean = MyCommon.IsTranslatedUEOffer(Offer.OfferID, MyCommon)
        Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)
        Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = IIf(MyCommon.Fetch_SystemOption(260) = "1", True, False)
        Dim bOfferEditable As Boolean = MyCommon.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, MyCommon, Offer.OfferID)

        If (Not rec.UIFormField = DisplayTypes.INPUT_FILE) AndAlso (rec.MediaData IsNot Nothing) Then
            MediaData = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(rec.MediaData))
        End If

        AssetId = "assetCh" & ChannelID & "Mt" & rec.MediaTypeID & "L" & rec.LanguageID

        Select Case rec.UIFormField
            Case DisplayTypes.SELECT_MULTIPLE
                FieldStr = "<select name=""" & AssetId & """ id=""" & AssetId & """ >"
            Case DisplayTypes.INPUT_FILE
                FieldStr = "<input type=""button"" id=""btn" & AssetId & """ value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""showGraphicSelector(" & ChannelID & ", " & rec.MediaTypeID & "," & rec.LanguageID & ");""" & IIf(rec.MediaData = "", "", "style="" display:none;""") & IIf(Logix.UserRoles.EditOffer = False Or Not m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID), "disabled=""disabled""", "") & " />"
                If (Logix.UserRoles.EditOffer = False Or m_EditOfferRegardlessOfBuyer Or IsOfferWaitingForApproval(Offer.OfferID)) Then
                    FieldStr &= "<img src=""show-image.aspx?caller=channels&ChannelID=" & ChannelID & "&OfferID=" & Offer.OfferID & "&MediaTypeID=" & rec.MediaTypeID & "&LanguageID=" & rec.LanguageID & "&t=" & Date.Now.Ticks & """ id=""" & AssetId & """onclick=""false""" & IIf(rec.MediaData = "", "style="" display:none;""", "") & " />"
                Else
                    FieldStr &= "<img src=""show-image.aspx?caller=channels&ChannelID=" & ChannelID & "&OfferID=" & Offer.OfferID & "&MediaTypeID=" & rec.MediaTypeID & "&LanguageID=" & rec.LanguageID & "&t=" & Date.Now.Ticks & """ id=""" & AssetId & """onclick=""showGraphicSelector(" & ChannelID & ", " & rec.MediaTypeID & "," & rec.LanguageID & ");""" & IIf(rec.MediaData = "", "style="" display:none;""", "") & " />"
                End If

            Case Else
                FieldStr = "<input type=""text"" id=""" & AssetId & """ name=""" & AssetId & """ class=""long"" value=""" & MediaData & """ " &
                           "       onkeyup=""selectChannel(" & ChannelID & ");"" onmousedown=""selectChannel(" & ChannelID & ");"" />"
        End Select

        Return FieldStr
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub Close_Page()
        Common.Close_LogixRT()
        Send_BodyEnd()
        Common = Nothing
        Logix = Nothing
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    ' Determine if a customer condition is set for the offer
    Private Function IsCustomerAssigned() As Boolean
        Dim dt As DataTable
        Dim Assigned As Boolean = False

        Common.QueryStr = "select CG.CustomerGroupID,Name,ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " &
                            "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " &
                            "where RewardOptionID=" & Offer.RewardOptionID & " and ICG.Deleted=0;"
        dt = Common.LRT_Select
        Assigned = (dt.Rows.Count > 0)

        Return Assigned
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    ' Determine if a product condition is set for the offer
    Private Function IsProductAssigned() As Boolean
        Dim dt As DataTable
        Dim Assigned As Boolean = False

        Common.QueryStr = "select count(*) NumRecs from CPE_IncentiveProductGroups with (NoLock) where Deleted=0 and RewardOptionID=" & Offer.RewardOptionID & ";"
        dt = Common.LRT_Select
        Assigned = (dt.Rows.Count > 0)

        Return Assigned
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    ' Determine if a accumulation message should be displayed
    Private Function IsAccumulationEnabled() As Boolean
        Dim dt As DataTable
        Dim AccumEnabled As Boolean = False
        ' select out to determine if we should show the option for a reward printed message which is phase2
        Common.QueryStr = "select PG.ProductGroupID, PG.Name, PT.Phrase as UnitDescription, ExcludedProducts, ProductComboID, " &
                            "QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod from CPE_IncentiveProductGroups as IPG with (NoLock) " &
                            "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " &
                            "left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " &
                            "left join CPE_UnitTypes as UT with (NoLock) on UT.UnitTypeID=IPG.QtyUnitType " &
                            "inner join PhraseText PT with (NoLock) on PT.PhraseID=UT.PhraseID " &
                            "where IPG.RewardOptionID=" & Offer.RewardOptionID & " and IPG.Deleted=0;"
        dt = Common.LRT_Select
        If (dt.Rows.Count > 0 AndAlso Common.NZ(dt.Rows(0).Item("AccumMin"), 0) > 0) Then
            AccumEnabled = True
        End If

        ' if it's not already enabled, then check for a customer and points reward being set.
        ' this causes the accumulation to be enabled as well.
        If (Not (AccumEnabled) AndAlso IsCustomerAssigned()) Then
            Common.QueryStr = "select count(*) as NumRecs from CPE_Deliverables with (NoLock) where DeliverableTypeID in (8, 11) " &
                                "and RewardOptionPhase=3 and RewardOptionID=" & Offer.RewardOptionID & " and Deleted=0;"
            dt = Common.LRT_Select
            If (dt.Rows.Count > 0 AndAlso Common.NZ(dt.Rows(0).Item("NumRecs"), 0) > 0) Then
                AccumEnabled = True
            End If
        End If

        ' Determine if the offer has a stored value discount -- this in part determines if accumulation notifications are made available.
        If (Not (AccumEnabled) AndAlso IsCustomerAssigned()) Then
            Common.QueryStr = "select count(*) NumRecs from CPE_Discounts as DI with (NoLock) " &
                                "inner join CPE_Deliverables as DE on DE.OutputID=DI.DiscountID " &
                                "where DI.AmountTypeID=7 and DE.RewardOptionID=" & Offer.RewardOptionID & ";"
            dt = Common.LRT_Select
            If (dt.Rows.Count > 0) Then
                AccumEnabled = (Common.NZ(dt.Rows(0).Item("NumRecs"), 0) > 0)
            End If
        End If

        Return AccumEnabled
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    ' Determine if this offer is a footer one.
    Private Function IsFooterOffer() As Boolean
        Dim MyCpe As New Copient.CPEOffer

        Return MyCpe.IsFooterOffer(Offer.OfferID)
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    ' Find the RewardOptionID from the OfferID
    Private Sub LoadRewardOptionID()
        Dim dt As DataTable

        Offer.RewardOptionID = 0

        Common.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) " &
                            "where IncentiveID=" & Offer.OfferID & " and TouchResponse=0 and Deleted=0;"
        dt = Common.LRT_Select
        If (dt.Rows.Count > 0) Then
            Offer.RewardOptionID = Common.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadOfferData()
        Dim dt As DataTable

        ' dig the offer info out of the database
        Common.QueryStr = "select Name, IsTemplate, FromTemplate, 0 as EngineSubTypeID, 0 as RewardOptionID, EngineID,NULL as BuyerID " &
                          "from Offers with (NoLock) " &
                          "where OfferID = " & Offer.OfferID & " " &
                          "union " &
                          "select CPE.IncentiveName as Name, CPE.IsTemplate, CPE.FromTemplate, CPE.EngineSubTypeID, RO.RewardOptionID, CPE.EngineID,buy.ExternalBuyerId as BuyerID " &
                          "from CPE_Incentives as CPE with (NoLock) " &
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID = CPE.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " &
                          "left outer join Buyers as buy with (nolock) on buy.BuyerId= CPE.BuyerId " &
                          "where CPE.IncentiveID=" & Offer.OfferID & ";"

        dt = Common.LRT_Select
        For Each row In dt.Rows
            If (Common.Fetch_UE_SystemOption(168) = "1" AndAlso Common.NZ(row.Item("BuyerID"), "") <> "") Then
                Offer.Name = "Buyer " + row.Item("BuyerID").ToString() + " -" + Common.NZ(row.Item("Name"), "").ToString()
            Else
                Offer.Name = Common.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
            End If
            'Offer.Name = Common.NZ(row.Item("Name"), "")
            Offer.RewardOptionID = Common.NZ(row.Item("RewardOptionID"), 0)
            Offer.IsTemplate = Common.NZ(row.Item("IsTemplate"), False)
            Offer.FromTemplate = Common.NZ(row.Item("FromTemplate"), False)
            Offer.EngineID = Common.NZ(row.Item("EngineID"), -1)
            Offer.EngineSubTypeID = Common.NZ(row.Item("EngineSubTypeID"), 0)
        Next

        Offer.HasAnyCustomer = CPEOffer_Has_AnyCustomer(Common, Offer.OfferID)
        LoadTemplatePermissions()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadTemplatePermissions()
        Dim dt As DataTable

        Common.QueryStr = "select DisAllow_Channels from TemplatePermissions with (NoLock) where OfferID=@OfferID"
        Common.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = Offer.OfferID
        dt = Common.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            Offer.TemplatePermissions.DisAllow_Channels = Common.NZ(dt.Rows(0).Item("DisAllow_Channels"), False)
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadAvailableLanguages()
        Dim dt As DataTable
        Dim Lang As Language

        Common.QueryStr = "select LanguageID, Name, PhraseTerm from Languages as LANG with (NoLock) where AvailableForCustFacing = 1;"
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Lang = New Language()
                Lang.ID = row.Item("LanguageID")
                Lang.Name = Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseTerm"), ""), LanguageID, Common.NZ(row.Item("Name"), ""))
                AvailableLanguages.Add(Lang)
            Next
        End If

        If AvailableLanguages Is Nothing OrElse AvailableLanguages.Count = 0 Then
            Lang = New Language()
            Lang.ID = 1
            Lang.Name = Copient.PhraseLib.Lookup("lang.name.en-US", LanguageID, "English")
            AvailableLanguages.Add(Lang)
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadLimitTypes()
        Dim TempQueryStr As String
        Dim dt As DataTable
        Dim t As New TimeType

        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        TempQueryStr = "select TimeTypeID,PhraseID from CPE_DistributionTimeTypes with (NoLock)"
        If Offer.HasAnyCustomer Then
            'Restrict time type to hours if the offer has an AnyCustomer condition since we can't carry the limits beyond a single transaction
            TempQueryStr &= " where TimeTypeID=2"
        End If

        Common.QueryStr = TempQueryStr
        dt = Common.LRT_Select
        For Each row As DataRow In dt.Rows
            t = New TimeType()
            t.ID = Common.NZ(row.Item("TimeTypeID"), 0)
            t.Name = Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseID"), 0), LanguageID)
            TimeTypes.Add(t)
        Next

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadChannels()
        Dim c As Channel
        Dim dt As DataTable
        Dim ChannelID As Integer

        Common.QueryStr = "select ChannelID, Name, PhraseTerm , UsesStartDate, UsesEndDate, UsesLimits " &
                          "from Channels with (NoLock) where Enabled=1;"
        dt = Common.LRT_Select()
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                ChannelID = row.Item("ChannelID")
                If ChannelCommon.IsChannelUsedByPromoEngine(ChannelID, Offer.EngineID) Then
                    c = New Channel()
                    c.ID = Common.NZ(row.Item("ChannelID"), 0)
                    c.Name = Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseTerm"), ""), LanguageID, Common.NZ(row.Item("Name"), ""))
                    c.UsesStartDate = Common.NZ(row.Item("UsesStartDate"), False)
                    c.UsesEndDate = Common.NZ(row.Item("UsesEndDate"), False)
                    c.UsesLimits = Common.NZ(row.Item("UsesLimits"), False)

                    LoadDates(c)

                    For Each l As Language In AvailableLanguages
                        LoadAssets(c, l.ID)
                    Next

                    ' check if the channel is POS and if so then load its specific assets
                    If c.ID = POS_CHANNEL_ID Then
                        LoadPrintedMessage(c)
                        LoadPosImgUrl(c)
                        LoadCashierMessage(c)
                        LoadAccumMessage(c)
                        LoadGraphic(c)
                        LoadLimit(c)
                        LoadPosNotificationCheckValue(c)
                    End If

                    c.Selected = IsChannelSelected(c)

                    Channels.Add(c)
                End If
            Next

            ' sort the channels by name after the phrase translation takes place
            Channels.Sort(Function(c1 As Channel, c2 As Channel)
                              Return c1.Name.CompareTo(c2.Name)
                          End Function)
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadPrintedMessage(ByRef c As Channel)
        Dim dt As DataTable

        ' printed message notifications are only available for the POS channel in the CPE
        Common.QueryStr = "select DEL.DeliverableID, PMT.MessageID, PMT.BodyText " &
                            "from CPE_Deliverables as DEL with (NoLock) " &
                            "inner join PrintedMessageTiers as PMT with (NoLock) on PMT.MessageID = DEL.OutputID " &
                            "where DEL.RewardOptionID = " & Offer.RewardOptionID & " and DEL.RewardOptionPhase = 1 " &
                            "  and DEL.Deleted = 0 and DEL.DeliverableTypeID = 4 and PMT.TierLevel = 1;"
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            c.PrintedMessage.DeliverableID = Common.NZ(dt.Rows(0).Item("DeliverableID"), 0)
            c.PrintedMessage.MessageID = Common.NZ(dt.Rows(0).Item("MessageID"), 0)
            c.PrintedMessage.BodyText = Common.NZ(dt.Rows(0).Item("BodyText"), "")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadPosImgUrl(ByRef c As Channel)
        Dim dt As DataTable

        ' ImgUrl message notifications are only available for the POS channel in the UE

        Common.QueryStr = "Select DEL.DeliverableID, PTT.PTPKID, PTT.Data " &
                        "from CPE_Deliverables as DEL with (NoLock) " &
                        "inner join PassThruTiers as PTT with (NoLock) on PTT.PTPKID = DEL.OutputID " &
                        "where DEL.RewardOptionID = " & Offer.RewardOptionID & " and DEL.RewardOptionPhase = 1 " &
                        "  and DEL.Deleted = 0 and DEL.DeliverableTypeID = 12 and PTT.TierLevel = 1;"

        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            c.PosImgURL.DeliverableID = Common.NZ(dt.Rows(0).Item("DeliverableID"), 0)
            c.PosImgURL.PTPKID = Common.NZ(dt.Rows(0).Item("PTPKID"), 0)
            c.PosImgURL.Data = Common.NZ(dt.Rows(0).Item("Data"), "")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadCashierMessage(ByRef c As Channel)
        Dim dt As DataTable

        ' printed message notifications are only available for the POS channel in the CPE
        Common.QueryStr = "select D.DeliverableID, CM.MessageID, CMT.Line1, CMT.Line2 " &
                            "from CPE_Deliverables D with (NoLock) " &
                            "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " &
                            "inner join CPE_CashierMessageTiers CMT with (NoLock) on CMT.MessageID = CM.MessageID " &
                            "where D.RewardOptionID=" & Offer.RewardOptionID & " and DeliverableTypeID=9 " &
                            "  and D.RewardOptionPhase=1 and CMT.TierLevel=1;"
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            c.CashierMessage.DeliverableID = Common.NZ(dt.Rows(0).Item("DeliverableID"), 0)
            c.CashierMessage.MessageID = Common.NZ(dt.Rows(0).Item("MessageID"), 0)
            c.CashierMessage.Line1 = Common.NZ(dt.Rows(0).Item("Line1"), "")
            c.CashierMessage.Line2 = Common.NZ(dt.Rows(0).Item("Line2"), "")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadAccumMessage(ByRef c As Channel)
        Dim dt As DataTable

        Common.QueryStr = "select DEL.DeliverableID, PM.MessageID, PMT.BodyText " &
                            "from CPE_Deliverables DEL with (NoLock) " &
                            "inner join PrintedMessages PM with (NoLock) on DEL.OutputID=PM.MessageID " &
                            "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID=PMT.MessageID " &
                            "where DEL.RewardOptionPhase=2 and DEL.RewardOptionID=" & Offer.RewardOptionID & " and DEL.DeliverableTypeID=4 and PMT.TierLevel=1;"
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            c.AccumMessage.DeliverableID = Common.NZ(dt.Rows(0).Item("DeliverableID"), 0)
            c.AccumMessage.MessageID = Common.NZ(dt.Rows(0).Item("MessageID"), 0)
            c.AccumMessage.BodyText = Common.NZ(dt.Rows(0).Item("BodyText"), "")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadGraphic(ByRef c As Channel)
        Dim dt As DataTable
        Dim gr As GraphicDeliverable

        c.Graphic = New List(Of GraphicDeliverable)

        Common.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, D.DeliverableID, D.ScreenCellID as CellID, D.DisallowEdit, " &
                            "OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName, OSA.ClientFileName " &
                            "from OnScreenAds as OSA with (NoLock) " &
                            "inner join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID " &
                            "inner join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " &
                            "inner join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID " &
                            "where D.RewardOptionID=" & Offer.RewardOptionID & " and OSA.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=1;"

        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                gr = New GraphicDeliverable
                gr.DeliverableID = Common.NZ(row.Item("DeliverableID"), 0)
                gr.OnScreenAdID = Common.NZ(row.Item("AdID"), 0)
                gr.Name = Common.NZ(row.Item("GraphicName"), "")
                gr.ImageType = Common.NZ(row.Item("ImageType"), 0)
                gr.CellSelectID = Common.NZ(row.Item("CellID"), 0)
                gr.URI = Common.NZ(row.Item("ClientFileName"), "")

                c.Graphic.Add(gr)
            Next
        Else
            gr = New GraphicDeliverable
            gr.DeliverableID = 0
            gr.OnScreenAdID = 0
            gr.Name = ""

            c.Graphic.Add(gr)
        End If

    End Sub


    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadLimit(ByRef c As Channel)
        Dim dt As DataTable

        If ChannelCommon.IsChannelUsedByPromoEngine(c.ID, Offer.EngineID) AndAlso c.ID = POS_CHANNEL_ID Then

            Common.QueryStr = "select P1DistQtyLimit, P1DistTimeType, P1DistPeriod " &
                              "from CPE_Incentives with (NoLock) where IncentiveID = " & Offer.OfferID & ";"

            dt = Common.LRT_Select
            If dt.Rows.Count > 0 Then
                c.Limit.P1DistQtyLimit = Common.NZ(dt.Rows(0).Item("P1DistQtyLimit"), 1)
                c.Limit.P1DistTimeType = Common.NZ(dt.Rows(0).Item("P1DistTimeType"), 0)
                c.Limit.P1DistPeriod = Common.NZ(dt.Rows(0).Item("P1DistPeriod"), -1)
                If c.Limit.P1DistQtyLimit = 0 AndAlso c.Limit.P1DistTimeType = 0 AndAlso c.Limit.P1DistPeriod = -1 Then
                    c.Limit.P1DistPeriod = 0
                End If
            End If
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadPosNotificationCheckValue(ByRef c As Channel)
        Dim dt As DataTable
        If ChannelCommon.IsChannelUsedByPromoEngine(c.ID, Offer.EngineID) AndAlso c.ID = POS_CHANNEL_ID Then
            Common.QueryStr = "select PosNotificationCheck from CPE_Incentives with (NoLock) where IncentiveID = " & Offer.OfferID & ";"
            dt = Common.LRT_Select
            If dt.Rows.Count > 0 Then
                c.PosNotificationCheck = Common.NZ(dt.Rows(0).Item("PosNotificationCheck"), 0)
            End If
        End If
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadDates(ByRef c As Channel)
        Dim dt As DataTable

        If c.UsesStartDate OrElse c.UsesEndDate Then

            Common.QueryStr = "select StartDate, EndDate from ChannelOffers with (NoLock) " &
                              "where OfferID = " & Offer.OfferID & " and ChannelID = " & c.ID & ";"

            dt = Common.LRT_Select
            If dt.Rows.Count > 0 Then
                c.StartDate = Common.NZ(dt.Rows(0).Item("StartDate"), Date.MinValue)
                c.EndDate = Common.NZ(dt.Rows(0).Item("EndDate"), Date.MinValue)
            End If
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub LoadAssets(ByRef c As Channel, ByVal LanguageID As Integer)
        Dim dt As DataTable
        Dim Asset As ChannelOfferAssestRec

        If c.Assets Is Nothing Then
            c.Assets = New List(Of ChannelOfferAssestRec)
        End If

        Common.QueryStr = "select CM.MediaTypeID, CMT.Name, CMT.PhraseTerm, CMT.FieldTypeID, COA.MediaData " &
                          "from ChannelMedia as CM with (NoLock) " &
                          "inner join ChannelMediaTypes as CMT with (NoLock) on CMT.MediaTypeID = CM.MediaTypeID " &
                          "inner join ChannelMediaEngines as CME with (NoLock) on CME.ChannelMediaID = CM.ChannelMediaID " &
                          "left join ChannelOfferAssets as COA with (NoLock) " &
                          "  on COA.MediaTypeID = CMT.MediaTypeID and COA.ChannelID = " & c.ID & " and COA.OfferID = " & Offer.OfferID & " " &
                          "     and COA.LanguageID = " & LanguageID & " " &
                          "where CM.ChannelID = " & c.ID & "and CM.Enabled=1 and CME.PromoEngineID = " & Offer.EngineID & " " &
                          "order by CM.DisplayOrder asc;"

        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                Asset = New ChannelOfferAssestRec
                Asset.MediaTypeID = Common.NZ(row.Item("MediaTypeID"), 0)
                Asset.MediaData = Common.NZ(row.Item("MediaData"), "")
                Asset.Name = Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseTerm"), "").ToString, LanguageID, Common.NZ(row.Item("Name"), "").ToString)
                Asset.LanguageID = LanguageID
                Asset.UIFormField = Common.NZ(row.Item("FieldTypeID"), DisplayTypes.UNKNOWN)
                c.Assets.Add(Asset)
            Next
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub RemoveChannel(ByVal ChannelID As Integer)
        Dim Success As Boolean = True
        Dim Message As String = ""
        Dim TempStr As String = ""
        Dim AssetDT As DataTable = Nothing
        Dim index As Integer

        Try
            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()

            ' remove all channel offer and asset records before creating the currently selected ones.
            DeleteChannelOffers(ChannelID)
            If ChannelID = POS_CHANNEL_ID Then
                ' grab all of the non-text assets we are going to need to hide on the page (need to return these in the ajax response)
                Common.QueryStr = "select OutputID, DeliverableTypeID " &
                                  "from CPE_Deliverables as D Inner Join CPE_RewardOptions as RO " &
                                  "on D.RewardOptionID=RO.RewardOptionID and D.Deleted=0 and RO.TouchResponse=0 and D.RewardOptionPhase=1 " &
                                  "where RO.IncentiveID=" & Offer.OfferID & ";"
                AssetDT = Common.LRT_Select
                Success = DeletePrintedMessage(Offer.OfferID, TempStr)
                Success = DeleteImageUrl(Offer.OfferID, TempStr)
                Success = DeleteAccumMessage(Offer.OfferID, TempStr)
                Success = DeleteCashierMessage(Offer.OfferID, TempStr)
                Success = DeleteGraphic(Offer.OfferID, TempStr)
                UpdatePosNotificationCheckValue(0) 'Zero is default value for PosNotificationCheckValue
            Else
                ' grab all of the non-text assets we are going to need to hide on the page (need to return these in the ajax response)
                Common.QueryStr = "select MediaTypeID, LanguageID from channelOfferAssets where ChannelID=" & ChannelID & " and mediaformatID<>1 and OfferID=" & Offer.OfferID & ";"
                AssetDT = Common.LRT_Select
                DeleteChannelOfferAssets(ChannelID)
            End If

            'Since the offer content for this channel is being deleted, we need to automatically deploy it so the channel knows to delete its cached copy of the offer
            ChannelCommon.DeployChannelForOffer(ChannelID, Offer.OfferID)
            Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Detokenize("offer-channels.deletedChannel", LanguageID, ChannelCommon.GetChannelName(ChannelID)))
            Message = Copient.PhraseLib.Lookup("term.removed", LanguageID)

        Catch ex As Exception
            Success = False
            Message = Copient.PhraseLib.Lookup("term.SaveFailed", LanguageID) & ": " & ex.ToString()

        Finally
            Common.QueryStr = IIf(Success, "COMMIT", "ROLLBACK") & " TRANSACTION;"
            Common.LRT_Execute()

            If Success Then SetOfferStatusFlagUpdate()

        End Try

        ' send back the response 
        Sendb("{")
        Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """,")
        Sendb(" ""DelAssets"": [ ")
        If Not (AssetDT Is Nothing) Then
            If Not (ChannelID = 1) Then  'for all of the channels except the POS channel
                For index = 0 To (AssetDT.Rows.Count - 1)
                    Sendb("{ ""assetID"": ""assetCh" & ChannelID & "Mt" & AssetDT.Rows(index)("MediaTypeID").ToString & "L" & AssetDT.Rows(index)("LanguageID").ToString & """ }" & IIf(index = AssetDT.Rows.Count - 1, " ", ", "))
                Next
            Else 'for the POS channel only
                For index = 0 To (AssetDT.Rows.Count - 1)
                    Sendb("{ ""assetID"": ""assetCh" & ChannelID & "Mt")
                    Select Case (AssetDT.Rows(index)("DeliverableTypeID"))
                        Case 1
                            Sendb("9") 'CPE_DeliverableTypeID 1 = ChannelMediaTypeID 9 - POS Graphic
                            Sendb("Ad" & AssetDT.Rows(index)("OutputID"))
                        Case 4
                            Sendb("1") 'CPE_DeliverableTypeID 4 = ChannelMediaTypeID 1 - POS Printed message
                        Case 9
                            Sendb("6")  'CPE_DeliverableTypeID 9 = ChannelMediaTypeID 6 - POS Cashier message
                        Case 12
                            Sendb("17")  'CPE_DeliverableTypeID 12 = ChannelMediaTypeID 17 - POS Image URL message
                        Case Else
                            Sendb("0")
                    End Select
                    Sendb(""" }")
                    Sendb(IIf(index = AssetDT.Rows.Count - 1, " ", ", "))
                Next
            End If
        End If
        Sendb("] ")

        If Not ChannelID = POS_CHANNEL_ID Then
            Sendb(", ")
            Sendb(GetDeployMsgJSON(ChannelID, Offer.OfferID, False))
        End If
        Sendb("}")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub SaveChannel(ByVal ChannelID As Integer)
        Dim Success As Boolean = True
        Dim Message As String = ""

        Try
            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()

            ' remove all channel offer and asset records before creating the currently selected ones.
            DeleteChannelOffers(ChannelID)

            If Not ChannelID = POS_CHANNEL_ID Then
                DeleteChannelOfferAssets_ForSave(ChannelID)
            End If

            SaveChannelOffer(ChannelID)
            SaveChannelOfferAssets(ChannelID)
            If ChannelID = POS_CHANNEL_ID Then
                SaveLimits(ChannelID)

                ' Update PosNotificationCheckValue
                Dim PosProductionConditionValue As Integer = 0
                Integer.TryParse(GetCgiValue("PosProductionConditionCheck"), PosProductionConditionValue)
                UpdatePosNotificationCheckValue(PosProductionConditionValue)
            End If

            Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Detokenize("offer-channels.editedChannel", LanguageID, ChannelCommon.GetChannelName(ChannelID)))
            Message = Copient.PhraseLib.Lookup("term.ChangesWereSaved", LanguageID)

        Catch ex As Exception
            Success = False
            Message = Copient.PhraseLib.Lookup("term.SaveFailed", LanguageID) & ": " & ex.ToString()

        Finally
            Common.QueryStr = IIf(Success, "COMMIT", "ROLLBACK") & " TRANSACTION;"
            Common.LRT_Execute()

            If Success Then SetOfferStatusFlagUpdate()

        End Try

        ' send back the response 
        Sendb("{")
        Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """")
        If Not (ChannelID = POS_CHANNEL_ID) Then
            Sendb(", " & GetDeployMsgJSON(ChannelID, Offer.OfferID, False))
        End If
        Sendb("}")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub SetOfferStatusFlagUpdate()

        If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.UE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
            Common.QueryStr = "Update CPE_Incentives set StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
        ElseIf Offer.EngineID = CommonInc.InstalledEngines.CM Then
            Common.QueryStr = "Update Offers set StatusFlag=1 where OfferId=" & Offer.OfferID & ";"
        End If
        Common.LRT_Execute()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Function GetDeployMsgJSON(ByVal ChannelID As Integer, ByVal OfferID As Long, DeploymentSuccessFlag As Boolean) As String
        Dim DeployMsg As String = ""
        Dim DeployMsgColor As String = "black"
        Dim ReturnVal As String = ""

        If DeploymentSuccessFlag Then
            DeployMsg = " " & Copient.PhraseLib.Lookup("term.deployed", LanguageID)
            DeployMsgColor = "#00FF00"
        End If
        If ChannelCommon.HasChannelChanged(ChannelID, OfferID) Then
            DeployMsg = " " & Copient.PhraseLib.Lookup("term.changesnotdeployed", LanguageID)
            DeployMsgColor = "#FF3333"
        End If
        If DeployMsg = "" AndAlso ChannelCommon.IsChannelWaitingDeployment(ChannelID, OfferID) Then
            DeployMsg = " " & Copient.PhraseLib.Lookup("term.awaitingdeployment", LanguageID)
            DeployMsgColor = "#00FF00"
        End If
        ReturnVal = ReturnVal & " ""DeployMsg"": """ & DeployMsg & ""","
        ReturnVal = ReturnVal & " ""DeployMsgColor"": """ & DeployMsgColor & """"
        Return ReturnVal

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub DeleteChannelOffers(ByVal ChannelID As Integer)

        Common.QueryStr = "dbo.pt_ChannelOffers_Delete"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = Offer.OfferID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub SaveChannelOffer(ByVal ChannelID As Integer)
        Dim StartDate, EndDate As Date
        Dim SubChannelList As List(Of Copient.ChannelCommon.SubChannelOfferRec)
        Dim OfferSubChannelList As List(Of Copient.ChannelCommon.SubChannelOfferRec) = New List(Of Copient.ChannelCommon.SubChannelOfferRec)
        Dim OfferSubChannelRec As Copient.ChannelCommon.SubChannelOfferRec
        Dim TempStr As String

        Date.TryParse(GetCgiValue("chStart" & ChannelID), Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, StartDate)
        Date.TryParse(GetCgiValue("chEnd" & ChannelID), Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, EndDate)

        Common.QueryStr = "dbo.pt_ChannelOffers_Insert"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = Offer.OfferID
        Common.LRTsp.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = IIf(StartDate = Date.MinValue, DBNull.Value, StartDate)
        Common.LRTsp.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = IIf(EndDate = Date.MinValue, DBNull.Value, EndDate)
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()

        ' update the eligibility dates for the POS channel with the start and end dates
        If ChannelID = POS_CHANNEL_ID Then
            Common.QueryStr = "dbo.pa_CPE_UpdateEligibilityDates"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = Offer.OfferID
            Common.LRTsp.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = IIf(StartDate = Date.MinValue, DBNull.Value, StartDate)
            Common.LRTsp.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = IIf(EndDate = Date.MinValue, DBNull.Value, EndDate)
            Common.LRTsp.ExecuteNonQuery()
            Common.Close_LRTsp()
        Else
            'save selection of any sub-channels
            SubChannelList = ChannelCommon.GetOfferSubChannels(ChannelID, Offer.OfferID)
            If SubChannelList.Count > 0 Then
                If Request.Form.GetValues("Ch" & ChannelID.ToString & "subchannels") IsNot Nothing Then
                    For Each TempStr In Request.Form.GetValues("Ch" & ChannelID.ToString & "subchannels")
                        OfferSubChannelRec = New Copient.ChannelCommon.SubChannelOfferRec
                        OfferSubChannelRec.SubChannelID = CInt(TempStr)
                        OfferSubChannelRec.Associated = True
                        OfferSubChannelList.Add(OfferSubChannelRec)
                    Next
                End If
                ChannelCommon.SetSubChannelsOffers(ChannelID, Offer.OfferID, OfferSubChannelList)
            End If
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub SaveLimits(ByVal ChannelID As Integer)
        Dim P1DistQtyLimit, P1DistTimeType, P1DistPeriod As Integer

        Integer.TryParse(GetCgiValue("limit1Ch" & ChannelID), P1DistQtyLimit)
        Integer.TryParse(GetCgiValue("P1DistTimeTypeCh" & ChannelID), P1DistTimeType)
        Integer.TryParse(GetCgiValue("limit1periodCh" & ChannelID), P1DistPeriod)

        Common.QueryStr = "update CPE_Incentives with (RowLock) " &
                          "  set P1DistQtyLimit=" & P1DistQtyLimit & ", " &
                          "      P1DistTimeType=" & P1DistTimeType & ", " &
                          "      P1DistPeriod=" & IIf(P1DistPeriod = -1, "null", P1DistPeriod) & ", " &
                          "      StatusFlag=1, " &
                          "      LastUpdate=getdate(), " &
                          "      LastUpdatedByAdminID=" & AdminUserID & " " &
                          "  where IncentiveID =" & Offer.OfferID & ";"
        Common.LRT_Execute()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  
    'Update PosNotificationCheck Value in CPE_Incentives table
    Private Sub UpdatePosNotificationCheckValue(PosProductionConditionValue As Integer)
        Common.QueryStr = "update CPE_Incentives with (RowLock) " &
                      "  set PosNotificationCheck=" & PosProductionConditionValue & ", " &
                      "      StatusFlag=1, " &
                      "      LastUpdate=getdate(), " &
                      "      LastUpdatedByAdminID=" & AdminUserID & " " &
                      "  where IncentiveID =" & Offer.OfferID & ";"
        Common.LRT_Execute()
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub DeleteChannelOfferAssets(ByVal ChannelID As Integer)

        ' remove all the offer assets for the channel/offer combination
        Common.QueryStr = "delete ChannelOfferAssets " &
                          "from ChannelOfferAssets as COA with (NoLock) " &
                          "inner join ChannelMediaTypes as CMT with (NoLock) on CMT.MediaTypeID = COA.MediaTypeID " &
                          "where COA.ChannelID=" & ChannelID & " and COA.OfferID=" & Offer.OfferID & ";"
        Common.LRT_Execute()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub DeleteChannelOfferAssets_ForSave(ByVal ChannelID As Integer)

        ' remove all the offer assets for the channel/offer combination
        Common.QueryStr = "delete ChannelOfferAssets " &
                          "from ChannelOfferAssets as COA with (NoLock) " &
                          "inner join MediaFormats as MF on COA.MediaFormatID=MF.MediaFormatID " &
                          "where MF.MediaFormatID=1 and COA.ChannelID=" & ChannelID & " and COA.OfferID=" & Offer.OfferID & ";"
        Common.LRT_Execute()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub SaveChannelOfferAssets(ByVal ChannelID As Integer)
        Dim MediaData As String = ""
        Dim MediaTypeID As Integer
        Dim MediaTypes() As String
        Dim insertStatus As Integer
        'Dim ms As System.IO.MemoryStream

        ' load all the asset media types for this channel
        MediaTypes = Request.Form.GetValues("assetCh" & ChannelID & "MediaTypeID")

        If MediaTypes IsNot Nothing Then
            For Each l As Language In AvailableLanguages
                For Each mtID As String In MediaTypes
                    Integer.TryParse(mtID, MediaTypeID)
                    MediaData = Common.NZ(GetCgiValue("assetCh" & ChannelID & "Mt" & mtID & "L" & l.ID), "").ToString()

                    ' Convert the media data to base 64 for storage of all asset values
                    If Not IsGraphicMediaType(mtID) Then
                        MediaData = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(MediaData))

                        If MediaTypeID > 0 AndAlso MediaData.Trim().Length > 0 Then
                            Common.QueryStr = "dbo.pt_ChannelOfferAssets_Insert"
                            Common.Open_LRTsp()
                            Common.LRTsp.Parameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
                            Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = Offer.OfferID
                            Common.LRTsp.Parameters.Add("@MediaTypeID", SqlDbType.Int).Value = MediaTypeID
                            Common.LRTsp.Parameters.Add("@MediaData", SqlDbType.NVarChar, -1).Value = MediaData
                            Common.LRTsp.Parameters.Add("@PreviewMediaData", SqlDbType.NVarChar, -1).Value = DBNull.Value
                            Common.LRTsp.Parameters.Add("@MediaFormatID", SqlDbType.Int).Value = GetMediaFormatID(mtID)
                            Common.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int).Value = l.ID
                            Common.LRTsp.Parameters.Add("@InsertStatus", SqlDbType.Int).Direction = ParameterDirection.Output
                            Common.LRTsp.ExecuteNonQuery()
                            insertStatus = Common.LRTsp.Parameters("@InsertStatus").Value
                            MyCommon.Close_LRTsp()

                            ' insertStatus = 1 means record inserted.
                            ' insertStatus = 2 means "Channel input provided to dbo.pt_ChannelOfferAssets_Insert is invalid, so record was not inserted."
                            ' insertStatus = 0 means "Channel input provided to dbo.pt_ChannelOfferAssets_Insert is duplicate.  Record already exists, so no record was inserted." 
                        End If
                    End If
                Next
            Next

        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub RemoveAsset(ByVal ChannelID As Integer, ByVal MediaTypeID As Integer, ByVal AssetLanguageID As Integer)
        Dim Success As Boolean = True
        Dim Message As String = ""

        Try
            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()

            Common.QueryStr = "dbo.pt_ChannelOfferAssets_Delete"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
            Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = Offer.OfferID
            Common.LRTsp.Parameters.Add("@MediaTypeID", SqlDbType.Int).Value = MediaTypeID
            Common.LRTsp.Parameters.Add("@LanguageID", SqlDbType.Int).Value = AssetLanguageID
            Common.LRTsp.ExecuteNonQuery()
            Common.Close_LRTsp()

            Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Detokenize("offer-channels.deletedChannelAsset", LanguageID, ChannelCommon.GetMediaTypeName(MediaTypeID), ChannelCommon.GetChannelName(ChannelID)))
            Message = Copient.PhraseLib.Lookup("offer-asset.removedValue", LanguageID)

        Catch ex As Exception
            Success = False
            Message = Copient.PhraseLib.Lookup("offer-asset.removedFailed", LanguageID) & ": " & ex.ToString()

        Finally
            Common.QueryStr = IIf(Success, "COMMIT", "ROLLBACK") & " TRANSACTION;"
            Common.LRT_Execute()
        End Try

        ' send back the response 
        Sendb("{")
        Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""MediaTypeID"": " & MediaTypeID & ",")
        Sendb(" ""LanguageID"": " & AssetLanguageID & ",")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """,")
        Sendb(GetDeployMsgJSON(ChannelID, Offer.OfferID, False))
        Sendb("}")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    'This function deletes the eligibility printed message related to the supplied OfferID
    Private Function DeletePrintedMessage(ByVal OfferID As Long, ByRef ResultMessage As String) As Boolean
        Dim Success As Boolean = False

        Try
            ResultMessage = ""

            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()

            If (OfferID > 0) Then
                Common.QueryStr = "Delete from PrintedMessageTiers " &
                                  "from PrintedMessageTiers as PMT Inner Join CPE_Deliverables as D " &
                                  "     on PMT.MessageID=D.OutputID and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=4 and PMT.TierLevel=1 " &
                                  "Inner Join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from PrintedMessages " &
                                  "from PrintedMessages as PM Inner Join CPE_Deliverables as D " &
                                  "     on PM.MessageID=D.OutputID and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=4 " &
                                  "Inner Join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from CPE_Deliverables " &
                                  "from CPE_Deliverables as D Inner join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "     and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=4 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletepmsg", LanguageID))
            End If
            Success = True
            Common.QueryStr = "COMMIT TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedValue", LanguageID)

        Catch ex As Exception
            Success = False
            Common.QueryStr = "ROLLBACK TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedFailed", LanguageID) & ": " & ex.ToString()

        End Try
        Return Success

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Function DeleteImageUrl(ByVal OfferID As Long, ByRef ResultMessage As String) As Boolean

        Dim Success As Boolean = False

        Try
            ResultMessage = ""

            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()

            If (OfferID > 0) Then
                Common.QueryStr = "Delete from PassThruTiers " &
                                  "from PassThruTiers as PTT Inner Join CPE_Deliverables as D " &
                                  "     on PTT.PTPKID=D.OutputID and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=12 and PTT.TierLevel=1 " &
                                  "Inner Join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from PassThrus " &
                                  "from PassThrus as PT Inner Join CPE_Deliverables as D " &
                                  "     on PT.PKID=D.OutputID and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=12 " &
                                  "Inner Join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from CPE_Deliverables " &
                                  "from CPE_Deliverables as D Inner join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "     and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=12 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deleteimgurlmsg", LanguageID))
            End If
            Success = True
            Common.QueryStr = "COMMIT TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedValue", LanguageID)

        Catch ex As Exception
            Success = False
            Common.QueryStr = "ROLLBACK TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedFailed", LanguageID) & ": " & ex.ToString()

        End Try
        Return Success

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    'This function deletes the eligibility accumulation message related to the supplied OfferID
    Private Function DeleteAccumMessage(ByVal OfferID As Long, ByRef ResultMessage As String) As Boolean
        Dim Success As Boolean = False

        Try
            ResultMessage = ""

            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()

            If (OfferID > 0) Then
                Common.QueryStr = "Delete from PrintedMessageTiers " &
                                  "from PrintedMessageTiers as PMT Inner Join CPE_Deliverables as D " &
                                  "     on PMT.MessageID=D.OutputID and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=4 and PMT.TierLevel=1 " &
                                  "Inner Join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from PrintedMessages " &
                                  "from PrintedMessages as PM Inner Join CPE_Deliverables as D " &
                                  "     on PM.MessageID=D.OutputID and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=4 " &
                                  "Inner Join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from CPE_Deliverables " &
                                  "from CPE_Deliverables as D Inner join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "     and D.RewardOptionPhase=2 and D.Deleted=0 and D.DeliverableTypeID=4 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletepmsg", LanguageID))
            End If
            Success = True
            Common.QueryStr = "COMMIT TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedValue", LanguageID)

        Catch ex As Exception
            Success = False
            Common.QueryStr = "ROLLBACK TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedFailed", LanguageID) & ": " & ex.ToString()

        End Try
        Return Success

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub RemovePrintedMessage(ByVal ChannelID As Integer, ByVal MediaTypeID As Integer, ByVal DeliverableID As Integer, ByVal MessageID As Integer)
        Dim Success As Boolean = True
        Dim Message As String = ""

        If MediaTypeID = MediaTypes.POS_Receipt_Message Then
            Success = DeletePrintedMessage(Offer.OfferID, Message)
        ElseIf MediaTypeID = MediaTypes.POS_Accumulation_Message Then
            Success = DeleteAccumMessage(Offer.OfferID, Message)
        End If

        ' send back the response 
        Sendb("{")
        Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""MediaTypeID"": " & MediaTypeID & ",")
        Sendb(" ""DeliverableID"": " & DeliverableID & ",")
        Sendb(" ""MessageID"": " & MessageID & ",")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """")
        Sendb("}")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub RemovePosImgUrl(ByVal ChannelID As Integer, ByVal MediaTypeID As Integer, ByVal DeliverableID As Integer, ByVal PTPKID As Integer)
        Dim Success As Boolean = True
        Dim Message As String = ""
        Success = DeleteImageUrl(Offer.OfferID, Message)

        ' send back the response 
        Sendb("{")
        Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""MediaTypeID"": " & MediaTypeID & ",")
        Sendb(" ""DeliverableID"": " & DeliverableID & ",")
        Sendb(" ""PTPKID"": " & PTPKID & ",")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """")
        Sendb("}")

        'Sendb("{")
        'Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        'Sendb(" ""ChannelID"": " & ChannelID & ",")
        'Sendb(" ""MediaTypeID"": " & MediaTypeID & ",")
        ''Sendb(" ""DeliverableID"": " & DeliverableID & ",")
        ''Sendb(" ""MessageID"": " & MessageID & ",")
        'Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """")
        'Sendb("}")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    'This function deletes the eligibility cashier message related to the supplied OfferID
    Private Function DeleteCashierMessage(ByVal OfferID As Long, ByRef ResultMessage As String) As Boolean
        Dim Success As Boolean = False

        Try
            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()

            If (OfferID > 0) Then
                Common.QueryStr = "delete from CPE_CashierMessageTiers with (RowLock) " &
                                  "from CPE_CashierMessageTiers as CMT Inner Join CPE_Deliverables as D on CMT.MessageID=D.OutputID " &
                                  "     and D.DeliverableTypeID=9 and CMT.TierLevel=1 and D.RewardOptionPhase=1 " &
                                  "Inner Join CPE_RewardOptions as RO on D.RewardOptionID=RO.RewardOptionID and RO.TouchResponse=0 " &
                                  "where RO.IncentiveID=" & OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from CPE_CashierMessages with (RowLock) " &
                                  "from CPE_CashierMessages as CM Inner Join CPE_Deliverables as D on CM.MessageID=D.OutputID " &
                                  "     and D.DeliverableTypeID=9 and D.RewardOptionPhase=1 " &
                                  "Inner Join CPE_RewardOptions as RO on D.RewardOptionID=RO.RewardOptionID and RO.TouchResponse=0 " &
                                  "where RO.IncentiveID=" & OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "delete from CPE_Deliverables with (RowLock) " &
                                  "from CPE_Deliverables as D Inner join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                                  "     and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=9 " &
                                  "Where RO.IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
                Common.LRT_Execute()
                Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletecashier", LanguageID))
            End If
            Success = True
            Common.QueryStr = "COMMIT TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedValue", LanguageID)

        Catch ex As Exception
            Success = False
            Common.QueryStr = "ROLLBACK TRANSACTION;"
            Common.LRT_Execute()
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedFailed", LanguageID) & ": " & ex.ToString()
        End Try

        Return Success

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub RemoveCashierMessage(ByVal ChannelID As Integer, ByVal MediaTypeID As Integer, ByVal DeliverableID As Integer, ByVal MessageID As Integer)
        Dim Success As Boolean = True
        Dim Message As String = ""

        Success = DeleteCashierMessage(Offer.OfferID, Message)

        ' send back the response 
        Sendb("{")
        Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""MediaTypeID"": " & MediaTypeID & ",")
        Sendb(" ""DeliverableID"": " & DeliverableID & ",")
        Sendb(" ""MessageID"": " & MessageID & ",")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """")
        Sendb("}")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    'This function deletes the eligibility graphic related to the supplied OfferID
    Private Function DeleteGraphic(ByVal OfferID As Long, ByRef ResultMessage As String) As Boolean

        Dim Success As Boolean = False
        Dim Done As Boolean
        Dim DT As DataTable

        Try
            Common.QueryStr = "BEGIN TRANSACTION;"
            Common.LRT_Execute()
            ' NOTE TO SELF - need to delete any CPE_DeliverableROIDs records related to the eligibility graphic ... the following query is NOT correct (not finished)
            Common.QueryStr = "delete from CPE_DeliverableROIDs with (RowLock) " &
                              "from CPE_DeliverableROIDS as DR Inner Join CPE_Deliverables as D on D.DeliverableID=DR.DeliverableID " &
                              "Inner Join CPE_RewardOptions as RO on D.RewardOptionID=RO.RewardOptionID and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=1 and RO.TouchResponse=0 " &
                              "Where RO.IncentiveID=" & OfferID & ";"
            Common.LRT_Execute()

            Common.QueryStr = "delete from CPE_Deliverables with (RowLock) " &
                              "from CPE_Deliverables as D Inner join CPE_RewardOptions as RO on RO.RewardOptionID=D.RewardOptionID and RO.TouchResponse=0 " &
                              "     and D.RewardOptionPhase=1 and D.Deleted=0 and D.DeliverableTypeID=1 " &
                              "Where RO.IncentiveID=" & OfferID & ";"
            Common.LRT_Execute()
            Common.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
            Common.LRT_Execute()

            'clean up any orphans created by deleting the deliverable
            Done = False
            While Not (Done)
                Common.QueryStr = "select RO.RewardOptionID from CPE_RewardOptions as RO with (NoLock) " &
                                    "where RO.IncentiveID=" & OfferID & " and RO.TouchResponse=1 and RO.Deleted=0 and RO.RewardOptionID not in " &
                                    "(select RewardOptionID from CPE_DeliverableROIDs as DR with (NoLock) where DR.IncentiveID=" & OfferID & " and DR.Deleted=0);"
                DT = Common.LRT_Select
                If (DT.Rows.Count > 0) Then
                    For Each row In DT.Rows
                        'before we delete any deliverables from the orphaned ROID, delete the DeliverableROIDs records for those Deliverables
                        Common.QueryStr = "delete from CPE_DeliverableROIDs with (RowLock) where DeliverableID in " &
                                          "(select DeliverableID from CPE_Deliverables with (NoLock) where RewardOptionID=" & Common.NZ(row.Item("RewardOptionID"), 0) & ");"
                        Common.LRT_Execute()
                        'now delete the Deliverables for this ROID - may create more orphaned ROIDS ... that's why we are looping
                        Common.QueryStr = "delete from CPE_Deliverables with (RowLock) where RewardOptionID=" & Common.NZ(row.Item("RewardOptionID"), 0) & ";"
                        Common.LRT_Execute()
                        'finally ... delete the orphaned ROID
                        Common.QueryStr = "update CPE_RewardOptions with (RowLock) set Deleted=1, LastUpdate=getdate() where RewardOptionID=" & Common.NZ(row.Item("RewardOptionID"), 0) & ";"
                        Common.LRT_Execute()
                    Next
                Else
                    Done = True
                End If
            End While
            Success = True
            Common.QueryStr = "COMMIT TRANSACTION;"
            Common.LRT_Execute()

            Common.Activity_Log(3, Offer.OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Notification.deletegraphic", LanguageID))
            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedValue", LanguageID)

        Catch ex As Exception
            Success = False
            Common.QueryStr = "ROLLBACK TRANSACTION;"
            Common.LRT_Execute()

            ResultMessage = Copient.PhraseLib.Lookup("offer-asset.removedFailed", LanguageID) & ": " & ex.ToString()
        End Try

        Return Success

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub RemoveGraphic(ByVal ChannelID As Integer, ByVal MediaTypeID As Integer, ByVal DeliverableID As Integer, ByVal OnScreenAdID As Integer)
        Dim Success As Boolean = True
        Dim Message As String = ""

        Success = DeleteGraphic(Offer.OfferID, Message)

        ' send back the response 
        Sendb("{")
        Sendb(" ""Status"": " & IIf(Success, "1", "0") & ",")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""MediaTypeID"": " & MediaTypeID & ",")
        Sendb(" ""DeliverableID"": " & DeliverableID & ",")
        Sendb(" ""OnScreenAdID"": " & OnScreenAdID & ",")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """")
        Sendb("}")

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Sub GetFormValues()
        Offer.OfferID = ParseLong("OfferID")
        Offer.RewardOptionID = ParseLong("RewardID")
        Offer.EngineID = -1
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Function ParseLong(ByVal TokenName As String) As Long
        Dim TempLong As Long

        Long.TryParse(GetCgiValue(TokenName), TempLong)
        Return TempLong
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Function DateToString(ByVal d As Date) As String
        Dim DateStr As String = ""

        If d > Date.MinValue Then
            DateStr = Logix.ToShortDateString(d, Common)
        End If

        Return DateStr
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Function CheckChannelDate(ByVal d As Date, ByVal name As String) As String
        If String.IsNullOrWhiteSpace(name) Then
            If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.UE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
                name = "StartDate"
            ElseIf Offer.EngineID = CommonInc.InstalledEngines.CM Then
                name = "ProdStartDate"
            Else
                name = "ProdStartDate"
            End If
        End If

        If ((d = Nothing) OrElse (d <= Date.MinValue)) Then
            If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.UE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
                Common.QueryStr = "SELECT " & name & " FROM CPE_Incentives WITH (NoLock) WHERE IncentiveID=" & Offer.OfferID & ";"
            ElseIf Offer.EngineID = CommonInc.InstalledEngines.CM Then
                Common.QueryStr = "SELECT " & name & " FROM Offers with (NoLock) WHERE OfferID=" & Offer.OfferID & ";"
            Else
                Common.QueryStr = "SELECT " & name & " FROM Offers with (NoLock) WHERE OfferID=" & Offer.OfferID & ";"
            End If

            Dim rst As DataTable = Common.LRT_Select()
            If rst.Rows.Count > 0 Then
                d = Common.NZ(rst.Rows(0).Item(name), "1/1/1900")
            End If
        End If

        Return DateToString(d)
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Private Function IsChannelSelected(ByVal c As Channel) As Boolean
        Dim Selected As Boolean = False

        ' check if any POS-specific properties of the channel are set
        Selected = c.PrintedMessage.DeliverableID > 0 OrElse c.CashierMessage.DeliverableID > 0 OrElse c.AccumMessage.DeliverableID OrElse (c.Graphic IsNot Nothing AndAlso c.Graphic.Count > 0 AndAlso c.Graphic(0).DeliverableID > 0)

        ' check if dates are set
        Selected = Selected OrElse (c.StartDate > Date.MinValue OrElse c.EndDate > Date.MinValue)

        ' check if any assets are currently assigned to this channel
        If Not Selected AndAlso (c.Assets IsNot Nothing AndAlso c.Assets.Count > 0) Then
            For Each rec As ChannelOfferAssestRec In c.Assets
                Selected = rec.MediaData IsNot Nothing AndAlso rec.MediaData.Trim.Length > 0
                If Selected Then Exit For
            Next
        End If

        Return Selected

    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function IsGraphicMediaType(ByVal MediaTypeID As Integer) As Boolean
        Dim dt As DataTable

        Common.QueryStr = "select MediaTypeID from ChannelMediaTypes " &
                          "where MediaTypeID=" & MediaTypeID & " and FieldTypeID=" & DisplayTypes.INPUT_FILE & ";"
        dt = Common.LRT_Select

        Return (dt.Rows.Count > 0)
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function FindAssetByLanguageID(ByVal Asset As ChannelOfferAssestRec, ByVal LanguageID As Integer) As Boolean
        Return (Asset.LanguageID = LanguageID)
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function AreValidDates(ByVal ChannelID As Integer, ByRef RetMsg As String) As Boolean
        Dim Valid As Boolean
        Dim StartDate, EndDate As Date

        Valid = Date.TryParse(GetCgiValue("chStart" & ChannelID), Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, StartDate)
        If Not Valid OrElse StartDate.Year < 1753 Then
            RetMsg = Copient.PhraseLib.Lookup("term.InvalidStartDate", LanguageID)
            Return False
        End If

        Valid = Date.TryParse(GetCgiValue("chEnd" & ChannelID), Common.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, EndDate)
        If Not Valid OrElse EndDate.Year < 1753 Then
            RetMsg = Copient.PhraseLib.Lookup("term.InvalidEndDate", LanguageID)
            Return False
        End If

        If StartDate > EndDate Then
            RetMsg = Copient.PhraseLib.Lookup("reports.startdate", LanguageID)
            Return False
        End If

        Return Valid
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function AreValidAssets(ByVal ChannelID As Integer, ByRef RetMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim MediaTypes() As String
        Dim MediaTypeID As Integer
        Dim MediaData As String

        ' load all the asset media types for this channel
        MediaTypes = Request.Form.GetValues("assetCh" & ChannelID & "MediaTypeID")

        If MediaTypes IsNot Nothing Then
            For Each l As Language In AvailableLanguages
                For Each mtID As String In MediaTypes
                    Integer.TryParse(mtID, MediaTypeID)
                    MediaData = Common.NZ(GetCgiValue("assetCh" & ChannelID & "Mt" & mtID & "L" & l.ID), "").ToString()

                    ' graphic media types save independent of other channel assets
                    If Not IsGraphicMediaType(mtID) AndAlso MediaTypeID > 0 Then
                        Valid = IsValidAsset(MediaData, MediaTypeID, RetMsg)
                        If Not Valid Then Return False
                    End If

                Next
            Next

        End If

        Return Valid
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function IsValidAsset(ByVal MediaData As String, ByVal MediaTypeID As Integer, ByRef RetMsg As String) As Boolean
        Dim Valid As Boolean = True
        Dim dt As DataTable
        Dim MediaLen As Integer = 0
        Dim MinVal, MaxVal As Integer
        Dim IsNumericField As Boolean = False
        Dim IsRequiredField As Boolean = False

        If MediaData IsNot Nothing Then MediaLen = MediaData.Length

        ' Check the media type for restrictions on all media that uses text (i.e. UIFieldTypes of 4, 6, 10)
        Common.QueryStr = "select MPT.Name, MPT.PhraseTerm, MPT.MinValue, MPT.MaxValue, CMT.FieldTypeID, " &
                          "isnull(MPT.IsNumeric,0) As IsNumeric, isnull(MPT.IsRequired,0) As IsRequired from MediaParamTypes as MPT with (NoLock) " &
                          "inner join ChannelMediaParams as CMP with (NoLock) on CMP.ParamTypeID = MPT.ParamTypeID " &
                          "inner join ChannelMedia as CM with (NoLock) on CM.ChannelMediaID = CMP.ChannelMediaID " &
                          "inner join ChannelMediaTypes as CMT with (NoLock) on CMT.MediaTypeID = CM.MediaTypeID " &
                          "where CM.MediaTypeID = " & MediaTypeID & " and CMT.FieldTypeID in (4, 6, 10, 12, 13);"
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            For Each row As DataRow In dt.Rows
                MinVal = CInt(Common.NZ(dt.Rows(0).Item("MinValue"), Integer.MinValue))
                MaxVal = CInt(Common.NZ(dt.Rows(0).Item("MaxValue"), Integer.MaxValue))
                IsNumericField = Common.NZ(dt.Rows(0).Item("IsNumeric"), 0)
                IsRequiredField = Common.NZ(dt.Rows(0).Item("IsRequired"), 0)
                Valid = (MediaLen >= MinVal AndAlso MediaLen <= MaxVal)

                If Not Valid Then
                    RetMsg = Copient.PhraseLib.Detokenize("offer-channels.minMaxViolation", LanguageID, Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseTerm"), ""), LanguageID, Common.NZ(row.Item("Name"), "")), MediaLen, MinVal, MaxVal)
                    Exit For
                Else
                    If IsNumericField AndAlso Not IsNumeric(MediaData) AndAlso MediaData <> "" AndAlso MediaData <> "0" Then
                        Valid = False
                        RetMsg = Copient.PhraseLib.Detokenize("offer-channels.numericviolation", LanguageID, Common.NZ(row.Item("Name"), ""))
                        Exit For
                    Else
                        If IsRequiredField AndAlso (MediaData = "" OrElse MediaData Is Nothing OrElse MediaData = "0") Then
                            Valid = False
                            RetMsg = Copient.PhraseLib.Detokenize("offer-channels.requiredviolation", LanguageID, Common.NZ(row.Item("Name"), ""))
                            Exit For
                        End If
                    End If
                End If
            Next
        End If

        Return Valid
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Sub HandleSave(ByVal ChannelID As Integer)
        Dim Valid As Boolean = True
        Dim RetMsg As String = ""

        Valid = AreValidDates(ChannelID, RetMsg)
        Valid = Valid AndAlso AreValidAssets(ChannelID, RetMsg)

        If Valid Then
            SaveChannel(ChannelID)
        Else
            ' send back failure JSON response.
            Sendb("{")
            Sendb(" ""Status"": 0,")
            Sendb(" ""ChannelID"": " & ChannelID & ",")
            Sendb(" ""Message"": """ & RetMsg.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """,")
            Sendb(GetDeployMsgJSON(ChannelID, Offer.OfferID, False))
            Sendb("}")
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Sub DeployChannel(ByVal ChannelID As Integer)
        Dim Success As Boolean = False
        Dim RetMsg As String = ""

        If CanDeployChannel(ChannelID, RetMsg) Then
            ChannelCommon.DeployChannelForOffer(ChannelID, Offer.OfferID)
            RetMsg = Copient.PhraseLib.Lookup("term.deployed", LanguageID)
            Success = True
        End If

        Sendb("{")
        Sendb(" ""Status"": 1,")
        Sendb(" ""ChannelID"": " & ChannelID & ",")
        Sendb(" ""Message"": """ & RetMsg.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """,")
        Sendb(GetDeployMsgJSON(ChannelID, Offer.OfferID, Success))
        Sendb("}")
    End Sub

    Sub SaveChannelLockedState()
        If (GetCgiValue("IsTemplate") = True) Then
            ' time to update the status bits for the templates
            Dim form_Disallow_Channels As Integer = 0
            If (GetCgiValue("IsChannelsLocked") <> String.Empty) Then
                If (GetCgiValue("IsChannelsLocked") = "1") Then
                    form_Disallow_Channels = 1
                End If
                Common.QueryStr = "update TemplatePermissions with (RowLock) set Disallow_Channels=@Disallow_Channels" &
                                    " where OfferID=@OfferID"
                Common.DBParameters.Add("@Disallow_Channels", SqlDbType.Bit).Value = form_Disallow_Channels
                Common.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = Common.Extract_Val(GetCgiValue("OfferID"))
                Common.ExecuteNonQuery(Copient.DataBases.LogixRT)
            End If
        End If
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Function CanDeployChannel(ByVal ChannelID As Integer, ByRef RetMsg As String) As Boolean
        Dim Deployable As Boolean = False
        Dim dt As DataTable

        ' in order for a channel to deploy, the POS must first be deployed.  Check the incentive shadow table to see if this is the case.
        Select Case Offer.EngineID
            Case CommonInc.InstalledEngines.CM
                Common.QueryStr = "select OfferID from CM_ST_Offers with (NoLock) where OfferID = " & Offer.OfferID
                dt = Common.LRT_Select
                Deployable = (dt.Rows.Count > 0)
            Case CommonInc.InstalledEngines.CPE, CommonInc.InstalledEngines.CAM, CommonInc.InstalledEngines.UE
                Common.QueryStr = "select IncentiveID from CPE_ST_Incentives with (NoLock) where IncentiveID = " & Offer.OfferID
                dt = Common.LRT_Select
                Deployable = (dt.Rows.Count > 0)
            Case Else
                Common.QueryStr = "select OfferID from CM_ST_Offers with (NoLock) where OfferID = " & Offer.OfferID
                dt = Common.LRT_Select
                Deployable = (dt.Rows.Count > 0)
        End Select

        If Not Deployable Then
            RetMsg = Copient.PhraseLib.Lookup("offer-channels.posNotDeployedFirst", LanguageID)
        End If

        'Verify if the offer channel is present in ChannelOffers table which is precondition for deployment of channel
        Common.QueryStr = "select 1 from ChannelOffers with (NoLock) where ChannelId=@ChannelID and OfferID = @OfferID"
        Common.DBParameters.Add("@ChannelID", SqlDbType.Int).Value = ChannelID
        Common.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = Offer.OfferID
        dt = Common.ExecuteQuery(DataBases.LogixRT)
        Deployable = (dt.Rows.Count = 1)
        If Not Deployable Then
            RetMsg = Copient.PhraseLib.Lookup("offer-channels.NoChannelInformationSaved", LanguageID)
        End If

        Return Deployable
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Function GetMediaFormatID(ByVal MediaTypeID As Integer) As Integer
        Dim FormatID As Integer = 0
        Dim dt As DataTable

        Common.QueryStr = "select MediaFormatID from MediaTypeFormats with (NoLock) where MediaTypeId = " & MediaTypeID
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            FormatID = Common.NZ(dt.Rows(0).Item("MediaFormatID"), 0)
        End If

        Return FormatID
    End Function

    '-------------------------------------------------------------------------------------------------------------  

    Sub HandleSaveLegacy()

        Dim LegacyStatus As LegacyStatusRec
        Dim UseLegacy As Boolean = False
        Dim DisplayOnWebKiosk As Boolean = False
        Dim DisabledOnCFW As Boolean = False
        Dim Message As String = ""
        Dim HideChannels = False
        Dim ForceChannelRule = False

        LegacyStatus = Get_Legacy_Status(Offer.OfferID, Offer.EngineID)
        If GetCgiValue("enablechannels") = "false" Then UseLegacy = True
        If GetCgiValue("DisabledOnCFW") <> "" Then DisabledOnCFW = True
        If GetCgiValue("DisplayOnWebKiosk") <> "" Then DisplayOnWebKiosk = True
        If (UseLegacy <> LegacyStatus.UseLegacy) OrElse (DisabledOnCFW <> LegacyStatus.DisabledOnCFW) OrElse (DisplayOnWebKiosk <> LegacyStatus.DisplayOnWebKiosk) Then
            'there's been a changed in the selection between using channels and using legacy web/kiosk methods
            If UseLegacy Then
                'the user is wanting to use the legacy web/kiosk outputs - see if there is any channel content in place, if so, they can't switch away from channels
                If HasChannelContent() Then
                    'The user can't switch back to using legacy web/kiosk delivery because the offer has channel content.  
                    'If the user wants to switch back to legacy delivery, then they will need to delete all of the channel content first.
                    Message = Copient.PhraseLib.Lookup("channels.cannotswitchlegacy", LanguageID) 'Can not switch to legacy delivery<BR>while channel content exists
                    ForceChannelRule = True
                Else
                    'switch the offer to using legacy delivery
                    If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.UE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
                        Common.QueryStr = "Update CPE_Incentives set DisabledOnCFW=" & IIf(DisabledOnCFW, "1", "0") & ", DisplayOnWebKiosk=" & IIf(DisplayOnWebKiosk, "1", "0") & ", UseLegacyWebKiosk=1, StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
                    ElseIf Offer.EngineID = CommonInc.InstalledEngines.CM Then
                        Common.QueryStr = "Update Offers set DisabledOnCFW=" & IIf(DisabledOnCFW, "1", "0") & ", DisplayOnWebKiosk=" & IIf(DisplayOnWebKiosk, "1", "0") & ", UseLegacyWebKiosk=1, StatusFlag=1 where OfferId=" & Offer.OfferID & ";"
                    End If
                    Common.LRT_Execute()
                    Message = Copient.PhraseLib.Lookup("term.changessaved", LanguageID) 'Changes saved
                End If
            Else 'we are using Channels
                If Offer.EngineID = CommonInc.InstalledEngines.CPE OrElse Offer.EngineID = CommonInc.InstalledEngines.UE OrElse Offer.EngineID = CommonInc.InstalledEngines.CAM Then
                    Common.QueryStr = "Update CPE_Incentives set DisabledOnCFW=0, DisplayOnWebKiosk=0, UseLegacyWebKiosk=0, StatusFlag=1 where IncentiveID=" & Offer.OfferID & ";"
                ElseIf Offer.EngineID = CommonInc.InstalledEngines.CM Then
                    Common.QueryStr = "Update Offers set DisabledOnCFW=0, DisplayOnWebKiosk=0, UseLegacyWebKiosk=0, StatusFlag=1 where OfferId=" & Offer.OfferID & ";"
                End If
                Common.LRT_Execute()
                Message = Copient.PhraseLib.Lookup("term.changessaved", LanguageID) 'Changes saved
            End If 'UseLegacy

        End If

        If UseLegacy And HasChannelContent() = False Then
            HideChannels = True
        End If
        Sendb("{")
        Sendb(" ""Message"": """ & Message.Replace(ControlChars.Quote, "\" & ControlChars.Quote) & """, ")
        Sendb(" ""hidechannels"": """ & IIf(HideChannels, "1", "0") & """, ")
        Sendb(" ""forcechannelrule"": """ & IIf(ForceChannelRule, "1", "0") & """")
        Sendb("}")
    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Private Function HasChannelContent() As Boolean

        Dim contentFound As Boolean = False
        Dim dt As DataTable

        Common.QueryStr = "SELECT count(*) AS NumRecs FROM ChannelOfferAssets WHERE OfferID=" & Offer.OfferID & ";"
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).Item("NumRecs") > 0 Then contentFound = True
        End If
        If Not (contentFound) Then
            Common.QueryStr = "SELECT count(*) AS NumRecs FROM ChannelOffers WHERE OfferID=" & Offer.OfferID & ";"
            dt = Common.LRT_Select
            If dt.Rows.Count > 0 Then
                If dt.Rows(0).Item("NumRecs") > 0 Then contentFound = True
            End If
        End If

        Return contentFound

    End Function

    '-------------------------------------------------------------------------------------------------------------  

</script>
<%
    Common.AppName = "offer-channels.aspx"
    Response.Expires = 0
    On Error GoTo ErrorTrap

    If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
    ChannelCommon = New Copient.ChannelCommon(Common)
    BrokerChannel = New Copient.BrokerChannel(Common)

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    AdminUserID = Verify_AdminUser(Common, Logix)
    BannersEnabled = (Common.Fetch_SystemOption(66) = "1")

    GetFormValues()
    LoadAvailableLanguages()

    If GetCgiValue("mode") = "SaveChannelLockedState" Then
        SaveChannelLockedState()
    End If

    LoadOfferData()

    Select Case GetCgiValue("mode")
        Case "SaveChannel"
            ' save changes to the specified channel and return the status of the save
            HandleSave(Common.Extract_Val(GetCgiValue("ChannelID")))

        Case "SaveLegacy"
            ' save changes to the specified channel and return the status of the save
            HandleSaveLegacy()

        Case "RemoveChannel"
            RemoveChannel(Common.Extract_Val(GetCgiValue("ChannelID")))

        Case "RemoveAsset"
            RemoveAsset(Common.Extract_Val(GetCgiValue("ChannelID")), Common.Extract_Val(GetCgiValue("MediaTypeID")), _
                        Common.Extract_Val(GetCgiValue("LanguageID")))

        Case "RemovePMsg"
            RemovePrintedMessage(Common.Extract_Val(GetCgiValue("ChannelID")), MediaTypes.POS_Receipt_Message, _
                                 Common.Extract_Val(GetCgiValue("DeliverableID")), Common.Extract_Val(GetCgiValue("MessageID")))

        Case "RemoveImgUrl"
            RemovePosImgUrl(Common.Extract_Val(GetCgiValue("ChannelID")), MediaTypes.POS_Image_Url,
                                 Common.Extract_Val(GetCgiValue("DeliverableID")), Common.Extract_Val(GetCgiValue("PTPKID")))

        Case "RemoveAMsg"
            RemovePrintedMessage(Common.Extract_Val(GetCgiValue("ChannelID")), MediaTypes.POS_Accumulation_Message, _
                                 Common.Extract_Val(GetCgiValue("DeliverableID")), Common.Extract_Val(GetCgiValue("MessageID")))

        Case "RemoveCMsg"
            RemoveCashierMessage(Common.Extract_Val(GetCgiValue("ChannelID")), MediaTypes.POS_Cashier_Message, _
                                 Common.Extract_Val(GetCgiValue("DeliverableID")), Common.Extract_Val(GetCgiValue("MessageID")))

        Case "RemoveGraphic"
            RemoveGraphic(Common.Extract_Val(GetCgiValue("ChannelID")), MediaTypes.POS_Graphic, _
                          Common.Extract_Val(GetCgiValue("DeliverableID")), Common.Extract_Val(GetCgiValue("OnScreenAdID")))

        Case "DeployChannel"
            DeployChannel(Common.Extract_Val(GetCgiValue("ChannelID")))

        Case Else
            ' load up the page for display
            LoadChannels()
            LoadLimitTypes()
            Send_Page()
    End Select

    If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
    Common = Nothing
    Logix = Nothing

    Response.End()

ErrorTrap:
    Response.Write("<pre>" & Common.Error_Processor() & "</pre>")
    If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
    Common = Nothing
    Logix = Nothing

%>
<script src="../javascript/jquery-1.10.2.min.js"></script>
<input type="hidden" id="savedTime" name="savedTime" value="<%=DateTime.Now()%>" />
<script type="text/javascript">
    $(document).ready(function () {
        var savedTimeVal = document.getElementById('savedTime');
        var offerIDVal = document.getElementById('OfferID');
        if (savedTimeVal != null && offerIDVal != null) {
            var savedTime = new Date(savedTimeVal.value).getTime();
            var presentTime = new Date().getTime();
            var seconds = (presentTime - savedTime) / 1000;
            if (seconds > 2) {
                $.support.cors = true;
                $.ajax({
                    type: "POST",
                    url: "/Connectors/AjaxProcessingFunctions.asmx/GetLockedSystemOptions",
                    data: JSON.stringify({ offerID: offerIDVal.value }),
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (data) {
                        if (data.d == "true") {
                            var searchstr = "offer-channels.aspx";
                            var regEx = new RegExp(searchstr, "ig");
                            window.location.href = window.location.href.replace(regEx, "UE/UEOffer-sum.aspx");
                        }
                    },
                });
            }
        }
    });
</script>
