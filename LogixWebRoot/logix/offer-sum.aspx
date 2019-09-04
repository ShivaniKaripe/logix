<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-sum.aspx 
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
  Dim MyExport As New Copient.ExportXml(MyCommon)
  Dim MyImport As New Copient.ImportXml(MyCommon)

  Dim OfferID As Long
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim rst3 As DataTable
  Dim rst4 As DataTable
  Dim rowCount As Integer
  Dim IsTemplate As Boolean
  Dim FromTemplate As Boolean = False
  Dim IsExpired As Boolean
  Dim DeployDate As String = ""
  Dim DeployBtnOffsetIE As Integer = 0
  Dim DeployBtnOffsetFF As Integer = 0
  Dim SourceOfferID As Integer = 0
  Dim LinksDisabled As Boolean = False
  Dim ShowActionButton As Boolean = False
  Dim NewUpdateLevel As Integer = 0
  Dim GCount As Integer = 0
  Dim AnyProductUsed As Boolean = False
  Dim AnyCustomerUsed As Boolean = False
  Dim AnyStoreUsed As Boolean = False
  Dim LongDate As New DateTime
  Dim LongDateZeroTime As New DateTime
  Dim TodayDateZeroTime As New DateTime
  Dim DaysDiff As Integer = 0
  Dim i As Integer = 0
  Dim x As Integer = 0
  Dim TierAmtCount As Integer = 0
  Dim EngineId As Integer = 0
  Dim EngineSubTypeID As Integer = 0
  Dim Cookie As HttpCookie = Nothing
  Dim BoxesValue As String = ""
  Dim counter As Integer = 0
  Dim ValidateIncentiveColor As String = "green"
  Dim ComponentColor As String = "green"
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannerNames As String() = Nothing
  Dim BannerIDs As Integer() = Nothing
  Dim BannerCt As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim StatusMessage As String = ""
  Dim bExportToEDW As Boolean = False
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim Popup As Boolean = False
  Dim FolderNames As String = ""
  
  Dim objTemp As Object
  Dim intNumDecimalPlaces As Integer
  Dim decFactor As Decimal
  Dim decTemp As Decimal
  Dim sTemp1 As String
  Dim sTemp2 As String
  Dim bNeedToFormat As Boolean
  Dim bShowCRM As Boolean
  Dim iCRMType As Integer
  Dim InboundCRMEngineID As Integer
  Dim CRMEngineID As Integer
  Dim PrefManInstalled As Boolean = False
  Dim lProductionID As Long = 0
  Dim iWorkflowStatus As Integer = 0
  Dim bWorkflowActive As Boolean = False
  Dim bUseTestDates As Boolean
  Dim bUseDisplayDates As Boolean = False
  Dim bProductionSystem As Boolean = True
  Dim bTestSystem As Boolean = False
  Dim bArchiveSystem As Boolean = False
  Dim bBuckParentOffer As Boolean
  Dim bBuckParentToBe As Boolean
  Dim bBuckChildOffer As Boolean
  Dim bEnableBuckOffers As Boolean
  Dim bUseOfferRedemptionThreshold As Boolean = False
  Dim iTierTypeId As Integer = 0
  Dim bStatus As Boolean
  Dim oBuckStatus As Copient.ImportXml.BuckOfferStatus
  
  CurrentRequest.Resolver.AppName = "offer-sum.aspx"
  Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()
  Dim m_customerGroup As ICustomerGroups = CurrentRequest.Resolver.Resolve(Of ICustomerGroups)()
  Dim bShelfLabelEnabled As Boolean = False
  Dim bCopyInboundCrmEngineID As Boolean = True
  Dim m_Logger As ILogger = CurrentRequest.Resolver.Resolve(Of ILogger)()
  Dim offerValidationLogFilePrefix As String = "OfferValidationLog"
    
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-sum.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  bUseTestDates = (MyCommon.Fetch_SystemOption(88) = "1")
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)

  bUseOfferRedemptionThreshold = IIf(MyCommon.Fetch_CM_SystemOption(83) = "1", True, False)
  
  If MyCommon.Fetch_CM_SystemOption(85) = "1" Then
    bUseDisplayDates = True
  Else
    bUseDisplayDates = False
  End If

  If MyCommon.Fetch_CM_SystemOption(105) = "1" Then
    bShelfLabelEnabled = True
  Else
    bShelfLabelEnabled = False
  End If

  If MyCommon.Fetch_CM_SystemOption(107) = "1" Then
    bCopyInboundCrmEngineID = True
  Else
    bCopyInboundCrmEngineID = False
  End If

  PrefManInstalled = MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER)
  If PrefManInstalled Then MyCommon.Open_PrefManRT()
  
  objTemp = MyCommon.Fetch_CM_SystemOption(41)
  If Not (Integer.TryParse(objTemp.ToString, intNumDecimalPlaces)) Then
    intNumDecimalPlaces = 0
  End If
  decFactor = (10 ^ intNumDecimalPlaces)
  
  OfferID = Request.QueryString("OfferID")
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-gen.aspx?new=New")
    GoTo done
  End If
  
  bEnableBuckOffers = (MyCommon.Fetch_CM_SystemOption(137) = "1")
  If bEnableBuckOffers Then
    oBuckStatus = MyImport.BuckOfferGetStatus(OfferID)
    Select Case oBuckStatus
      Case Copient.ImportXml.BuckOfferStatus.BuckParentNoChildren,
        Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren,
        Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly,
        Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
        bBuckParentOffer = True
        bBuckChildOffer = False
        bBuckParentToBe = False
      Case Copient.ImportXml.BuckOfferStatus.BuckChildPaper,
       Copient.ImportXml.BuckOfferStatus.BuckChildDigital
        bBuckParentOffer = False
        bBuckChildOffer = True
        bBuckParentToBe = False
      Case Copient.ImportXml.BuckOfferStatus.BuckTiered
        bBuckParentOffer = False
        bBuckChildOffer = False
        bBuckParentToBe = True
      Case Else
        bBuckParentOffer = False
        bBuckChildOffer = False
        bBuckParentToBe = False
        If oBuckStatus = Copient.ImportXml.BuckOfferStatus.ErrorOccurred Then
          infoMessage = MyImport.GetErrorMsg()
        End If
    End Select
  Else
    bBuckParentOffer = False
    bBuckChildOffer = False
    bBuckParentToBe = False
    oBuckStatus = Copient.ImportXml.BuckOfferStatus.NotBuckOffer
  End If
 
  StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
  If Not Integer.TryParse(MyCommon.Fetch_SystemOption(25), iCRMType) Then iCRMType = 0
  
  ' load up all the folder names to which this offer is assigned.
  MyCommon.QueryStr = "select distinct FI.FolderID, F.FolderName from FolderItems as FI with (NoLock) " & _
                      "inner join Folders as F with (NoLock) on F.FolderID = FI.FolderID " & _
                      "where LinkID=" & OfferID & " and LinkTypeID=1;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    For Each row In rst.Rows
      If FolderNames <> "" Then FolderNames &= " <br />"
      FolderNames &= "<a href=""javascript:openPopup('/logix/folder-browse.aspx?Action=NavigateToFolder&OfferID=" & OfferID & _
                     "&FolderID=" & MyCommon.NZ(row.Item("FolderID"), "0") & "');"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("FolderName"), ""), 25) & "</a>"
    Next
  Else
    FolderNames = "<a href=""javascript:openPopup('/logix/folder-browse.aspx?OfferID=" & OfferID & "');"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</a>"
  End If
  
  bTestSystem = (MyCommon.Fetch_CM_SystemOption(77) = "1")
  bArchiveSystem = (MyCommon.Fetch_CM_SystemOption(77) = "2")
  If bTestSystem Or bArchiveSystem Then
    bProductionSystem = False
  Else
    bProductionSystem = True
  End If
  bWorkflowActive = (MyCommon.Fetch_CM_SystemOption(74) = "1")
    
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("offer-new.aspx")
  ElseIf (Request.QueryString("OfferFromTemp") <> "") Then
    ' dbo.pc_CreateOfferFromTemplate @TemplateID bigint, @OfferID bigint OUTPUT
    MyCommon.QueryStr = "dbo.pc_Create_CM_OfferFromTemplate"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@TemplateID", SqlDbType.NVarChar, 200).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@CopyInboundCRM", SqlDbType.Bit).Value = IIf(bCopyInboundCrmEngineID, 1, 0)
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    SourceOfferID = OfferID
    OfferID = MyCommon.LRTsp.Parameters("@OfferID").Value
    MyCommon.Close_LRTsp()
   
    CreateNewLocalPromotionVariables(OfferID, MyCommon)
    If bUseDisplayDates Then
      'Updating TemplatePermission table with the Disallow_DisplayDates based on the SystemOption #85
      UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 85)
      SaveOfferDisplayDates(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineId)
    End If
     
    If (bUseOfferRedemptionThreshold) Then
            
            'Updating TemplatePermission table with the Disallow_OfferRedempThreshold based on the SystemOption #83
            UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 83)

            SaveOfferThresholdPerHourValue(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineId)
     End If
    
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("offer.createdfromtemplate", LanguageID) & ": " & SourceOfferID)
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-gen.aspx?OfferID=" & OfferID)
    GoTo done
  ElseIf (Request.QueryString("saveastemp") <> "") Then
    ' dbo.pc_CreateTemplateFromOffer @OfferID bigint, @TemplateID bigint OUTPUT
    MyCommon.QueryStr = "dbo.pc_Create_CM_TemplateFromOffer"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.NVarChar, 200).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@CopyInboundCRM", SqlDbType.Bit).Value = IIf(bCopyInboundCrmEngineID, 1, 0)
    MyCommon.LRTsp.Parameters.Add("@TemplateID", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    SourceOfferID = OfferID
    OfferID = MyCommon.LRTsp.Parameters("@TemplateID").Value
    MyCommon.Close_LRTsp()

    CreateNewLocalPromotionVariables(OfferID, MyCommon)
    If bUseDisplayDates Then
      'Updating TemplatePermission table with the Disallow_DisplayDates based on the SystemOption #85
      UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 85)
      SaveOfferDisplayDates(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineId)
    End If
       
    If (bUseOfferRedemptionThreshold) Then
            'Updating TemplatePermission table with the Disallow_OfferRedempThreshold based on the SystemOption #83
            UpdateTemplatePermissions(MyCommon, SourceOfferID, OfferID, 83)
            SaveOfferThresholdPerHourValue(MyCommon, SourceOfferID, OfferID, AdminUserID, EngineId)
    End If
        
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("templates.createdfromoffer", LanguageID) & ": " & SourceOfferID)
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-gen.aspx?OfferID=" & OfferID)
    GoTo done
  ElseIf (Request.QueryString("deploy") <> "") Then
    Dim bImmediateDeploy As Boolean = False
    Dim bAutodetermineDeploy As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(111) = "1", True, False)
    Dim CRMSendStatus As String = ""
    
    bStatus = MyExport.ValidateOfferForDeploy(OfferID, LanguageID, False, False)
    If bAutodetermineDeploy Then
      Dim OfferStartDate As Date
      objTemp = MyCommon.Fetch_CM_SystemOption(112)
      Dim NoDaysInAdvDeferDeployment As Integer
      If Not (Integer.TryParse(objTemp.ToString, NoDaysInAdvDeferDeployment)) Then
        NoDaysInAdvDeferDeployment = 0
      End If
      MyCommon.QueryStr = "select isnull(ProdStartDate,0) as ProdStartDate, CRMEngineID from Offers with (NoLock) where Deleted=0 and OfferID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        If Not IsDBNull(rst.Rows(0).Item("ProdStartDate")) Then
          Date.TryParse(rst.Rows(0).Item("ProdStartDate"), OfferStartDate)
          Dim DeployableDate As Date = Now().AddDays(NoDaysInAdvDeferDeployment)
          If OfferStartDate <= DeployableDate Then
            bImmediateDeploy = True 'if Offer start date less than or equal to (current date + NoDaysInAdvDeferDeployment) then, deploy it immediately.
          End If
        End If
        CRMEngineID = MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), 0)
      End If
      
      If (CRMEngineID > 2) Then
        CRMSendStatus = ", CRMSendStatus=0 "
      End If
    If bStatus Then
        If bImmediateDeploy Then
          MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1 " & CRMSendStatus & " where OfferID=" & OfferID
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
        Else
          'defer deployment		
          MyCommon.QueryStr = "update Offers with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1 " & CRMSendStatus & " where OfferID=" & OfferID
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deferdeploy", LanguageID))
        End If
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
        SetLastDeployValidationMessage(MyCommon, OfferID, Copient.PhraseLib.Lookup("term.validationsuccessful", LanguageID))
        m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.CM, OfferID, Copient.PhraseLib.Lookup("term.validationsuccessful", LanguageID)), offerValidationLogFilePrefix)
        GoTo done
      Else
        infoMessage = MyExport.GetErrorMsg
        SetLastDeployValidationMessage(MyCommon, OfferID, "<font color=""red"">" & infoMessage & "</font>")
        m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.CM, OfferID, infoMessage), offerValidationLogFilePrefix)
      End If
    Else
      If bStatus Then

        MyCommon.QueryStr = "select CRMEngineID from Offers with (NoLock) where Deleted=0 and OfferID=" & OfferID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          CRMEngineID = MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), 0)
        End If
      
        If (CRMEngineID > 2) Then
          CRMSendStatus = ", CRMSendStatus=0 "
        End If

        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1 " & CRMSendStatus & " where OfferID=" & OfferID
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
      SetLastDeployValidationMessage(MyCommon, OfferID, Copient.PhraseLib.Lookup("term.validationsuccessful", LanguageID))
      m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.CM, OfferID, Copient.PhraseLib.Lookup("term.validationsuccessful", LanguageID)), offerValidationLogFilePrefix)
      GoTo done
    Else
      infoMessage = MyExport.GetErrorMsg
      SetLastDeployValidationMessage(MyCommon, OfferID, "<font color=""red"">" & infoMessage & "</font>")
      m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.CM, OfferID, infoMessage), offerValidationLogFilePrefix)
    End If
    End If
  ElseIf (Request.QueryString("sendoutbound") <> "") Then
    Dim iExtInterfaceTypeID As Integer = -1

    MyCommon.QueryStr = "select CRMEngineID from Offers with (NoLock) where Deleted=0 and OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      If Not Integer.TryParse(rst.Rows(0).Item(0), CRMEngineID) Then CRMEngineID = 0
    End If
    If CRMEngineID > 0 Then
      MyCommon.QueryStr = "select ExtInterfaceTypeID from ExtCRMInterfaces with (NoLock) where Deleted=0 and Active=1 and ExtInterfaceID=" & CRMEngineID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        If Not Integer.TryParse(rst.Rows(0).Item(0), iExtInterfaceTypeID) Then iExtInterfaceTypeID = -1
      End If
    End If
    Select Case iExtInterfaceTypeID
      Case 0
        ' CRM
        MyCommon.QueryStr = "update Offers with (RowLock) set LastCRMSendDate=getdate(), CRMEngineUpdateLevel=CRMEngineUpdateLevel+1,CRMSendStatus=1,CRMSendToExport=1 where OfferID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-sendtocrm", LanguageID))
      Case 1
        ' Old Teradata CRM
        MyCommon.QueryStr = "update Offers with (RowLock) set LastCRMSendDate=getdate(), CRMEngineUpdateLevel=CRMEngineUpdateLevel+1,CRMSendStatus=1 where OfferID=" & OfferID
        MyCommon.LRT_Execute()
    
        ' create an entry, if necessary, for use in TCRM agent processing 
        MyCommon.QueryStr = "select LinkID from CRMEngineUpdateLevels with (NoLock) where EngineID=1 and ItemType=1 and LinkID=" & OfferID
        rst = MyCommon.LRT_Select
        If rst.Rows.Count = 0 Then
          MyCommon.QueryStr = "insert into CRMEngineUpdateLevels with (RowLock) (EngineID, LinkID, ItemType, LastUpdateLevel, LastUpdate) " & _
                              "  values (1, " & OfferID & ",1,0,getdate());"
          MyCommon.LRT_Execute()
        End If
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-sendtocrm", LanguageID))
      Case Else
        infoMessage = Copient.PhraseLib.Lookup("term.invalidoutbound", LanguageID) & ": " & CRMEngineID
    End Select
  ElseIf (Request.QueryString("delete") <> "") Then
    bStatus = True
    If bEnableBuckOffers Then
      If bBuckParentOffer Then
        Select Case oBuckStatus
          Case Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren
            bStatus = MyImport.BuckChildOffersDelete(OfferID, 0, AdminUserID, LanguageID)
            If Not bStatus Then
              infoMessage = MyImport.GetErrorMsg
            End If
          Case Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly
            bStatus = MyImport.BuckChildOffersDelete(OfferID, 1, AdminUserID, LanguageID)
            If Not bStatus Then
              infoMessage = MyImport.GetErrorMsg
            End If
          Case Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
            bStatus = MyImport.BuckChildOffersDelete(OfferID, 2, AdminUserID, LanguageID)
            If Not bStatus Then
              infoMessage = MyImport.GetErrorMsg
            End If
        End Select
        If infoMessage = "" Then
          bStatus = MyImport.BuckParentOfferDelete(OfferID, AdminUserID, LanguageID)
          If Not bStatus Then
            infoMessage = MyImport.GetErrorMsg
          End If
        End If
      End If
    End If ' buck offers enabled
    
    If bStatus Then
      ' normal offer delete path
    Dim optInGroup As CustomerGroup = m_Offer.GetOfferDefaultCustomerGroup(OfferID, EngineId)
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,Deleted=1,UpdateLevel=UpdateLevel+1 where OfferID=" & OfferID
    MyCommon.LRT_Execute()
    
        'Mark Client ID deleted if this is enternal offer.
        MyCommon.QueryStr = "dbo.pt_ExtOfferID_Delete"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.NVarChar, 20).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineId
        MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()
    
    'Remove any Offer Eligibility Conditions associated with the Offer.    
    m_Offer.DeleteOfferEligibleConditions(OfferID, EngineId)
    If (optInGroup IsNot Nothing) Then
      m_customerGroup.DeleteCustomerGroup(optInGroup.CustomerGroupID)
    End If

    'remove the banners assigned to this offer
    If (MyCommon.Fetch_SystemOption(66) = "1") Then
      MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = " & OfferID
      MyCommon.LRT_Execute()
    End If

    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-delete", LanguageID))
    Response.Status = "301 Moved Permanently"
    MyCommon.QueryStr = "select IsTemplate from Offers with (NoLock) where OfferID=" & OfferID
    rst = MyCommon.LRT_Select()
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    If (IsTemplate) Then
      Response.AddHeader("Location", "temp-list.aspx")
    Else
      Response.AddHeader("Location", "offer-list.aspx")
    End If
    GoTo done
    End If
  ElseIf (Request.QueryString("delete_paper") <> "") Then
    If bEnableBuckOffers Then
      bStatus = MyImport.BuckChildOffersDelete(OfferID, 1, AdminUserID, LanguageID)
      If Not bStatus Then
        infoMessage = MyImport.GetErrorMsg
      End If
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
      GoTo done
    End If ' allow buck tiers
  ElseIf (Request.QueryString("delete_digital") <> "") Then
    If bEnableBuckOffers Then
      bStatus = MyImport.BuckChildOffersDelete(OfferID, 2, AdminUserID, LanguageID)
      If Not bStatus Then
        infoMessage = MyImport.GetErrorMsg
      End If
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
      GoTo done
    End If ' allow buck tiers
  ElseIf (Request.QueryString("export") <> "") Or (Request.QueryString("exportCME") <> "") Or (Request.QueryString("exportCRM") <> "") Then
    Dim sFileFullPathName As String
    sFileFullPathName = MyCommon.Fetch_SystemOption(29) & "\Offer" & Request.QueryString("OfferID").ToString & ".xml"
    If (Request.QueryString("export") <> "") Then
            bStatus = MyExport.ExportOfferLoad(Request.QueryString("OfferID").ToString, sFileFullPathName, False, True)
    Else
      If (Request.QueryString("exportCRM") <> "") Then
        bStatus = MyExport.ExportOfferCrm(Request.QueryString("OfferID").ToString, sFileFullPathName)
      Else
        bStatus = MyExport.GenOfferXml(Request.QueryString("OfferID").ToString, sFileFullPathName, True, False)
      End If
    End If
    If Not bStatus Then
      infoMessage = MyExport.GetStatusMsg
      If infoMessage = "" Then
        infoMessage = MyExport.GetErrorMsg
      End If
    Else
      If (MyExport.GetFileType = Copient.ExportXmlCPE.FileTypeEnum.XML_FORMAT) Then
        Dim oRead As System.IO.StreamReader
        Dim LineIn As String
        Dim Bom As String = ChrW(65279)
        oRead = System.IO.File.OpenText(sFileFullPathName)
        Response.ContentEncoding = Encoding.Unicode
        Response.ContentType = "application/octet-stream"
        Response.AddHeader("Content-Disposition", "attachment; filename=" & "Offer" & Request.QueryString("OfferID").ToString & ".xml")
        
        'force little endian fffe bytes at front, why?  i dont know but is required.
        Sendb(Bom)
        While oRead.Peek <> -1
          LineIn = oRead.ReadLine()
          Send(LineIn)
        End While
        oRead.Close()
      ElseIf (MyExport.GetFileType = Copient.ExportXmlCPE.FileTypeEnum.GZ_FORMAT) Then
        Dim fs As System.IO.FileStream
        Dim br As System.IO.BinaryReader
        sFileFullPathName += ".gz"
        Response.ContentType = "application/octet-stream"
        Response.AddHeader("Content-Disposition", "attachment; filename=" & "Offer" & Request.QueryString("OfferID").ToString & ".xml.gz")
        fs = System.IO.File.Open(sFileFullPathName, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
        br = New System.IO.BinaryReader(fs)
        Response.BinaryWrite(br.ReadBytes(fs.Length))
        fs.Flush()
        br.Close()
        fs.Close()
        fs = Nothing
      End If
      System.IO.File.Delete(sFileFullPathName)
      MyCommon.Activity_Log(3, Request.QueryString("OfferID"), AdminUserID, Copient.PhraseLib.Lookup("offer.exported", LanguageID))
      GoTo done
    End If
  ElseIf (Request.QueryString("exportedw") <> "") Then
    Dim sFileFullPathName As String
    sFileFullPathName = MyCommon.Fetch_SystemOption(73).Trim
    If sFileFullPathName.Length > 0 Then
      If Right(sFileFullPathName, 1) <> "\" Then
        sFileFullPathName = sFileFullPathName & "\"
      End If
      sFileFullPathName = sFileFullPathName & "Offer" & Request.QueryString("OfferID") & "_" & Now.ToString("yyyy-MM-dd_HHmmss") & ".xml"
      bStatus = MyExport.ExportOfferLoad(Request.QueryString("OfferID").ToString, sFileFullPathName, True)
      If Not bStatus Then
        infoMessage = MyExport.GetStatusMsg
      Else
        StatusMessage = Copient.PhraseLib.Lookup("cpeoffer-sum.exportedwok", LanguageID)

        MyCommon.QueryStr = "dbo.pa_UpdateOfferFeeds"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = Request.QueryString("OfferID")
        MyCommon.LRTsp.Parameters.Add("@LastFeed", SqlDbType.DateTime).Value = Now()
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      End If
    Else
      infoMessage = Copient.PhraseLib.Lookup("term.invalidfilepath", LanguageID)
    End If
  ElseIf (Request.QueryString("deferdeploy") <> "") Then
    Dim CRMSendStatus As String = ""

    bStatus = MyExport.ValidateOfferForDeploy(OfferID, LanguageID, False, False)
    If bStatus Then
      MyCommon.QueryStr = "select CRMEngineID from Offers with (NoLock) where Deleted=0 and OfferID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        CRMEngineID = MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), 0)
      End If
      
      If (CRMEngineID > 2) Then
        CRMSendStatus = ", CRMSendStatus=0 "
      End If

      MyCommon.QueryStr = "update Offers with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1 " & CRMSendStatus & " where OfferID=" & OfferID
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deferdeploy", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
      GoTo done
    Else
      infoMessage = MyExport.GetErrorMsg
    End If
  ElseIf (Request.QueryString("canceldeploy") <> "") Then
    ' check if the offer is still in awaiting deployment status
    MyCommon.QueryStr = "select StatusFlag, DeployDeferred, EngineID from Offers with (NoLock) where OfferID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      ' update status to modified (1) if offer is still awaiting deployment, otherwise alert user that offer was already deployed.
      If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 2) Or (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = True) Then
        EngineId = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
        MyCommon.QueryStr = "select LastUpdateLevel from PromoEngineUpdateLevels with (NoLock) " & _
                            "where LinkID=" & OfferID & " and EngineID=" & EngineId & " and ItemType=1;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          NewUpdateLevel = MyCommon.NZ(rst.Rows(0).Item("LastUpdateLevel"), 0)
        End If
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", DeployDeferred=0, UpdateLevel=" & NewUpdateLevel & " where OfferID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.canceldeploy", LanguageID))
        If bProductionSystem And bWorkflowActive Then
          ' set Workflow Status back to Ready to Deploy, since the previous update trigger reset it.
          MyCommon.QueryStr = "update Offers with (RowLock) set WorkflowStatus=3 where OfferID=" & OfferID
          MyCommon.LRT_Execute()
        End If
        infoMessage = Copient.PhraseLib.Lookup("term.deploymentcanceled", LanguageID)
      ElseIf (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 0) Then
        infoMessage = Copient.PhraseLib.Lookup("message.alreadydeployed", LanguageID)
      Else
        infoMessage = Copient.PhraseLib.Lookup("message.unablecanceldeployment", LanguageID)
      End If
    End If
  ElseIf (Request.QueryString("copyoffer") <> "") Then
    MyCommon.QueryStr = "dbo.pc_Copy_CM_Offer"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.BigInt).Value = Request.QueryString("OfferID")
    MyCommon.LRTsp.Parameters.Add("@CopyInboundCRM", SqlDbType.Bit).Value = IIf(bCopyInboundCrmEngineID, 1, 0)
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    OfferID = MyCommon.LRTsp.Parameters("@OfferID").Value
    MyCommon.Close_LRTsp()
    
    If (OfferID > 0) Then
      CreateNewLocalPromotionVariables(OfferID, MyCommon)
      If bUseDisplayDates Then
        'Updating TemplatePermission table with the Disallow_DisplayDates based on the SystemOption #85
        UpdateTemplatePermissions(MyCommon, Request.QueryString("OfferID"), OfferID, 85)
        SaveOfferDisplayDates(MyCommon, Request.QueryString("OfferID"), OfferID, AdminUserID, EngineId)
      End If
           
     If (bUseOfferRedemptionThreshold) Then
                'Updating TemplatePermission table with the Disallow_OfferRedempThreshold based on the SystemOption #83
                UpdateTemplatePermissions(MyCommon, Request.QueryString("OfferID"), OfferID, 83)
                SaveOfferThresholdPerHourValue(MyCommon, Request.QueryString("OfferID"), OfferID, AdminUserID, EngineId)
      End If
            
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-copy", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
      GoTo done
    Else
      OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    End If
  ElseIf (Request.QueryString("preValidate") <> "") Then
    MyCommon.QueryStr = "update Offers with (RowLock) set WorkflowStatus=1 where OfferID=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.prevalidate", LanguageID))
    infoMessage = MyExport.SendWorkflowOutbound(OfferID, 1, AdminUserID, LanguageID)
    If StatusCode <> Copient.LogixInc.STATUS_FLAGS.STATUS_DEVELOPMENT AndAlso StatusCode <> Copient.LogixInc.STATUS_FLAGS.STATUS_SCHEDULED Then
      If StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_TESTING Then
        If bUseTestDates Then
          infoMessage = Copient.PhraseLib.Lookup("term.revalidationrequired", LanguageID)
        End If
      Else
        infoMessage = Copient.PhraseLib.Lookup("term.revalidationrequired", LanguageID)
      End If
    End If
  ElseIf (Request.QueryString("postValidate") <> "") Then
    bStatus = MyExport.ValidateOfferForDeploy(OfferID, LanguageID, False, True)
    If bStatus Then
      bStatus = MyExport.TransferOfferToTest(OfferID)
      If bStatus Then
        MyCommon.QueryStr = "update Offers with (RowLock) set WorkflowStatus=2 where OfferID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.postvalidate", LanguageID))
        infoMessage = MyExport.SendWorkflowOutbound(OfferID, 2, AdminUserID, LanguageID)
      Else
        infoMessage = MyExport.GetErrorMsg
      End If
    Else
      infoMessage = MyExport.GetErrorMsg
    End If
  ElseIf (Request.QueryString("readyToDeploy") <> "") Then
    MyCommon.QueryStr = "update Offers with (RowLock) set WorkflowStatus=3 where OfferID=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("term.readytodeploy", LanguageID))
    infoMessage = MyExport.SendWorkflowOutbound(OfferID, 3, AdminUserID, LanguageID)
  ElseIf (Request.QueryString("create_paper") <> "") Then
    bStatus = MyImport.BuckChildOffersCreate(OfferID, 1, AdminUserID, LanguageID)
    If Not bStatus Then
      infoMessage = MyImport.GetErrorMsg
    End If
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
    GoTo done
  ElseIf (Request.QueryString("create_digital") <> "") Then
    bStatus = MyImport.BuckChildOffersCreate(OfferID, 2, AdminUserID, LanguageID)
    If Not bStatus Then
      infoMessage = MyImport.GetErrorMsg
    End If
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-sum.aspx?OfferID=" & OfferID)
    GoTo done
  End If
  
  MyCommon.QueryStr = "select OfferID,IsTemplate,FromTemplate,ExtOfferID,Name,o.Description,o.OfferCategoryID, " & _
                      "CG.description as cgDescription,CG.OfferCategoryID,OfferTypeID,isnull(ProdStartDate,0) as ProdStartDate, " & _
                      "isnull(ProdEndDate,0) as ProdEndDate,TestStartDate,TestEndDate,TierTypeID,NumTiers, " & _
                      "isnull(DistPeriod,0) as DistPeriod,isnull(DistPeriodLimit,0) as DistPeriodLimit,ExportToEDW, " & _
                      "DistPeriodVarID,EmployeeFiltering,NonEmployeesOnly,CRMRestricted,CRMEngineID,O.LastUpdate, " & _
                      "StatusFlag,PriorityLevel,O.EngineID,CMOADeployStatus,CMOADeployRpt,CMOARptDate, " & _
                      "CMOADeploySuccessDate,PE.Description as eDescription,DeployDeferred, " & _
                      "AU1.FirstName + ' ' + AU1.LastName as CreatedBy, AU2.FirstName + ' ' + AU2.LastName as LastUpdatedBy, " & _
                      "InboundCRMEngineID, WorkflowStatus, ProductionID " & _
                      "from Offers as O with (nolock) " & _
                      "left join OfferCategories as CG with (NoLock) on CG.offerCategoryID=O.OffercategoryID " & _
                      "left join PromoEngines as PE with (NoLock) on PE.EngineID=O.EngineID " & _
                      "left join AdminUsers as AU1 with (NoLock) on AU1.AdminUserID = O.CreatedByAdminID " & _
                      "left join AdminUsers as AU2 with (NoLock) on AU2.AdminUserID = O.LastUpdatedByAdminID " & _
                      "where OfferID=" & OfferID & " and O.Deleted=0;"
  rst = MyCommon.LRT_Select()
  If rst.Rows.Count > 0 Then
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    EngineId = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    LinksDisabled = IIf(MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 2, True, False)
    If (Not LinksDisabled) Then
      LinksDisabled = (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = True)
    End If
    If Popup Then
      LinksDisabled = True
    End If
    InboundCRMEngineID = MyCommon.NZ(rst.Rows(0).Item("InboundCRMEngineID"), 0)
    CRMEngineID = MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), 0)
    iWorkflowStatus = MyCommon.NZ(rst.Rows(0).Item("WorkflowStatus"), 0)
    lProductionID = MyCommon.NZ(rst.Rows(0).Item("ProductionID"), 0)
    
    ' Use shadow table for Archive export
    MyCommon.QueryStr = "select ExportToEDW from CM_ST_Offers with (nolock) where OfferID=" & OfferID & " and Deleted=0"
    rst2 = MyCommon.LRT_Select()
    If rst2.Rows.Count > 0 Then
      bExportToEDW = MyCommon.NZ(rst2.Rows(0).Item("ExportToEDW"), False)
    End If
    
  Else
    infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
  End If

  ' get the banner assigned to this offer
  If (MyCommon.Fetch_SystemOption(66) = "1") Then
    MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from BannerOffers BO with (NoLock) " & _
                          "inner join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                          "where BO.OfferID = " & Request.QueryString("OfferID")
    rst2 = MyCommon.LRT_Select
    BannerCt = rst2.Rows.Count
    If (BannerCt > 0) Then
      ReDim BannerNames(BannerCt - 1)
      ReDim BannerIDs(BannerCt - 1)
      For i = 0 To BannerNames.GetUpperBound(0)
        BannerNames(i) = MyCommon.SplitNonSpacedString(MyCommon.NZ(rst2.Rows(i).Item("Name"), ""), 25)
        BannerIDs(i) = MyCommon.NZ(rst2.Rows(i).Item("BannerID"), -1)
      Next
    Else
      ReDim BannerNames(0)
      ReDim BannerIDs(0)
      BannerNames(0) = Copient.PhraseLib.Lookup("term.all", LanguageID)
      BannerIDs(i) = -1
    End If
  End If
  
  If iCRMType > 0 And CRMEngineID <> 0 Then
    bShowCRM = True
  Else
    bShowCRM = False
  End If

  Send_HeadBegin("term.offer", "term.summary", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  If Popup Then
    If (IsTemplate) Then
      Send_BodyBegin(13)
    Else
      Send_BodyBegin(3)
    End If
  Else
    If (IsTemplate) Then
      Send_BodyBegin(11)
    Else
      Send_BodyBegin(1)
    End If
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 2)
    If (rst.Rows.Count = 0) Then
      Send_Subtabs(Logix, 23, 3, , OfferID)
    Else
      If (IsTemplate) Then
        Send_Subtabs(Logix, 22, 3, , OfferID)
      Else
        Send_Subtabs(Logix, 21, 3, , OfferID)
      End If
    End If
  End If
  
  ' Get the user's preference from the cookie for collapsing/showing the boxes
  Cookie = Request.Cookies("BoxesCollapsed")
  If Not (Cookie Is Nothing) Then
    BoxesValue = Cookie.Value
    If (BoxesValue Is Nothing OrElse BoxesValue.Trim = "") Then
      BoxesValue = "0"
    End If
  Else
    BoxesValue = "0"
  End If
  
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "perm.offers-access")
    Send_BodyEnd()
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "perm.offers-access-templates")
    Send_BodyEnd()
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
  
  If (rst.Rows.Count < 1) Then
    Send("")
    Send("<div id=""intro"">")
    Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & Request.QueryString("OfferID") & "</h1>")
    Send("</div>")
    Send("<div id=""main"">")
    Send("    <div id=""infobar"" class=""red-background"">")
    Send("        " & infoMessage)
    Send("    </div>")
    Send("</div>")
    Send("</div>")
    Send("</body>")
    Send("</html>")
    GoTo done
  End If
  
  If (MyCommon.NZ(rst.Rows(0).Item("prodEndDate"), Now().AddDays(-1)) < Today) Then
    IsExpired = True
  End If
%>
<script type="text/javascript">
    function LoadDocument(url) { 
        location = url; 
    }
    
    var divElems = new Array("generalbody", "periodbody", "limitsbody","deploymentbody", 
                             "locationbody", "notificationbody", "conditionbody", "rewardbody", "validationbody","displaybody", "shelflabelbody");
    var divVals  = new Array(1, 2, 4, 8, 16, 32, 64, 128, 256);
    var divImages = new Array("imgGeneral", "imgPeriod", "imgLimits", "imgDeployment",  
                              "imgLocations", "imgNotifications", "imgOptInConditions", "imgConditions", "imgRewards", "imgValidation","imgDisplay", "imgShelfLabel");
    var boxesValue = <% Sendb(BoxesValue) %>;
    
    function updateCookie() {
        updateBoxesCookie(divElems, divVals);
    }
    
    function collapseBoxes() {
        updatePageBoxes(divElems, divVals, divImages, boxesValue);
    }

    function showDiv(elemName) {
        var elem = document.getElementById(elemName);
        
        if (elem != null) {
            elem.style.display = (elem.style.display == "none") ? "block" : "none";
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
	
function toggleDialog(elemName, shown) {
    var elem = document.getElementById(elemName);
    var fadeElem = document.getElementById('OfferfadeDiv');
    
    if (elem != null) {
      elem.style.display = (shown) ? 'block' : 'none';
    }
    
    if (fadeElem != null) {
      fadeElem.style.display = (shown) ? 'block' : 'none';
    }
    
  }
    function assignNoofDuplicateOffers(shown) {
	 var elem = document.getElementById('DuplicateNoofOffer');
      var fadeElem = document.getElementById('OfferfadeDiv');
      if (elem != null) {
        elem.style.display = (shown) ? 'block' : 'none';
      }
      if (fadeElem != null) {
        fadeElem.style.display = (shown) ? 'block' : 'none';
      }
	  if (shown)  {
	   document.getElementById('txtDuplicateOffersCnt').value='1';
	   document.getElementById('txtDuplicateOffersCnt').focus();
	   ClearNoOfDuplicateOfferserror();
	   return false;
	  } else {
	  return true;
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
		   xmlhttpPost_OfferDuplicateString('OfferFeeds.aspx', 'Mode=NoOfDuplicateOffers&OfferID=' + document.mainform.OfferID.value + '&EngineID=0&DuplicateCnt=' + dupOffersCntvalue,'NoOfDuplicateOffers');
			//  // hide the new popup
			toggleDialog('DuplicateNoofOffer',false);
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
	
    function xmlhttpPost_OfferDuplicateString(strURL, qryStr, action) {
      var xmlHttpReq = false;
      var self = this;
      // Mozilla/Safari
      if (window.XMLHttpRequest) {
        self.xmlHttpReq = new XMLHttpRequest();
      }
      // IE
      else if (window.ActiveXObject) {
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
      }
      self.xmlHttpReq.open('POST', strURL + '?' + qryStr, true);
      self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	  self.xmlHttpReq.setRequestHeader("Content-length", qryStr.length);
	  self.xmlHttpReq.setRequestHeader("Connection", "close");
      self.xmlHttpReq.onreadystatechange = function() {
      if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
        if (action == 'NoOfDuplicateOffers') {
		  handleresponseDuplicateOffers(self.xmlHttpReq.responseText);
         }
       }
	  }
	  self.xmlHttpReq.send(qryStr);
      return false;
    }	

    function handleresponseDuplicateOffers(responseText) {
	  if (responseText.substring(0, 2) == 'OK') {
		window.location = 'offer-list.aspx';
	  } else {
	    window.location = 'offer-gen.aspx?' + responseText; 
	  }
    }	
	
    function setComponentsColor(color) {
        var elem = document.getElementById("linkComponent");

        if (elem != null) {
            elem.style.color = color;
        }
    }
    
    function toggleDropdown() {
        if (document.getElementById("actionsmenu") != null) {
            bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
            if (bOpen) {
                document.getElementById("actionsmenu").style.visibility = 'visible';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
            } else {
                document.getElementById("actionsmenu").style.visibility = 'hidden';
                document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
            }
        }
    }
</script>
<form action="#" id="mainform" name="mainform">
<input type="hidden" name="OfferID" value="<%sendb(offerid) %>" />
<%
  Dim oName As String = MyCommon.NZ(rst.Rows(0).Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
  Send("<div id=""intro"">")
  If (IsTemplate) Then
    Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & MyCommon.NZ(rst.Rows(0).Item("OfferID"), 0) & ": " & MyCommon.TruncateString(oName, 43) & "</h1>")
  Else
    Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & MyCommon.NZ(rst.Rows(0).Item("OfferID"), 0) & ": " & MyCommon.TruncateString(oName, 43) & "</h1>")
  End If
  Send(vbTab & "<div id=""controls""" & IIf(Popup, " style=""display:none;""", "") & ">")
  ShowActionButton = (Logix.UserRoles.CreateTemplate And Not IsTemplate) OrElse (Logix.UserRoles.CRUDOfferFromTemplate And IsTemplate) _
                      OrElse (((Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate)) Or (Logix.UserRoles.ExportOffer)) _
                      OrElse (Logix.UserRoles.SendOffersToCRM And Not IsTemplate And bShowCRM) _
                      OrElse ((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not IsTemplate) _
                      OrElse ((MyCommon.Fetch_SystemOption(73) <> "") And bExportToEDW) _
                      OrElse (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) _
                      OrElse (Logix.UserRoles.CreateOfferFromBlank) _
                      OrElse (Logix.UserRoles.EditFolders) _
                      OrElse (MyCommon.Fetch_CM_SystemOption(12) = "1") _
                      OrElse (bWorkflowActive AndAlso Logix.UserRoles.AssignPreValidate AndAlso Not IsTemplate AndAlso iWorkflowStatus <> 1) _
                      OrElse (bWorkflowActive AndAlso Logix.UserRoles.AssignPostValidate AndAlso Not IsTemplate AndAlso iWorkflowStatus = 1) _
                      OrElse (bWorkflowActive AndAlso Logix.UserRoles.AssignReadyToDeploy AndAlso Not IsTemplate AndAlso iWorkflowStatus = 2)

  If (Not LinksDisabled OrElse IsTemplate) Then
    If (ShowActionButton) Then
      Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " &#9660;"" onclick=""toggleDropdown();"" />")
      Send("<div class=""actionsmenu"" id=""actionsmenu"">")
      If (Logix.UserRoles.EditFolders) Then
        Send_AssignFolders(OfferID)
      End If
      If (Logix.UserRoles.CreateOfferFromBlank And Not bArchiveSystem) Then
        If oBuckStatus = Copient.ImportXml.BuckOfferStatus.NotBuckOffer Then
        Send_CopyOffer(IsTemplate)
      End If
      End If
      
      If bWorkflowActive Then
        If (Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate) Then
          If bBuckParentOffer Then
            Select Case oBuckStatus
              Case Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren,
               Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly,
               Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
                Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & "Are you sure that you want to delete this Buck offer, points programs, and all the child offers?" & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & " />")
              Case Else
                Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & "Are you sure that you want to delete this Buck offer and points programs?" & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & " />")
            End Select
            Select Case oBuckStatus
              Case Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren
                Send("<input type=""submit""  class=""regular"" id=""delete_paper"" name=""delete_paper"" onclick=""if(confirm('" & "Are you sure that you want to delete the paper child offers?" & "')){}else{return false}"" value=""" & "Delete paper child offers" & """" & " />")
                Send("<input type=""submit""  class=""regular"" id=""delete_digital"" name=""delete_digital"" onclick=""if(confirm('" & "Are you sure that you want to delete all the digital child offers?" & "')){}else{return false}"" value=""" & "Delete digital child offers" & """" & " />")
              Case Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly
                Send("<input type=""submit""  class=""regular"" id=""delete_paper"" name=""delete_paper"" onclick=""if(confirm('" & "Are you sure that you want to delete the paper child offers?" & "')){}else{return false}"" value=""" & "Delete paper child offers" & """" & " />")
              Case Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
                Send("<input type=""submit""  class=""regular"" id=""delete_digital"" name=""delete_digital"" onclick=""if(confirm('" & "Are you sure that you want to delete all the digital child offers?" & "')){}else{return false}"" value=""" & "Delete digital child offers" & """" & " />")
            End Select
          ElseIf Not bBuckChildOffer Then
            If (MyCommon.Fetch_CM_SystemOption(113) = "0") Then
          Send_Delete()
        End If
          End If
        End If
      Else
        ' proceed normally
        If ((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not IsTemplate And Not bBuckParentToBe) Then
          If (MyCommon.Fetch_SystemOption(145) = "1") Then
            Dim condt As DataTable
            MyCommon.QueryStr = "select ConditionID, ConditionTypeID from OfferConditions OC with (NoLock) " & _
                                "inner join Offers O with (NoLock) on O.OfferID = OC.OfferID " & _
                                         "where OC.Deleted = 0 and O.Deleted = 0 and OC.OfferID =" & OfferID & " "
            condt = MyCommon.LRT_Select
            If (condt.Rows.Count > 0) Then
              If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
              Send_DeferDeploy()
              End If
            Else
              If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
              Send_DeferDeployConditional()
            End If
            End If
          Else
            If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
            Send_DeferDeploy()
          End If
        End If
        End If
        If (Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate) Then
          If bBuckParentOffer Then
            Select Case oBuckStatus
              Case Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren,
               Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly,
               Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
                Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & "Are you sure that you want to delete this Buck offer, points programs, and all the child offers?" & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & " />")
              Case Else
                Send("<input type=""submit"" accesskey=""d"" class=""regular"" id=""delete"" name=""delete"" onclick=""if(confirm('" & "Are you sure that you want to delete this Buck offer and points programs?" & "')){}else{return false}"" value=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """" & " />")
            End Select
            Select Case oBuckStatus
              Case Copient.ImportXml.BuckOfferStatus.BuckParentBothChildren
                Send("<input type=""submit""  class=""regular"" id=""delete_paper"" name=""delete_paper"" onclick=""if(confirm('" & "Are you sure that you want to delete the paper child offers?" & "')){}else{return false}"" value=""" & "Delete paper child offers" & """" & " />")
                Send("<input type=""submit""  class=""regular"" id=""delete_digital"" name=""delete_digital"" onclick=""if(confirm('" & "Are you sure that you want to delete all the digital child offers?" & "')){}else{return false}"" value=""" & "Delete digital child offers" & """" & " />")
              Case Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly
                Send("<input type=""submit""  class=""regular"" id=""delete_paper"" name=""delete_paper"" onclick=""if(confirm('" & "Are you sure that you want to delete the paper child offers?" & "')){}else{return false}"" value=""" & "Delete paper child offers" & """" & " />")
              Case Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
                Send("<input type=""submit""  class=""regular"" id=""delete_digital"" name=""delete_digital"" onclick=""if(confirm('" & "Are you sure that you want to delete all the digital child offers?" & "')){}else{return false}"" value=""" & "Delete digital child offers" & """" & " />")
            End Select
          ElseIf Not bBuckChildOffer Then
            If (MyCommon.Fetch_CM_SystemOption(113) = "0") Then
          Send_Delete()
        End If
          End If
        End If
        If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate And Not bBuckParentToBe) Then
          ' To implement Request for Offer Validation
          If (MyCommon.Fetch_SystemOption(145) = "1") Then
            Dim condt As DataTable
            MyCommon.QueryStr = "select ConditionID, ConditionTypeID from OfferConditions OC with (NoLock) " & _
                                "inner join Offers O with (NoLock) on O.OfferID = OC.OfferID " & _
                                         "where OC.Deleted = 0 and O.Deleted = 0 and OC.OfferID =" & OfferID & " "
            condt = MyCommon.LRT_Select
            If (condt.Rows.Count > 0) Then
              Send_Deploy()
            Else
              If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
              Send_DeployConditional()
            End If
            End If
          Else
            Send_Deploy()
          End If
        End If
      End If
      
      If (Logix.UserRoles.ExportOffer) Then
        If oBuckStatus = Copient.ImportXml.BuckOfferStatus.NotBuckOffer Then
        Send_Export()
        End If
        If (MyCommon.Fetch_CM_SystemOption(12) = "1" And Not bBuckParentToBe) Then
          Send_ExportCME()
          Send_ExportCRM()
        End If
        If ((MyCommon.Fetch_SystemOption(73) <> "") And bExportToEDW And Not bArchiveSystem And Not bBuckParentToBe) Then
          Send_ExportToEDW()
        End If
      End If
      If (Logix.UserRoles.CreateOfferFromBlank And Not bArchiveSystem) Then
        Send_New()
      End If
      If (Logix.UserRoles.CRUDOfferFromTemplate And IsTemplate And Not bArchiveSystem) Then
        Send_OfferFromTemp()
      End If
      If (Logix.UserRoles.CreateTemplate And Not IsTemplate And Not bArchiveSystem And oBuckStatus = Copient.ImportXml.BuckOfferStatus.NotBuckOffer) Then
        Send_Saveastemp()
      End If
      If (Logix.UserRoles.SendOffersToCRM And Not IsTemplate And Not bArchiveSystem And bShowCRM And oBuckStatus = Copient.ImportXml.BuckOfferStatus.NotBuckOffer) Then
        If InboundCRMEngineID = 0 Then
          ' all "internal" offers may be be sent outbound
          Send_SendOutbound()
        Else
          ' check Cm option to see if one may send a CRM external offer outbound (allow round trip?)
          Dim iSendExternalOutbound As Integer = 0
          objTemp = MyCommon.Fetch_CM_SystemOption(48)
          If Not (Integer.TryParse(objTemp.ToString, iSendExternalOutbound)) Then iSendExternalOutbound = 0
          If iSendExternalOutbound <> 0 Then
            ' only allow if option is set to true
            Send_SendOutbound()
          End If
        End If
      End If
      
      ' check Workflow option, if set then display workflow action buttons
      If bWorkflowActive And Not bBuckParentToBe Then
        If bProductionSystem Then
          If (Logix.UserRoles.AssignPreValidate AndAlso Not IsTemplate AndAlso iWorkflowStatus <> 1 AndAlso StatusCode <> Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED AndAlso Not OfferDeployed(MyCommon, OfferId)) Then
            Send_PreValidate()
          End If
          If (Logix.UserRoles.AssignPostValidate AndAlso Not IsTemplate AndAlso iWorkflowStatus = 1) Then
            Send_PostValidate()
          End If
          If (Logix.UserRoles.AssignReadyToDeploy AndAlso Not IsTemplate AndAlso iWorkflowStatus = 2) Then
            Send_ReadyToDeploy()
          End If
        End If
        If ((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not IsTemplate And Not bArchiveSystem And Not bBuckParentToBe) Then
          ' Ready to Deploy?
          If iWorkflowStatus = 3 Or bTestSystem Or IsExpired Then
            If (MyCommon.Fetch_SystemOption(145) = "1") Then
              Dim condt As DataTable
              MyCommon.QueryStr = "select ConditionID, ConditionTypeID from OfferConditions OC with (NoLock) " & _
                                  "inner join Offers O with (NoLock) on O.OfferID = OC.OfferID " & _
                                           "where OC.Deleted = 0 and O.Deleted = 0 and OC.OfferID =" & OfferID & " "
              condt = MyCommon.LRT_Select
              If (condt.Rows.Count > 0) Then
                If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
                Send_DeferDeploy()
                End If
              Else
                If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
                Send_DeferDeployConditional()
              End If
              End If
            Else
              If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
              Send_DeferDeploy()
            End If
          End If
        End If
        End If
        If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate And Not bArchiveSystem And Not bBuckParentToBe) Then
          ' Ready to Deploy?
          If iWorkflowStatus = 3 Or bTestSystem Or IsExpired Then
            ' To implement Request for Offer Validation
            If (MyCommon.Fetch_SystemOption(145) = "1") Then
              Dim condt As DataTable
              MyCommon.QueryStr = "select ConditionID, ConditionTypeID from OfferConditions OC with (NoLock) " & _
                                  "inner join Offers O with (NoLock) on O.OfferID = OC.OfferID " & _
                                  "where OC.Deleted = 0 and O.Deleted = 0 and OC.OfferID =" & OfferID & " "
              condt = MyCommon.LRT_Select
              If (condt.Rows.Count > 0) Then
                Send_Deploy()
              Else
                If (MyCommon.Fetch_CM_SystemOption(111) = "0") Then
                Send_DeployConditional()
              End If
              End If
            Else
              Send_Deploy()
            End If
          End If
        End If
      End If
      
      Select Case oBuckStatus
        Case Copient.ImportXml.BuckOfferStatus.BuckParentNoChildren, Copient.ImportXml.BuckOfferStatus.BuckTiered
          Send("<input type=""submit"" class=""regular"" id=""create_paper"" name=""create_paper"" value=""" & "Create paper child offers" & """" & " />")
          Send("<input type=""submit"" class=""regular"" id=""create_digital"" name=""create_digital"" value=""" & "Create digital child offers" & """" & " />")
        Case Copient.ImportXml.BuckOfferStatus.BuckParentPaperChildrenOnly
          Send("<input type=""submit"" class=""regular"" id=""create_digital"" name=""create_digital"" value=""" & "Create digital child offers" & """" & " />")
        Case Copient.ImportXml.BuckOfferStatus.BuckParentDigitalChildrenOnly
          Send("<input type=""submit"" class=""regular"" id=""create_paper"" name=""create_paper"" value=""" & "Create paper child offers" & """" & " />")
      End Select

      If (MyCommon.Fetch_CM_SystemOption(113) = "1") Then
        If (Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate) Then
          If Not (bBuckParentOffer Or bBuckChildOffer) Then
            Send_Delete()
          End If
        End If
      End If
      Send("</div>")
    End If
  Else
    If ((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not IsTemplate) Then
      Send_CancelDeploy()
    End If
  End If
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_NotesButton(3, OfferID, AdminUserID)
    End If
  End If
  Send(vbTab & "</div>")
  Send("</div>")
%>
<div id="main">
  <%
    If Not IsTemplate Then
      If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) <> 2) Then
        If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) > 0) Then
          If (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = False) Then
            modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
            Send("<div id=""modbar"">" & modMessage & "</div>")
          End If
        End If
      End If
    End If
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
      
    ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
    If (Not IsTemplate AndAlso modMessage = "") Then
      MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where CreatedDate = LastUpdate and OfferID=" & OfferID
      rst3 = MyCommon.LRT_Select
      If (rst3.Rows.Count = 0) Then
        Send_Status(OfferID)
      End If
    End If
  %>
  <div id="column1">
    <div class="box" id="general">
      <%  
        If (LinksDisabled) Then
          Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</span></h2>")
        Else
          Send("<h2 style=""float:left;""><span><a href=""offer-gen.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</a></span></h2>")
        End If
        Send_BoxResizer("generalbody", "imgGeneral", Copient.PhraseLib.Lookup("term.general", LanguageID), True)
      %>
      <div id="generalbody">
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.general", LanguageID))%>"
          cellpadding="0" cellspacing="0">
          <tr>
            <td style="width: 95px;">
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.NZ(rst.Rows(0).Item("OfferID"), 0))%>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.NZ(rst.Rows(0).Item("ExtOfferID"), ""))%>
            </td>
          </tr>
          <% If Not bProductionSystem Then%>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.productionid", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.NZ(rst.Rows(0).Item("ProductionID"), ""))%>
            </td>
          </tr>
          <% End If%>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.engine", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                Select Case EngineId
                  Case 0
                    Send(Copient.PhraseLib.Lookup("term.cm", LanguageID))
                  Case 1
                    Send(Copient.PhraseLib.Lookup("term.catalina", LanguageID))
                  Case 4
                    Send(Copient.PhraseLib.Lookup("term.dp", LanguageID))
                  Case Else
                    Send(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End Select
              %>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.crmengine", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                MyCommon.QueryStr = "select ExtInterfaceID, Name, PhraseID from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=" & CRMEngineID & ";"
                rst4 = MyCommon.LRT_Select()
                If rst4.Rows.Count > 0 Then
                  If Not IsDBNull(rst4.Rows(0).Item("PhraseID")) Then
                    Sendb(Copient.PhraseLib.Lookup(rst4.Rows(0).Item("PhraseID"), LanguageID))
                  Else
                    Sendb(MyCommon.NZ(rst4.Rows(0).Item("Name"), ""))
                  End If
                End If
              %>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                sTemp1 = Logix.GetOfferStatus(OfferID, LanguageID)
                Sendb(sTemp1)
              %>
            </td>
          </tr>
          <% If iWorkflowStatus = 2 Then%>
          <tr>
            <td>
              <b>
                <% Sendb(" ")%>
              </b>
            </td>
            <td>
              <%
                sTemp2 = MyExport.GetStatusForTransferToTest(OfferID)
                Sendb(sTemp2)
              %>
            </td>
          </tr>
          <% End If%>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("Name"), 0), 25))%>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("Description"), ""), 25))%>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.folders", LanguageID))%>
                :</b>
            </td>
            <td id="folderNames">
              <% Sendb(FolderNames)%>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.category", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                'Sendb("<a href=""javascript:openPopup('offer-timeline.aspx?Category=" & MyCommon.NZ(row.Item("OfferCategoryID"), "") & "')"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("cgDescription"), ""), 25) & "</a>")
                If MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), 0) > 0 Then
                  Send("<a href=""category-edit.aspx?OfferCategoryID=" & MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("cgDescription"), ""), 25) & "</a>")
                Else
                  Send(Copient.PhraseLib.Lookup("term.none", LanguageID))
                End If
              %>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.priority", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.NZ(rst.Rows(0).Item("priorityLevel"), ""))%>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                If MyCommon.NZ(rst.Rows(0).Item("distPeriodLimit"), 0) = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))
                Else
                  If MyCommon.NZ(rst.Rows(0).Item("distPeriodLimit"), 0) = 1 Then
                    Sendb(Copient.PhraseLib.Lookup("term.once", LanguageID))
                  ElseIf MyCommon.NZ(rst.Rows(0).Item("distPeriodLimit"), 0) = 2 Then
                    Sendb(Copient.PhraseLib.Lookup("term.twice", LanguageID))
                  Else
                    Sendb(MyCommon.NZ(rst.Rows(0).Item("distPeriodLimit"), 0))
                  End If
                  Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                  If (MyCommon.NZ(rst.Rows(0).Item("distPeriod"), 0) > 0) Then
                    If (MyCommon.NZ(rst.Rows(0).Item("distPeriod"), 0) = 1) Then
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase))
                    Else
                      Sendb(MyCommon.NZ(rst.Rows(0).Item("distPeriod"), 0) & " " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase))
                    End If
                  Else
                    If (MyCommon.NZ(rst.Rows(0).Item("distPeriod"), 0) = 0) Then
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.transaction", LanguageID), VbStrConv.Lowercase))
                    ElseIf (MyCommon.NZ(rst.Rows(0).Item("distPeriod"), 0) = -1) Then
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.customer", LanguageID), VbStrConv.Lowercase))
                    Else
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.unknown", LanguageID), VbStrConv.Lowercase))
                    End If
                  End If
                End If
              %>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.tiers", LanguageID))%>
                :</b>
            </td>
            <td>
              <% 
                Sendb(MyCommon.NZ(rst.Rows(0).Item("numtiers"), ""))
                If oBuckStatus <> Copient.ImportXml.BuckOfferStatus.NotBuckOffer Then
                  Sendb(" - " & oBuckStatus.ToString)
                End If
              %>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.reporting", LanguageID))%>
                :</b>
            </td>
            <%
              If (Logix.UserRoles.AccessReports = True) AndAlso (Popup = False) Then
                Sendb("<td><a href=""reports-detail.aspx?OfferID=" & MyCommon.NZ(rst.Rows(0).Item("OfferID"), -1) & "&amp;Start=" & MyCommon.NZ(rst.Rows(0).Item("prodStartDate"), "1/1/1900") & "&amp;End=" & MyCommon.NZ(rst.Rows(0).Item("prodEndDate"), "1/1/1900") & "&amp;Status=" & Copient.PhraseLib.Lookup("offer.status" & MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0), LanguageID) & """>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & "</a></td>")
              Else
                Sendb("<td>" & Copient.PhraseLib.Lookup("term.enabled", LanguageID) & "</td>")
              End If
            %>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.createdby", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CreatedBy"), ""), 25))%>
            </td>
          </tr>
          <tr>
            <td>
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.lastupdatedby", LanguageID))%>
                :</b>
            </td>
            <td>
              <% Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("LastUpdatedBy"), ""), 25))%>
            </td>
          </tr>
        </table>
      </div>
    </div>
    <% If (Not IsTemplate) Then%>
    <div class="box" id="validation">
      <h2 style="float: left;">
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID))%>
        </span>
      </h2>
      <%  Send_BoxResizer("validationbody", "imgValidation", Copient.PhraseLib.Lookup("term.validation", LanguageID), True)%>
      <div id="validationbody">
        <%
          Dim dtValid, dtComponents As DataTable
          Dim rowOK(), rowWaiting(), rowWatches(), rowWarnings() As DataRow
          Dim rowComp As DataRow
          Dim GraceHours As Integer
          Dim GraceHoursWarn As Integer
          Dim iOfferLocations As Integer
          Dim AllComponentsValid As Boolean = True
            
          objTemp = MyCommon.Fetch_CM_SystemOption(10)
          If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
            GraceHours = 4
          End If
            
          objTemp = MyCommon.Fetch_CM_SystemOption(11)
          If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
            GraceHoursWarn = 24
          End If
            
          MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_Offer"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
          MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
          MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn
            
          dtValid = MyCommon.LRTsp_select()
          iOfferLocations = dtValid.Rows.Count
            
          rowOK = dtValid.Select("Status=0", "LocationName")
          rowWaiting = dtValid.Select("Status=1", "LocationName")
          rowWatches = dtValid.Select("Status=2", "LocationName")
          rowWarnings = dtValid.Select("Status=3", "LocationName")
          MyCommon.Close_LRTsp()
            
          ValidateIncentiveColor = IIf(rowWarnings.Length > 0, "red", "green")
            
          Send("<a href=""javascript:showDiv('divOffer');"" style=""color:" & ValidateIncentiveColor & ";""><b>+ " & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</b></a><br />")
          Send("<div id=""divOffer"" style=""margin-left:10px;display:none;"">")
          If Popup Then
            Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & " of " & iOfferLocations & ")<br />")
            Send(Copient.PhraseLib.Lookup("cgroup-edit.waitlocations", LanguageID) & " (" & rowWaiting.Length & " of " & iOfferLocations & ")<br />")
            Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & " of " & iOfferLocations & ")<br />")
            Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & " of " & iOfferLocations & ")<br />")
          Else
            Send("<a id=""validLink" & OfferID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=in&id=" & OfferID & "&level=0&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
            Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iOfferLocations & ")</a><br />")
            Send("<a id=""waitingLink" & OfferID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=in&id=" & OfferID & "&level=1&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
            Send(Copient.PhraseLib.Lookup("cgroup-edit.waitlocations", LanguageID) & " (" & rowWaiting.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iOfferLocations & ")</a><br />")
            Send("<a id=""watchLink" & OfferID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=in&id=" & OfferID & "&level=2&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
            Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iOfferLocations & ")</a><br />")
            Send("<a id=""warningLink" & OfferID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=in&id=" & OfferID & "&level=3&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
            Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & " " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " " & iOfferLocations & ")</a><br />")
          End If
          Send("</div>")
            
          MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_OfferComponents"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
          dtComponents = MyCommon.LRTsp_select
          MyCommon.Close_LRTsp()
          Send("<a id=""linkComponent"" href=""javascript:showDiv('divComponents');"" style=""color:green;""><br class=""half"" /><b>+ " & Copient.PhraseLib.Lookup("term.components", LanguageID) & "</b><br /></a>")
          Send("<div id=""divComponents"" style=""display:none;"">")
            
          For Each rowComp In dtComponents.Rows
            Send("<div style=""margin-left:10px;"">")
            WriteComponent(MyCommon, rowComp, ComponentColor)
            AllComponentsValid = AllComponentsValid AndAlso (ComponentColor.ToUpper = "GREEN")
            Send("</div>")
          Next
          Send("</div>")
            
          ' Update the Offer Validation Summary table with the most current validation information
          If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
            MyCommon.Open_LogixRT()
          End If
          MyCommon.QueryStr = "dbo.pa_UpdateValidationSummary"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
          MyCommon.LRTsp.Parameters.Add("@ValidLocations", SqlDbType.Int).Value = rowOK.Length
          MyCommon.LRTsp.Parameters.Add("@WatchLocations", SqlDbType.Int).Value = rowWaiting.Length + rowWatches.Length
          MyCommon.LRTsp.Parameters.Add("@WarningLocations", SqlDbType.Int).Value = rowWarnings.Length
          MyCommon.LRTsp.Parameters.Add("@ComponentsValid", SqlDbType.Bit).Value = IIf(AllComponentsValid, 1, 0)
          MyCommon.LRTsp.ExecuteNonQuery()
          MYCommon.Close_LRTsp()
        %>
      </div>
      <hr class="hidden" />
    </div>
    <% End If%>
    <div class="box" id="period">
      <h2 style="float: left;">
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>
        </span>
      </h2>
      <% Send_BoxResizer("periodbody", "imgPeriod", Copient.PhraseLib.Lookup("term.period", LanguageID), True)%>
      <div id="periodbody">
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>"
          cellpadding="0" cellspacing="0">
          <% If bUseTestDates Then%>
          <tr>
            <td style="width: 85px; vertical-align: top;">
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.test", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                LongDate = MyCommon.NZ(rst.Rows(0).Item("testStartDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                Sendb(" - ")
                LongDate = MyCommon.NZ(rst.Rows(0).Item("testEndDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("testStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("testEndDate"), "1/1/1900")) + 1 & " ")
                If DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("testStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("testEndDate"), "1/1/1900")) + 1 = 1 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                Else
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                End If
                Sendb("<br />")
                Dim testStartDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(rst.Rows(0).Item("testStartDate"), "1/1/1900"))
                Dim testEndDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(rst.Rows(0).Item("testEndDate"), "1/1/1900"))
                If testStartDateDiff > 1 Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeginsIn", LanguageID, testStartDateDiff) & " ")
                ElseIf testStartDateDiff = 1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeginsTomorrow", LanguageID) & " ")
                ElseIf testStartDateDiff = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganToday", LanguageID) & " ")
                ElseIf testStartDateDiff = -1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganYesterday", LanguageID) & " ")
                Else
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeganDaysAgo", LanguageID, Math.Abs(testStartDateDiff)) & " ")
                End If
                If testEndDateDiff > 1 Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndsIn", LanguageID, testEndDateDiff))
                ElseIf testEndDateDiff = 1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsTomorrow", LanguageID))
                ElseIf testEndDateDiff = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsToday", LanguageID))
                ElseIf testEndDateDiff = -1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndedYesterday", LanguageID))
                Else
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndedDaysAgo", LanguageID, Math.Abs(testEndDateDiff)))
                End If
              %>
            </td>
          </tr>
          <% End If%>
          <tr>
            <td style="width: 85px; vertical-align: top;">
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                LongDate = MyCommon.NZ(rst.Rows(0).Item("prodStartDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                Sendb(" - ")
                LongDate = MyCommon.NZ(rst.Rows(0).Item("prodEndDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("prodStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("prodEndDate"), "1/1/1900")) + 1 & " ")
                If DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("prodStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("prodEndDate"), "1/1/1900")) + 1 = 1 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                Else
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                End If
                Sendb("<br />")
                Dim StartDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(rst.Rows(0).Item("prodStartDate"), "1/1/1900"))
                Dim EndDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(rst.Rows(0).Item("prodEndDate"), "1/1/1900"))
                If StartDateDiff > 1 Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeginsIn", LanguageID, StartDateDiff) & " ")
                ElseIf StartDateDiff = 1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeginsTomorrow", LanguageID) & " ")
                ElseIf StartDateDiff = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganToday", LanguageID) & " ")
                ElseIf StartDateDiff = -1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganYesterday", LanguageID) & " ")
                Else
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeganDaysAgo", LanguageID, Math.Abs(StartDateDiff)) & " ")
                End If
                If EndDateDiff > 1 Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndsIn", LanguageID, EndDateDiff))
                ElseIf EndDateDiff = 1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsTomorrow", LanguageID))
                ElseIf EndDateDiff = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsToday", LanguageID))
                ElseIf EndDateDiff = -1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndedYesterday", LanguageID))
                Else
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndedDaysAgo", LanguageID, Math.Abs(EndDateDiff)))
                End If
              %>
            </td>
          </tr>
        </table>
      </div>
    </div>
    <% If bUseDisplayDates Then%>
    <div class="box" id="displaydates">
      <h2 style="float: left;">
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.displaydates", LanguageID))%>
        </span>
      </h2>
      <% Send_BoxResizer("displaybody", "imgDisplay", Copient.PhraseLib.Lookup("term.display", LanguageID), True)%>
      <div id="displaybody">
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>"
          cellpadding="0" cellspacing="0">
          <%MyCommon.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & OfferID & " "
            Dim dtODisp As New DataTable
            dtODisp = MyCommon.LRT_Select()
            If dtODisp.Rows.Count > 0 Then
              If Not IsDBNull(dtODisp.Rows(0).Item("DisplayStartDate")) And Not IsDBNull(dtODisp.Rows(0).Item("DisplayEndDate")) Then%>
          <tr>
            <td style="width: 85px; vertical-align: top;">
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.display", LanguageID))%>
                :</b>
            </td>
            <td>
              <%
                LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                Sendb(" - ")
                LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900"), MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")) + 1 & " ")
                If DateDiff(DateInterval.Day, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900"), MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")) + 1 = 1 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                Else
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                End If
                Sendb("<br />")
                LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(LongDate.ToString("HH:mm:ss")) Else Sendb("?")
                Sendb(" - ")
                LongDate = MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900")
                If LongDate > "1/1/1900" Then Sendb(LongDate.ToString("HH:mm:ss")) Else Sendb("?")
                Sendb("<br />")
                    
                Dim DispStartDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), "1/1/1900"))
                Dim DispEndDateDiff As Integer = DateDiff(DateInterval.Day, Today, MyCommon.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), "1/1/1900"))
                If DispStartDateDiff > 1 Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeginsIn", LanguageID, DispStartDateDiff) & " ")
                ElseIf DispStartDateDiff = 1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeginsTomorrow", LanguageID) & " ")
                ElseIf DispStartDateDiff = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganToday", LanguageID) & " ")
                ElseIf DispStartDateDiff = -1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.BeganYesterday", LanguageID) & " ")
                Else
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.BeganDaysAgo", LanguageID, Math.Abs(DispStartDateDiff)) & " ")
                End If
                If DispEndDateDiff > 1 Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndsIn", LanguageID, DispEndDateDiff))
                ElseIf DispEndDateDiff = 1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsTomorrow", LanguageID))
                ElseIf DispEndDateDiff = 0 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndsToday", LanguageID))
                ElseIf DispEndDateDiff = -1 Then
                  Sendb(Copient.PhraseLib.Lookup("offer-sum.EndedYesterday", LanguageID))
                Else
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.EndedDaysAgo", LanguageID, Math.Abs(DispEndDateDiff)))
                End If
              %>
            </td>
          </tr>
          <% Else%>
          <tr>
            <td style="width: 85px; vertical-align: top;">
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%>
              </b>
            </td>
            <td>
              &nbsp;
            </td>
          </tr>
          <%End If
          Else%>
          <tr>
            <td style="width: 85px; vertical-align: top;">
              <b>
                <% Sendb(Copient.PhraseLib.Lookup("term.none", LanguageID))%>
              </b>
            </td>
            <td>
              &nbsp;
            </td>
          </tr>
          <%  End If%>
        </table>
      </div>
    </div>
    <% End If%>
    <% If (Not IsTemplate) Then%>
    <div class="box" id="deployment">
      <h2>
        <span>
          <%Sendb(Copient.PhraseLib.Lookup("term.deployment", LanguageID))%>
        </span>
      </h2>
      <h3>
        <%Sendb(Copient.PhraseLib.Lookup("term.lastattempted", LanguageID))%>
        :
      </h3>
      <%
        LongDate = MyCommon.NZ(rst.Rows(0).Item("CMOARptDate"), "1/1/1900")
        If LongDate > "1/1/1900" Then
          DaysDiff = DateDiff("d", LongDate.ToShortDateString(), DateTime.Today.ToShortDateString())
          Sendb("&nbsp;" & Logix.ToLongDateTimeString(LongDate, MyCommon))
          If DaysDiff < 0 Then
          ElseIf DaysDiff = 0 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
          ElseIf DaysDiff = 1 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
          Else
            Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
          End If
        Else
          Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
        End If
        Sendb("<br />")
      %>
      <br class="half" />
      <h3>
        <%Sendb(Copient.PhraseLib.Lookup("term.lastsuccessful", LanguageID))%>
        :
      </h3>
      <%
        LongDate = MyCommon.NZ(rst.Rows(0).Item("CMOADeploySuccessDate"), "1/1/1900")
        If LongDate > "1/1/1900" Then
          DaysDiff = DateDiff("d", LongDate.ToShortDateString(), DateTime.Today.ToShortDateString())
          Sendb("&nbsp;" & Logix.ToLongDateTimeString(LongDate, MyCommon))
          If DaysDiff < 0 Then
          ElseIf DaysDiff = 0 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
          ElseIf DaysDiff = 1 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
          Else
            Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
          End If
        Else
          Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
        End If
        Sendb("<br />")
      %>
    <br class="half" />
    <h3>
        <%Sendb(Copient.PhraseLib.Lookup("term.lastvalidationmessage", LanguageID))%>
        :
    </h3>
    <%
        Sendb(GetLastDeployValidationMessage(MyCommon, OfferID))
        Sendb("<br />")
    %>
      <br class="half" />
      <h3>
        <%Sendb(Copient.PhraseLib.Lookup("term.laststatus", LanguageID))%>
        :
      </h3>
      <span <%
        If  CMS.Utilities.NZ(rst.Rows(0).Item("CMOADeployRpt"),"").ToLower() <> "export okay" Then
          Sendb(" class='red'")
        End If
        %> >
      <%Sendb("&nbsp;" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CMOADeployRpt"), ""), 25))%>
      </span>
      <br />
      <br class="half" />
      <% If bShowCRM Then%>
      <h3>
        <%Sendb(Copient.PhraseLib.Lookup("offer-sum.crmlastsent", LanguageID))%>
        :
      </h3>
      <%
        MyCommon.QueryStr = "select LastCRMSendDate from offers with (NoLock) where OfferId=" & OfferID & ";"
        rst4 = MyCommon.LRT_Select
        If rst4.Rows.Count > 0 Then
          LongDate = MyCommon.NZ(rst4.Rows(0).Item("LastCRMSendDate"), "1/1/1900")
        Else
          LongDate = "1/1/1900"
        End If
        TodayDateZeroTime = New Date(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0)
        LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
        If LongDate > "1/1/1900" Then
          DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
          Sendb("&nbsp;" & Logix.ToLongDateTimeString(LongDate, MyCommon))
          If DaysDiff < 0 Then
          ElseIf DaysDiff = 0 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
          ElseIf DaysDiff = 1 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
          Else
            Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
          End If
        Else
          Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
        End If
        Sendb("<br />")
      %>
      <br class="half" />
      <h3>
        <%Sendb(Copient.PhraseLib.Lookup("term.laststatus", LanguageID))%>
        :
      </h3>
      <%
        Dim iStatus As Integer = 0
        MyCommon.QueryStr = "select CRMSendStatus from Offers with (NoLock) where OfferId=" & OfferID & ";"
        rst4 = MyCommon.LRT_Select
        If rst4.Rows.Count > 0 Then
          iStatus = MyCommon.NZ(rst4.Rows(0).Item("CRMSendStatus"), 0)
        End If
        If CRMEngineID > 2 Then
          Select Case iStatus
            Case 0
            Case 1
              Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.waitcrmexport", LanguageID))
            Case 2
              Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.crmwaitack", LanguageID))
            Case 3
              Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.ok", LanguageID))
            Case Else
              Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.error", LanguageID))
          End Select
        Else
          ' Teradata
          Select Case iStatus
            Case 0
            Case 1
              Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.inprogress", LanguageID))
            Case 2
              Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.ok", LanguageID))
            Case Else
              Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.error", LanguageID))
          End Select
        End If
      %>
      <br />
      <br class="half" />
      <h3>
        <%Sendb(Copient.PhraseLib.Lookup("offer-sum.crmlastreceived", LanguageID))%>
        :
      </h3>
      <%
        MyCommon.QueryStr = "select LastUpdate from CRMImportQueue with (NoLock) where OfferID=" & OfferID & " order by LastUpdate desc;"
        rst4 = MyCommon.LRT_Select
        If rst4.Rows.Count > 0 Then
          LongDate = MyCommon.NZ(rst4.Rows(0).Item("LastUpdate"), "1/1/1900")
        Else
          LongDate = "1/1/1900"
        End If
        LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
        If LongDate > "1/1/1900" Then
          DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
          Sendb("&nbsp;" & Logix.ToLongDateTimeString(LongDate, MyCommon))
          If DaysDiff < 0 Then
          ElseIf DaysDiff = 0 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.today", LanguageID), VbStrConv.Lowercase) & ")")
          ElseIf DaysDiff = 1 Then
            Sendb(" (" & StrConv(Copient.PhraseLib.Lookup("term.yesterday", LanguageID), VbStrConv.Lowercase) & ")")
          Else
            Sendb(" (" & DaysDiff & " " & StrConv(Copient.PhraseLib.Lookup("term.daysago", LanguageID), VbStrConv.Lowercase) & ")")
          End If
        Else
          Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
        End If
        Sendb("<br />")
      %>
      <br class="half" />
      <% End If%>
      <hr class="hidden" />
    </div>
    <% End If%>
  </div>
  <div id="gutter">
  </div>
  <div id="column2">
    <div class="box" id="locations">
      <%
        If (LinksDisabled) Then
          Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.locations", LanguageID) & "</span></h2>")
        Else
          Send("<h2 style=""float:left;""><span><a href=""offer-loc.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.locations", LanguageID) & "</a></span></h2>")
        End If
        Send_BoxResizer("locationbody", "imgLocations", Copient.PhraseLib.Lookup("term.locations", LanguageID), True)
      %>
      <div id="locationbody">
        <%
          If (BannersEnabled) Then
            Sendb("<h3>" & Copient.PhraseLib.Lookup(IIf(BannerCt > 1, "term.banners", "term.banner"), LanguageID) & "</h3>")

            Send("<ul class=""condensed"">")
            For i = 0 To BannerNames.GetUpperBound(0)
              If (BannerIDs(i) > -1) Then
                Sendb("<li><a href=""banner-edit.aspx?BannerID=" & BannerIDs(i) & """>" & BannerNames(i) & "</a></li>")
              Else
                Sendb("<li>" & BannerNames(i) & "</li>")
              End If
            Next
            Send("</ul>")
          End If
        %>
        <h3>
          <% Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID))%>
        </h3>
        <%
          GCount = 0
          MyCommon.QueryStr = "select OL.LocationGroupID,LG.Name from OfferLocations as OL with (NoLock) " & _
                              "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=OL.LocationGroupID " & _
                              "where Excluded=0 and OL.Deleted=0 and OL.OfferID=" & OfferID & " order by LG.Name"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("LocationGroupID"), -1) = 1) Then
              AnyStoreUsed = True
            End If
          Next
          If rst.Rows.Count > 0 And Not AnyStoreUsed Then
            For Each row In rst.Rows
              MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), -1) & " And Deleted = 0"
              rst2 = MyCommon.LRT_Select()
              For Each row2 In rst2.Rows
                GCount = GCount + (row2.Item("GCount"))
              Next
            Next
            Sendb("&nbsp;" & GCount & " ")
            If GCount <> 1 Then
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.stores", LanguageID), VbStrConv.Lowercase))
            Else
              Sendb(StrConv(Copient.PhraseLib.Lookup("term.store", LanguageID), VbStrConv.Lowercase))
            End If
            Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
            Sendb(rst.Rows.Count & " ")
            If rst.Rows.Count <> 1 Then
              Send(StrConv(Copient.PhraseLib.Lookup("term.storegroups", LanguageID), VbStrConv.Lowercase) & ":<br />")
            Else
              Send(StrConv(Copient.PhraseLib.Lookup("term.storegroup", LanguageID), VbStrConv.Lowercase) & ":<br />")
            End If
          End If
          Send("<ul class=""condensed"">")
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("LocationGroupID"), -1) > 1) Then
              Sendb("<li><a href=""lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
            Else
              Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
            End If
            MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), -1) & " And Deleted = 0"
            rst2 = MyCommon.LRT_Select()
            For Each row2 In rst2.Rows
              If (MyCommon.NZ(row.Item("LocationGroupID"), -1) > 1) Then
                Sendb(" (" & row2.Item("GCount") & ")")
              End If
            Next
          Next
          If (rst.Rows.Count = 0) Then
            Sendb("<li>" & Copient.PhraseLib.Lookup("term.none", LanguageID))
          End If
          ' Check for and display any excluded store groups
          MyCommon.QueryStr = "select OL.LocationGroupID,LG.Name from OfferLocations as OL with (NoLock) inner join LocationGroups as LG with (NoLock) on " & _
                              "LG.LocationGroupID=OL.LocationGroupID where Excluded=1 and OL.deleted=0 and OL.OfferID=" & OfferID
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
            For Each row In rst.Rows
              MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), -1) & " And Deleted=0"
              rst2 = MyCommon.LRT_Select()
              For Each row2 In rst2.Rows
                GCount = GCount + (row2.Item("GCount"))
              Next
              Sendb("<a href=""lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
              Sendb(" (" & GCount & ")")
            Next
          End If
          Send("</li>")
          Send("</ul>")
        %>
        <h3>
          <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
        </h3>
        <%
          MyCommon.QueryStr = "select OT.TerminalTypeID as TID,T.Name from OfferTerminals as OT with (NoLock) left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                              "where Excluded=0 and OfferID=" & OfferID
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          Sendb("<ul class=""condensed"">")
          If rowCount > 0 Then
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CPE Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CM Terminals") Then
                Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
              Else
                Sendb("<li><a href=""terminal-edit.aspx?TerminalID=" & MyCommon.NZ(row.Item("TID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
              End If
            Next
          Else
            Sendb("<li>" & Copient.PhraseLib.Lookup("term.none", LanguageID))
          End If
          ' Check for any display any excluded terminals
          MyCommon.QueryStr = "select OT.TerminalTypeID as TID,T.Name from OfferTerminals as OT with (NoLock) left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                              "where Excluded=1 and OfferID=" & OfferID
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
            For Each row In rst.Rows
              x = x + 1
              Sendb("<a href=""terminal-edit.aspx?TerminalID=" & MyCommon.NZ(row.Item("TID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
              If x = (rst.Rows.Count - 1) Then
                Send(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
              ElseIf x < rst.Rows.Count Then
                Send(", ")
              Else
                Send("")
              End If
            Next
          Else
          End If
          Send("</li>")
          Send("</ul>")
        %>
        <hr class="hidden" />
      </div>
    </div>
          <div class="box" id="offernotifications">
      <%  
          If (LinksDisabled) Then
            Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.channels", LanguageID) & "</span></h2>")
          Else
            Send("<h2 style=""float:left;""><span><a href=""offer-channels.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.channels", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("notificationbody", "imgNotifications", "Notifications", True)
      %>
      <div id="notificationbody">
        <%
          MyCommon.QueryStr = "select distinct CO.ChannelID, CH.Name, CH.PhraseTerm, CO.StartDate, CO.EndDate  " & _
                              "from ChannelOffers as CO with (NoLock) " & _
                              "inner join Channels as CH with (NoLock) on CH.ChannelID = CO.ChannelID " & _
                              "where CO.OfferID = " & OfferID & "and StartDate is not null and EndDate is not null " & _
                              "union " & _
                              "select distinct CO.ChannelID, CH.Name, CH.PhraseTerm, CO.StartDate, CO.EndDate  " & _
                              "from ChannelOfferAssets as COA with (NoLock) " & _
                              "inner join ChannelOffers as CO with (NoLock) on CO.ChannelID = COA.ChannelID " & _
                              "inner join Channels as CH with (NoLock) on CH.ChannelID = CO.ChannelID " & _
                              "where CO.OfferID = " & OfferID & " and ISNULL(MediaData, '') <> '' " & _
                              "order by ChannelID; "
          rst2 = MyCommon.LRT_Select()
          If rst2.Rows.Count > 0 Then
            For Each row2 In rst2.Rows
              Send("<h3>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("PhraseTerm"), "").ToString, LanguageID, MyCommon.NZ(row2.Item("Name"), "")) & "</h3>")
              Sendb("<span style=""margin-left: 10px;"">")
              LongDate = MyCommon.NZ(row2.Item("StartDate"), "1/1/1900")
              If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
              Sendb(" - ")
              LongDate = MyCommon.NZ(row2.Item("EndDate"), "1/1/1900")
              If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
              Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(row2.Item("StartDate"), "1/1/1900"), MyCommon.NZ(row2.Item("EndDate"), "1/1/1900")) + 1 & " ")
              If DateDiff(DateInterval.Day, MyCommon.NZ(row2.Item("StartDate"), "1/1/1900"), MyCommon.NZ(row2.Item("EndDate"), "1/1/1900")) + 1 = 1 Then
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
              Else
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
              End If
              Send("</span>")
              Send("<br />")
              Send("<br class=""half"" />")
            Next
          Else
            Send(Copient.PhraseLib.Lookup("CPEoffer-sum.noChannelsConfigured", LanguageID))
          End If
          
        %>
        <hr class="hidden" />
      </div>
    </div>
<%      
      
      Send("<div class=""box"" id=""OptInConditions"">")
      If (LinksDisabled) Then
        Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.optinconditions", LanguageID) & "</span></h2>")
      Else
        Send("<h2 style=""float:left;""><span><a href=""Offer-con.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.optinconditions", LanguageID) & "</a></span></h2>")
      End If
  Send_BoxResizer("optinconditionbody", "imgOptInConditions", Copient.PhraseLib.Lookup("term.optinconditions", LanguageID), True)
      Send("<div id=""optinconditionbody"">")
      Dim conditionlength As Integer
      Dim customers As List(Of CMS.AMS.Models.Customer)
      Dim Offer As CMS.AMS.Models.Offer = m_Offer.GetOffer(OfferID, LoadOfferOptions.AllEligibilityConditions)
      ' Customer conditions
      If Offer.EligibleCustomerGroupConditions IsNot Nothing Then
        Send("<h3>" & Copient.PhraseLib.Lookup("term.customerconditions", LanguageID) & "</h3>")
        Send("<ul class=""condensed"">")
        Sendb("<li>")
        conditionlength = Offer.EligibleCustomerGroupConditions.IncludeCondition.Count
        For Each includeCondition As CustomerConditionDetails In Offer.EligibleCustomerGroupConditions.IncludeCondition
          If includeCondition.CustomerGroupID = 1 OrElse includeCondition.CustomerGroupID = 2 OrElse includeCondition.CustomerGroupID = 3 OrElse includeCondition.CustomerGroupID = 4 Then
            Sendb(MyCommon.SplitNonSpacedString(includeCondition.CustomerGroup.Name, 25))
            If conditionlength > 1 Then
              Sendb(" and ")
            Else
              Sendb(" ")
            End If
          Else
            Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & includeCondition.CustomerGroupID & """>" & MyCommon.SplitNonSpacedString(includeCondition.CustomerGroup.Name, 25) & "</a>")
            customers = m_customerGroup.GetCustomersByGroupID(includeCondition.CustomerGroupID)
            If customers.Count > 0 Then
              Sendb(" (" & customers.Count & ") " & IIf(conditionlength > 1, " or ", "") & "")
            ElseIf conditionlength > 1 Then
              Sendb(" or ")
            Else
              Sendb(" ")
            End If
          End If
          conditionlength = conditionlength - 1
        Next
          
        conditionlength = Offer.EligibleCustomerGroupConditions.ExcludeCondition.Count
        If conditionlength > 0 Then
          Sendb(Copient.PhraseLib.Lookup("term.excluding", LanguageID) & " ")
        End If
        For Each excludeCondition As CustomerConditionDetails In Offer.EligibleCustomerGroupConditions.ExcludeCondition
          If excludeCondition.CustomerGroupID = 1 OrElse excludeCondition.CustomerGroupID = 2 OrElse excludeCondition.CustomerGroupID = 3 OrElse excludeCondition.CustomerGroupID = 4 Then
            Sendb(MyCommon.SplitNonSpacedString(excludeCondition.CustomerGroup.Name, 25))
            If conditionlength > 1 Then
              Sendb(" and ")
            Else
              Sendb(" ")
            End If
          Else
            Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & excludeCondition.CustomerGroupID & """>" & MyCommon.SplitNonSpacedString(excludeCondition.CustomerGroup.Name, 25) & "</a>")
            customers = m_customerGroup.GetCustomersByGroupID(excludeCondition.CustomerGroupID)
            If customers.Count > 0 Then
              Sendb(" (" & customers.Count & ") " & IIf(conditionlength > 1, " and ", "") & "")
            ElseIf conditionlength > 1 Then
              Sendb(" and ")
            End If
          End If
          conditionlength = conditionlength - 1
        Next
        Send("</li>")
        Send("</ul>")
        Send("<br class=""half"" />")
      
        ' Points Eligibility Conditions
        If (Offer.EligiblePointsProgramConditions.Count > 0) Then
          Send("<h3>" & Copient.PhraseLib.Lookup("term.pointsconditions", LanguageID) & "</h3>")
          Send("<ul class=""condensed"">")
          For Each pointscondition As CMS.AMS.Models.PointsCondition In Offer.EligiblePointsProgramConditions
            Sendb("<li>")
            Sendb(CMS.Utilities.NZ(pointscondition.Quantity, 0))
            If (pointscondition.ProgramID > 0) Then
              If Popup Then
                Sendb(" " & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramName, ""), 25))
              Else
                Sendb(" <a href=""point-edit.aspx?ProgramGroupID=" & CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramID, "") & """>" & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramName, ""), 25) & "</a>")
              End If
            ElseIf (pointscondition.ProgramID = 0 AndAlso pointscondition.RequiredFromTemplate = False) Then
              Sendb(" <span class=""red"">")
              Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
              Sendb("</span>")
            Else
              Sendb(MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(pointscondition.PointsProgram.ProgramName, ""), 25))
            End If
            Send("</li>")
          Next
          Send("</ul>")
          Send("<br class=""half"" />")
        End If
      
        ' Stored Value Eligibility Conditions
        If (Offer.EligibleSVProgramConditions.Count > 0) Then
          Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalueconditions", LanguageID) & "</h3>")
          Send("<ul class=""condensed"">")
          For Each storedvaluecondition As CMS.AMS.Models.SVCondition In Offer.EligibleSVProgramConditions
            Sendb("<li>")
            Sendb(CMS.Utilities.NZ(storedvaluecondition.Quantity, 0))
            If (storedvaluecondition.SVProgramID > 0) Then
              If Popup Then
                Sendb(" " & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(storedvaluecondition.SVProgram.ProgramName, ""), 25))
              Else
                Sendb(" <a href=""SV-edit.aspx?ProgramGroupID=" & CMS.Utilities.NZ(storedvaluecondition.SVProgram.SVProgramID, "") & """>" & MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(storedvaluecondition.SVProgram.ProgramName, ""), 25) & "</a>")
              End If
            ElseIf (storedvaluecondition.SVProgramID = 0 AndAlso storedvaluecondition.RequiredFromTemplate = False) Then
              Sendb(" <span class=""red"">")
              Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
              Sendb("</span>")
            Else
              Sendb(MyCommon.SplitNonSpacedString(CMS.Utilities.NZ(storedvaluecondition.SVProgram.ProgramName, ""), 25))
            End If
            Send("</li>")
          Next
          Send("</ul>")
          Send("<br class=""half"" />")
        End If
      
      Else
        Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
      End If
      Send("<hr class=""hidden"" />")
      Send("</div>")
      Send("</div>")
    %>
    <div class="box" id="offerconditions">
      <% 
        If (LinksDisabled) Then
          Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & "</span></h2>")
        Else
          Send("<h2 style=""float:left;""><span><a href=""offer-con.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & "</a></span></h2>")
        End If
        Send_BoxResizer("conditionbody", "imgConditions", Copient.PhraseLib.Lookup("term.conditions", LanguageID), True)
      %>
      <div id="conditionbody">
        <%
          ' Customer conditions
          MyCommon.QueryStr = "select OfferID,Tiered,O.ConditionID,C.Description as ConditionDescription,O.ConditionTypeID, " & _
                              "O.GrantTypeID,G.Description as GrantDescription,G.PhraseID as GrantPhraseID,QtyUnitType,ConditionOrder,LinkID,CT.AmtRequired," & _
                              "CG.Name as Name,CG.PhraseID as CGPhraseID,CG.CustomerGroupID,CG.NewCardholders,ExcludedID,CGE.Name as ExcludedName,O.RequiredFromTemplate, " & _
                              "U.Description as UnitDescription,J.Description as JoinDescription,QtyUnitType from OfferConditions as O " & _
                              "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                              "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                              "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                              "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                              "left join CustomerGroups as CG with (nolock) on O.LinkID=CG.CustomerGroupID " & _
                              "left join CustomerGroups as CGE with (nolock) on O.ExcludedID=CGE.CustomerGroupID " & _
                              "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                              "where O.OfferID=" & OfferID & " and O.deleted=0 and O.ConditionTypeID=1 order by Tiered,ConditionOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          GCount = 0
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.customerconditions", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("<li>")
              If (MyCommon.NZ(row.Item("CustomerGroupID"), -1) > 2) And (MyCommon.NZ(row.Item("NewCardholders"), False) = False) Then
                Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>")
                If (MyCommon.NZ(row.Item("CGPhraseID"), 0) > 0) Then
                  Sendb(Copient.PhraseLib.Lookup(row.Item("CGPhraseID"), LanguageID))
                Else
                  Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                End If
                Send("</a>")
              ElseIf (IsDBNull(row.Item("CustomerGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                Sendb("<span class=""red"">* ")
                Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))
                Sendb(" " & Copient.PhraseLib.Lookup("term.by", LanguageID))
                Sendb(" " & Copient.PhraseLib.Lookup("term.template", LanguageID))
                Sendb("</span>")
              Else
                If MyCommon.NZ(row.Item("Name"), "") = "" Then
                  Sendb("<i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
                Else
                  Sendb(MyCommon.NZ(row.Item("Name"), ""))
                End If
              End If
              MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID = " & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & " And Deleted = 0"
              rst2 = MyCommon.LXS_Select()
              For Each row2 In rst2.Rows
                If (MyCommon.NZ(row.Item("CustomerGroupID"), -1) > 2) And (MyCommon.NZ(row.Item("NewCardholders"), False) = False) Then
                  Sendb(" (" & row2.Item("GCount") & ") ")
                Else
                  Sendb(" ")
                End If
              Next
              If MyCommon.NZ(row.Item("ExcludedID"), 0) > 0 Then
                Sendb(Copient.PhraseLib.Lookup("term.excluding", LanguageID) & " ")
                If MyCommon.NZ(row.Item("ExcludedID"), 0) <= 2 Then
                  Sendb(MyCommon.NZ(row.Item("ExcludedName"), ""))
                Else
                  Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("ExcludedID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ExcludedName"), ""), 25) & "</a>")
                  MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID = " & MyCommon.NZ(row.Item("ExcludedID"), -1) & " And Deleted = 0"
                  rst3 = MyCommon.LXS_Select()
                  Sendb(" (" & rst3.Rows(0).Item("GCount") & ") ")
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If

          ' Preference conditions
          If PrefManInstalled Then
            MyCommon.QueryStr = "select ConditionID, LinkID as PreferenceID from OfferConditions with (NoLock) " & _
                                "where OfferID=" & OfferID & " and ConditionTypeID=100 and Deleted=0"
            rst = MyCommon.LRT_Select
            i = 1
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.preferenceconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Send("  <li>")
                Send_Preference_Details(MyCommon, MyCommon.NZ(row.Item("PreferenceID"), 0))
                Send_Preference_Info(MyCommon, MyCommon.NZ(row.Item("ConditionID"), 0))
                  
                If i < rst.Rows.Count Then
                  Send(" <br /><i>" & Copient.PhraseLib.Lookup("term.and", LanguageID).ToLower & "</i> ")
                End If
                    
                Send("  </li>")
                i += 1
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
          End If
            
          ' Product conditions
          MyCommon.QueryStr = "select OfferID,Tiered,O.ConditionID,O.GrantTypeID," & _
                              "G.Description as GrantDescription,G.PhraseID as GrantPhraseID,O.QtyUnitType,ConditionOrder,CT.AmtRequired,RequiredFromTemplate, " & _
                              "PG.Name as Name,PG.ProductGroupID,PGE.Name as ExcludedName,PGE.ProductGroupID as ExcludedID," & _
                              "U.Description as UnitDescription,J.Description as JoinDescription from OfferConditions as O " & _
                              "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                              "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                              "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                              "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                              "left join ProductGroups as PG with (nolock) on O.LinkID=PG.ProductGroupID " & _
                              "left join ProductGroups as PGE with (nolock) on O.ExcludedID=PGE.ProductGroupID " & _
                              "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                              "where O.OfferID=" & OfferID & " and O.deleted=0 and O.ConditionTypeID=2 order by Tiered,ConditionOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          GCount = 0
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("ProductGroupID"), -1) = 1) Then
              AnyProductUsed = True
            End If
          Next
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.productconditions", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("<li>")
              ' Show required amounts
              If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.exactly", LanguageID) & " ")
              End If
              If MyCommon.NZ(row.Item("AmtRequired"), 0) > 0 Then
                If MyCommon.NZ(row.Item("QtyUnitType"), -1) = 1 Then
                  Sendb(Int(MyCommon.NZ(row.Item("AmtRequired"), -1)) & " ")
                  If MyCommon.NZ(row.Item("AmtRequired"), -1) = 1 Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.item", LanguageID), VbStrConv.Lowercase))
                  Else
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.items", LanguageID), VbStrConv.Lowercase))
                  End If
                ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 2 Then
                  Sendb(FormatCurrency(MyCommon.NZ(row.Item("AmtRequired"), -1)))
                ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 3 Then
                  Sendb(MyCommon.NZ(row.Item("AmtRequired"), -1) & " ")
                  If MyCommon.NZ(row.Item("AmtRequired"), -1) = 1 Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.pound", LanguageID), VbStrConv.Lowercase))
                  Else
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.pounds", LanguageID), VbStrConv.Lowercase))
                  End If
                ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 4 Then
                  Sendb(MyCommon.NZ(row.Item("AmtRequired"), -1) & " ")
                  If MyCommon.NZ(row.Item("AmtRequired"), -1) = 1 Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.gallon", LanguageID), VbStrConv.Lowercase))
                  Else
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.gallons", LanguageID), VbStrConv.Lowercase))
                  End If
                ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 6 Or MyCommon.NZ(row.Item("QtyUnitType"), -1) = 7 Or MyCommon.NZ(row.Item("QtyUnitType"), -1) = 10 Then
                  Sendb(FormatCurrency(MyCommon.NZ(row.Item("AmtRequired"), -1)))
                End If
              Else
                MyCommon.QueryStr = "Select AmtRequired from ConditionTiers with (NoLock) where ConditionID=" & MyCommon.NZ(row.Item("ConditionID"), -1)
                rst2 = MyCommon.LRT_Select
                TierAmtCount = 0
                If rst2.Rows.Count = 0 Then
                  Sendb("0")
                End If
                For Each row2 In rst2.Rows
                  If MyCommon.NZ(row.Item("QtyUnitType"), -1) = 1 Then
                    Sendb(Int(MyCommon.NZ(row2.Item("AmtRequired"), -1)) & " ")
                    If MyCommon.NZ(row2.Item("AmtRequired"), -1) = 1 Then
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.item", LanguageID), VbStrConv.Lowercase))
                    Else
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.items", LanguageID), VbStrConv.Lowercase))
                    End If
                  ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 2 Then
                    Sendb(FormatCurrency(MyCommon.NZ(row2.Item("AmtRequired"), -1)))
                  ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 3 Then
                    Sendb(MyCommon.NZ(row2.Item("AmtRequired"), -1) & " ")
                    If MyCommon.NZ(row2.Item("AmtRequired"), -1) = 1 Then
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.pound", LanguageID), VbStrConv.Lowercase))
                    Else
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.pounds", LanguageID), VbStrConv.Lowercase))
                    End If
                  ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 4 Then
                    Sendb(MyCommon.NZ(row2.Item("AmtRequired"), -1) & " ")
                    If MyCommon.NZ(row2.Item("AmtRequired"), -1) = 1 Then
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.gallon", LanguageID), VbStrConv.Lowercase))
                    Else
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.gallons", LanguageID), VbStrConv.Lowercase))
                    End If
                  ElseIf MyCommon.NZ(row.Item("QtyUnitType"), -1) = 6 Or MyCommon.NZ(row.Item("QtyUnitType"), -1) = 7 Or MyCommon.NZ(row.Item("QtyUnitType"), -1) = 10 Then
                    Sendb(FormatCurrency(MyCommon.NZ(row2.Item("AmtRequired"), -1)))
                  End If
                  TierAmtCount = TierAmtCount + 1
                  If TierAmtCount < rst2.Rows.Count Then
                    Sendb(" / ")
                  End If
                Next
              End If
              If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 2 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.ormore", LanguageID) & " ")
              End If
              Sendb(" " & Copient.PhraseLib.Lookup("term.of", LanguageID) & " ")
              ' Find and show the population of the group
              If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & MyCommon.NZ(row.Item("ProductGroupID"), -1) & " And Deleted = 0"
                rst2 = MyCommon.LRT_Select()
                For Each row2 In rst2.Rows
                  If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                    Sendb(" (" & row2.Item("GCount") & ") ")
                  Else
                    Sendb(" ")
                  End If
                Next
              ElseIf (MyCommon.NZ(row.Item("ProductGroupID"), -1) = 1) Then
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.anyproduct", LanguageID), VbStrConv.Lowercase))
              Else
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID), VbStrConv.Lowercase))
              End If
              ' Show excluded group and its population
              If MyCommon.NZ(row.Item("ExcludedID"), 0) > 0 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & " ")
                Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ExcludedID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ExcludedName"), ""), 25) & "</a>")
                MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & MyCommon.NZ(row.Item("ExcludedID"), "-1") & " And Deleted = 0"
                rst3 = MyCommon.LRT_Select()
                Sendb(" (" & rst3.Rows(0).Item("GCount") & ") ")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Points conditions
          MyCommon.QueryStr = "select OfferID,Tiered,O.ConditionID,PTS.ProgramName,PTS.ProgramID,O.GrantTypeID,G.Description as GrantDescription,G.PhraseID as GrantPhraseID,ConditionOrder,CT.AmtRequired," & _
                              "U.Description as UnitDescription,J.Description as JoinDescription,QtyUnitType, RequiredFromTemplate from OfferConditions as O " & _
                              "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                              "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                              "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                              "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                              "left join PointsPrograms as PTS with (nolock) on O.LinkID=PTS.ProgramID " & _
                              "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                              "where O.OfferID=" & OfferID & " and O.deleted=0 and O.ConditionTypeID=3 order by Tiered,ConditionOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.pointsconditions", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("<li>")
              If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.exactly", LanguageID) & " ")
              End If
              MyCommon.QueryStr = "Select AmtRequired from ConditionTiers with (NoLock) where ConditionID=" & MyCommon.NZ(row.Item("ConditionID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("0")
              Else
                For Each row2 In rst2.Rows
                  Sendb(Int(MyCommon.NZ(row2.Item("AmtRequired"), -1)))
                  TierAmtCount = TierAmtCount + 1
                  If TierAmtCount < rst2.Rows.Count Then
                    Sendb(" / ")
                  End If
                Next
                If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 2 Then
                  Sendb(" " & Copient.PhraseLib.Lookup("term.ormore", LanguageID) & " ")
                End If
                Sendb(" <a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Stored Value conditions
          MyCommon.QueryStr = "select OfferID,Tiered,O.ConditionID,SVP.SVProgramID as ProgramID,SVP.Name as ProgramName,SVP.SVTypeID,O.GrantTypeID,G.Description as GrantDescription,G.PhraseID as GrantPhraseID,ConditionOrder,CT.AmtRequired," & _
                              "U.Description as UnitDescription,J.Description as JoinDescription,QtyUnitType, RequiredFromTemplate from OfferConditions as O " & _
                              "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                              "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                              "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                              "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                              "left join StoredValuePrograms as SVP with (nolock) on O.LinkID=SVP.SVProgramID " & _
                              "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                              "where O.OfferID=" & OfferID & " and O.deleted=0 and O.ConditionTypeID=6 order by Tiered,ConditionOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalueconditions", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              bNeedToFormat = False
              If intNumDecimalPlaces > 0 Then
                If Int(MyCommon.NZ(row.Item("SVTypeID"), 0)) = 1 Then
                  bNeedToFormat = True
                End If
              End If
              Sendb("<li>")
              If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.exactly", LanguageID) & " ")
              End If
              MyCommon.QueryStr = "Select AmtRequired from ConditionTiers with (NoLock) where ConditionID=" & MyCommon.NZ(row.Item("ConditionID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("0")
              Else
                For Each row2 In rst2.Rows
                  If bNeedToFormat Then
                    decTemp = (Int(MyCommon.NZ(row2.Item("AmtRequired"), 0)) * 1.0) / decFactor
                    sTemp1 = FormatNumber(decTemp, intNumDecimalPlaces)
                    Send(sTemp1)
                  Else
                    Sendb(Int(MyCommon.NZ(row2.Item("AmtRequired"), 0)))
                  End If
                  TierAmtCount = TierAmtCount + 1
                  If TierAmtCount < rst2.Rows.Count Then
                    Sendb(" / ")
                  End If
                Next
                If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 2 Then
                  Sendb(" " & Copient.PhraseLib.Lookup("term.ormore", LanguageID) & " ")
                End If
                Sendb(" <a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Advanced Limit conditions
          MyCommon.QueryStr = "select OfferID,Tiered,O.ConditionID,AL.Name,AL.LimitID,O.GrantTypeID,G.Description as GrantDescription,G.PhraseID as GrantPhraseID,ConditionOrder,CT.AmtRequired," & _
                              "U.Description as UnitDescription,J.Description as JoinDescription,QtyUnitType, RequiredFromTemplate from OfferConditions as O " & _
                              "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                              "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                              "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                              "left join UnitTypes as U with (nolock) on O.QtyUnitType=U.UnitTypeID " & _
                              "left join CM_AdvancedLimits as AL with (nolock) on O.LinkID=AL.LimitID " & _
                              "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                              "where O.OfferID=" & OfferID & " and O.deleted=0 and O.ConditionTypeID=7 order by Tiered,ConditionOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.advlimitconditions", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("<li>")
              If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.exactly", LanguageID) & " ")
              End If
              MyCommon.QueryStr = "Select AmtRequired from ConditionTiers with (NoLock) where ConditionID=" & MyCommon.NZ(row.Item("ConditionID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("0")
              Else
                For Each row2 In rst2.Rows
                  Sendb(Int(MyCommon.NZ(row2.Item("AmtRequired"), -1)))
                  TierAmtCount = TierAmtCount + 1
                  If TierAmtCount < rst2.Rows.Count Then
                    Sendb(" / ")
                  End If
                Next
                If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 2 Then
                  Sendb(" " & Copient.PhraseLib.Lookup("term.ormore", LanguageID) & " ")
                End If
                Sendb(" <a href=""CM-advlimit-edit.aspx?LimitID=" & MyCommon.NZ(row.Item("LimitID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Tender conditions
          MyCommon.QueryStr = "select OfferID,Tiered,O.ConditionID,ConditionOrder,CT.AmtRequired,CTT.TenderTypeID,TT.Description,TT.ExtTenderCode," & _
                              "O.GrantTypeID,G.Description as GrantDescription,G.PhraseID as GrantPhraseID,J.Description as JoinDescription from OfferConditions as O with (nolock)" & _
                              "left join GrantTypes as G with (nolock) on O.GrantTypeID=G.GrantTypeID " & _
                              "left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                              "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                              "left join ConditionTenderTypes as CTT with (nolock) on CTT.ConditionID=O.ConditionID " & _
                              "left join TenderTypes as TT with (nolock) on TT.TenderTypeID=CTT.TenderTypeID " & _
                              "left join ConditionTiers as CT with (nolock) on O.ConditionID=CT.ConditionID and CT.TierLevel=0 " & _
                              "where O.OfferID=" & OfferID & " and O.deleted=0 and O.ConditionTypeID=4 order by ConditionOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.tenderconditions", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("<li>")
              If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 1 Then
                Sendb(Copient.PhraseLib.Lookup("term.exactly", LanguageID) & " ")
              End If
              MyCommon.QueryStr = "Select AmtRequired from ConditionTiers with (NoLock) where ConditionID=" & MyCommon.NZ(row.Item("ConditionID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              For Each row2 In rst2.Rows
                Sendb(FormatCurrency(MyCommon.NZ(row2.Item("AmtRequired"), -1)))
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb(" / ")
                End If
              Next
              If MyCommon.NZ(row.Item("GrantTypeID"), -1) = 2 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.ormore", LanguageID))
              End If
              If MyCommon.NZ(row.Item("Description"), "") <> "" Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.in", LanguageID))
                Sendb(" <a href=""tender.aspx"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Description"), ""), 25) & "</a>")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Day/Time conditions
          MyCommon.QueryStr = "select OfferID,Tiered,ConditionOrder, " & _
                              "CT.StartHour,CT.StartMinute,CT.EndHour,CT.EndMinute,CT.Sunday,CT.Monday,CT.Tuesday,CT.Wednesday,CT.Thursday,CT.Friday,CT.Saturday, " & _
                              "J.Description as JoinDescription from OfferConditions as O " & _
                              "with (nolock) left join ConditionTypes as C with (nolock) on O.ConditionTypeID=C.ConditionTypeID " & _
                              "left join JoinTypes as J with (nolock) on O.JoinTypeID=J.JoinTypeID " & _
                              "left join ConditionTimes as CT with (nolock) on CT.ConditionID=O.ConditionID " & _
                              "where O.OfferID=" & OfferID & " and O.deleted=0 and O.ConditionTypeID=5 order by ConditionOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.timeconditions", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("<li>")
              Dim DaysTrue As Integer = 0
              If MyCommon.NZ(row.Item("Sunday"), False) And MyCommon.NZ(row.Item("Monday"), False) And MyCommon.NZ(row.Item("Tuesday"), False) And MyCommon.NZ(row.Item("Wednesday"), False) _
              And MyCommon.NZ(row.Item("Thursday"), False) And MyCommon.NZ(row.Item("Friday"), False) And MyCommon.NZ(row.Item("Saturday"), False) Then
                Send(Copient.PhraseLib.Lookup("term.everyday", LanguageID))
                DaysTrue = 7
              Else
                If (MyCommon.NZ(row.Item("Sunday"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.sunday", LanguageID) & " ")
                  DaysTrue = DaysTrue + 1
                End If
                If (MyCommon.NZ(row.Item("Monday"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.monday", LanguageID) & " ")
                  DaysTrue = DaysTrue + 1
                End If
                If (MyCommon.NZ(row.Item("Tuesday"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.tuesday", LanguageID) & " ")
                  DaysTrue = DaysTrue + 1
                End If
                If (MyCommon.NZ(row.Item("Wednesday"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.wednesday", LanguageID) & " ")
                  DaysTrue = DaysTrue + 1
                End If
                If (MyCommon.NZ(row.Item("Thursday"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.thursday", LanguageID) & " ")
                  DaysTrue = DaysTrue + 1
                End If
                If (MyCommon.NZ(row.Item("Friday"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.friday", LanguageID) & " ")
                  DaysTrue = DaysTrue + 1
                End If
                If (MyCommon.NZ(row.Item("Saturday"), False)) Then
                  Sendb(Copient.PhraseLib.Lookup("term.saturday", LanguageID) & "&nbsp;")
                  DaysTrue = DaysTrue + 1
                End If
              End If
              If (DaysTrue > 0) Then
                Sendb("<br />")
              End If
              Dim StartHour, EndHour, StartMinute, EndMinute As String
              StartHour = MyCommon.NZ(row.Item("StartHour"), "").ToString.PadLeft(2, "0")
              EndHour = MyCommon.NZ(row.Item("EndHour"), "").ToString.PadLeft(2, "0")
              StartMinute = MyCommon.NZ(row.Item("StartMinute"), "").ToString.PadLeft(2, "0")
              EndMinute = MyCommon.NZ(row.Item("EndMinute"), "").ToString.PadLeft(2, "0")
              If (StartHour = "00") And (StartMinute = "00") And (EndHour = "00") And (EndMinute = "00") Then
                If DaysTrue = 0 Then
                  Sendb("<i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
                End If
              Else
                Sendb(StartHour & ":" & StartMinute)
                Sendb(" - ")
                Sendb(EndHour & ":" & EndMinute)
              End If
              Send("</li>")
            Next
            Send("  </ul>")
          End If
            
          If (counter = 0) Then
            Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
          End If
          counter = 0
        %>
        <hr class="hidden" />
      </div>
    </div>
    <div class="box" id="offerrewards">
      <%  
        If (LinksDisabled) Then
          Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & "</span></h2>")
        Else
          Send("<h2 style=""float:left;""><span><a href=""offer-rew.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & "</a></span></h2>")
        End If
        Send_BoxResizer("rewardbody", "imgRewards", Copient.PhraseLib.Lookup("term.rewards", LanguageID), True)
      %>
      <div id="rewardbody">
        <%
          ' Discount rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardAmountTypeID, " & _
                              "R.TriggerQty,R.RewardLimit,R.RewardLimitTypeID,R.ApplyToLimit,R.SponsorID, " & _
                              "R.UseSpecialPricing, R.SPRepeatAtOccur, R.RewardDistPeriod, R.RewardDistLimitVarID, " & _
                              "R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGNAME,EPG.Name as EPGNAME " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=1 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.discounts", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select RewardAmount from RewardTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              For Each row2 In rst2.Rows
                If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 10 Then
                  Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) <> 7 Then
                  If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 2 Or MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 9 Then
                    Sendb(row2.Item("RewardAmount") & "%")
                  Else
                    Sendb(FormatCurrency(row2.Item("RewardAmount")))
                  End If
                  TierAmtCount = TierAmtCount + 1
                  If TierAmtCount < rst2.Rows.Count Then
                    Sendb(" / ")
                  End If
                End If
              Next
              If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 7 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.free", LanguageID) & " ")
              ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 5 Then
                Sendb(" ")
              Else
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & " ")
              End If
              If MyCommon.NZ(row.Item("PGID"), -1) = 0 Then
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.entiretransaction", LanguageID), VbStrConv.Lowercase))
              ElseIf MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                Sendb(StrConv(MyCommon.NZ(row.Item("PGName"), ""), VbStrConv.Lowercase))
              Else
                Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
              End If
              If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 3 OrElse  MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 6  Then
                Sendb(" (" & Copient.PhraseLib.Lookup("term.perunitweight", LanguageID) & ")")
              ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 4 Then
                Sendb(" (" & Copient.PhraseLib.Lookup("term.perunitvolume", LanguageID) & ")")
              End If
              If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
              End If
              If MyCommon.NZ(row.Item("RewardLimit"), -1) = 0 Then
                Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.unlimited", LanguageID), VbStrConv.Lowercase))
              Else
                Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & " ")
                If MyCommon.NZ(row.Item("RewardLimitTypeID"), -1) = 3 Then
                  Sendb(MyCommon.NZ(row.Item("RewardLimit"), -1) & " " & StrConv(Copient.PhraseLib.Lookup("term.pounds", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardLimitTypeID"), -1) = 4 Then
                  Sendb(MyCommon.NZ(row.Item("RewardLimit"), -1) & " " & StrConv(Copient.PhraseLib.Lookup("term.gallons", LanguageID), VbStrConv.Lowercase))
                Else
                  Sendb(FormatCurrency(MyCommon.NZ(row.Item("RewardLimit"), -1)))
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Points rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardAmountTypeID,RP.ProgramID,PP.ProgramName," & _
                              "R.ApplyToLimit,R.TriggerQty,R.RewardLimit,R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGNAME,EPG.Name as EPGNAME " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "left join RewardPoints as RP with (NoLock) on RP.RewardPointsID=R.LinkID " & _
                              "left join PointsPrograms as PP with (NoLock) on PP.ProgramID=RP.ProgramID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=2 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select RewardAmount from RewardTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              For Each row2 In rst2.Rows
                Sendb(Int(MyCommon.NZ(row2.Item("RewardAmount"), 0)))
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb("/")
                End If
              Next
              If MyCommon.NZ(row.Item("ProgramName"), "") = "" Then
                Sendb(" <i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
              Else
                Sendb(" <a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
              End If
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 1 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.item", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 2 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.dollar", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 3 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.pound", LanguageID), VbStrConv.Lowercase))
                End If
                Send("<br />")
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(MyCommon.NZ(row.Item("PGName"), "") & " ")
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("RewardLimit"), -1) > 0 Then
                  Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
                End If
                If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                  Send("<br />")
                  Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                  If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                    Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetsPoints", LanguageID, MyCommon.NZ(row.Item("ApplyToLimit"), -1)))
                  End If
                  Sendb(")")
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Tender Points rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardAmountTypeID,RP.ProgramID,PP.ProgramName," & _
                              "R.ApplyToLimit,R.TriggerQty,R.RewardLimit,R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGNAME,EPG.Name as EPGNAME " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "left join RewardPoints as RP with (NoLock) on RP.RewardPointsID=R.LinkID " & _
                              "left join PointsPrograms as PP with (NoLock) on PP.ProgramID=RP.ProgramID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=13 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.tenderpoints", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select RewardAmount from RewardTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              For Each row2 In rst2.Rows
                Sendb(Int(MyCommon.NZ(row2.Item("RewardAmount"), 0)))
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb("/")
                End If
              Next
              If MyCommon.NZ(row.Item("ProgramName"), "") = "" Then
                Sendb(" <i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
              Else
                Sendb(" <a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
              End If
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 1 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.item", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 2 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.dollar", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 3 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.pound", LanguageID), VbStrConv.Lowercase))
                End If
                Send("<br />")
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(MyCommon.NZ(row.Item("PGName"), "") & " ")
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("RewardLimit"), -1) > 0 Then
                  Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
                End If
                If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                  Send("<br />")
                  Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                  If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                    Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetsPoints", LanguageID, MyCommon.NZ(row.Item("ApplyToLimit"), -1)))
                  End If
                  Sendb(")")
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Stored Value rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardAmountTypeID,RSV.ProgramID,RSV.SVTypeID,SV.Name as ProgramName," & _
                              "R.ApplyToLimit,R.TriggerQty,R.RewardLimit,R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGNAME,EPG.Name as EPGNAME " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "left join CM_RewardStoredValues as RSV with (NoLock) on RSV.RewardStoredValuesID=R.LinkID " & _
                              "left join StoredValuePrograms as SV with (NoLock) on SV.SVProgramID=RSV.ProgramID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=10 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              bNeedToFormat = False
              If intNumDecimalPlaces > 0 Then
                If Int(MyCommon.NZ(row.Item("SVTypeID"), 0)) = 1 Then
                  bNeedToFormat = True
                End If
              End If
              Sendb("  <li>")
              MyCommon.QueryStr = "select RewardAmount from RewardTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              For Each row2 In rst2.Rows
                If bNeedToFormat Then
                  decTemp = (Int(MyCommon.NZ(row2.Item("RewardAmount"), 0)) * 1.0) / decFactor
                  sTemp1 = FormatNumber(decTemp, intNumDecimalPlaces)
                  Send(sTemp1)
                Else
                  Sendb(Int(MyCommon.NZ(row2.Item("RewardAmount"), 0)))
                End If
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb("/")
                End If
              Next
              If MyCommon.NZ(row.Item("ProgramName"), "") = "" Then
                Sendb(" <i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
              Else
                Sendb(" <a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
              End If
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 1 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.item", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 2 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.dollar", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 3 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.pound", LanguageID), VbStrConv.Lowercase))
                End If
                Send("<br />")
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(MyCommon.NZ(row.Item("PGName"), "") & " ")
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("RewardLimit"), -1) > 0 Then
                  Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
                End If
                If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                  Send("<br />")
                  Sendb(" (" & MyCommon.NZ(row.Item("TriggerQty"), -1) & " items required")
                  If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                    Sendb(", " & MyCommon.NZ(row.Item("ApplyToLimit"), -1) & " " & IIf(MyCommon.NZ(row.Item("ApplyToLimit"), -1) = 1, "gets", "get") & " value")
                  End If
                  Sendb(")")
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If

          ' Advanced Limit rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardAmountTypeID,AL.LimitID,AL.Name," & _
                              "R.ApplyToLimit,R.TriggerQty,R.RewardLimit,R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGNAME,EPG.Name as EPGNAME " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "left join CM_RewardAdvancedLimits as RAL with (NoLock) on RAL.RewardAdvLimitID=R.LinkID " & _
                              "left join CM_AdvancedLimits as AL with (NoLock) on AL.LimitID=RAL.LimitID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=12 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.advlimits", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select RewardAmount from RewardTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              For Each row2 In rst2.Rows
                Sendb(Int(MyCommon.NZ(row2.Item("RewardAmount"), 0)))
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb("/")
                End If
              Next
              If MyCommon.NZ(row.Item("Name"), "") = "" Then
                Sendb(" <i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
              Else
                Sendb(" <a href=""CM-advlimit-edit.aspx?LimitID=" & MyCommon.NZ(row.Item("LimitID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
              End If
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                If MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 1 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.item", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 2 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.dollar", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardAmountTypeID"), -1) = 3 Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.pound", LanguageID), VbStrConv.Lowercase))
                End If
                Send("<br />")
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(MyCommon.NZ(row.Item("PGName"), "") & " ")
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("RewardLimit"), -1) > 0 Then
                  Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
                End If
                If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                  Send("<br />")
                  Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                  If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                    Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetsPoints", LanguageID, MyCommon.NZ(row.Item("ApplyToLimit"), -1)))
                  End If
                  Sendb(")")
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Printed message rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardLimit,R.TriggerQty,R.ApplyToLimit," & _
                              "R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGName,EPG.Name as EPGName " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=3 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.printedmessages", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select BodyText from PrintedMessageTiers with (nolock) where MessageID=" & MyCommon.NZ(row.Item("LinkID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              Dim Details As StringBuilder
              If rst2.Rows.Count = 0 Then
                Sendb("<i>" & Copient.PhraseLib.Lookup("term.empty", LanguageID) & "</i>")
              End If
              For Each row2 In rst2.Rows
                Details = New StringBuilder(200)
                Details.Append(ReplaceTags(MyCommon.NZ(row2.Item("BodyText"), "")))
                If (Details.ToString().Length > 41) Then
                  Details = Details.Remove(38, (Details.Length - 38))
                  Details.Append("...")
                End If
                Details.Replace(vbCrLf, "<br />")
                If Details.ToString = "" Then
                  Sendb("<i>" & Copient.PhraseLib.Lookup("term.empty", LanguageID) & "</i>")
                Else
                  Sendb("""" & MyCommon.SplitNonSpacedString(Details.ToString, 25) & """")
                End If
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb(" /<br />")
                End If
              Next
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Send("<br />")
                Sendb(Copient.PhraseLib.Lookup("term.requires", LanguageID) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(StrConv(MyCommon.NZ(row.Item("PGName"), ""), VbStrConv.Lowercase))
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & row.Item("PGID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                Sendb(".")
              End If
              If MyCommon.NZ(row.Item("RewardLimit"), -1) > 0 Then
                Send(" &nbsp;" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
              End If
              If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                Send("<br />")
                Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetsPoints", LanguageID, MyCommon.NZ(row.Item("ApplyToLimit"), -1)))
                End If
                Sendb(")")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Cashier message rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardLimit,R.TriggerQty,R.ApplyToLimit," & _
                              "R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGName,EPG.Name as EPGName " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "left join CashierMessages as CM with (NoLock) on LinkID=CM.MessageID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=4 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.cashiermessages", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select Line1Text from CashierMessageTiers with (nolock) where MessageID=" & MyCommon.NZ(row.Item("LinkID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("<i>" & Copient.PhraseLib.Lookup("term.empty", LanguageID) & "</i>")
              End If
              For Each row2 In rst2.Rows
                If MyCommon.NZ(row2.Item("Line1Text"), "") <> "" Then
                  Sendb("""" & MyCommon.SplitNonSpacedString(row2.Item("Line1Text"), 25) & """")
                Else
                  Sendb("")
                End If
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb(" /<br />")
                End If
              Next
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Send("<br />")
                Sendb(Copient.PhraseLib.Lookup("term.requires", LanguageID) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(StrConv(MyCommon.NZ(row.Item("PGName"), ""), VbStrConv.Lowercase))
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                Sendb(".")
              End If
              If MyCommon.NZ(row.Item("RewardLimit"), 0) > 0 Then
                Send(" &nbsp;" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
              End If
              If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                Send("<br />")
                Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetsPoints", LanguageID, MyCommon.NZ(row.Item("ApplyToLimit"), -1)))
                End If
                Sendb(")")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Group membership rewards
          MyCommon.QueryStr = "select R.RewardID,R.Tiered,R.RewardOrder,R.RewardTypeID " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join RewardTiers as RTiers with (nolock) on R.RewardID=RTiers.RewardID and RTiers.TierLevel=0 " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and (R.RewardTypeID=5 or R.RewardTypeID=6) order by RewardTypeID,RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.membership", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "Select RC.RewardID,RC.CustomerGroupID,CG.Name,CG.CustomerGroupID from RewardCustomerGroupTiers as RC with (NoLock) " & _
                                  "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=RC.CustomerGroupID " & _
                                  "where RC.RewardID=" & row.Item("RewardID") & " order by TierLevel"
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("<i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
              Else
                For Each row2 In rst2.Rows
                  Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & row2.Item("CustomerGroupID") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row2.Item("Name"), ""), 25) & "</a>")
                  TierAmtCount = TierAmtCount + 1
                  If TierAmtCount < rst2.Rows.Count Then
                    Sendb(" / ")
                  End If
                Next
                If MyCommon.NZ(row.Item("RewardTypeID"), -1) = 5 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.granted", LanguageID), VbStrConv.Lowercase))
                ElseIf MyCommon.NZ(row.Item("RewardTypeID"), -1) = 6 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.revoked", LanguageID), VbStrConv.Lowercase))
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Catalina coupon rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardLimit,R.TriggerQty,R.ApplyToLimit," & _
                              "R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGName,EPG.Name as EPGName " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=8 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.catalinacoupons", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select XmlText, TierLevel from RewardXmlTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("<i>" & Copient.PhraseLib.Lookup("term.empty", LanguageID) & "</i>")
              End If
              For Each row2 In rst2.Rows
                Sendb(Copient.PhraseLib.Lookup("term.mclu", LanguageID) & " " & row2.Item("XmlText"))
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb(" / ")
                End If
              Next
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Send("<br />")
                Sendb(Copient.PhraseLib.Lookup("term.requires", LanguageID) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(StrConv(MyCommon.NZ(row.Item("PGName"), ""), VbStrConv.Lowercase))
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                Sendb(".")
              End If
              If MyCommon.NZ(row.Item("RewardLimit"), 0) > 0 Then
                Send("<br />" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
              End If
              If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                Send("<br />")
                Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetCoupon", LanguageID, row.Item("ApplyToLimit")))
                End If
                Sendb(")")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Cents Off fuel rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardAmountTypeID,RSV.ProgramID,SV.Name as ProgramName," & _
                              "R.ApplyToLimit,R.TriggerQty,R.RewardLimit,R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGNAME,EPG.Name as EPGNAME " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "left join CM_RewardStoredValues as RSV with (NoLock) on RSV.RewardStoredValuesID=R.LinkID " & _
                              "left join StoredValuePrograms as SV with (NoLock) on SV.SVProgramID=RSV.ProgramID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=14 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.centsofffuel", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              If intNumDecimalPlaces > 0 Then
                decTemp = (Int(MyCommon.NZ(row.Item("TriggerQty"), 0)) * 1.0) / decFactor
                sTemp1 = FormatNumber(decTemp, intNumDecimalPlaces)
              Else
                sTemp1 = Int(MyCommon.NZ(row.Item("TriggerQty"), 0)).ToString()
              End If
                
              decTemp = (Int(MyCommon.NZ(row.Item("ApplyToLimit"), 0)) * 1.0) / 100.0
              sTemp2 = decTemp.ToString("0.00")

              Send("$" & sTemp2 & " " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & " " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " " & sTemp1 & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & " ")
              If MyCommon.NZ(row.Item("ProgramName"), "") = "" Then
                Sendb(" <i>" & Copient.PhraseLib.Lookup("term.nothing", LanguageID) & "</i>")
              Else
                Sendb(" <a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
              End If
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Send("<br />")
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(MyCommon.NZ(row.Item("PGName"), "") & " ")
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("RewardLimit"), -1) > 0 Then
                  decTemp = MyCommon.NZ(row.Item("RewardLimit"), 0.0)
                  sTemp2 = decTemp.ToString("0.00")
                  Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & " $" & sTemp2)
                End If
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If


          ' Generic XML passthrough rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardLimit,R.TriggerQty,R.ApplyToLimit," & _
                              "R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGName,EPG.Name as EPGName " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=7 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.xmlpassthroughs", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select XmlText, TierLevel from RewardXmlTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("<i>" & Copient.PhraseLib.Lookup("term.empty", LanguageID) & "</i>")
              End If
              'For Each row2 In rst2.Rows
              '  Sendb("""" & row2.Item("XmlText") & """")
              '  TierAmtCount = TierAmtCount + 1
              '  If TierAmtCount < rst2.Rows.Count Then
              '    Sendb(" / ")
              '  End If
              'Next
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Send("<br />")
                Sendb(Copient.PhraseLib.Lookup("term.requires", LanguageID) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(StrConv(MyCommon.NZ(row.Item("PGName"), ""), VbStrConv.Lowercase))
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                Sendb(".")
              End If
              If MyCommon.NZ(row.Item("RewardLimit"), 0) > 0 Then
                Send("<br />" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
              End If
              If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                Send("<br />")
                Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetReward", LanguageID, row.Item("ApplyToLimit")))
                End If
                Sendb(")")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' Bin range rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardLimit,R.TriggerQty,R.ApplyToLimit," & _
                              "R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGName,EPG.Name as EPGName " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=9 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.binranges", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              MyCommon.QueryStr = "select XmlText, TierLevel from RewardXmlTiers with (nolock) where RewardID=" & MyCommon.NZ(row.Item("RewardID"), -1)
              rst2 = MyCommon.LRT_Select
              TierAmtCount = 0
              If rst2.Rows.Count = 0 Then
                Sendb("<i>" & Copient.PhraseLib.Lookup("term.empty", LanguageID) & "</i>")
              End If
              For Each row2 In rst2.Rows
                Sendb(row2.Item("XmlText"))
                TierAmtCount = TierAmtCount + 1
                If TierAmtCount < rst2.Rows.Count Then
                  Sendb(" / ")
                End If
              Next
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Send("<br />")
                Sendb(Copient.PhraseLib.Lookup("term.requires", LanguageID) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(StrConv(MyCommon.NZ(row.Item("PGName"), ""), VbStrConv.Lowercase))
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                Sendb(".")
              End If
              If MyCommon.NZ(row.Item("RewardLimit"), 0) > 0 Then
                Send("<br />" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
              End If
              If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                Send("<br />")
                Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetReward", LanguageID, row.Item("ApplyToLimit")))
                End If
                Sendb(")")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If
            
          ' auto Print Gift Receipt rewards
          MyCommon.QueryStr = "select R.RewardID,R.OfferID,R.Tiered,R.RewardOrder,R.RewardTypeID,R.LinkID,R.RewardLimit,R.TriggerQty,R.ApplyToLimit," & _
                              "R.ProductGroupID as PGID,R.ExcludedProdGroupID as EPGID,PG.Name as PGName,EPG.Name as EPGName " & _
                              "from OfferRewards as R with (NoLock) " & _
                              "left join AmountTypes as AMT with (NoLock) on R.RewardAmountTypeID=AMT.AmountTypeID " & _
                              "left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=R.ProductGroupID " & _
                              "left join ProductGroups as EPG with (NoLock) on EPG.ProductGroupID=R.ExcludedProdGroupID " & _
                              "left join RewardTypes as RT with (NoLock) on RT.RewardTypeID=R.RewardTypeID " & _
                              "where OfferID=" & OfferID & " and R.deleted=0 and R.RewardTypeID=11 order by RewardOrder"
          rst = MyCommon.LRT_Select
          rowCount = rst.Rows.Count
          If rowCount > 0 Then
            counter = counter + 1
            Send("<h3>" & Copient.PhraseLib.Lookup("term.autogiftreceipt", LanguageID) & "</h3>")
            Send("  <ul class=""condensed"">")
            For Each row In rst.Rows
              Sendb("  <li>")
              Sendb("<i>" & Copient.PhraseLib.Lookup("term.autogiftreceipt", LanguageID))
              If MyCommon.NZ(row.Item("PGID"), -1) > 0 Then
                Send("<br />")
                Sendb(Copient.PhraseLib.Lookup("term.requires", LanguageID) & " ")
                If MyCommon.NZ(row.Item("PGID"), -1) = 1 Then
                  Sendb(StrConv(MyCommon.NZ(row.Item("PGName"), ""), VbStrConv.Lowercase))
                Else
                  Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("PGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("PGName"), ""), 25) & "</a>")
                End If
                If MyCommon.NZ(row.Item("EPGID"), -1) > 0 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase))
                  Sendb(" <a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("EPGID"), -1) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("EPGName"), ""), 25) & "</a>")
                End If
                Sendb(".")
              End If
              If MyCommon.NZ(row.Item("RewardLimit"), 0) > 0 Then
                Send("<br />" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " " & Int(MyCommon.NZ(row.Item("RewardLimit"), -1)) & ".")
              End If
              If MyCommon.NZ(row.Item("TriggerQty"), 1) > 1 Then
                Send("<br />")
                Sendb(" (" & Copient.PhraseLib.Detokenize("offer-sum.ItemsRequired", LanguageID, MyCommon.NZ(row.Item("TriggerQty"), -1)))
                If (MyCommon.NZ(row.Item("TriggerQty"), 0) > MyCommon.NZ(row.Item("ApplyToLimit"), 0)) Then
                  Sendb(Copient.PhraseLib.Detokenize("offer-sum.GetCoupon", LanguageID, row.Item("ApplyToLimit")))
                End If
                Sendb(")")
              End If
              Send("</li>")
            Next
            Send("  </ul>")
            Send("  <br class=""half"" />")
          End If

          If (counter = 0) Then
            Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
          End If
          counter = 0
        %>
        <hr class="hidden" />
      </div>
    </div>
    <% If bShelfLabelEnabled Then%>
    <div class="box" id="shelflabel">
      <%      
        Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.shelflabel", LanguageID) & "</span></h2>&nbsp")
        Send_BoxResizer("shelflabelbody", "imgShelfLabel", Copient.PhraseLib.Lookup("term.shelflabel", LanguageID), True)
      %>
      <div id="shelflabelbody">
        <% Sendb("<h3>" & Copient.PhraseLib.Lookup("term.lastshelflabelextract", LanguageID) & ":</h3>")%>
        <%
          MyCommon.QueryStr = "select LastShelfLabelExtract from OfferAccessoryFields with (NoLock) where OfferID=" & OfferID & ";"
          rst4 = MyCommon.LRT_Select()
          If rst4.Rows.Count > 0 Then
            LongDate = MyCommon.NZ(rst4.Rows(0).Item("LastShelfLabelExtract"), "1/1/1900")
            If LongDate > "1/1/1900" Then 
              Sendb(Logix.ToShortDateTimeString(LongDate, MyCommon))
            Else 
              Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
            End If
          Else
            Sendb(Copient.PhraseLib.Lookup("term.never", LanguageID))
          End If
        %>
      </div>
    </div>
    <% End If%>
  </div>
  <br clear="all" />
</div>
</form>
<div id="OfferfadeDiv">
</div>
<div id="DuplicateNoofOffer" class="folderdialog" style="position: absolute; top: 200px;
  left: 400px; width: 400px; height: 150px">
  <div class="foldertitlebar">
    <span class="dialogtitle">
      <% Sendb(Copient.PhraseLib.Lookup("term.newfromtemp", LanguageID))%></span> <span
        class="dialogclose" onclick="toggleDialog('DuplicateNoofOffer', false);">X</span>
  </div>
  <div class="dialogcontents">
    <div id="DuplicateOffererror" style="display: none; color: red;">
    </div>
    <table style="width:90%">
      <tr>
        <td>
          &nbsp;
        </td>
      </tr>
      <tr>
        <td>
          <label for="infoStart">
            <% Sendb(Copient.PhraseLib.Lookup("term.duplicateOfferstoCreate", LanguageID).Replace("99", MyCommon.NZ(MyCommon.Fetch_SystemOption(184), 0).ToString()))%></label>
          <input type="text" style="width: 20px" id="txtDuplicateOffersCnt" name="txtDuplicateOffersCnt"
            maxlength="2" value="" />
        </td>
      </tr>
      <tr>
        <td>
          &nbsp;
        </td>
      </tr>
	  <tr align="right">
        <td>
          <input type="button" name="btnOk" id="btnOk" value="<% Sendb(Copient.PhraseLib.Lookup("term.ok", LanguageID))%>"
            onclick="addDuplicateOfferscount();" />
          <input type="button" name="btnCancel" id="btnCancel" value="<% Sendb(Copient.PhraseLib.Lookup("term.cancel", LanguageID))%>"
            onclick="toggleDialog('DuplicateNoofOffer', false);" />
        </td>
       </tr>	  
    </table>
  </div>
</div>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<script runat="server">
  Const ANNIVERSARY_DATE_OP As Integer = 2
 
 Function GetLastDeployValidationMessage(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long) As String
    Dim lastDeployValidationMsg As String = String.Empty
    Dim dt As DataTable
    MyCommon.QueryStr = "Select LastDeployValidationMessage From Offers " & _
                          " Where OfferId=@OfferId" 
    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If Not dt is Nothing And dt.Rows.Count=1 Then
       lastDeployValidationMsg = dt.Rows(0)(0).ToString()
    End If
    Return lastDeployValidationMsg
  End Function	
  Sub SetLastDeployValidationMessage(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long, ByVal Message As String)
   MyCommon.QueryStr = "Update Offers " & _
                      "  Set LastDeployValidationMessage=@Message " & _
                      "  where OfferId=@OfferId" 
   MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
   MyCommon.DBParameters.Add("@Message", SqlDbType.NVarChar).Value = Message
   MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
  End Sub	

  Sub WriteComponent(ByRef Common As Copient.CommonInc, ByVal rowComp As DataRow, ByRef ComponentColor As String)
    Dim RecordType As String = ""
    Dim ID As Integer
    Dim StoredProcName As String = ""
    Dim IDParmName As String = ""
    Dim TypeCode As String = ""
    Dim PageName As String = ""
    Dim dtValid As DataTable
    Dim rowOK(), rowWaiting(), rowWatches(), rowWarnings() As DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceHoursWarn As Integer
    Dim ShowSubReport As Boolean = True
    Dim iGroupLocations As Integer
    
    objTemp = Common.Fetch_CM_SystemOption(10)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If
    
    objTemp = Common.Fetch_CM_SystemOption(11)
    If Not (Integer.TryParse(objTemp.ToString, GraceHoursWarn)) Then
      GraceHoursWarn = 24
    End If

    RecordType = Common.NZ(rowComp.Item("RecordType"), "")
    ID = Common.NZ(rowComp.Item("ID"), -1)
    
    Select Case RecordType
      Case "term.customergroup"
        StoredProcName = "dbo.pa_CM_ValidationReport_CustGroup"
        IDParmName = "@CustomerGroupID"
        TypeCode = "cg"
        PageName = "cgroup-edit.aspx?CustomerGroupID="
        ShowSubReport = IIf(ID = 1 OrElse ID = 2, False, True)
      Case "term.productgroup"
        StoredProcName = "dbo.pa_CM_ValidationReport_ProdGroup"
        IDParmName = "@ProductGroupID"
        TypeCode = "pg"
        PageName = "pgroup-edit.aspx?ProductGroupID="
        ShowSubReport = IIf(ID = 1, False, True)
    End Select
    
    Common.QueryStr = StoredProcName
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add(IDParmName, SqlDbType.Int).Value = ID
    Common.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
    Common.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn
    
    dtValid = Common.LRTsp_select()
    iGroupLocations = dtValid.Rows.Count
    
    rowOK = dtValid.Select("Status=0", "LocationName")
    rowWaiting = dtValid.Select("Status=1", "LocationName")
    rowWatches = dtValid.Select("Status=2", "LocationName")
    rowWarnings = dtValid.Select("Status=3", "LocationName")
    
    If (ShowSubReport AndAlso ComponentColor <> "red") Then
      ComponentColor = IIf(rowWarnings.Length > 0, "red", "green")
    End If
    
    Send("<div style=""margin-left:10px;"">")
    Sendb(Copient.PhraseLib.Lookup(Common.NZ(rowComp.Item("RecordType"), ""), LanguageID) & " #" & ID & ": ")
    If (ShowSubReport) Then
      Send("<a href=""" & PageName & ID & """>" & Common.SplitNonSpacedString(Common.NZ(rowComp.Item("Name"), "&nbsp;"), 20) & "</a>")
    Else
      Send(Common.SplitNonSpacedString(Common.NZ(rowComp.Item("Name"), "&nbsp;"), 20))
    End If
    
    If (ShowSubReport) Then
      Send("<div style=""margin-left:20px;"">")
      Send("<a id=""validLink" & ID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=0&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
      Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & " of " & iGroupLocations & ")</a><br />")
      Send("<a id=""waitingLink" & ID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=1&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
      Send(Copient.PhraseLib.Lookup("cgroup-edit.waitlocations", LanguageID) & " (" & rowWaiting.Length & " of " & iGroupLocations & ")</a><br />")
      Send("<a id=""watchLink" & ID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=2&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
      Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & " of " & iGroupLocations & ")</a><br />")
      Send("<a id=""warningLink" & ID & """ href=""javascript:openPopup('CM-validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=3&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
      Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & " of " & iGroupLocations & ")</a><br />")
      Send("</div>")
    End If
    
    Send("</div>")
    Common.Close_LRTsp()
  End Sub
  
  Sub CreateNewLocalPromotionVariables(ByVal lOfferId As Long, ByRef Mycommon As Copient.CommonInc)
    Dim lRewardID As Long
    Dim lPromoVarId As Long
    Dim rst As DataTable
    Dim row As DataRow
  
    ' create local promotion variables for this new offer
    Mycommon.QueryStr = "select OfferID from Offers with (NoLock) where OfferID=" & lOfferId & _
                        " and DistPeriodLimit > 0.00 and DistPeriod <> 0 and DistPeriodVarID=0;"
    rst = Mycommon.LRT_Select
    For Each row In rst.Rows
      Mycommon.Open_LogixXS()
      Mycommon.QueryStr = "dbo.pc_DistributionVar_Create"
      Mycommon.Open_LXSsp()
      Mycommon.LXSsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = lOfferId
      Mycommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
      Mycommon.LXSsp.ExecuteNonQuery()
      lPromoVarId = Mycommon.LXSsp.Parameters("@VarID").Value
      Mycommon.Close_LXSsp()
      Mycommon.Close_LogixXS()
      Mycommon.QueryStr = "update Offers with (RowLock) set DistPeriodVarID=" & lPromoVarId & " where OfferID=" & lOfferId & ";"
      Mycommon.LRT_Execute()
    Next
    
    ' create local promotion variables for this new offer's rewards
    Mycommon.QueryStr = "select RewardID from OfferRewards with (NoLock) where OfferID=" & lOfferId & _
                        " and RewardLimit > 0.00 and RewardDistPeriod <> 0 and RewardDistLimitVarID=0;"
    rst = Mycommon.LRT_Select
    For Each row In rst.Rows
      lRewardID = Mycommon.NZ(row.Item("RewardID"), 0)
      Mycommon.Open_LogixXS()
      Mycommon.QueryStr = "dbo.pc_RewardLimitVar_Create"
      Mycommon.Open_LXSsp()
      Mycommon.LXSsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = lRewardID
      Mycommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
      Mycommon.LXSsp.ExecuteNonQuery()
      lPromoVarId = Mycommon.LXSsp.Parameters("@VarID").Value
      Mycommon.Close_LXSsp()
      Mycommon.Close_LogixXS()
      Mycommon.QueryStr = "update OfferRewards with (RowLock) set RewardDistLimitVarID=" & lPromoVarId & " where RewardID=" & lRewardID & ";"
      Mycommon.LRT_Execute()
    Next
  End Sub

  Sub Send_Preference_Details(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long)
    Dim dt As DataTable
    Dim IntegrationVals As New Copient.CommonInc.IntegrationValues
    Dim PrefPageName As String = ""
    Dim Tokens As String = ""
    Dim RootURI As String = ""
    
    Common.QueryStr = "select UserCreated, Name as PrefName " & _
                      "from Preferences as PREF with (NoLock) " & _
                      "where PREF.PreferenceID=" & PreferenceID & " and PREF.Deleted=0;"
    dt = Common.PMRT_Select
    If dt.Rows.Count > 0 Then
      If (Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER, IntegrationVals)) Then
        PrefPageName = IIf(Common.NZ(dt.Rows(0).Item("UserCreated"), False), "prefscustom-edit.aspx", "prefsstd-edit.aspx")
          
        RootURI = IntegrationVals.HTTP_RootURI
        If RootURI IsNot Nothing AndAlso RootURI.Length > 0 AndAlso Right(RootURI, 1) <> "/" Then
          RootURI &= "/"
        End If
        
        Tokens = "SendToURI="
        Sendb("  <a href=""authtransfer.aspx?SendToURI=" & RootURI & "UI/" & PrefPageName & "?prefid=" & PreferenceID & """>")
        Send(Common.NZ(dt.Rows(0).Item("PrefName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)) & "</a>")
      End If
    End If
  End Sub

  Sub Send_Preference_Info(ByRef Common As Copient.CommonInc, ByVal ConditionID As Integer)
    Dim dt As DataTable
    Dim PreferenceID As Long = 0
    Dim ComboText As String = ""
    Dim i As Integer = 0
    Dim CellCount As Integer = 0
    Dim ValueSent As Boolean = False
    Dim AndComboed As Boolean = True
    
    ' find all the tier values
    Common.QueryStr = "select CPV.PKID, CPV.PreferenceID, CPV.Value, CPV.ValueComboTypeID, CPV.DateOperatorTypeID, " & _
                      "  case when POT.PhraseID is null then POT.Description" & _
                      "  else Convert(nvarchar(200), PT.Phrase) end as OperatorText " & _
                      "from CM_ConditionPreferenceValues as CPV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = CPV.OperatorTypeID " & _
                      "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID & " " & _
                      "where CPV.ConditionID=" & ConditionID
    dt = Common.LRT_Select
    For i = 0 To dt.Rows.Count - 1
      AndComboed = (Common.NZ(dt.Rows(i).Item("ValueComboTypeID"), 2) = 1)
      PreferenceID = Common.NZ(dt.Rows(i).Item("PreferenceID"), 0)
      
      If ValueSent Then Send(" " & Copient.PhraseLib.Lookup(IIf(AndComboed, "term.and", "term.or"), LanguageID) & " ")

      If Common.NZ(dt.Rows(i).Item("DateOperatorTypeID"), 0) > 0 Then
        Send(Get_Date_Display_Text(Common, dt.Rows(i).Item("PKID")))
      Else
        Send(Common.NZ(dt.Rows(i).Item("OperatorText"), "") & " " & Get_Preference_Value(Common, PreferenceID, Common.NZ(dt.Rows(i).Item("Value"), "")))
      End If

      If i < dt.Rows.Count - 1 Then
        Send(" <i>" & ComboText.ToLower & "</i> ")
      End If

      ValueSent = True
    Next
  End Sub

  Function OfferDeployed(ByRef Common As Copient.CommonInc, ByVal OfferID As Long) As Boolean
    
    Dim Deployed As Boolean = False
    Dim dt As DataTable

     Common.QueryStr = "select CMOADeploySuccessDate, StatusFlag, DeployDeferred from Offers with (NoLock) where OfferId=" & OfferID
     dt = Common.LRT_Select
     If Not IsDBNull(dt.Rows(0).Item("CMOADeploySuccessDate")) Then
       Deployed = True
     End If
	 
	 If (Common.NZ(dt.Rows(0).Item("StatusFlag"), -1) <> 2) Then
        If (Common.NZ(dt.Rows(0).Item("StatusFlag"), 0) > 0) Then
          If (Common.NZ(dt.Rows(0).Item("DeployDeferred"), False) = False) Then
            Deployed = False
          End If
        End If
      End If
	 
    Return Deployed 

  End Function
  
  Function Get_Preference_Value(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal Value As String) As String
    Dim TempLong As Long = 0
    Dim dt As DataTable
    
    Common.QueryStr = "select DataTypeID from Preferences with (NoLock) where PreferenceID=" & PreferenceID & " and Deleted=0;"
    dt = Common.PMRT_Select
    If dt.Rows.Count > 0 Then
      Select Case Common.NZ(dt.Rows(0).Item("DataTypeID"), 0)
        Case 1 ' list
          ' lookup to see if this is a preference with list items, if so get the list item name
          Common.QueryStr = "select case when UPT.PhraseID is null then PLI.Name " & _
                            "       else CONVERT(nvarchar(200), UPT.Phrase) end as PhraseText " & _
                            "from Preferences as PREF with (NoLock) " & _
                            "inner join PreferenceListItems as PLI with (NoLock) on PLI.PreferenceID = PREF.PreferenceID " & _
                            "left join UserPhraseText as UPT with (NoLock) on UPT.PhraseID = PLI.NamePhraseID " & _
                            "where PREF.Deleted=0 and PREF.DataTypeID=1 and PREF.PreferenceID=" & PreferenceID & _
                            "  and PLI.Value=N'" & Value & "';"
          dt = Common.PMRT_Select
          If dt.Rows.Count > 0 Then
            Value = Common.NZ(dt.Rows(0).Item("PhraseText"), Value)
          End If
        Case 5 ' boolean
          Value = Copient.PhraseLib.Lookup(IIf(Value = "1", "term.true", "term.false"), LanguageID)
      End Select
      
    End If

    Return Value
  End Function
  
  Function Get_Date_Display_Text(ByRef Common As Copient.CommonInc, ByVal ValuePKID As Integer) As String
    Dim DisplayText As String = ""
    Dim dt As DataTable
    Dim ValueModifier As String = ""
    Dim Offset, DaysBefore, DaysAfter As Integer
    
    Common.QueryStr = "select CPV.Value, CPV.ValueModifier, CPV.ValueTypeID, POT.PhraseID as OperatorPhraseID, CPV.DaysBefore, CPV.DaysAfter, " & _
                      "CPV.DateOperatorTypeID, PDOT.PhraseID as DateOpPhraseID " & _
                      "from CM_ConditionPreferenceValues as CPV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = CPV.OperatorTypeID " & _
                      "inner join CPE_PrefDateOperatorTypes as PDOT with (NoLock) on PDOT.PrefDateOperatorTypeID = CPV.DateOperatorTypeID " & _
                      "where PKID=" & ValuePKID & ";"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      DisplayText = Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("DateOpPhraseID"), ""), LanguageID) & " "
      DisplayText &= Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("OperatorPhraseID"), ""), LanguageID) & " "
      If Common.NZ(dt.Rows(0).Item("ValueTypeID"), 0) = 1 Then
        DisplayText &= "[" & Copient.PhraseLib.Lookup("term.currentdate", LanguageID).ToLower & "]"
        ValueModifier = Common.NZ(dt.Rows(0).Item("ValueModifier"), "")
        If ValueModifier <> "" AndAlso Integer.TryParse(ValueModifier, Offset) Then
          ValueModifier = " " & IIf(Offset < 0, " - ", " + ") & Math.Abs(Offset)
        End If
        DisplayText &= ValueModifier
      Else
        DisplayText &= " " & Common.NZ(dt.Rows(0).Item("Value"), "")
      End If

      If Common.NZ(dt.Rows(0).Item("DateOperatorTypeID"), 0) = ANNIVERSARY_DATE_OP Then
        DaysBefore = Common.NZ(dt.Rows(0).Item("DaysBefore"), 0)
        DaysAfter = Common.NZ(dt.Rows(0).Item("DaysAfter"), 0)

        If DaysBefore > 0 AndAlso DaysAfter > 0 Then
          DisplayText &= " (-" & DaysBefore & " / +" & DaysAfter & ")"
        ElseIf DaysBefore > 0 AndAlso DaysAfter = 0 Then
          DisplayText &= " (-" & DaysBefore & ")"
        ElseIf DaysBefore = 0 AndAlso DaysAfter > 0 Then
          DisplayText &= " (+" & DaysAfter & ")"
        End If
      End If
    End If
    
    Return DisplayText
  End Function

  Sub UpdateTemplatePermissions(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal systemOption As Integer)
       
    Dim dtTempPermission As New DataTable
    Dim Disallow_DisplayDates As Integer = Integer.MinValue
    Dim Disallow_OfferRedempThreshold As Integer = Integer.MinValue
        
    If (systemOption = 85) Then
      Common.QueryStr = "SELECT Disallow_DisplayDates from TemplatePermissions with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
      dtTempPermission = Common.LRT_Select()
      If dtTempPermission.Rows.Count > 0 Then
        Disallow_DisplayDates = Convert.ToInt16(dtTempPermission.Rows(0).Item("Disallow_DisplayDates"))
      End If
      Common.QueryStr = "UPDATE TemplatePermissions with (RowLock) Set Disallow_DisplayDates=" & Disallow_DisplayDates & " WHERE OfferID = " & OfferID
      Common.LRT_Execute()
    ElseIf (systemOption = 83) Then
      Common.QueryStr = "SELECT Disallow_OfferRedempThreshold from TemplatePermissions  with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
      dtTempPermission = Common.LRT_Select()
      If dtTempPermission.Rows.Count > 0 Then
        Disallow_OfferRedempThreshold = Convert.ToInt16(dtTempPermission.Rows(0).Item("Disallow_OfferRedempThreshold"))
      End If
      Common.QueryStr = "UPDATE TemplatePermissions  with (RowLock) Set Disallow_OfferRedempThreshold=" & Disallow_OfferRedempThreshold & " WHERE OfferID = " & OfferID
      Common.LRT_Execute()
    End If
  End Sub

    Sub SaveOfferThresholdPerHourValue(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal AdminUserID As Integer, ByVal EngineId As Integer)
        Dim OfferRedemptionThresholdperHour As Integer = 0
        Common.QueryStr = "SELECT RedemThresholdPerHour FROM offerAccessoryFields with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
        Dim dtRedemption As New DataTable
        dtRedemption = Common.LRT_Select()
        If dtRedemption.Rows.Count > 0 Then
            OfferRedemptionThresholdperHour = Common.NZ(dtRedemption.Rows(0).Item("RedemThresholdPerHour"), 0)
        End If
        Common.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        Common.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = DBNull.Value
        Common.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = DBNull.Value
        'Updating only OfferRedemptionThresholdperHour because it depends on CM SystemOption #83,  'pa_UpdateOfferAccessoryFields' contains logic to insert/update based on the engineID and systemoption passed     
        Common.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = OfferRedemptionThresholdperHour
        Common.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
        Common.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineId
        Common.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 83
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    End Sub


  Sub SaveOfferDisplayDates(ByRef Common As Copient.CommonInc, ByVal SourceOfferID As Long, ByVal OfferID As Long, ByVal AdminUserID As Integer, ByVal EngineId As Integer)
        
    Common.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & SourceOfferID & ";"
    Dim dtODisp As New DataTable
    Dim startDate As String = ""
    Dim endDate As String = ""
    dtODisp = Common.LRT_Select()
    If dtODisp.Rows.Count > 0 Then
      startDate = Common.NZ(dtODisp.Rows(0).Item("DisplayStartDate"), Nothing)
      endDate = Common.NZ(dtODisp.Rows(0).Item("DisplayEndDate"), Nothing)
    End If
    Common.QueryStr = "dbo.pa_UpdateOfferAccessoryFields"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    Common.LRTsp.Parameters.Add("@DisplayStartTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(startDate), DBNull.Value, startDate)
    Common.LRTsp.Parameters.Add("@DisplayEndTime", SqlDbType.DateTime).Value = IIf(String.IsNullOrEmpty(endDate), DBNull.Value, endDate)
                            
    Common.LRTsp.Parameters.Add("@RedemThresholdPerHour", SqlDbType.BigInt).Value = DBNull.Value
    Common.LRTsp.Parameters.Add("@AdminUserId", SqlDbType.BigInt).Value = AdminUserID
    Common.LRTsp.Parameters.Add("@EngineID", SqlDbType.BigInt).Value = EngineId
                
    Common.LRTsp.Parameters.Add("@OptionID", SqlDbType.BigInt).Value = 85
    Common.LRTsp.ExecuteNonQuery()
    Common.Close_LRTsp()

  End Sub

</script>
<script type="text/javascript" language="javascript">
  collapseBoxes();
  setComponentsColor('<% Sendb(ComponentColor) %>');
</script>
<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
    <%
     If (StatusMessage <> "") Then 
       Send("alert('" & StatusMessage & "');")  
     End If
    %>
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
  Send_BodyEnd()
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  If PrefManInstalled Then MyCommon.Close_PrefManRT()
  Logix = Nothing
  MyCommon = Nothing
%>
