<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.Contract" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-sum.aspx 
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
  Dim rst4 As DataTable
  Dim row4 As DataRow
  Dim rstTiers As DataTable
  Dim OfferID As Long = Request.QueryString("OfferID")
  Dim IsTemplate As Boolean
  Dim IsTemplateVal As String = "Not"
  Dim FromTemplate As Boolean = False
  Dim ActiveSubTab As Integer = 24
  Dim roid As Integer = 0
  Dim i As Integer = 0
  Dim Days As String = ""
  Dim Times As String = ""
  Dim Tenders As String = ""
  Dim Cookie As HttpCookie = Nothing
  Dim BoxesValue As String = ""
  Dim ShowCRM As Boolean = True
  Dim ValidateIncentiveColor As String = "green"
  Dim ComponentColor As String = "green"
  Dim IsDeployable As Boolean = False
  Dim LinksDisabled As Boolean = False
  Dim Expired As Boolean = False
  Dim TempDateTime As New DateTime
  Dim ShowActionButton As Boolean = False
  Dim SourceOfferID As Integer
  Dim DeployBtnIEOffset As Integer = 0
  Dim DeployBtnFFOffset As Integer = 0
  Dim NewUpdateLevel As Integer = 0
  Dim GCount As Integer = 0
  Dim AnyProductUsed As Boolean = False
  Dim AnyCustomerUsed As Boolean = False
  Dim AnyStoreUsed As Boolean = False
  Dim LongDate As New DateTime
  Dim LongDateZeroTime As New DateTime
  Dim TodayDateZeroTime As New DateTime
  Dim DaysDiff As Integer = 0
  Dim rowCount As Integer = 0
  Dim counter As Integer = 0
  Dim ErrorMsg As String = ""
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannerNames As String() = Nothing
  Dim BannerIDs As Integer() = Nothing
  Dim BannerCt As Integer = 0
  Dim BannersEnabled As Boolean = False
  Dim StatusMessage As String = ""
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim TenderList As String = ""
  Dim TenderValue As String = ""
  Dim TenderRequired As Boolean
  Dim TenderExcluded As Boolean
  Dim TenderExcludedAmt As Object = Nothing
  Dim TempPriority As Integer = 0
  Dim Popup As Boolean = False
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim OfferImported As Boolean = False
  Dim ImportMessage As String = ""
  Dim ProdGroupList As String = "-1"
  Dim EngineID As Integer = 2
  Dim EngineSubTypeID As Integer = 0
  Dim FolderNames As String = ""
  Dim x As Integer = 0
  Dim ActivityLogMsg As String
  Dim CollisionThreshold As Integer = 0
  Dim CollisionScope As Integer = 0
  Dim AssociatedProducts As Integer = 0
  Dim HasBundleDiscount As Boolean = False

  Const POS_CHANNEL_ID As Integer = 1
  
  CurrentRequest.Resolver.AppName = "CPEoffer-sum.aspx"
  Dim m_Offer As IOffer = CurrentRequest.Resolver.Resolve(Of IOffer)()
  Dim m_customerGroup As ICustomerGroups = CurrentRequest.Resolver.Resolve(Of ICustomerGroups)()
  Dim m_Logger As ILogger = CurrentRequest.Resolver.Resolve(Of ILogger)()
  Dim offerValidationLogFilePrefix As String = "OfferValidationLog"

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-sum.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Open_PrefManRT()
  End If
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If Not IsValidOffer(OfferID, CMS.AMS.Models.Engines.CPE, MyCommon) Then
    Server.Transfer("server-message.aspx?ErrorMessage=" & Copient.PhraseLib.Lookup("term.error.url-tweaking", LanguageID), False)
  End If

  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)
  CollisionThreshold = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(143))
  If CollisionThreshold <= 0 Then
    CollisionThreshold = 1
  End If
  CollisionScope = MyCommon.Extract_Val(MyCommon.Fetch_CPE_SystemOption(144))

  MyCommon.QueryStr = "select RewardOptionID, TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    roid = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
    TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
  End If
  
  MyCommon.QueryStr = "select EngineSubTypeID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
  End If
  
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
  
  If (Request.QueryString("new") <> "") Then
    If Request.QueryString("IsTemplate") <> "" Then
      IsTemplate = (Request.QueryString("IsTemplate") = "IsTemplate")
      If IsTemplate Then
        Response.Redirect("offer-new.aspx?NewTemplate=Yes&new=New")
      Else
        Response.Redirect("offer-new.aspx")
      End If
    End If
  ElseIf (Request.QueryString("OfferFromTemp") <> "") Then
        Try
            ' dbo.pc_CreateOfferFromTemplate @TemplateID bigint, @OfferID bigint OUTPUT
            MyCommon.QueryStr = "dbo.pc_Create_CPE_OfferFromTemplate"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            SourceOfferID = OfferID
            OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
            MyCommon.Close_LRTsp()
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("offer.createdfromtemplate", LanguageID) & ": " & SourceOfferID)
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "CPEoffer-gen.aspx?OfferID=" & OfferID)
            GoTo done
        Catch ex As Exception
            If ex.Message = "error.couldnot-processoffers" Then
                infoMessage = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                infoMessage = ex.Message
            End If
        End Try
  ElseIf (Request.QueryString("saveastemp") <> "") Then
        Try
            MyCommon.QueryStr = "dbo.pc_Create_CPE_TemplateFromOffer"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            SourceOfferID = OfferID
            OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
            MyCommon.Close_LRTsp()
            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("templates.createdfromoffer", LanguageID) & ": " & SourceOfferID)
            Response.Status = "301 Moved Permanently"
            Response.AddHeader("Location", "CPEoffer-gen.aspx?OfferID=" & OfferID)
            GoTo done
        Catch ex As Exception
            If ex.Message = "error.couldnot-processoffers" Then
                infoMessage = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                infoMessage = ex.Message
            End If
        End Try
  ElseIf (Request.QueryString("deploy") <> "") Then
    IsDeployable = IsDeployableOffer(Logix, MyCommon, OfferID, roid, ErrorMsg)
    If (IsDeployable) Then
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
      MyCommon.LRT_Execute()
      ActivityLogMsg = "[history.offer-deploy]"
      If GetCgiValue("deploytransreqskip") = "1" Then
        ActivityLogMsg = ActivityLogMsg & vbCrLf & "<BR>" & "- [history.offer-deploy.trasnrequired]" & " " & CheckForTranslationDeployError(MyCommon, roid)
      End If
      'Prevent the ActivityLogMessage from being longer than ActivityLog.Description can hold
      ActivityLogMsg = Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID)
      If ActivityLogMsg.Length > 1000 Then ActivityLogMsg = Left(ActivityLogMsg, 995) & " ..."
      'Write the message to the ActivityLog
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "CPEoffer-sum.aspx?OfferID=" & OfferID)
      SetLastDeployValidationMessage(MyCommon, OfferID, "term.validationsuccessful")
      m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.CPE, OfferID, Copient.PhraseLib.Lookup("term.validationsuccessful", LanguageID)), offerValidationLogFilePrefix)
      GoTo done
    Else
      infoMessage = ErrorMsg
      If EngineSubTypeID = 1 Then
        infoMessage &= " " & Copient.PhraseLib.Lookup("CPEoffer-sum.InstantWinRequired", LanguageID)
      End If
      SetLastDeployValidationMessage(MyCommon, OfferID, "<font color=""red"">" & infoMessage & "</font>")
      m_Logger.WriteInfo(String.Format("Engine:{0}, OfferId:{1}, Message:{2}", Engines.CPE, OfferID, infoMessage), offerValidationLogFilePrefix)
    End If
  ElseIf (Request.QueryString("sendoutbound") <> "") Then
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastCRMSendDate=getdate(), CRMEngineUpdateLevel=CRMEngineUpdateLevel+1, CRMSendToExport=1 where IncentiveID=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-sendtocrm", LanguageID))
  ElseIf (Request.QueryString("delete") <> "") Then
    Dim optInGroup As CustomerGroup = m_Offer.GetOfferDefaultCustomerGroup(OfferID, EngineID)
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2, Deleted=1, LastUpdate=getdate(), UpdateLevel=UpdateLevel+1 where IncentiveID=" & OfferID
    MyCommon.LRT_Execute()
    'Mark the shadow table offer as deleted as well.
    MyCommon.QueryStr = "update CPE_ST_Incentives with (RowLock) set Deleted=1, LastUpdate=getdate(), UpdateLevel = (select UpdateLevel from CPE_Incentives with (NoLock)where IncentiveID=" & OfferID & ") where IncentiveID=" & OfferID
    MyCommon.LRT_Execute()
        
    'Mark Client ID deleted if this is enternal offer.
    MyCommon.QueryStr = "dbo.pt_ExtOfferID_Delete"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.NVarChar, 20).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
    MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()
    'Remove any Offer Eligibility Conditions associated with the Offer.    
    m_Offer.DeleteOfferEligibleConditions(OfferID, EngineID)
    If (optInGroup IsNot Nothing) Then
      m_customerGroup.DeleteCustomerGroup(optInGroup.CustomerGroupID)
    End If
    'Remove the banners assigned to this offer
    If (BannersEnabled) Then
      MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = " & OfferID
      MyCommon.LRT_Execute()
    End If
    'Also remove/update the triggers for any associated EIW conditions
    MyCommon.QueryStr = "update CPE_EIWTriggers with (RowLock) set Removed=1, LastUpdate=getdate() where RewardOptionID=" & roid & " and Removed=0;"
    MyCommon.LRT_Execute()
    'Record activity
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-delete", LanguageID))
    Response.Status = "301 Moved Permanently"
    MyCommon.QueryStr = "select IsTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select()
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    If (IsTemplate) Then
      Response.AddHeader("Location", "/logix/temp-list.aspx")
    Else
      Response.AddHeader("Location", "/logix/offer-list.aspx")
    End If
    GoTo done
  ElseIf (Request.QueryString("export") <> "") Then
    Dim MyExport As New Copient.ExportXmlCPE
    Dim bStatus As Boolean
    Dim bProduction As Boolean
    Dim sFileFullPathName As String
    bProduction = True ' uses production start/end date
    sFileFullPathName = MyCommon.Fetch_SystemOption(29) & "\Offer" & Request.QueryString("OfferID") & ".gz"
    bStatus = MyExport.GenerateOfferXML(Request.QueryString("OfferID"), sFileFullPathName, bProduction)
    If Not bStatus Then
      infoMessage = MyExport.ErrorMessage
      'display error?
    Else
      If (MyExport.GetFileType = Copient.ExportXmlCPE.FileTypeEnum.XML_FORMAT) Then
        Dim oRead As System.IO.StreamReader
        Dim LineIn As String
        Dim Bom As String = ChrW(65279)
        oRead = System.IO.File.OpenText(sFileFullPathName)
        Response.ContentEncoding = Encoding.UTF8
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
        Response.ClearContent()
        Response.ClearHeaders()
        Response.Clear()
        Dim fs As System.IO.FileStream
        Dim br As System.IO.BinaryReader
        Response.ContentType = "application/octet-stream"
        Response.AddHeader("Content-Disposition", "attachment; filename=" & "Offer" & Request.QueryString("OfferID").ToString & ".gz")
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
    Dim MyExport As New Copient.ExportXmlCPE
    Dim bStatus As Boolean
    Dim bProduction As Boolean
    Dim EDWFilePath As String = ""
    Dim EDWFileName As String = ""
    
    bProduction = True ' uses production start/end date
    EDWFilePath = MyCommon.Fetch_SystemOption(73).Trim
    If (Right(MyCommon.Fetch_SystemOption(73), 1) <> "\") Then EDWFilePath &= "\"
    EDWFileName = "Offer" & Request.QueryString("OfferID") & "_" & Now.ToString("yyyy-MM-dd_HHmmss")
    
    MyExport.SetFileType(Copient.ExportXmlCPE.FileTypeEnum.XML_FORMAT)
    MyExport.SetTableType(Copient.ExportXmlCPE.TableTypeEnum.DEPLOYED)
    bStatus = MyExport.GenerateOfferXML(Request.QueryString("OfferID"), EDWFilePath & EDWFileName & ".xml", bProduction)
    If Not bStatus Then
      infoMessage = Copient.PhraseLib.Lookup("cpeoffer-sum.exportedwalert", LanguageID)
    Else
      ' write out a check file to state that the exported file is ready.
      System.IO.File.WriteAllText(EDWFilePath & EDWFileName & ".ok", EDWFilePath & EDWFileName & ".xml" & ControlChars.Tab & "OK")
      
      MyCommon.QueryStr = "dbo.pa_UpdateOfferFeeds"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = Request.QueryString("OfferID")
      MyCommon.LRTsp.Parameters.Add("@LastFeed", SqlDbType.DateTime).Value = Now()
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      
      StatusMessage = Copient.PhraseLib.Lookup("cpeoffer-sum.exportedwok", LanguageID)
    End If
  ElseIf (Request.QueryString("deferdeploy") <> "") Then
    IsDeployable = IsDeployableOffer(Logix, MyCommon, OfferID, roid, ErrorMsg)
    If (IsDeployable) Then
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID=" & OfferID
      MyCommon.LRT_Execute()
      ActivityLogMsg = "[history.offer-deploy]"
      If GetCgiValue("deploytransreqskip") = "1" Then
        ActivityLogMsg = ActivityLogMsg & vbCrLf & "<BR>" & "- [history.offer-deploy.trasnrequired]" & " " & CheckForTranslationDeployError(MyCommon, roid)
      End If
      'Prevent the ActivityLogMessage from being longer than ActivityLog.Description can hold
      ActivityLogMsg = Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID)
      If ActivityLogMsg.Length > 1000 Then ActivityLogMsg = Left(ActivityLogMsg, 995) & " ..."
      'Write the message to the ActivityLog
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.DecodeEmbededTokens(ActivityLogMsg, LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "CPEoffer-sum.aspx?OfferID=" & OfferID)
      GoTo done
    Else
      infoMessage = ErrorMsg
    End If
  ElseIf (Request.QueryString("canceldeploy") <> "") Then
    ' check if the offer is still in awaiting deployment status
    MyCommon.QueryStr = "select StatusFlag, DeployDeferred from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      ' update status to modified (1) if offer is still awaiting deployment, otherwise alert user that offer was already deployed.
      If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 2) Or (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = True) Then
        MyCommon.QueryStr = "select LastUpdateLevel from PromoEngineUpdateLevels with (NoLock) " & _
                            "where LinkID=" & OfferID & " and EngineID=2 and ItemType=1;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          NewUpdateLevel = MyCommon.NZ(rst.Rows(0).Item("LastUpdateLevel"), 0)
        End If
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, DeployDeferred=0, UpdateLevel=" & NewUpdateLevel & " where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.canceldeploy", LanguageID))
        infoMessage = Copient.PhraseLib.Lookup("term.deploymentcanceled", LanguageID)
      ElseIf (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 0) Then
        infoMessage = Copient.PhraseLib.Lookup("message.alreadydeployed", LanguageID)
      Else
        infoMessage = Copient.PhraseLib.Lookup("message.unablecanceldeployment", LanguageID)
      End If
    End If
    ElseIf (Request.QueryString("copyoffer") <> "") Then
        Try
            MyCommon.QueryStr = "select IsTemplate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
            rst = MyCommon.LRT_Select()
            IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
            If IsTemplate Then
                MyCommon.QueryStr = "dbo.pc_Create_CPE_TemplateFromOffer"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.NVarChar, 200).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@CreatedByAdminId", SqlDbType.Int).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                SourceOfferID = OfferID
                OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
                MyCommon.Close_LRTsp()
            Else
                MyCommon.QueryStr = "dbo.pc_Copy_CPE_Offer"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.BigInt).Value = Request.QueryString("OfferID")
                MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 2
                MyCommon.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = EngineSubTypeID
                MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.BigInt).Value = AdminUserID
                MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
                'MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set FromTemplate=0 where IncentiveID=" & OfferID & ";"
                MyCommon.Close_LRTsp()
            End If
            If (OfferID > 0) Then
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-copy", LanguageID))
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "CPEoffer-sum.aspx?OfferID=" & OfferID)
                GoTo done
            Else
                OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
            End If
        Catch ex As Exception
            If ex.Message = "error.couldnot-processoffers" Then
                infoMessage = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                infoMessage = ex.Message
            End If
        End Try
        
  End If
  
  If (Request.QueryString("OfferID") <> "") Then
    MyCommon.QueryStr = "select IncentiveID, OID.EngineID, PE.PhraseID as EnginePhraseID, PEST.PhraseID as EngineSubTypePhraseID, " & _
                        " IsTemplate, ClientOfferID, IncentiveName, CPE.Description, FromTemplate, " & _
                        " PromoClassID, CRMEngineID, Priority, StartDate, EndDate, EveryDOW, VendorCouponCode, EligibilityStartDate, " & _
                        " EligibilityEndDate, TestingStartDate, TestingEndDate, P1DistQtyLimit, P1DistTimeType, " & _
                        " P1DistPeriod, P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, " & _
                        " P3DistTimeType, P3DistPeriod, EnableImpressRpt, EnableRedeemRpt, CPE.CreatedDate, CPE.LastUpdate, CPE.Deleted, ExportToEDW, " & _
                        " CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt, CRMRestricted, StatusFlag, DeployDeferred, " & _
                        " OC.OfferCategoryID, OC.Description as CategoryName, " & _
                        " AU1.FirstName + ' ' + AU1.LastName as CreatedBy, AU2.FirstName + ' ' + AU2.LastName as LastUpdatedBy, OID.Imported, " & _
                        " CPE.EngineSubTypeID " & _
                        " from CPE_Incentives as CPE with (NoLock) " & _
                        " left join OfferIDs as OID with (NoLock) on OID.OfferID=CPE.IncentiveID " & _
                        " left join PromoEngines as PE with (NoLock) on PE.EngineID=OID.EngineID " & _
                        " left join PromoEngineSubTypes as PEST with (NoLock) on PEST.PromoEngineID=OID.EngineID and PEST.SubTypeID=OID.EngineSubTypeID " & _
                        " left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                        " left join AdminUsers as AU1 with (NoLock) on AU1.AdminUserID = CPE.CreatedByAdminID " & _
                        " left join AdminUsers as AU2 with (NoLock) on AU2.AdminUserID = CPE.LastUpdatedByAdminID " & _
                        " where IncentiveID=" & Request.QueryString("OfferID") & " and CPE.Deleted=0;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count < 1 Then
      infoMessage = Copient.PhraseLib.Lookup("error.deleted", LanguageID)
    Else
      IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
      FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
      LinksDisabled = IIf(MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), -1) = 2, True, False)
      If (Not LinksDisabled) Then
        LinksDisabled = (MyCommon.NZ(rst.Rows(0).Item("DeployDeferred"), False) = True)
      End If
      If Popup Then
        LinksDisabled = True
      End If
      TempDateTime = MyCommon.NZ(rst.Rows(0).Item("EndDate"), New Date(1900, 1, 1))
      'TempDateTime = TempDateTime.AddDays(1)
      If TempDateTime < Today() Then
        Expired = True
      End If
      OfferImported = MyCommon.NZ(rst.Rows(0).Item("Imported"), False)
      EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
    End If
    
    ' get the banner assigned to this offer
    If (BannersEnabled) Then
      MyCommon.QueryStr = "select BAN.BannerID, BAN.Name from BannerOffers BO with (NoLock) " & _
                          "inner join Banners BAN with (NoLock) on BAN.BannerID = BO.BannerID " & _
                          "where BO.OfferID = " & Request.QueryString("OfferID") & " order by BAN.Name;"
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
    
    'Determine if the offer has a group-level conditional (bundle) discount
    MyCommon.QueryStr = "select DiscountID from CPE_Discounts as DIS with (NoLock) " & _
                        "inner join CPE_Deliverables as DEL on DEL.OutputID=DIS.DiscountID " & _
                        "where DIS.DiscountTypeID=4 and DEL.RewardOptionID=" & roid & " and DEL.Deleted=0 and DIS.Deleted=0;"
    rst4 = MyCommon.LRT_Select
    If rst4.Rows.Count > 0 Then
      HasBundleDiscount = True
    End If
    
    ' find how many products are associated to the offer, for use in collision detection
    If (CollisionScope = 1) OrElse (CollisionScope = 2 AndAlso HasBundleDiscount) Then
      MyCommon.QueryStr = "select count(PGI.ProductID) as ProductCount " & _
                          "from ProdGroupItems as PGI with (NoLock) " & _
                          "inner join CPE_IncentiveProductGroups as IPG with (NoLock) on IPG.ProductGroupID=PGI.ProductGroupID and IPG.Deleted=0 and PGI.Deleted=0 and IPG.ExcludedProducts=0 and IPG.Disqualifier=0 " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                          "where RO.RewardOptionID=" & roid & ";"
    ElseIf (CollisionScope = 2) Then
      MyCommon.QueryStr = "select count(PGI.ProductID) as ProductCount " & _
                          "from ProdGroupItems as PGI with (NoLock) " & _
                          "inner join CPE_Discounts as DIS with (NoLock) on DIS.DiscountedProductGroupID=PGI.ProductGroupID and DIS.Deleted=0 and PGI.Deleted=0 " & _
                          "inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID=DIS.DiscountID and DEL.Deleted=0 " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=DEL.RewardOptionID " & _
                          "where RO.RewardOptionID=" & roid & ";"
    End If
    rst4 = MyCommon.LRT_Select
    AssociatedProducts = rst4.Rows(0).Item("ProductCount")
    
    ' check import status
    If OfferImported Then
      ProdGroupList = GetAssociatedProductGroupIDs(roid, MyCommon)
  
      ' are there any pending records for a newly-imported offer's product group
      MyCommon.QueryStr = "select PG.ProductGroupID, PG.Name, PIQ.StatusFlag from ProductGroups as PG with (NoLock) " & _
                          "inner join ProdInsertQueue as PIQ with (NoLock) on PIQ.ProductGroupID = PG.ProductGroupID " & _
                          "where(PIQ.ProductGroupID in (" & ProdGroupList & ") And PG.CreatedDate = PG.LastUpdate) " & _
                          "and PG.UpdateLevel=0 and PG.Deleted=0;"
      rst2 = MyCommon.LRT_Select
      If rst2.Rows.Count > 0 Then
        ImportMessage &= Copient.PhraseLib.Lookup("offer-import.pending-import", LanguageID)
        For Each row2 In rst2.Rows
          ImportMessage &= MyCommon.NZ(row2.Item("ProductGroupID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "<br />"
        Next
        ImportMessage &= "<a href=""CPEOffer-sum.aspx?OfferID=" & OfferID & """>[" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "]</a>"
      End If

      ' are there any failed product group imports for a newly-imported offer's product group
      MyCommon.QueryStr = "select ProductGroupID, Name, LastLoadMsg from ProductGroups " & _
                          "where(ProductGroupID in (" & ProdGroupList & ")  And CreatedDate = LastUpdate And UpdateLevel = 0 And Deleted = 0) " & _
                          "and LastLoadMsg like '%uploaded file%';"
      rst2 = MyCommon.LRT_Select
      If rst2.Rows.Count > 0 Then
        ImportMessage &= Copient.PhraseLib.Lookup("offer-import.failed-group-import", LanguageID)
        For Each row2 In rst2.Rows
          ImportMessage &= "<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row2.Item("ProductGroupID"), "") & """>" & _
                           MyCommon.NZ(row2.Item("ProductGroupID"), "") & " - " & MyCommon.NZ(row2.Item("Name"), "") & "</a><br />"
        Next
      End If
    End If
    
  End If
  
  ShowCRM = (MyCommon.Fetch_SystemOption(25) <> "0")
  MyCommon.QueryStr = "select CRMEngineID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
  rst4 = MyCommon.LRT_Select
  If rst4.Rows.Count > 0 Then
    If MyCommon.Extract_Val(MyCommon.NZ(rst4.Rows(0).Item("CRMEngineID"), 0)) = 0 Then
      ShowCRM = False
    End If
  End If
  
  If (IsTemplate) Then
    ActiveSubTab = 25
    IsTemplateVal = "IsTemplate"
  Else
    ActiveSubTab = 24
    IsTemplateVal = "Not"
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
      Send_BodyBegin(14)
    Else
      Send_BodyBegin(4)
    End If
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 2)
    If (rst.Rows.Count < 1) Then
      Send_Subtabs(Logix, 23, 3, , OfferID)
    Else
      Send_Subtabs(Logix, ActiveSubTab, 3, , OfferID)
    End If
  End If
  
  If (rst.Rows.Count < 1) Then
    Send("")
    Send("<div id=""intro"">")
    Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & Request.QueryString("OfferID") & "</h1>")
    Send("</div>")
    Send("<div id=""main"">")
    Send("  <div id=""infobar"" class=""red-background"">")
    Send("    " & infoMessage)
    Send("  </div>")
    Send("</div>")
    Send("</div>")
    Send("</body>")
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { }")
    Send("</script>")
    Send("</html>")
    GoTo done
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
  If (Logix.UserRoles.AccessInstantWinOffers = False AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "perm.offers-accessinstantwin")
    Send_BodyEnd()
    GoTo Done
  End If
  
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
%>
<script type="text/javascript">
  var divElems = new Array("generalbody", "periodbody", "limitsbody","deploymentbody", 
                           "locationbody", "notificationbody", "conditionbody", "rewardbody", "validationbody");
  var divVals  = new Array(1, 2, 4, 8, 16, 32, 64, 128, 256);
  var divImages = new Array("imgGeneral", "imgPeriod", "imgLimits", "imgDeployment",  
                            "imgLocations", "imgNotifications", "imgOptInConditions", "imgConditions", "imgRewards", "imgValidation");
  var boxesValue = <% Sendb(BoxesValue)%>;
  var DeferDeploy = false;
  
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
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▲';
      } else {
        document.getElementById("actionsmenu").style.visibility = 'hidden';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID))%>▼';
      }
    }
  }

  function confirmDeployCollision(deferred) {
    // Runs when the button (deploy or deferdeploy) in the actions menu is clicked.
    var confirmString = '';

    if (deferred) {
      confirmString = '<%Sendb(Copient.PhraseLib.Lookup("confirm.deferdeploy", LanguageID))%>';
    } else {
      confirmString = '<%Sendb(Copient.PhraseLib.Lookup("confirm.deploy", LanguageID))%>';
    }

    if (confirm(confirmString)) {
      if (deferred) {
        DeferDeploy = true;
      } else {
        DeferDeploy = false;
      }
      showConfirmationBox();
    } else {
      return false
    }
  }
  
  function showConfirmationBox() {
    // Reveals (and shows the appropriate buttons in) a box asking if the user wants to run collision detection
    var confirmationBox = document.getElementById("confirming");
    var collisionsBox = document.getElementById("loading");
    var deployButton = document.getElementById("confirmingDeploy");
    var deferDeployButton = document.getElementById("confirmingDeferDeploy");

    if (DeferDeploy) {
      deployButton.style.display = 'none';
      deferDeployButton.style.display = 'inline';
    } else {
      deployButton.style.display = 'inline';
      deferDeployButton.style.display = 'none';
    }

    if (confirmationBox != null && collisionsBox != null) {
      confirmationBox.style.display = 'block';
      collisionsBox.style.display = 'none';
    }
  }

  function loadCollisions() {
    showCollisionsBox();
    xmlhttpPost('OfferFeeds.aspx?Mode=GetProductCollisions&OfferID=<%Sendb(OfferID)%>', 'GetProductCollisions');
  }

  function showCollisionsBox() {
    var collisionsBox = document.getElementById("loading");
    var confirmationBox = document.getElementById("confirming");
    var deployButton = document.getElementById("collisionDeploy");
    var deferDeployButton = document.getElementById("collisionDeferDeploy");

    if (DeferDeploy) {
      deployButton.style.display = 'none';
      deferDeployButton.style.display = 'inline';
    } else {
      deployButton.style.display = 'inline';
      deferDeployButton.style.display = 'none';
    }

    if (collisionsBox != null && confirmationBox != null) {
      collisionsBox.style.display = 'block';
      confirmationBox.style.display = 'none';
    }
  }

  function xmlhttpPost(strURL, action) {
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
    self.xmlHttpReq.onreadystatechange = function() {
      if (self.xmlHttpReq != null && self.xmlHttpReq.readyState == 4) {
        if (action = "GetProductCollisions") {
          updateCollisionsBox(self.xmlHttpReq.responseText);
        }
      }
    }
    self.xmlHttpReq.send();
  }

  function updateCollisionsBox(responseMsg) {
    var collisionsContent = document.getElementById("collisionsContent");

    if (responseMsg.replace(/(\r\n|\n|\r)/gm,"") != '') {
      collisionsContent.innerHTML = responseMsg;
    } else {
      if (DeferDeploy) {
        window.location = 'CPEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>&deferdeploy=<%Sendb(Copient.PhraseLib.Lookup("term.deferdeployment", LanguageID))%>';
      } else {
        window.location = 'CPEoffer-sum.aspx?OfferID=<%Sendb(OfferID)%>&deploy=<%Sendb(Copient.PhraseLib.Lookup("term.deploy", LanguageID))%>';
      }
    }
  }
</script>
<form id="mainform" name="mainform" action="#">
  <input type="hidden" name="OfferID" id="OfferID" value="<% Sendb(OfferID)%>" />
  <input type="hidden" name="IsTemplate" id="IsTemplate" value="<% Sendb(IsTemplateVal)%>" />
  <div id="intro">
    <%
      Dim oName As String = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
      If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(oName, 43) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & ": " & MyCommon.TruncateString(oName, 43) & "</h1>")
      End If
      Send(vbTab & "<div id=""controls""" & IIf(Popup, " style=""display:none;""", "") & ">")
	
      'If the user has permission for one of the action buttons, display the action button. 
      ShowActionButton = (Logix.UserRoles.CreateTemplate And Not IsTemplate) OrElse (Logix.UserRoles.CRUDOfferFromTemplate And IsTemplate) _
                          OrElse ((Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate)) _
                          OrElse (Logix.UserRoles.SendOffersToCRM And Not IsTemplate And ShowCRM) OrElse (MyCommon.Fetch_SystemOption(73) <> "") _
        OrElse ((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not IsTemplate) _
        OrElse (Logix.UserRoles.CreateOfferFromBlank) _
        OrElse (Logix.UserRoles.EditFolders) _
        OrElse (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) _
        OrElse (Logix.UserRoles.ExportOffer) _
        OrElse (Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate)
      If (Not LinksDisabled OrElse IsTemplate) Then
        If (ShowActionButton = True) Then
          Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " &#9660;"" onclick=""toggleDropdown();"" />")
          Send("<div class=""actionsmenu"" id=""actionsmenu"">")
          If (Logix.UserRoles.EditFolders) Then
            Send_AssignFolders(OfferID)
          End If
          If (Logix.UserRoles.CreateOfferFromBlank) Then
            Send_CopyOffer(IsTemplate)
          End If
          If ((Logix.UserRoles.DeployNonTemplateOffers OrElse Logix.UserRoles.DeployTemplateOffers) And Not IsTemplate) Then
            If (Expired And MyCommon.Fetch_CPE_SystemOption(80) = "1") Then
            Else
              If (MyCommon.Fetch_CPE_SystemOption(141) = "1") AndAlso (AssociatedProducts > 0) Then
                Send_DeferDeployCollision()
              Else
                Send_DeferDeploy()
              End If
            End If
          End If
          If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) Then
            If (Expired And MyCommon.Fetch_CPE_SystemOption(80) = "1") Then
            Else
              If (MyCommon.Fetch_CPE_SystemOption(142) = "1") AndAlso (AssociatedProducts > 0) Then
                Send_DeployCollision()
              Else
                Send_Deploy()
              End If
            End If
          End If
          If (Logix.UserRoles.ExportOffer) Then
            Send_Export()
          End If
          If (MyCommon.NZ(rst.Rows(0).Item("ExportToEDW"), False) AndAlso MyCommon.Fetch_SystemOption(73).Trim <> "") Then
            Send("<input type=""submit"" id=""exportedw"" name=""exportedw"" value=""" & Copient.PhraseLib.Lookup("term.exporttoarchive", LanguageID) & """ />")
          End If
          If (Logix.UserRoles.CreateOfferFromBlank) Then
            Send_New()
          End If
          If (Logix.UserRoles.CRUDOfferFromTemplate And IsTemplate) Then
            Send_OfferFromTemp()
          End If
          If (Logix.UserRoles.CreateTemplate And Not IsTemplate) Then
            Send_Saveastemp()
          End If
          If (Logix.UserRoles.SendOffersToCRM And Not IsTemplate And ShowCRM) Then
            If (Expired And MyCommon.Fetch_CPE_SystemOption(80) = "1") Then
            Else
              Send_SendOutbound()
            End If
          End If
          If (Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate) Then
            Send_Delete()
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
    %>
  </div>
  <div id="main">
    <%
      If (Expired And MyCommon.Fetch_CPE_SystemOption(80) = "1" And Not IsTemplate) Then
        LinksDisabled = True
      End If
      StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
      If Not IsTemplate Then
        If (MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0) <> 2) Then
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
      If rst.Rows.Count < 1 Then
        GoTo done
      End If
      If (Not IsTemplate AndAlso modMessage = "") Then
        MyCommon.QueryStr = "select incentiveId from CPE_Incentives with (NoLock) where CreatedDate=LastUpdate and UpdateLevel=0 and IncentiveID=" & OfferID & ";"
        rst3 = MyCommon.LRT_Select
        If (rst3.Rows.Count = 0) Then
          Send_Status(OfferID, 2)
        End If
      End If
    
      ' send a message bar if this is an imported offer with a message set for feedback from the import process
      If OfferImported AndAlso ImportMessage <> "" Then
        Send("<div id=""modbar"" style=""background-color:#cc6600;"">" & ImportMessage & "</div>")
      End If
    %>
    <div id="column1">
      <div class="box" id="general">
        <%  
          If (LinksDisabled) Then
            Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</span></h2>")
          Else
            Send("<h2 style=""float:left;""><span><a href=""CPEoffer-gen.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("generalbody", "imgGeneral", Copient.PhraseLib.Lookup("term.general", LanguageID), True)
          Send("<div id=""generalbody"">")
          Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.general", LanguageID) & """ cellpadding=""0"" cellspacing=""0"">")
          Send("    <tr>")
          Send("      <td style=""width:130px;""><b>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0) & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("ClientOfferID"), "") & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.roid", LanguageID) & ":</b></td>")
          Send("      <td>" & roid & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.engine", LanguageID) & ":</b></td>")
          Sendb("      <td>")
          Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(rst.Rows(0).Item("EnginePhraseID"), 0), LanguageID))
          If MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0) > 0 Then
            Sendb(" " & Copient.PhraseLib.Lookup(MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0), LanguageID))
          End If
          Send("</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.vendor-coupon-code", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("VendorCouponCode"), "") & "</td>")
          Send("    </tr>")
          If (MyCommon.Fetch_SystemOption(25) <> "0") Then
            Send("<tr>")
            Send("  <td>")
            Send("    <b>" & Copient.PhraseLib.Lookup("term.outbound", LanguageID) & ":</b>")
            Send("  </td>")
            Send("  <td>")
            MyCommon.QueryStr = "select ExtInterfaceID, Name, PhraseID from ExtCRMInterfaces with (NoLock) " & _
                                "where ExtInterfaceID=" & MyCommon.NZ(rst.Rows(0).Item("CRMEngineID"), 0)
            rst4 = MyCommon.LRT_Select()
            If rst4.Rows.Count > 0 Then
              If Not IsDBNull(rst4.Rows(0).Item("PhraseID")) Then
                Sendb(Copient.PhraseLib.Lookup(rst4.Rows(0).Item("PhraseID"), LanguageID))
              Else
                Sendb(MyCommon.NZ(rst4.Rows(0).Item("Name"), ""))
              End If
            End If
            Send("  </td>")
            Send("</tr>")
          End If
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.status", LanguageID) & ":</b></td>")
          Send("      <td>" & Logix.GetOfferStatus(OfferID, LanguageID) & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), ""), 25) & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("Description"), ""), 25) & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.folders", LanguageID) & ":</b></td>")
          Send("      <td id=""folderNames"">" & FolderNames & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.category", LanguageID) & ":</b></td>")
          'Send("      <td><a href=""javascript:openPopup('offer-timeline.aspx?Category=" & MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), "") & "')"">" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CategoryName"), ""), 25) & "</a></td>")
          If MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), 0) > 0 Then
            Send("      <td><a href=""category-edit.aspx?OfferCategoryID=" & MyCommon.NZ(rst.Rows(0).Item("OfferCategoryID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CategoryName"), ""), 25) & "</a></td>")
          Else
            Send("      <td>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</td>")
          End If
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.priority", LanguageID) & ":</b></td>")
          TempPriority = MyCommon.NZ(rst.Rows(0).Item("Priority"), 0)
          If (TempPriority >= 900 AndAlso TempPriority <= 999) Then
            TempPriority = 900
          End If
          MyCommon.QueryStr = "select PhraseID from CPE_IncentivePriorities with (NoLock) where PriorityID=" & TempPriority & ";"
          rst4 = MyCommon.LRT_Select()
          If rst4.Rows.Count > 0 Then
            Send("      <td>" & Copient.PhraseLib.Lookup(rst4.Rows(0).Item("PhraseID"), LanguageID) & IIf(TempPriority = 900, " (" & MyCommon.NZ(rst.Rows(0).Item("Priority"), 0) - 900 & ")", "") & "</td>")
          End If
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.tiers", LanguageID) & ":</b></td>")
          Send("      <td>" & TierLevels & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.impression", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.reporting", LanguageID), VbStrConv.Lowercase) & ":</b></td>")
          If (MyCommon.NZ(rst.Rows(0).Item("EnableImpressRpt"), False) = True) Then
            If (Logix.UserRoles.AccessReports = True) AndAlso (Popup = False) Then
              Send("      <td><a href=""reports-detail.aspx?OfferID=" & MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0) & "&amp;Start=" & MyCommon.NZ(rst.Rows(0).Item("Startdate"), "1/1/1900") & "&amp;End=" & MyCommon.NZ(rst.Rows(0).Item("Enddate"), "1/1/1900") & "&amp;Status=" & Copient.PhraseLib.Lookup("offer.status" & MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0), LanguageID) & """>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & "</a></td>")
            Else
              Send("      <td>" & Copient.PhraseLib.Lookup("term.enabled", LanguageID) & "</td>")
            End If
          Else
            Send("      <td>" & Copient.PhraseLib.Lookup("term.disabled", LanguageID) & "</td>")
          End If
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.redemption", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.reporting", LanguageID), VbStrConv.Lowercase) & ":</b></td>")
          If (MyCommon.NZ(rst.Rows(0).Item("EnableRedeemRpt"), False) = True) Then
            If (Logix.UserRoles.AccessReports = True) AndAlso (Popup = False) Then
              Send("      <td><a href=""reports-detail.aspx?OfferID=" & MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0) & "&amp;Start=" & MyCommon.NZ(rst.Rows(0).Item("Startdate"), "1/1/1900") & "&amp;End=" & MyCommon.NZ(rst.Rows(0).Item("Enddate"), "1/1/1900") & "&amp;Status=" & Copient.PhraseLib.Lookup("offer.status" & MyCommon.NZ(rst.Rows(0).Item("StatusFlag"), 0), LanguageID) & """>" & Copient.PhraseLib.Lookup("term.available", LanguageID) & "</a></td>")
            Else
              Send("      <td>" & Copient.PhraseLib.Lookup("term.enabled", LanguageID) & "</td>")
            End If
          Else
            Send("      <td>" & Copient.PhraseLib.Lookup("term.disabled", LanguageID) & "</td>")
          End If
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.createdby", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CreatedBy"), ""), 25) & "</td>")
          Send("    </tr>")
          Send("    <tr>")
          Send("      <td><b>" & Copient.PhraseLib.Lookup("term.lastupdatedby", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("LastUpdatedBy"), ""), 25) & "</td>")
          Send("    </tr>")
          Send("  </table>")
          Send("</div>")
        %>
      </div>
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
            <tr>
              <td style="width: 85px;">
                <b>
                  <% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>:</b>
              </td>
              <td>
                <%
                  LongDate = MyCommon.NZ(rst.Rows(0).Item("StartDate"), New Date(1900, 1, 1))
                  If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                  Sendb(" - ")
                  LongDate = MyCommon.NZ(rst.Rows(0).Item("EndDate"), New Date(1900, 1, 1))
                  If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                  Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")) + 1 & " ")
                  If DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")) + 1 = 1 Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                  Else
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                  End If
                %>
              </td>
            </tr>
            <tr>
              <td>
                <b>
                  <% Sendb(Copient.PhraseLib.Lookup("term.testing", LanguageID))%>:</b>
              </td>
              <td>
                <%
                  LongDate = MyCommon.NZ(rst.Rows(0).Item("TestingStartDate"), "1/1/1900")
                  If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                  Sendb(" - ")
                  LongDate = MyCommon.NZ(rst.Rows(0).Item("TestingEndDate"), "1/1/1900")
                  If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                  Sendb(" (" & DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("TestingStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("TestingEndDate"), "1/1/1900")) + 1 & " ")
                  If DateDiff(DateInterval.Day, MyCommon.NZ(rst.Rows(0).Item("TestingStartDate"), "1/1/1900"), MyCommon.NZ(rst.Rows(0).Item("TestingEndDate"), "1/1/1900")) + 1 = 1 Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase) & ")")
                  Else
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase) & ")")
                  End If
                %>
              </td>
            </tr>
          </table>
        </div>
      </div>
      <div class="box" id="limits">
        <h2 style="float: left;">
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
          </span>
        </h2>
        <% Send_BoxResizer("limitsbody", "imgLimits", Copient.PhraseLib.Lookup("term.limits", LanguageID), True)%>
        <div id="limitsbody">
          <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>"
            cellpadding="0" cellspacing="0">
			<%-- Commented out to remove the references of Eligibility in offer summary page which is removed in offer general page. CLOUDSOL-1398
            <tr>
              <td style="width: 85px;">
                <b>
                  <% Sendb(Copient.PhraseLib.Lookup("term.eligibility", LanguageID))%>:</b>
              </td>
              <td>
                <%
                  If (MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 0) Then
                    Sendb(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))
                  ElseIf (MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 3650) AndAlso (MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0) = 1) AndAlso (MyCommon.NZ(rst.Rows(0).Item("P1DistTimeType"), 1) = 1) Then
                    Sendb(Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID))
                  ElseIf (MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 1 AndAlso (MyCommon.NZ(rst.Rows(0).Item("P1DistTimeType"), 0) = 2)) Then
                    Sendb(Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID))
                  ElseIf (MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P1DistTimeType"), 0) = 2) Then
                    Sendb(MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0))
                    Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.transaction", LanguageID), VbStrConv.Lowercase))
                  Else
                    Sendb(MyCommon.NZ(rst.Rows(0).Item("P1DistQtyLimit"), 0))
                    Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                    Send(MyCommon.NZ(rst.Rows(0).Item("P1DistPeriod"), 0))
                    MyCommon.QueryStr = "Select Description, PhraseID from CPE_DistributionTimeTypes with (NoLock) where TimeTypeId = " & MyCommon.NZ(rst.Rows(0).Item("P1DistTimeType"), 1)
                    rst2 = MyCommon.LRT_Select
                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup(rst2.Rows(0).Item("PhraseID"), LanguageID), VbStrConv.Lowercase))
                  End If
                %>
              </td>
            </tr>--%>
            <%-- Commented out 8/6 per Vince's advice. ALM
            <tr>
              <td><b><% Sendb(Copient.PhraseLib.Lookup("term.accumulation", LanguageID))%>:</b></td>
              <td><%
              If (MyCommon.NZ(rst.Rows(0).Item("P2DistPeriod"), 0) = 3650) Then
                Send(Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID))
              Else
                Send(MyCommon.NZ(rst.Rows(0).Item("P2DistQtyLimit"), 0) & " " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                Send(MyCommon.NZ(rst.Rows(0).Item("P2DistPeriod"), 0) & " ")
                MyCommon.QueryStr = "Select Description from CPE_DistributionTimeTypes with (NoLock) where TimeTypeId = " & MyCommon.NZ(rst.Rows(0).Item("P2DistTimeType"), 1)
                rst2 = MyCommon.LRT_Select
                Send(MyCommon.NZ(rst2.Rows(0).Item("Description"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)))
              End If
              %></td>
            </tr>
            --%>
            <tr>
              <td>
                <b>
                  <% Sendb(Copient.PhraseLib.Lookup("term.reward", LanguageID))%>:</b>
              </td>
              <td>
                <%
                  If (MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 0) Then
                    Send(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))
                  ElseIf (MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 3650) AndAlso (MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) = 1) AndAlso (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 1) = 1) Then
                    Send(Copient.PhraseLib.Lookup("term.onceperoffer", LanguageID))
                  ElseIf (MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 1 AndAlso (MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 2)) Then
                    Send(Copient.PhraseLib.Lookup("term.oncepertransaction", LanguageID))
                  ElseIf (MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0) = 1 AndAlso MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 0) = 2) Then
                    Sendb(MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0))
                    Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.transaction", LanguageID), VbStrConv.Lowercase))
                  Else
                    Sendb(MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0))
                    Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
                    Send(MyCommon.NZ(rst.Rows(0).Item("P3DistPeriod"), 0))
                    MyCommon.QueryStr = "Select Description, PhraseID from CPE_DistributionTimeTypes with (NoLock) where TimeTypeId = " & MyCommon.NZ(rst.Rows(0).Item("P3DistTimeType"), 1)
                    rst2 = MyCommon.LRT_Select
                    Sendb(" " & StrConv(Copient.PhraseLib.Lookup(rst2.Rows(0).Item("PhraseID"), LanguageID), VbStrConv.Lowercase))
                  End If
                %>
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
        <%  Send_BoxResizer("validationbody", "imgValidation", Copient.PhraseLib.Lookup("term.validationreport", LanguageID), True)%>
        <div id="validationbody">
          <%
            Dim dtValid, dtComponents As DataTable
            Dim rowOK(), rowWatches(), rowWarnings() As DataRow
            Dim rowComp As DataRow
            Dim objTemp As Object
            Dim GraceHours As Integer
            Dim GraceCount As Double
            Dim AllComponentsValid As Boolean = True
            
            objTemp = MyCommon.Fetch_CPE_SystemOption(41)
            If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
              GraceHours = 4
            End If
                    
            objTemp = MyCommon.Fetch_CPE_SystemOption(42)
            If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
              GraceCount = 0.1D
            End If
            
            MyCommon.QueryStr = "dbo.pa_ValidationReport_Incentive"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.Int).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
            MyCommon.LRTsp.Parameters.Add("@GraceCount", SqlDbType.Decimal, 2).Value = GraceCount
            
            dtValid = MyCommon.LRTsp_select()
            
            RemoveInactiveLocations(MyCommon, dtValid, OfferID)
            
            rowOK = dtValid.Select("Status=0", "LocationName")
            rowWatches = dtValid.Select("Status=1", "LocationName")
            rowWarnings = dtValid.Select("Status=2", "LocationName")
            MyCommon.Close_LRTsp()
            
            ValidateIncentiveColor = IIf(rowWarnings.Length > 0, "red", "green")
            
            Send("<a href=""javascript:showDiv('divOffer');"" style=""color:" & ValidateIncentiveColor & ";""><b>+ " & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</b></a><br />")
            Send("<div id=""divOffer"" style=""margin-left:10px;display:none;"">")
            If Popup Then
              Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")<br />")
              Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")<br />")
              Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")<br />")
            Else
              Send("<a id=""validLink" & OfferID & """ href=""javascript:openPopup('validation-report.aspx?type=in&id=" & OfferID & "&level=0');"">")
              Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")")
              Send("</a><br />")
              Send("<a id=""watchLink" & OfferID & """ href=""javascript:openPopup('validation-report.aspx?type=in&id=" & OfferID & "&level=1');"">")
              Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")")
              Send("</a><br />")
              Send("<a id=""warningLink" & OfferID & """ href=""javascript:openPopup('validation-report.aspx?type=in&id=" & OfferID & "&level=2');"">")
              Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")")
              Send("</a><br />")
            End If
            Send("</div>")
            
            MyCommon.QueryStr = "dbo.pa_ValidationReport_OfferComponents"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
            dtComponents = MyCommon.LRTsp_select
            
            Send("<a id=""linkComponent"" href=""javascript:showDiv('divComponents');"" style=""color:green;""><br class=""half"" /><b>+ " & Copient.PhraseLib.Lookup("term.components", LanguageID) & "</b></a><br />")
            Send("<div id=""divComponents"" style=""display:none;"">")
            For Each rowComp In dtComponents.Rows
              Send("<div style=""margin-left:10px;"">")
              WriteComponent(MyCommon, rowComp, ComponentColor, Popup)
              AllComponentsValid = AllComponentsValid AndAlso (ComponentColor.ToUpper = "GREEN")
              Send("</div>")
            Next
            Send("</div>")
            
            ' Update the Offer Validation Summary table with the most current validation information
            MyCommon.QueryStr = "dbo.pa_UpdateValidationSummary"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            MyCommon.LRTsp.Parameters.Add("@ValidLocations", SqlDbType.Int).Value = rowOK.Length
            MyCommon.LRTsp.Parameters.Add("@WatchLocations", SqlDbType.Int).Value = rowWatches.Length
            MyCommon.LRTsp.Parameters.Add("@WarningLocations", SqlDbType.Int).Value = rowWarnings.Length
            MyCommon.LRTsp.Parameters.Add("@ComponentsValid", SqlDbType.Bit).Value = IIf(AllComponentsValid, 1, 0)
            MyCommon.LRTsp.ExecuteNonQuery()
          %>
        </div>
        <hr class="hidden" />
      </div>
      <% End If%>
      <% If (Not IsTemplate) Then%>
      <div class="box" id="deployment">
        <h2 style="float: left;">
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.deployment", LanguageID))%>
          </span>
        </h2>
        <% Send_BoxResizer("deploymentbody", "imgDeployment", Copient.PhraseLib.Lookup("term.deployment", LanguageID), True)%>
        <div id="deploymentbody">
          <h3>
            <%Sendb(Copient.PhraseLib.Lookup("term.lastattempted", LanguageID))%>:
          </h3>
          <%
            TodayDateZeroTime = New Date(DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day, 0, 0, 0)
            LongDate = MyCommon.NZ(rst.Rows(0).Item("CPEOARptDate"), "1/1/1900")
            LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
            If LongDate > "1/1/1900" Then
              DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
              Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
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
            <%Sendb(Copient.PhraseLib.Lookup("term.lastsuccessful", LanguageID))%>:
          </h3>
          <%
            LongDate = MyCommon.NZ(rst.Rows(0).Item("CPEOADeploySuccessDate"), "1/1/1900")
            LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
            If LongDate > "1/1/1900" Then
              DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
              Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
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
            Dim validationMessage As String = GetLastDeployValidationMessage(MyCommon, OfferID)
            If (validationMessage <> "") Then
              If (validationMessage <> "term.validationsuccessful") Then
                'Checking ErrMgrTerms applicable in Offer Summary, because validationMessage is getting terms sometimes and actual term related text sometimes --AL-8825
                Dim ErrMsgTerms As String = "term.validationsuccessful, term.ReqTransFailed, deploy, deferdeploy, " &
                    "offer-sum.uomprecisionallowed, offer-sum.currhasbeenexceeded, offer-sum.currprecisionallowed, offer-sum.currhasbeenexceeded, " &
                    "offer-sum.unsupporteduom, offer-sum.undefineduom, offer-sum.unsupportedcurrency, offer-sum.nonselectedcurrency, " &
                    "cpeoffer-sum.deployalertforlockout, UEoffer-sum.tier-setup-invalid, offer-sum.required-incomplete," &
                    "cpeoffer-sum.deployalertforexpire, UEoffer-sum.deployalert"
                If ErrMsgTerms.Contains(validationMessage) Then
                  Sendb("<font color=""red"">" & Copient.PhraseLib.Lookup(validationMessage, LanguageID) & "</font>")
                Else
                  Sendb("<font color=""red"">" & validationMessage & "</font>")
                End If
              Else
                Sendb(Copient.PhraseLib.Lookup(validationMessage, LanguageID))
              End If
            End If
            Sendb("<br />")
          %>

          <br class="half" />
          <h3>
            <%Sendb(Copient.PhraseLib.Lookup("offer-sum.crmlastsent", LanguageID))%>:
          </h3>
          <%
            MyCommon.QueryStr = "select LastCRMSendDate from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
            rst4 = MyCommon.LRT_Select
            If rst4.Rows.Count > 0 Then
              LongDate = MyCommon.NZ(rst4.Rows(0).Item("LastCRMSendDate"), "1/1/1900")
            Else
              LongDate = "1/1/1900"
            End If
            LongDateZeroTime = New Date(LongDate.Year, LongDate.Month, LongDate.Day, 0, 0, 0)
            If LongDate > "1/1/1900" Then
              DaysDiff = DateDiff(DateInterval.Day, LongDateZeroTime, TodayDateZeroTime)
              Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
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
            <%Sendb(Copient.PhraseLib.Lookup("offer-sum.crmlastreceived", LanguageID))%>:
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
              Sendb("&nbsp;" & Logix.ToShortDateTimeString(LongDate, MyCommon))
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
          <hr class="hidden" />
        </div>
      </div>
      <% End If%>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="locationSum">
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
                  If Popup Then
                    Sendb("<li>" & BannerNames(i) & "</li>")
                  Else
                    Sendb("<li><a href=""banner-edit.aspx?BannerID=" & BannerIDs(i) & """>" & BannerNames(i) & "</a></li>")
                  End If
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
            MyCommon.QueryStr = "select OL.LocationGroupID,LG.Name,LG.PhraseID from OfferLocations as OL with (NoLock) " & _
                                "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=OL.LocationGroupID " & _
                                "where Excluded=0 and OL.Deleted=0 and OL.OfferID=" & OfferID & " order by LG.Name"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("LocationGroupID"), 0) = 1) Then
                AnyStoreUsed = True
              End If
            Next
            If rst.Rows.Count > 0 And Not AnyStoreUsed Then
              For Each row In rst.Rows
                MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), 0) & " And Deleted = 0"
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
              If IsDBNull(row.Item("PhraseID")) Then
                If Popup Then
                  Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                Else
                  Sendb("<li><a href=""lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                End If
              Else
                If (row.Item("PhraseID") = 0) Then
                  If Popup Then
                    Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                  Else
                    Sendb("<li><a href=""lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                  End If
                Else
                  Sendb("<li>" & MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                End If
              End If
              MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), 0) & " And Deleted = 0"
              rst2 = MyCommon.LRT_Select()
              For Each row2 In rst2.Rows
                If (MyCommon.NZ(row.Item("LocationGroupID"), 0) > 1) Then
                  Sendb(" (" & row2.Item("GCount") & ")")
                End If
              Next
            Next
            If (rst.Rows.Count = 0) Then
              Sendb("<li>" & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
            ' Check for and display any excluded store groups
            MyCommon.QueryStr = "select OL.LocationGroupID,LG.Name,LG.PhraseID from OfferLocations as OL with (NoLock) inner join LocationGroups as LG with (NoLock) on " & _
                                "LG.LocationGroupID=OL.LocationGroupID where Excluded=1 and OL.deleted=0 and OL.OfferID=" & OfferID
            rst = MyCommon.LRT_Select
            rowCount = rst.Rows.Count
            If rowCount > 0 Then
              Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
              For Each row In rst.Rows
                MyCommon.QueryStr = "select count(*) as GCount from LocGroupItems with (NoLock) where LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), 0) & " And Deleted=0"
                rst2 = MyCommon.LRT_Select()
                For Each row2 In rst2.Rows
                  GCount = GCount + (row2.Item("GCount"))
                Next
                If Not Popup Then
                  Sendb("<a href=""lgroup-edit.aspx?LocationGroupId=" & MyCommon.NZ(row.Item("LocationGroupID"), 0) & """>")
                End If
                If IsDBNull(row.Item("PhraseID")) Then
                  Sendb(MyCommon.NZ(row.Item("Name"), ""))
                Else
                  If (row.Item("PhraseID") = 0) Then
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                  Else
                    Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                  End If
                End If
                If Not Popup Then
                  Sendb("</a> ")
                End If
                Sendb("(" & GCount & ") ")
              Next
            End If
            Send("</li>")
            Send("</ul>")
          %>
          <h3>
            <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
          </h3>
          <% 
            MyCommon.QueryStr = "select OT.TerminalTypeID as TID,T.Name,T.PhraseID from OfferTerminals as OT with (NoLock) left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                "where OfferID=" & OfferID & " and Excluded=0 order by T.NAme"
            rst = MyCommon.LRT_Select
            Send("<ul class=""condensed"">")
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CPE Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CM Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All DP Terminals") Then
                If IsDBNull(row.Item("PhraseID")) Then
                  Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                Else
                  Sendb("<li>" & MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25))
                End If
              Else
                If Popup Then
                  Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                Else
                  Sendb("<li><a href=""terminal-edit.aspx?TerminalID=" & MyCommon.NZ(row.Item("TID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                End If
              End If
            Next
            If (rst.Rows.Count = 0) Then
              Sendb("<li>" & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
            ' Check for and display any excluded terminals
            MyCommon.QueryStr = "select OT.TerminalTypeID as TID,T.Name,T.PhraseID from OfferTerminals as OT with (NoLock) left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                "where OfferID=" & OfferID & " and Excluded=1 order by T.NAme"
            rst = MyCommon.LRT_Select
            rowCount = rst.Rows.Count
            If rowCount > 0 Then
              Sendb("&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
              For Each row In rst.Rows
                x = x + 1
                If (MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CPE Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All CM Terminals" OrElse MyCommon.NZ(row.Item("Name"), "").ToString.Trim = "All DP Terminals") Then
                  If IsDBNull(row.Item("PhraseID")) Then
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), 0), 25))
                  Else
                    Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID), 25))
                  End If
                Else
                  If Popup Then
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                  Else
                    Sendb("<a href=""terminal-edit.aspx?TerminalID=" & MyCommon.NZ(row.Item("TID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a> ")
                  End If
                End If
                If x = (rst.Rows.Count - 1) Then
                  Send(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                ElseIf x < rst.Rows.Count Then
                  Send(", ")
                Else
                  Send("")
                End If
              Next
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

                Select Case MyCommon.NZ(row2.Item("ChannelID"), 0)
                  Case POS_CHANNEL_ID
                    Send("<div style=""margin-left: 10px;"">")
                    ' Printed message notifications
                    MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID " & _
                                        "from CPE_Deliverables as D with (NoLock) " & _
                                        "inner join PrintedMessages as PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                        "inner join PrintedMessageTiers as PMT with (NoLock) on PM.MessageID=PMT.MessageID " & _
                                        "where D.Deleted=0 and D.RewardOptionPhase=1 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel=1;"
                    rst = MyCommon.LRT_Select()
                    If (rst.Rows.Count > 0) Then
                      counter = counter + 1
                      Dim Details As StringBuilder
                      Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.printedmessages", LanguageID) & "</h3>")
                      Send("<ul class=""condensed"">")
                      For Each row In rst.Rows
                        Details = New StringBuilder(200)
                        Details.Append(ReplaceTags(MyCommon.NZ(row.Item("BodyText"), "")))
                        If (Details.ToString().Length > 80) Then
                          Details = Details.Remove(77, (Details.Length - 77))
                          Details.Append("...")
                        End If
                        Details.Replace(vbCrLf, "<br />")
                        Send("<li>""" & HttpUtility.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, 25)) & """</li>")
                      Next
                      Send("</ul>")
                      Send("<br class=""half"" />")
                    End If
            
                    ' Cashier message notifications
                    MyCommon.QueryStr = "select D.DeliverableID, CM.MessageID, CMT.Line1, CMT.Line2 " & _
                                        "from CPE_Deliverables as D with (NoLock) " & _
                                        "inner join CPE_CashierMessages as CM with (NoLock) on D.OutputID=CM.MessageID " & _
                                        "left join CPE_CashierMessageTiers as CMT with (NoLock) on CMT.MessageID=CM.MessageID " & _
                                        "where D.RewardOptionID=" & roid & " and D.Deleted=0 and DeliverableTypeID=9 and D.RewardOptionPhase=1;"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                      counter = counter + 1
                      Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.cashiermessages", LanguageID) & "</h3>")
                      Send("<ul class=""condensed"">")
                      For Each row In rst.Rows
                        Send("<li>""" & MyCommon.NZ(row.Item("Line1"), "") & "<br />" & MyCommon.NZ(row.Item("Line2"), "") & """</li>")
                      Next
                      Send("</ul>")
                      Send("<br class=""half"" />")
                    End If
            
                    ' Graphics notifications
                    MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, OSA.Width as Width, OSA.Height as Height, OSA.GraphicSize as Size, OSA.ImageType as Type, " & _
                                        "D.DeliverableID, D.ScreenCellID as CellID, OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " & _
                                        "from OnScreenAds as OSA with (NoLock) Inner Join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID and D.RewardOptionID=" & roid & " " & _
                                        "and OSA.Deleted=0 and D.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=1 " & _
                                        "Inner Join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " & _
                                        "Inner Join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID;"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                      counter = counter + 1
                      Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & "</h3>")
                      Send("<ul class=""condensed"">")
                      For Each row In rst.Rows
                        If Popup Then
                          Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "&nbsp;")
                        Else
                          Sendb("<li><a href=""graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>&nbsp;")
                        End If
                        Sendb("(" & MyCommon.NZ(row.Item("Width"), "") & " x " & MyCommon.NZ(row.Item("Height"), ""))
                        If MyCommon.NZ(row.Item("Type"), 0) = 1 Then
                          Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                        ElseIf MyCommon.NZ(row.Item("Type"), 0) = 2 Then
                          Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
                        End If
                        Send(")</li>")
                      Next
                      Send("</ul>")
                      Send("<br class=""half"" />")
                    End If
            
                    ' Touchpoint notifications
                    MyCommon.QueryStr = "select RO.Name, RO.RewardOptionID, TA.OnScreenAdID as ParentAdID " & _
                                        "from CPE_RewardOptions RO with (NoLock) inner join CPE_DeliverableROIDs DR with (NoLock) on RO.RewardOptionID = DR.RewardOptionID " & _
                                        "inner join CPE_Deliverables D with (NoLock) on D.DeliverableID = DR.DeliverableID inner join TouchAreas TA with (NoLock) on DR.AreaID = TA.AreaID " & _
                                        "where RO.Deleted=0 and DR.Deleted=0 and TA.Deleted=0 and D.Deleted = 0 and RO.IncentiveID=" & OfferID & " and RO.TouchResponse=1 and D.RewardOptionPhase=1 order by RO.rewardoptionid;"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                      counter = counter + 1
                      Dim tpROID As Integer = 0
                      Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.touchpoints", LanguageID) & "</h3>")
                      Send("<ul class=""condensed"">")
                      For Each row In rst.Rows
                        tpROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)
                        Send("<li>")
                        If Popup Then
                          Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "<br />")
                        Else
                          Send("<a href=""graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("ParentAdID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a><br />")
                        End If
                        SetSummaryOnly(True)
                        Send_TouchpointRewards(OfferID, tpROID, 1, TierLevels)
                        Send("</li>")
                      Next
                      Send("</ul>")
                      Send("<br class=""half"" />")
                    End If
            
                    'Accumulation notifications
                    MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID " & _
                                        "from CPE_deliverables D with (NoLock) inner join PrintedMessages PM with (NoLock) on D.OutputID = PM.MessageID " & _
                                        "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID = PMT.MessageID " & _
                                        "where D.Deleted = 0 and D.RewardOptionPhase=2 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel = 1;"
                    rst = MyCommon.LRT_Select()
                    If (rst.Rows.Count > 0) Then
                      counter = counter + 1
                      Dim Details As StringBuilder
                      Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.accumulationmessage", LanguageID) & "</h3>")
                      Send("<ul class=""condensed"">")
                      For Each row In rst.Rows
                        Details = New StringBuilder(200)
                        Details.Append(ReplaceTags(MyCommon.NZ(row.Item("BodyText"), "")))
                        If (Details.ToString().Length > 80) Then
                          Details = Details.Remove(77, (Details.Length - 77))
                          Details.Append("...")
                        End If
                        Details.Replace(vbCrLf, "<br />")
                        Send("<li>""" & HttpUtility.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, 25)) & """</li>")
                      Next
                      Send("</ul>")
                    End If
            
                    If (counter = 0) Then
                      Send("<h3 style=""color: #606060;"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
                    End If
                    counter = 0
                  
                    Send("</div>")
                End Select
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
          Send("<h2 style=""float:left;""><span><a href=""CPEoffer-con.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.optinconditions", LanguageID) & "</a></span></h2>")
        End If
        Send_BoxResizer("optinconditionbody", "imgOptInConditions", Copient.PhraseLib.Lookup("term.optinconditions", LanguageID), True)
        Send("<div id=""optinconditionbody"">")
        Dim conditionlength As Integer
        Dim customers As List(Of CMS.AMS.Models.Customer)
        Dim Offer As Models.Offer = m_Offer.GetOffer(OfferID, LoadOfferOptions.AllEligibilityConditions)
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
            For Each pointscondition As Models.PointsCondition In Offer.EligiblePointsProgramConditions
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
            For Each storedvaluecondition As Models.SVCondition In Offer.EligibleSVProgramConditions
              Sendb("<li>")
              Sendb(CMS.Utilities.NZ(storedvaluecondition.Quantity, 0))
            
              If storedvaluecondition.SVProgram.SVType.SVTypeID > 1 Then
                Sendb(" ($" & Math.Round(storedvaluecondition.SVProgram.Value * storedvaluecondition.Quantity, storedvaluecondition.SVProgram.SVType.ValuePrecision) & ")")
              End If
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
            Send("<h2 style=""float:left;""><span><a href=""CPEoffer-con.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("conditionbody", "imgConditions", Copient.PhraseLib.Lookup("term.conditions", LanguageID), True)
        %>
        <div id="conditionbody">
          <%
            Dim icounter As Integer 'inclusion counter
            Dim xcounter As Integer 'exclusion counter
            ' Customer conditions
            MyCommon.QueryStr = "select CG.CustomerGroupID, Name, ExcludedUsers, PhraseID, ICG.RequiredFromTemplate, " & _
                                "CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                "where RewardOptionID=" & roid & " and ICG.Deleted=0 and ExcludedUsers=0"
            rst = MyCommon.LRT_Select
            i = 1
            GCount = 0
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("CustomerGroupID"), 99999) <= 2) Then
                AnyCustomerUsed = True
              End If
            Next
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.customerconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              Sendb("<li>")
              For Each row In rst.Rows
                If IsDBNull(row.Item("CustomerGroupID")) Then
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                Else
                  If IsDBNull(row.Item("PhraseID")) Then
                    If Popup Then
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25))
                    Else
                      Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25) & "</a>")
                    End If
                  Else
                    If (row.Item("PhraseID") = 0) Then
                      If Popup Then
                        Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(MyCommon.NZ(row.Item("Name"), ""), "&nbsp;"), 25))
                      Else
                        Sendb("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25) & "</a>")
                      End If
                    Else
                      Sendb(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                    End If
                  End If
                End If
                i = i + 1
                MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & " and Deleted = 0"
                rst2 = MyCommon.LXS_Select()
                For Each row2 In rst2.Rows
                  If Not (row2.Item("GCount") = 0) AndAlso Not MyCommon.NZ(row.Item("AnyCardholder"), False) AndAlso Not MyCommon.NZ(row.Item("AnyCustomer"), False) AndAlso Not MyCommon.NZ(row.Item("NewCardholders"), False) AndAlso Not MyCommon.NZ(row.Item("AnyCAMCardholder"), False) Then
                    Sendb(" (" & row2.Item("GCount") & ") ")
                    If (icounter <= rst.Rows.Count - 2) Then
                      Sendb(" or ")
                    Else
                      Sendb(" ")
                    End If
                  ElseIf (IsDBNull(row.Item("CustomerGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                    Sendb(" <span class=""red"">")
                    Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                    Sendb("</span>")
                  Else
                    Sendb(" ")
                  End If
                  icounter = icounter + 1
                Next
              Next
              ' Check for any display any excluded customer groups
              MyCommon.QueryStr = "select CG.CustomerGroupID,Name,PhraseID,ExcludedUsers from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                  "left join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                  "where RewardOptionID=" & roid & " and ICG.Deleted=0 and ExcludedUsers=1"
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                Sendb(Copient.PhraseLib.Lookup("term.excluding", LanguageID) & " ")
                For Each row In rst.Rows
                  If IsDBNull(row.Item("PhraseID")) Then
                    If Popup Then
                      Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                    Else
                      Send("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                    End If
                  Else
                    If (row.Item("PhraseID") = 0) Then
                      If Popup Then
                        Send(MyCommon.SplitNonSpacedString(row.Item("Name"), 25))
                      Else
                        Send("<a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                      End If
                    Else
                      Send(MyCommon.SplitNonSpacedString(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID), 25))
                    End If
                  End If
                  MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID = " & MyCommon.NZ(MyCommon.NZ(row.Item("CustomerGroupID"), 0), "-1") & " And Deleted = 0"
                  rst2 = MyCommon.LXS_Select()
                  For Each row2 In rst2.Rows
                    If (MyCommon.NZ(row.Item("CustomerGroupID"), -1) > 2) Then
                      Sendb(" (" & row2.Item("GCount") & ") ")
                      If (xcounter <= rst.Rows.Count - 2) Then
                        Sendb(" and ")
                      Else
                        Sendb(" ")
                      End If
                    End If
                    xcounter = xcounter + 1
                  Next
                Next
              End If
              Send("</li>")
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Attribute conditions
            MyCommon.QueryStr = "select IA.IncentiveAttributeID, DisallowEdit, RequiredFromTemplate, RO.AttributeComboID, " & _
                                "IAT.AttributeTypeID, AT.Description, IAT.AttributeValues " & _
                                "from CPE_IncentiveAttributes as IA with (NoLock) " & _
                                "left join CPE_IncentiveAttributeTiers as IAT with (NoLock) on IAT.IncentiveAttributeID=IA.IncentiveAttributeID " & _
                                "left join CPE_RewardOptions as RO with (NoLock) on IA.RewardOptionID=RO.RewardOptionID " & _
                                "left join AttributeTypes as AT with (NoLock) on AT.AttributeTypeID=IAT.AttributeTypeID " & _
                                "where IA.RewardOptionID=" & roid & " and IA.Deleted=0;"
            rst = MyCommon.LRT_Select
            i = 1
            GCount = 0
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.attributeconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                ' spit out the AttributeComboID
                If (i > 1) Then
                  If (MyCommon.NZ(row.Item("AttributeComboID"), 0) = 0) Then
                    ' single
                  ElseIf (row.Item("AttributeComboID") = 1) Then
                    ' and
                    Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "</i>")
                  Else
                    ' or
                    Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & "</i>")
                  End If
                End If
                Sendb("<li>")
                If IsDBNull(row.Item("AttributeTypeID")) Then
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                Else
                  Sendb("<a href=""attribute-edit.aspx?AttributeTypeID=" & MyCommon.NZ(row.Item("AttributeTypeID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Description"), ""), 25) & "</a>")
                End If
                If (IsDBNull(row.Item("AttributeTypeID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(" <span class=""red"">")
                  Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                  Sendb("</span>")
                End If
                
                Send("</li>")
                i = i + 1
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Product conditions
            MyCommon.QueryStr = "select PG.ProductGroupID, PG.Name, PG.PhraseID, UT.PhraseID, ExcludedProducts, ProductComboID, " & _
                                " QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,RequiredFromTemplate, IPG.IncentiveProductGroupID " & _
                                " from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionId " & _
                                " left join CPE_UnitTypes as UT with (NoLock)on UT.UnitTypeID=IPG.QtyUnitType " & _
                                " where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and Disqualifier=0;"
            rst = MyCommon.LRT_Select
            i = 1
            GCount = 0
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("ProductGroupID"), -1) = 1) Then
                AnyProductUsed = True
              End If
            Next
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.productconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                ' spit out the ProductComboID
                If (i > 1 And MyCommon.NZ(row.Item("ExcludedProducts"), False) = False) Then
                  If (MyCommon.NZ(row.Item("ProductComboID"), 0) = 0) Then
                    ' single
                  ElseIf (row.Item("ProductComboID") = 1) Then
                    ' and
                    Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "</i>")
                  Else
                    ' or
                    Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & "</i>")
                  End If
                End If
                If (MyCommon.NZ(row.Item("ExcludedProducts"), False) = True) Then
                  Sendb("<li>" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ": ")
                Else
                  Sendb("<li>")
                End If
                If IsDBNull(row.Item("ProductGroupID")) Then
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                Else
                  If IsDBNull(row.Item("PhraseID")) Then
                    If Popup Then
                      Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                    Else
                      Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                    End If
                  Else
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25))
                  End If
                End If
                If (IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(" <span class=""red"">")
                  Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                  Sendb("</span>")
                End If
                
                MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & " And Deleted = 0"
                rst2 = MyCommon.LRT_Select()
                For Each row2 In rst2.Rows
                  If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                    Send(" (" & row2.Item("GCount") & ")")
                  End If
                Next
                
                ' QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UnitDescription
                If (MyCommon.NZ(row.Item("QtyForIncentive"), 0) > 0 And Not MyCommon.NZ(row.Item("ExcludedProducts"), False)) Then
                  Send("<br />")
                  MyCommon.QueryStr = "select Quantity from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID=" & MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0) & " and RewardOptionID=" & roid & " order by TierLevel;"
                  rstTiers = MyCommon.LRT_Select
                  If rstTiers.Rows.Count > 0 Then
                    t = 1
                    For Each row4 In rstTiers.Rows
                      If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Then
                        Sendb(Int(MyCommon.NZ(row4.Item("Quantity"), 0)))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Then
                        Sendb(FormatCurrency(MyCommon.NZ(row4.Item("Quantity"), 0)))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                        Sendb(Math.Round(MyCommon.NZ(row4.Item("Quantity"), 0), 3) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                             Sendb(Copient.PhraseLib.Lookup("term.qty1atprice", LanguageID) & " " & FormatCurrency(MyCommon.NZ(row4.Item("Quantity"), 0)))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 10) Then
                                      Sendb(Math.Round(MyCommon.NZ(row4.Item("Quantity"), 0), 2))
                                End If
                      If t < TierLevels Then
                        Sendb(" / ")
                      End If
                      t = t + 1
                    Next
                    If MyCommon.NZ(row.Item("QtyUnitType"), 0) <> 4 Then
                      Send(" " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                    End If
                  End If
                  
                  Send("<br />")
                  If MyCommon.NZ(row.Item("AccumLimit"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumPeriod"), 0) <> 0 OrElse MyCommon.NZ(row.Item("AccumMin"), 0) <> 0 Then
                    ' There's at least some accumulation data set, so display it:
                    ' Limit value
                    If row.Item("AccumLimit") > 0 Then
                      Sendb(Copient.PhraseLib.Lookup("term.limit", LanguageID) & " ")
                      If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                        Sendb(Math.Truncate(MyCommon.NZ(row.Item("AccumLimit"), 0)))
                      ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                        Sendb(FormatCurrency(MyCommon.NZ(row.Item("AccumLimit"), 0)))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                        Sendb(MyCommon.NZ(row.Item("AccumLimit"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                      End If
                    Else
                      Sendb(Copient.PhraseLib.Lookup("term.nolimit", LanguageID))
                    End If
                    ' Period value
                    If row.Item("AccumPeriod") > 0 Then
                      Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.every", LanguageID), VbStrConv.Lowercase) & " ")
                      If row.Item("AccumPeriod") <= 1 Then
                        Sendb(StrConv(Copient.PhraseLib.Lookup("term.day", LanguageID), VbStrConv.Lowercase))
                      Else
                        Sendb(row.Item("AccumPeriod") & " " & StrConv(Copient.PhraseLib.Lookup("term.days", LanguageID), VbStrConv.Lowercase))
                      End If
                    End If
                    ' Minimum value
                    If row.Item("AccumMin") > 0 Then
                      Sendb(", " & StrConv(Copient.PhraseLib.Lookup("term.minimum", LanguageID), VbStrConv.Lowercase) & " ")
                      If MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1 Then
                        Send(Math.Truncate(MyCommon.NZ(row.Item("AccumMin"), 0)))
                      ElseIf MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2 Then
                        Send(FormatCurrency(MyCommon.NZ(row.Item("AccumMin"), 0)))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Then
                        Send(MyCommon.NZ(row.Item("AccumMin"), 0) & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                      End If
                    Else
                      Send(", " & StrConv(Copient.PhraseLib.Lookup("term.nominimum", LanguageID), VbStrConv.Lowercase))
                    End If
                  End If
                End If
                Send("</li>")
                i = i + 1
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Product disqualifiers
            MyCommon.QueryStr = "select PG.ProductGroupID, PG.Name, PG.PhraseID, UT.PhraseID, ExcludedProducts, ProductComboID, " & _
                                " QtyForIncentive, QtyUnitType, AccumMin, AccumLimit, AccumPeriod, RequiredFromTemplate, IPG.IncentiveProductGroupID " & _
                                " from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                " left join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                " left join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionId " & _
                                " left join CPE_UnitTypes as UT with (NoLock)on UT.UnitTypeID=IPG.QtyUnitType " & _
                                " where IPG.RewardOptionID=" & roid & " and IPG.Deleted=0 and Disqualifier=1;"
            rst = MyCommon.LRT_Select
            i = 1
            GCount = 0
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("ProductGroupID"), -1) = 1) Then
                AnyProductUsed = True
              End If
            Next
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.productdisqualifiers", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                ' spit out the ProductComboID
                If (i > 1) Then
                  Send("&nbsp;<i>" & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & "</i>")
                End If
                Sendb("<li>")
                If IsDBNull(row.Item("PhraseID")) Then
                  If Popup Then
                    Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                  Else
                    Sendb("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                  End If
                ElseIf (IsDBNull(row.Item("ProductGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(" <span class=""red"">")
                  Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                  Sendb("</span>")
                Else
                  Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), "&nbsp;"), 25))
                End If
                
                MyCommon.QueryStr = "select count(*) as GCount from ProdGroupItems with (NoLock) where ProductGroupID = " & MyCommon.NZ(row.Item("ProductGroupID"), "-1") & " And Deleted = 0"
                rst2 = MyCommon.LRT_Select()
                For Each row2 In rst2.Rows
                  If (MyCommon.NZ(row.Item("ProductGroupID"), -1) > 1) Then
                    Send(" (" & row2.Item("GCount") & ")")
                  End If
                Next
                
                ' QtyForIncentive,QtyUnitType,AccumMin,AccumLimit,AccumPeriod,UnitDescription
                If (MyCommon.NZ(row.Item("QtyForIncentive"), 0) > 0 And Not MyCommon.NZ(row.Item("ExcludedProducts"), False)) Then
                  Send("<br />")
                  MyCommon.QueryStr = "select Quantity from CPE_IncentiveProductGroupTiers where IncentiveProductGroupID=" & MyCommon.NZ(row.Item("IncentiveProductGroupID"), 0) & " and RewardOptionID=" & roid & " order by TierLevel;"
                  rstTiers = MyCommon.LRT_Select
                  If rstTiers.Rows.Count > 0 Then
                    t = 1
                    For Each row4 In rstTiers.Rows
                      If (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 1) Then
                        'Sendb(row4.Item("Quantity"))
                        'BZ4596 Fix - Display integer value for Product disqualifier in CPE Offer Summary
                        Sendb(Int(MyCommon.NZ(row4.Item("Quantity"), 0)))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 2) Then
                        Sendb(FormatCurrency(row4.Item("Quantity")))
                      ElseIf (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 3) Or (MyCommon.NZ(row.Item("QtyUnitType"), 0) = 4) Then
                        Sendb(row4.Item("Quantity") & " " & Copient.PhraseLib.Lookup("term.lbsgals", LanguageID))
                      End If
                      t = t + 1
                    Next
                    Send(" " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase))
                  End If
                End If
                Send("</li>")
                i = i + 1
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Points conditions
            MyCommon.QueryStr = "Select IPG.ProgramID, ProgramName, QtyForIncentive, RequiredFromTemplate from CPE_IncentivePointsGroups as IPG with (NoLock) " & _
                                "left join PointsPrograms as PP with (NoLock) on PP.ProgramID=IPG.ProgramID " & _
                                "where IPG.Deleted=0 and RewardOptionID=" & roid & ";"
            rst = MyCommon.LRT_Select
            i = 1
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.pointsconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Sendb("<li>")
                MyCommon.QueryStr = "select Quantity from CPE_IncentivePointsGroupTiers where RewardOptionID=" & roid & " order by TierLevel;"
                rstTiers = MyCommon.LRT_Select
                If rstTiers.Rows.Count > 0 Then
                  t = 1
                  For Each row4 In rstTiers.Rows
                    Sendb(MyCommon.NZ(row4.Item("Quantity"), 0))
                    If t < TierLevels Then
                      Sendb(" / ")
                    End If
                    t = t + 1
                  Next
                End If
                If (MyCommon.NZ(row.Item("ProgramID"), -1) > 0) Then
                  If Popup Then
                    Sendb(" " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25))
                  Else
                    Sendb(" <a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25) & "</a>")
                  End If
                ElseIf (IsDBNull(row.Item("ProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(" <span class=""red"">")
                  Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                  Sendb("</span>")
                Else
                  Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25))
                End If
                Send("</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Store value conditions
            MyCommon.QueryStr = "select ISVP.SVProgramID, SVP.Name, SVP.Value, SVP.SVTypeID, SVT.ValuePrecision, QtyForIncentive, RequiredFromTemplate from CPE_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                                "left join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=ISVP.SVProgramID " & _
                                "left join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " & _
                                "where ISVP.Deleted=0 and RewardOptionID=" & roid
            rst = MyCommon.LRT_Select
            i = 1
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalueconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Sendb("<li>")
                MyCommon.QueryStr = "select Quantity from CPE_IncentiveStoredValueProgramTiers where RewardOptionID=" & roid & " order by TierLevel;"
                rstTiers = MyCommon.LRT_Select
                If rstTiers.Rows.Count > 0 Then
                  t = 1
                  For Each row4 In rstTiers.Rows
                    Sendb(CInt(MyCommon.NZ(row4.Item("Quantity"), "0")))
                    If MyCommon.NZ(row.Item("SVTypeID"), 0) > 1 Then
                      Sendb(" ($" & Math.Round(MyCommon.NZ(row.Item("Value"), 0) * MyCommon.NZ(row4.Item("Quantity"), 0), MyCommon.NZ(row.Item("ValuePrecision"), 0)) & ")")
                    End If
                    If t < TierLevels Then
                      Sendb(" / ")
                    End If
                    t = t + 1
                  Next
                End If
                If (MyCommon.NZ(row.Item("SVProgramID"), -1) > 0) Then
                  If Popup Then
                    Sendb(" " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                  Else
                    Sendb(" <a href=""SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                  End If
                ElseIf (IsDBNull(row.Item("SVProgramID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(" <span class=""red"">")
                  Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                  Sendb("</span>")
                Else
                  Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                End If
                Send("</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Tender conditions
            MyCommon.QueryStr = "select ITT.IncentiveTenderID, ITT.TenderTypeID, ITT.Value, TT.Name, RequiredFromTemplate, RO.ExcludedTender, RO.ExcludedTenderAmtRequired " & _
                                "from CPE_IncentiveTenderTypes as ITT with (NoLock) " & _
                                "left join CPE_TenderTypes as TT with (NoLock) on TT.TenderTypeID=ITT.TenderTypeID " & _
                                "left join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ITT.RewardOptionID " & _
                                "where ITT.Deleted=0 and ITT.RewardOptionID=" & roid & ";"
            rst = MyCommon.LRT_Select
            i = 1
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.tenderconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                If TenderList <> "" Then
                  TenderList &= ", "
                End If
                If TenderValue <> "" Then
                  TenderValue &= ", "
                End If
                TenderList &= MyCommon.NZ(row.Item("Name"), "")
                TenderValue &= FormatCurrency(MyCommon.NZ(row.Item("Value"), 0), 3)
                TenderRequired = MyCommon.NZ(row.Item("RequiredFromTemplate"), False)
                TenderExcluded = MyCommon.NZ(row.Item("ExcludedTender"), False)
                TenderExcludedAmt = MyCommon.NZ(row.Item("ExcludedTenderAmtRequired"), 0)
              Next
              If TenderExcluded Then
                Sendb("<li>")
                Sendb(FormatCurrency(TenderExcludedAmt, 3) & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.allbut", LanguageID), VbStrConv.Lowercase) & " ")
                Sendb("<a href=""tender-engines.aspx"">" & MyCommon.SplitNonSpacedString(TenderList, 25) & "</a>")
                Send("</li>")
              Else
                For Each row In rst.Rows
                  Sendb("<li>")
                  MyCommon.QueryStr = "select Value from CPE_IncentiveTenderTypeTiers where RewardOptionID=" & roid & " and IncentiveTenderID=" & MyCommon.NZ(row.Item("IncentiveTenderID"), 0) & ";"
                  rstTiers = MyCommon.LRT_Select
                  If rstTiers.Rows.Count > 0 Then
                    t = 1
                    For Each row4 In rstTiers.Rows
                      Sendb(FormatCurrency(MyCommon.NZ(row4.Item("Value"), "0"), 3))
                      If t < TierLevels Then
                        Sendb(" / ")
                      End If
                      t = t + 1
                    Next
                  End If
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & " ")
                  If (MyCommon.NZ(row.Item("ExcludedTender"), False) = True) Then
                    Sendb(StrConv(Copient.PhraseLib.Lookup("term.allbut", LanguageID), VbStrConv.Lowercase) & " ")
                  End If
                  Sendb("<a href=""tender-engines.aspx"">" & MyCommon.NZ(row.Item("Name"), "") & "</a>")
                  If rst.Rows.Count > 1 And i < rst.Rows.Count Then
                    Sendb(", <i>" & StrConv(Copient.PhraseLib.Lookup("term.or", LanguageID), VbStrConv.Lowercase) & "</i>")
                  End If
                  Send("</li>")
                  i += 1
                Next
              End If
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Days conditions
            MyCommon.QueryStr = "select DOWID, PhraseID from CPE_DaysOfWeek DW with (NoLock)"
            rst = MyCommon.LRT_Select
            MyCommon.QueryStr = "select DOWID from CPE_IncentiveDOW with (NoLock) where Deleted=0 and IncentiveID=" & OfferID
            rst2 = MyCommon.LRT_Select
            For Each row In rst.Rows
              If rst2.Rows.Count >= 7 Then
                Days = Copient.PhraseLib.Lookup("term.everyday", LanguageID)
              Else
                For Each row2 In rst2.Rows
                  If (MyCommon.NZ(row2.Item("DOWID"), 0) = MyCommon.NZ(row.Item("DOWID"), 0)) Then
                    If (Days = "") Then
                      Days = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
                    Else
                      Days = Days & ", " & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID)
                    End If
                  End If
                Next
              End If
            Next
            If (Days.Trim.Length > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.dayconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              Send("<li>" & Days & "</li>")
              Send("</ul>")
              Send("<br class=""half"" />")
            End If

            ' Time conditions
            MyCommon.QueryStr = "select StartTime, EndTime from CPE_IncentiveTOD with (NoLock) where IncentiveID=" & OfferID
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              For i = 0 To rst.Rows.Count - 1
                If (i > 0) Then Times &= "; "
                Times &= MyCommon.NZ(rst.Rows(i).Item("StartTime"), "") & " - " & MyCommon.NZ(rst.Rows(i).Item("EndTime"), "")
              Next
            End If
            If (Times.Trim.Length > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.timeconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              Send("<li>" & Times & "</li>")
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Instant win conditions
            MyCommon.QueryStr = "select IncentiveInstantWinID,OddsOfWinning,NumPrizesAllowed,RandomWinners,RequiredFromTemplate from CPE_IncentiveInstantWin with (NoLock) " & _
                                "where Deleted=0 and RewardOptionID=" & roid & ";"
            rst = MyCommon.LRT_Select
            i = 1
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.instantwinconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Sendb("<li>")
                Sendb("1:" & MyCommon.NZ(row.Item("OddsOfWinning"), "0") & " ")
                If MyCommon.NZ(row.Item("RandomWinners"), False) Then
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.random", LanguageID), VbStrConv.Lowercase) & " ")
                Else
                  Sendb(StrConv(Copient.PhraseLib.Lookup("term.fixed", LanguageID), VbStrConv.Lowercase) & " ")
                End If
                Sendb(StrConv(Copient.PhraseLib.Lookup("term.odds", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.on", LanguageID), VbStrConv.Lowercase) & " ")
                Sendb(MyCommon.NZ(row.Item("NumPrizesAllowed"), "0") & " " & StrConv(Copient.PhraseLib.Lookup("term.prizes", LanguageID), VbStrConv.Lowercase))
                If (MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(" <span class=""red"">")
                  Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                  Sendb("</span>")
                End If
                Send("</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Trigger code (aka PLU) conditions
            MyCommon.QueryStr = "select IncentivePLUID, PLU, PerRedemption, CashierMessage, RequiredFromTemplate " & _
                                "from CPE_IncentivePLUs with (NoLock) " & _
                                "where RewardOptionID=" & roid & " order by PLU;"
            rst = MyCommon.LRT_Select
            i = 1
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.triggercodeconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Sendb("<li>")
                If MyCommon.NZ(row.Item("PLU"), "") <> "" Then
                  Sendb(MyCommon.NZ(row.Item("PLU"), Copient.PhraseLib.Lookup("term.undefined", LanguageID)))
                Else
                  Sendb(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
                End If
                If (MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                  Sendb(" <span class=""red"">")
                  Sendb("<small>(" & Copient.PhraseLib.Lookup("term.requiredbytemplate", LanguageID) & ")</small>")
                  Sendb("</span>")
                End If
                Send("</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Enterprise instant win conditions
            MyCommon.QueryStr = "select IncentiveEIWID, NumberOfPrizes, EIW.FrequencyID, EIWF.Description, DisallowEdit, RequiredFromTemplate " & _
                                "from CPE_IncentiveEIW as EIW with (NoLock) " & _
                                "inner join CPE_IncentiveEIWFrequency as EIWF on EIWF.FrequencyID=EIW.FrequencyID " & _
                                "where RewardOptionID=" & roid & ";"
            rst = MyCommon.LRT_Select
            i = 1
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.epriseinstantwinconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Send("  <li>")
                MyCommon.QueryStr = "select count(*) as TriggerCount from CPE_EIWTriggers where IncentiveEIWID=" & MyCommon.NZ(row.Item("IncentiveEIWID"), 0) & " and RewardOptionID=" & roid & " and Removed=0;"
                rst2 = MyCommon.LRT_Select
                If rst2.Rows(0).Item("TriggerCount") = 0 Then
                  Sendb("    " & Copient.PhraseLib.Lookup("term.no", LanguageID))
                Else
                  Sendb("    " & MyCommon.NZ(rst2.Rows(0).Item("TriggerCount"), 0))
                End If
                If rst2.Rows(0).Item("TriggerCount") = 1 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.trigger", LanguageID), VbStrConv.Lowercase))
                Else
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.triggers", LanguageID), VbStrConv.Lowercase))
                End If
                Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase))
                Sendb(" " & MyCommon.NZ(row.Item("NumberOfPrizes"), 0))
                If MyCommon.NZ(row.Item("NumberOfPrizes"), 0) = 1 Then
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.prize", LanguageID), VbStrConv.Lowercase))
                Else
                  Sendb(" " & StrConv(Copient.PhraseLib.Lookup("term.prizes", LanguageID), VbStrConv.Lowercase))
                End If
                Send(" " & StrConv(MyCommon.NZ(row.Item("Description"), ""), VbStrConv.Lowercase))
                Send("  </li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Card type conditions
            Dim CardTypes As String = ""
            MyCommon.QueryStr = "select CardTypeID from CPE_IncentiveCardTypes with (NoLock) where Deleted=0 and RewardOptionID=" & roid & ";"
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              MyCommon.QueryStr = "select Description, PhraseID from CardTypes with (NoLock) where CardTypeID=" & MyCommon.NZ(row.Item("CardTypeID"), -1) & ";"
              rst2 = MyCommon.LXS_Select
              If rst2.Rows.Count > 0 Then
                If CardTypes <> "" Then
                  CardTypes &= ", "
                End If
                If MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0) > 0 Then
                  CardTypes &= Copient.PhraseLib.Lookup(MyCommon.NZ(rst2.Rows(0).Item("PhraseID"), 0), LanguageID)
                Else
                  CardTypes &= MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
                End If
              Else
                If CardTypes <> "" Then
                  CardTypes &= ", "
                End If
                CardTypes &= Copient.PhraseLib.Lookup("term.unknown", LanguageID)
              End If
            Next
            If (CardTypes.Trim.Length > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.cardtypeconditions", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              Send("<li>" & CardTypes & "</li>")
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Preference conditions
            If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
              MyCommon.QueryStr = "select CIP.PreferenceID, RO.PreferenceComboID, CIP.IncentivePrefsID from CPE_IncentivePrefs as CIP with (NoLock) " & _
                                  "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID = CIP.RewardOptionID " & _
                                  "where CIP.RewardOptionID = " & roid
              rst = MyCommon.LRT_Select
              i = 1
              If (rst.Rows.Count > 0) Then
                counter = counter + 1
                Send("<h3>" & Copient.PhraseLib.Lookup("term.preferenceconditions", LanguageID) & "</h3>")
                Send("<ul class=""condensed"">")
                For Each row In rst.Rows
                  Send("  <li>")
                  Send_Preference_Details(MyCommon, MyCommon.NZ(row.Item("PreferenceID"), 0))
                  Send_Preference_Info(MyCommon, MyCommon.NZ(row.Item("IncentivePrefsID"), 0), roid, TierLevels)
                  
                  If i < rst.Rows.Count Then
                    Send(" <br /><i>" & Copient.PhraseLib.Lookup("term." & IIf(MyCommon.NZ(row.Item("PreferenceComboID"), 1) = 1, "and", "or"), LanguageID).ToLower & "</i> ")
                  End If
                    
                  Send("  </li>")
                  i += 1
                Next
                Send("</ul>")
                Send("<br class=""half"" />")
              End If
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
            Send("<h2 style=""float:left;""><span><a href=""CPEoffer-rew.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("rewardbody", "imgRewards", Copient.PhraseLib.Lookup("term.rewards", LanguageID), True)
        %>
        <div id="rewardbody">
          <%
            ' Discount reward
            MyCommon.QueryStr = "select DISC.DiscountID, DISC.DiscountTypeID, DISC.Name, DISC.DiscountTypeId, DISC.ReceiptDescription, DISC.DiscountBarcode, DISC.VoidBarcode, " & _
                                "DISC.DiscountAmount, DISC.DiscountedProductGroupID as SelectedPG, DISC.ItemLimit, DISC.WeightLimit, DISC.IsWeightTotal, DISC.DollarLimit, DISC.ExcludedProductGroupID as ExcludedPG, " & _
                                "DISC.DiscountAmount, DISC.ChargebackDeptID, DISC.AmountTypeID, DISC.L1Cap, DISC.L2DiscountAmt, DISC.L2AmountTypeID, DISC.L2Cap, DISC.L3DiscountAmt, DISC.L3AmountTypeID, " & _
                                "DISC.DecliningBalance, DISC.RetroactiveDiscount, DISC.UserGroupID, DISC.BestDeal, DISC.AllowNegative, DISC.ComputeDiscount, DISC.SVProgramID, " & _
                                "D.DeliverableID, AT.AmountTypeID, AT.PhraseID as AmountPhraseID, DT.PhraseID as DiscountPhraseID " & _
                                "from CPE_Deliverables D with (NoLock) inner join CPE_Discounts DISC with (NoLock) on D.OutputID = DISC.DiscountID " & _
                                "left join CPE_AmountTypes AT with (NoLock) on AT.AmountTypeID = DISC.AmountTypeID " & _
                                "left join CPE_DiscountTypes DT with (NoLock) on DT.DiscountTypeID = DISC.DiscountTypeID " & _
                                "where D.Deleted = 0 and DISC.Deleted = 0 and D.RewardOptionPhase=3 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=2;"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Dim Details As StringBuilder
              Send("<h3>" & Copient.PhraseLib.Lookup("term.discounts", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Details = New StringBuilder(200)
                MyCommon.QueryStr = "select DiscountAmount, ItemLimit, WeightLimit, IsWeightTotal, DollarLimit from CPE_DiscountTiers where DiscountID=" & MyCommon.NZ(row.Item("DiscountID"), 0) & ";"
                rstTiers = MyCommon.LRT_Select
                t = 1
                Select Case (MyCommon.NZ(row.Item("AmountTypeID"), 0))
                  Case 1, 5
                    If rstTiers.Rows.Count > 0 Then
                      For Each row4 In rstTiers.Rows
                        Details.Append("$" & Format(MyCommon.NZ(row4.Item("DiscountAmount"), ""), "#####0.00#"))
                        If t < rstTiers.Rows.Count Then
                          Details.Append(" / ")
                        End If
                        t += 1
                      Next
                      Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                    End If
                  Case 3
                    If rstTiers.Rows.Count > 0 Then
                      For Each row4 In rstTiers.Rows
                        Details.Append(Math.Round(MyCommon.NZ(row4.Item("DiscountAmount"), ""), 2) & "% ")
                        If t < rstTiers.Rows.Count Then
                          Details.Append(" / ")
                        End If
                        t += 1
                      Next
                      Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                    End If
                  Case 4
                    Details.Append(Copient.PhraseLib.Lookup("term.free", LanguageID) & "&nbsp;")
                  Case 2, 6
                    If rstTiers.Rows.Count > 0 Then
                      For Each row4 In rstTiers.Rows
                        Details.Append("$" & Math.round(MyCommon.NZ(row4.Item("DiscountAmount"), ""), 2) & "&nbsp;")
                        If t < rstTiers.Rows.Count Then
                          Details.Append(" / ")
                        End If
                        t += 1
                      Next
                    End If
                  Case 7
                    MyCommon.QueryStr = "select SVProgramID, Name from StoredValuePrograms with (NoLock) where SVProgramID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & ";"
                    rst3 = MyCommon.LRT_Select
                    If (rst3.Rows.Count > 0) Then
                      Details.Append(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & " (<a href=""sv-edit.aspx?ProgramGroupID=" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("SVProgramID"), 0), 25) & """>" & MyCommon.NZ(rst3.Rows(0).Item("Name"), "") & "</a>) " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase) & " ")
                    End If
                  Case 8
                    Details.Append(Copient.PhraseLib.Lookup("term.specialpricing", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.for", LanguageID), VbStrConv.Lowercase) & " ")
                  Case Else
                    If rstTiers.Rows.Count > 0 Then
                      For Each row4 In rstTiers.Rows
                        Details.Append(MyCommon.NZ(row4.Item("DiscountAmount"), "") & "&nbsp;")
                        If t < rstTiers.Rows.Count Then
                          Details.Append(" / ")
                        End If
                        t += 1
                      Next
                    End If
                End Select
                
                If MyCommon.NZ(row.Item("SelectedPG"), 0) = 0 Then
                  If MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 4 Then
                    'Bundled pricing discount -- get products from conditions
                    MyCommon.QueryStr = "select PG.ProductGroupID, PG.Name, PG.PhraseID, PG.AnyProduct " & _
                                        "from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                        "inner join ProductGroups as PG with (NoLock) on PG.ProductGroupID=IPG.ProductGroupID " & _
                                        "inner join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID " & _
                                        "where RO.IncentiveID = " & OfferID & " And IPG.Deleted = 0 And Disqualifier = 0 And ProductComboID = 1 And QtyUnitType = 1 And QtyForIncentive = 1 " & _
                                        "order by AnyProduct DESC, Name;"
                    rst3 = MyCommon.LRT_Select
                    If rst3.Rows.Count > 1 Then
                      i = 1
                      For Each row3 In rst3.Rows
                        If MyCommon.NZ(row3.Item("PhraseID"), 0) > 0 Then
                          Details.Append(Copient.PhraseLib.Lookup(row3.Item("PhraseID"), LanguageID))
                        Else
                          Details.Append("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row3.Item("ProductGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                        End If
                        If i < rst3.Rows.Count Then
                          Details.Append(" / ")
                        End If
                        i += 1
                      Next
                    End If
                  Else
                    Details.Append(StrConv(Copient.PhraseLib.Lookup("term.nothing", LanguageID), VbStrConv.Lowercase))
                  End If
                Else
                  MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID = " & row.Item("SelectedPG")
                  rst3 = MyCommon.LRT_Select()
                  For Each row3 In rst3.Rows
                    If MyCommon.NZ(row.Item("SelectedPG"), 0) = 1 Then
                      Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase))
                    Else
                      If Popup Then
                        Details.Append(MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25))
                      Else
                        Details.Append("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("SelectedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                      End If
                    End If
                  Next
                End If
                If MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 2 Then
                  Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.department", LanguageID), VbStrConv.Lowercase) & ")")
                ElseIf MyCommon.NZ(row.Item("DiscountTypeID"), 0) = 3 Then
                  Details.Append(" (" & StrConv(Copient.PhraseLib.Lookup("term.basket", LanguageID), VbStrConv.Lowercase) & ")")
                End If
                
                If MyCommon.NZ(row.Item("ExcludedPG"), 0) = 0 Then
                Else
                  MyCommon.QueryStr = "select Name from ProductGroups with (NoLock) where ProductGroupID = " & row.Item("ExcludedPG")
                  rst3 = MyCommon.LRT_Select()
                  For Each row3 In rst3.Rows
                    Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.excluding", LanguageID), VbStrConv.Lowercase) & " ")
                    If MyCommon.NZ(row.Item("ExcludedPG"), 0) = 1 Then
                      Details.Append(StrConv(MyCommon.NZ(row3.Item("Name"), ""), VbStrConv.Lowercase))
                    Else
                      If Popup Then
                        Details.Append(MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25))
                      Else
                        Details.Append("<a href=""pgroup-edit.aspx?ProductGroupID=" & MyCommon.NZ(row.Item("ExcludedPG"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row3.Item("Name"), ""), 25) & "</a>")
                      End If
                    End If
                  Next
                End If
                
                If MyCommon.NZ(row.Item("L1Cap"), 0) > 0 Then
                  Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.upto", LanguageID), VbStrConv.Lowercase) & " " & FormatCurrency(row.Item("L1Cap")))
                End If
                
                If MyCommon.NZ(row.Item("ItemLimit"), 0) = 0 And MyCommon.NZ(row.Item("WeightLimit"), 0) = 0 And MyCommon.NZ(row.Item("DollarLimit"), 0) = 0 Then
                  Details.Append(",&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.unlimited", LanguageID), VbStrConv.Lowercase))
                Else
                  If MyCommon.NZ(row.Item("DiscountTypeID"), 0) <> 3 Then
                    Details.Append(",&nbsp;" & StrConv(Copient.PhraseLib.Lookup("term.limit", LanguageID), VbStrConv.Lowercase) & "&nbsp;")
                    If rstTiers.Rows.Count > 0 Then
                      If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 Then
                        t = 1
                        For Each row4 In rstTiers.Rows
                          Details.Append(MyCommon.NZ(row4.Item("ItemLimit"), ""))
                          If t < rstTiers.Rows.Count Then
                            Details.Append(" / ")
                          End If
                          t += 1
                        Next
                        Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.items", LanguageID), VbStrConv.Lowercase))
                      End If
                      If MyCommon.NZ(rstTiers.Rows(0).Item("DollarLimit"), 0) > 0 Then
                        If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 Then
                          Details.Append(" (")
                        End If
                        t = 1
                        For Each row4 In rstTiers.Rows
                          Details.Append(FormatCurrency(MyCommon.NZ(row4.Item("DollarLimit"), 0)))
                          If t < rstTiers.Rows.Count Then
                            Details.Append(" / ")
                          End If
                          t += 1
                        Next
                        If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 Then
                          Details.Append(")")
                        End If
                      End If
                      If MyCommon.NZ(rstTiers.Rows(0).Item("WeightLimit"), 0) > 0 Then
                        If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 OrElse MyCommon.NZ(rstTiers.Rows(0).Item("DollarLimit"), 0) > 0 Then
                          Details.Append(" (")
                        End If
                        t = 1
                        For Each row4 In rstTiers.Rows
                          Details.Append(Math.Round(MyCommon.NZ(row4.Item("WeightLimit"), ""),3))
                          Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.lbsgals", LanguageID), VbStrConv.Lowercase))
                          Details.Append(" " & IIf(row4.Item("IsWeightTotal") = True, Copient.PhraseLib.Lookup("term.total", LanguageID).ToLower(), Copient.PhraseLib.Lookup("term.peritem", LanguageID).ToLower()))
                          If t < rstTiers.Rows.Count Then
                            Details.Append(" / ")
                          End If
                          t += 1
                        Next
                        If MyCommon.NZ(rstTiers.Rows(0).Item("ItemLimit"), 0) > 0 OrElse MyCommon.NZ(rstTiers.Rows(0).Item("DollarLimit"), 0) > 0 Then
                          Details.Append(")")
                        End If
                      End If
                    End If
                  End If
                End If
                ' If there are multiple levels, this will display their details on a second line.
                If (MyCommon.NZ(row.Item("L2DiscountAmt"), 0) > 0) And MyCommon.NZ(row.Item("L2AmountTypeID"), 0) = 3 Then
                  Details.Append("<br />(" & Copient.PhraseLib.Lookup("term.over", LanguageID) & " " & FormatCurrency(MyCommon.NZ(row.Item("L1Cap"), "0")) & ", ")
                  Details.Append(row.Item("L2DiscountAmt") & "% " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase))
                  If (MyCommon.NZ(row.Item("L2Cap"), 0) > 0) Then
                    Details.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.upto", LanguageID), VbStrConv.Lowercase) & " " & FormatCurrency(row.Item("L2Cap")))
                  End If
                  Details.Append(")")
                  If (MyCommon.NZ(row.Item("L3DiscountAmt"), 0) > 0) And MyCommon.NZ(row.Item("L3AmountTypeID"), 0) = 3 Then
                    Details.Append("<br />(" & Copient.PhraseLib.Lookup("term.over", LanguageID) & " " & FormatCurrency(MyCommon.NZ(row.Item("L2Cap"), "0")) & ", ")
                    Details.Append(row.Item("L3DiscountAmt") & "% " & StrConv(Copient.PhraseLib.Lookup("term.off", LanguageID), VbStrConv.Lowercase) & ")")
                  End If
                End If
                Send("<li>" & Details.ToString & "</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Printed message rewards
            MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, D.DeliverableID " & _
                                "from CPE_Deliverables D with (NoLock) inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                "where D.Deleted=0 and D.RewardOptionPhase=3 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4;"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Dim Details As StringBuilder
              Send("<h3>" & Copient.PhraseLib.Lookup("term.printedmessages", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                MyCommon.QueryStr = "select PMT.BodyText from PrintedMessageTiers PMT with (NoLock) " & _
                                    "where MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & " order by TierLevel;"
                rstTiers = MyCommon.LRT_Select()
                If rstTiers.Rows.Count > 0 Then
                  For Each row4 In rstTiers.Rows
                    Details = New StringBuilder(200)
                    Details.Append(ReplaceTags(MyCommon.NZ(row4.Item("BodyText"), "")))
                    If (Details.ToString().Length > 80) Then
                      Details = Details.Remove(77, (Details.Length - 77))
                      Details.Append("...")
                    End If
                    Details.Replace(vbCrLf, "<br />")
                    Send("<li>""" & HttpUtility.HtmlEncode(MyCommon.SplitNonSpacedString(Details.ToString, 25)) & """</li>")
                  Next
                End If
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Cashier message rewards
            MyCommon.QueryStr = "select D.DeliverableID, CM.MessageID from CPE_Deliverables D with (NoLock) " & _
                                "inner join CPE_CashierMessages CM with (NoLock) on D.OutputID=CM.MessageID " & _
                                "where D.RewardOptionID=" & roid & " and D.Deleted=0 and DeliverableTypeID=9 and D.RewardOptionPhase=3;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.cashiermessages", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                MyCommon.QueryStr = "select Line1, Line2, Beep, BeepDuration from CPE_CashierMessageTiers as CMT with (NoLock) " & _
                                    "where MessageID=" & MyCommon.NZ(row.Item("MessageID"), 0) & " order by TierLevel;"
                rstTiers = MyCommon.LRT_Select()
                If rstTiers.Rows.Count > 0 Then
                  For Each row4 In rstTiers.Rows
                    Send("<li>""" & MyCommon.NZ(row4.Item("Line1"), "") & "<br />" & MyCommon.NZ(row4.Item("Line2"), "") & """</li>")
                  Next
                End If
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Franking rewards
            MyCommon.QueryStr = "select D.DeliverableID, FM.FrankID from CPE_Deliverables D with (NoLock) " & _
                                "inner join CPE_FrankingMessages FM with (NoLock) on D.OutputID=FM.FrankID " & _
                                "where D.RewardOptionID=" & roid & " and D.Deleted=0 and DeliverableTypeID=10 and D.RewardOptionPhase=3;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.frankingmessages", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                MyCommon.QueryStr = "select FrankID, OpenDrawer, ManagerOverride, FrankFlag, FrankingText, Line1, Line2, Beep, BeepDuration " & _
                                    "from CPE_FrankingMessageTiers as FMT with (NoLock) " & _
                                    "where FrankID=" & MyCommon.NZ(row.Item("FrankID"), 0) & " order by TierLevel;"
                rstTiers = MyCommon.LRT_Select()
                If rstTiers.Rows.Count > 0 Then
                  For Each row4 In rstTiers.Rows
                    If MyCommon.NZ(row4.Item("FrankingText"), "") = "" Then
                      Send("<li>")
                    Else
                      Send("<li>""" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row4.Item("FrankingText"), ""), 25) & """<br />")
                    End If
                    Sendb(IIf(MyCommon.NZ(row4.Item("OpenDrawer"), False) = True, Copient.PhraseLib.Lookup("term.opendrawer", LanguageID) & ",&nbsp;", Copient.PhraseLib.Lookup("term.closeddrawer", LanguageID) & ",&nbsp;"))
                    Sendb(IIf(MyCommon.NZ(row4.Item("ManagerOverride"), False) = True, StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.required", LanguageID), VbStrConv.Lowercase) & ",&nbsp;", StrConv(Copient.PhraseLib.Lookup("term.override", LanguageID), VbStrConv.Lowercase) & " " & StrConv(Copient.PhraseLib.Lookup("term.notrequired", LanguageID), VbStrConv.Lowercase) & ",&nbsp;"))
                    If (MyCommon.NZ(row4.Item("FrankFlag"), 0) = 0) Then
                      Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                    ElseIf (MyCommon.NZ(row4.Item("FrankFlag"), 0) = 1) Then
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.only", LanguageID), VbStrConv.Lowercase))
                    ElseIf (MyCommon.NZ(row4.Item("FrankFlag"), 0) = 2) Then
                      Sendb(Copient.PhraseLib.Lookup("term.posdata", LanguageID) & " ")
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
                      Sendb(StrConv(Copient.PhraseLib.Lookup("term.frankingmessage", LanguageID), VbStrConv.Lowercase) & " ")
                    End If
                    Send("</li>")
                  Next
                End If
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Points rewards
            MyCommon.QueryStr = "select PP.ProgramName as Name, PP.ProgramID, D.DeliverableID as PKID, DP.PKID as DPPKID " & _
                                "from PointsPrograms as PP with (NoLock) " & _
                                "inner join CPE_DeliverablePoints as DP with (NoLock) on PP.ProgramID=DP.ProgramID and DP.Deleted=0 and PP.Deleted=0 " & _
                                "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DP.DeliverableID and D.Deleted=0 and D.RewardOptionPhase=3 " & _
                                "where D.RewardOptionID=" & roid & " order by PP.ProgramName;"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.points", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                MyCommon.QueryStr = "select Quantity from CPE_DeliverablePointTiers as DPT with (NoLock) " & _
                                    "where DPPKID=" & MyCommon.NZ(row.Item("DPPKID"), 0) & " order by TierLevel;"
                rstTiers = MyCommon.LRT_Select()
                If rstTiers.Rows.Count > 0 Then
                  Sendb("<li>")
                  t = 1
                  For Each row4 In rstTiers.Rows
                    Sendb(MyCommon.NZ(row4.Item("Quantity"), "") & " ")
                    If t < TierLevels Then
                      Sendb(" / ")
                    End If
                    t = t + 1
                  Next
                End If
                If Popup Then
                  Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                Else
                  Sendb("<a href=""point-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("ProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                End If
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Stored value rewards
            MyCommon.QueryStr = "select distinct DSV.PKID, DSV.DeliverableID, SVP.Name, SVP.Value, SVP.SVTypeID, SVT.ValuePrecision, SVP.SVProgramID from CPE_Deliverables as D with (NoLock) " & _
                                "Inner Join CPE_DeliverableStoredValue as DSV with (NoLock) on DSV.PKID=D.OutputID and IsNull(DSV.Deleted,0)=0 " & _
                                "Inner Join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and SVP.Deleted=0 " & _
                                "Inner Join SVTypes as SVT with (NoLock) on SVP.SVTypeID=SVT.SVTypeID " & _
                                "where D.RewardOptionPhase = 3 And D.DeliverableTypeID = 11 And D.RewardOptionID =" & roid
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.storedvalue", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Sendb("<li>")
                MyCommon.QueryStr = "select Quantity from CPE_DeliverableStoredValueTiers with (NoLock) where DSVPKID=" & MyCommon.NZ(row.Item("PKID"), 0) & ";"
                rstTiers = MyCommon.LRT_Select
                If rstTiers.Rows.Count > 0 Then
                  t = 1
                  For Each row4 In rstTiers.Rows
                    Sendb(CInt(MyCommon.NZ(row4.Item("Quantity"), "0")) & " ")
                    If MyCommon.NZ(row.Item("SVTypeID"), 0) > 1 Then
                      Sendb(" ($" & Math.Round(MyCommon.NZ(row.Item("Value"), 0) * MyCommon.NZ(row4.Item("Quantity"), 0), MyCommon.NZ(row.Item("ValuePrecision"), 0)) & ")")
                    End If
                    If t < TierLevels Then
                      Sendb(" / ")
                    End If
                    t = t + 1
                  Next
                End If
                If Popup Then
                  Sendb(" " & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25))
                Else
                  Sendb(" <a href=""SV-edit.aspx?ProgramGroupID=" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), ""), 25) & "</a>")
                End If
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            'Group membership rewards
            MyCommon.QueryStr = "select D.DeliverableID, D.DeliverableTypeID, D.RewardOptionID as ROID " & _
                                "from CPE_Deliverables as D with (NoLock) " & _
                                "where D.Deleted=0 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID IN (5) and D.RewardOptionPhase=3;"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.groupmembership", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                MyCommon.QueryStr = "select DCGT.CustomerGroupID, CG.Name from CPE_DeliverableCustomerGroupTiers as DCGT with (NoLock) " & _
                                    "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=DCGT.CustomerGroupID " & _
                                    "where DeliverableID=" & MyCommon.NZ(row.Item("DeliverableID"), 0) & ";"
                rstTiers = MyCommon.LRT_Select
                If rstTiers.Rows.Count > 0 Then
                  For Each row4 In rstTiers.Rows
                    If (MyCommon.NZ(row4.Item("CustomerGroupID"), 0) > 2) Then
                      If Popup Then
                        Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row4.Item("Name"), ""), 25) & "&nbsp;")
                      Else
                        Sendb("<li><a href=""cgroup-edit.aspx?CustomerGroupID=" & MyCommon.NZ(row4.Item("CustomerGroupID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row4.Item("Name"), ""), 25) & "</a>&nbsp;")
                      End If
                    Else
                      Sendb("<li>" & MyCommon.SplitNonSpacedString(row.Item("Name"), 25) & "&nbsp;")
                    End If
                    MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID=" & MyCommon.NZ(row4.Item("CustomerGroupID"), -1) & " And Deleted=0;"
                    rst2 = MyCommon.LXS_Select()
                    For Each row2 In rst2.Rows
                      If (MyCommon.NZ(row4.Item("CustomerGroupID"), -1) > 2) Then
                        Send("(" & row2.Item("GCount") & ")")
                      End If
                    Next
                    Send("</li>")
                  Next
                End If
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Graphics rewards
            MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, D.DeliverableID, D.ScreenCellID as CellID, OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " & _
                                "from OnScreenAds as OSA with (NoLock) Inner Join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID and D.RewardOptionID=" & roid & " and OSA.Deleted=0 and D.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=3 " & _
                                "Inner Join ScreenCells as SC with (NoLock) on D.ScreenCellID=SC.CellID " & _
                                "Inner Join ScreenLayouts as SL with (NoLock) on SL.LayoutID=SC.LayoutID;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Send("<h3>" & Copient.PhraseLib.Lookup("term.graphics", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                If Popup Then
                  Sendb("<li>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "&nbsp;")
                Else
                  Sendb("<li><a href=""graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("AdId"), "") & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("GraphicName"), ""), 25) & "</a>&nbsp;")
                End If
                Sendb("(" & MyCommon.NZ(row.Item("Width"), "") & " x " & MyCommon.NZ(row.Item("Height"), ""))
                If MyCommon.NZ(row.Item("ImageType"), "") = 1 Then
                  Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                ElseIf MyCommon.NZ(row.Item("ImageType"), "") = 2 Then
                  Sendb("&nbsp;" & Copient.PhraseLib.Lookup("term.gif", LanguageID))
                End If
                Send(")</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Touchpoint rewards
            MyCommon.QueryStr = "select RO.Name, RO.RewardOptionID, TA.OnScreenAdID as ParentAdID " & _
                                "from CPE_RewardOptions RO with (NoLock) inner join CPE_DeliverableROIDs DR with (NoLock) on RO.RewardOptionID = DR.RewardOptionID " & _
                                "inner join CPE_Deliverables D with (NoLock) on D.DeliverableID = DR.DeliverableID inner join TouchAreas TA with (NoLock) on DR.AreaID = TA.AreaID " & _
                                "where RO.Deleted=0 and DR.Deleted=0 and TA.Deleted=0 and D.Deleted = 0 and RO.IncentiveID=" & OfferID & " and " & _
                                "RO.TouchResponse=1 and D.RewardOptionPhase=3 order by RO.rewardoptionid;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Dim tpROID As Integer = 0
              Send("<h3>" & Copient.PhraseLib.Lookup("term.touchpoints", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                tpROID = MyCommon.NZ(row.Item("RewardOptionID"), 0)
                Send("<li>")
                If Popup Then
                  Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "<br />")
                Else
                  Send("<a href=""graphic-edit.aspx?OnScreenAdId=" & MyCommon.NZ(row.Item("ParentAdID"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID)), 25) & "</a><br />")
                End If
                SetSummaryOnly(True)
                Send_TouchpointRewards(OfferID, tpROID, 3, TierLevels)
                Send("</li>")
              Next
              Send("</ul>")
            End If
            
            ' Pass-thru rewards
            MyCommon.QueryStr = "select PassThruRewardID, Name, PhraseID from PassThruRewards with (NoLock) order by Name;"
            rst2 = MyCommon.LRT_Select
            If rst2.Rows.Count > 0 Then
              For Each row2 In rst2.Rows
                
                MyCommon.QueryStr = "select D.DeliverableID, D.DisallowEdit, DPT.PKID, DPT.PassThruRewardID, PTR.Name, PTR.PhraseID, PTR.LSInterfaceID, PTR.ActionTypeID " & _
                                    "from CPE_Deliverables as D with (NoLock) " & _
                                    "inner join PassThrus as DPT with (NoLock) on DPT.PKID=D.OutputID " & _
                                    "inner join PassThruRewards as PTR with (NoLock) on PTR.PassThruRewardID=DPT.PassThruRewardID " & _
                                    "where D.RewardOptionID=" & roid & " and DPT.PassThruRewardID=" & MyCommon.NZ(row2.Item("PassThruRewardID"), 0) & " and D.Deleted=0 and D.DeliverableTypeID=12 " & _
                                    "order by Name;"
                rst = MyCommon.LRT_Select()
                If (rst.Rows.Count > 0) Then
                  counter = counter + 1
                  If IsDBNull(row2.Item("PhraseID")) Then
                    Send("<h3>" & MyCommon.NZ(row2.Item("Name"), Copient.PhraseLib.Lookup("term.passthrureward", LanguageID)) & "</h3>")
                  Else
                    Send("<h3>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</h3>")
                  End If
                  Send("<ul class=""condensed"">")
                  For Each row In rst.Rows
                    If IsDBNull(row.Item("PhraseID")) Then
                      Send("  <li>" & MyCommon.NZ(row2.Item("Name"), "") & "</li>")
                    Else
                      Send("  <li>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</li>")
                    End If
                  Next
                  Send("</ul>")
                  Send("<br class=""half"" />")
                End If
                
              Next
            End If
            
            If (counter = 0) Then
              Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
            End If
            counter = 0
          %>
          <hr class="hidden" />
        </div>
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
<!-- #Include virtual="/include/graphic-reward.inc" -->
<script runat="server">
  Sub WriteComponent(ByRef MyCommon As Copient.CommonInc, ByVal rowComp As DataRow, ByRef ComponentColor As String, ByVal Popup As Boolean)
    Dim RecordType As String = ""
    Dim ID As Integer
    Dim StoredProcName As String = ""
    Dim IDParmName As String = ""
    Dim TypeCode As String = ""
    Dim PageName As String = ""
    Dim dtValid As DataTable
    Dim rowOK(), rowWatches(), rowWarnings() As DataRow
    Dim objTemp As Object
    Dim GraceHours As Integer
    Dim GraceCount As Double
    Dim ShowSubReport As Boolean = True
    Dim ShowLinks As Boolean
    
    If Popup Then
      ShowLinks = False
    Else
      ShowLinks = True
    End If
    
    objTemp = MyCommon.Fetch_CPE_SystemOption(41)
    If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
      GraceHours = 4
    End If
    
    objTemp = MyCommon.Fetch_CPE_SystemOption(42)
    If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
      GraceCount = 0.1D
    End If
    
    RecordType = MyCommon.NZ(rowComp.Item("RecordType"), "")
    ID = MyCommon.NZ(rowComp.Item("ID"), -1)
    
    Select Case RecordType
      Case "term.customergroup"
        StoredProcName = "dbo.pa_ValidationReport_CustGroup"
        IDParmName = "@CustomerGroupID"
        TypeCode = "cg"
        PageName = "cgroup-edit.aspx?CustomerGroupID="
        ShowSubReport = IIf(ID = 1 OrElse ID = 2, False, True)
      Case "term.productgroup"
        StoredProcName = "dbo.pa_ValidationReport_ProdGroup"
        IDParmName = "@ProductGroupID"
        TypeCode = "pg"
        PageName = "pgroup-edit.aspx?ProductGroupID="
        ShowSubReport = IIf(ID = 1, False, True)
      Case "term.graphics"
        StoredProcName = "dbo.pa_ValidationReport_Graphic"
        IDParmName = "@OnScreenAdID"
        TypeCode = "gr"
        PageName = "graphic-edit.aspx?OnScreenAdID="
    End Select
    
    MyCommon.QueryStr = StoredProcName
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add(IDParmName, SqlDbType.Int).Value = ID
    MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
    MyCommon.LRTsp.Parameters.Add("@GraceCount", SqlDbType.Decimal, 2).Value = GraceCount
    
    dtValid = MyCommon.LRTsp_select()
    
    rowOK = dtValid.Select("Status=0", "LocationName")
    rowWatches = dtValid.Select("Status=1", "LocationName")
    rowWarnings = dtValid.Select("Status=2", "LocationName")
    
    If (ShowSubReport AndAlso ComponentColor <> "red") Then
      ComponentColor = IIf(rowWarnings.Length > 0, "red", "green")
    End If
    
    Send("<div style=""margin-left:10px;"">")
    Sendb(Copient.PhraseLib.Lookup(MyCommon.NZ(rowComp.Item("RecordType"), ""), LanguageID) & " #" & ID & ": ")
    If (ShowSubReport) Then
      If (ShowLinks) Then
        Send("<a href=""" & PageName & ID & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rowComp.Item("Name"), "&nbsp;"), 20) & "</a>")
      Else
        Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(rowComp.Item("Name"), "&nbsp;"), 20))
      End If
    Else
      Send(MyCommon.SplitNonSpacedString(MyCommon.NZ(rowComp.Item("Name"), "&nbsp;"), 20))
    End If
    
    If (ShowSubReport) Then
      Send("<div style=""margin-left:20px;"">")
      If ShowLinks Then
        Send("<a id=""validLink" & ID & """ href=""javascript:openPopup('validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=0');"">")
        Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")")
        Send("</a><br />")
        Send("<a id=""watchLink" & ID & """ href=""javascript:openPopup('validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=1');"">")
        Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")")
        Send("</a><br />")
        Send("<a id=""warningLink" & ID & """ href=""javascript:openPopup('validation-report.aspx?type=" & TypeCode & "&id=" & ID & "&level=2');"">")
        Send("  " & Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")")
        Send("</a><br />")
      Else
        Send(Copient.PhraseLib.Lookup("cgroup-edit.validlocations", LanguageID) & " (" & rowOK.Length & ")<br />")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.watchlocations", LanguageID) & " (" & rowWatches.Length & ")<br />")
        Send(Copient.PhraseLib.Lookup("cgroup-edit.warninglocations", LanguageID) & " (" & rowWarnings.Length & ")<br />")
      End If

      Send("</div>")
    End If
    
    Send("</div>")
    MyCommon.Close_LRTsp()
  End Sub
  
  Sub RemoveInactiveLocations(ByRef MyCommon As Copient.CommonInc, ByRef dt As DataTable, ByVal IncentiveID As Integer)
    Dim dtLoc As DataTable
    Dim row As DataRow
    Dim LocationID As Integer
    Dim LocGroupList As String = ""
    Dim IsAllLocs As Boolean
    Dim LocTable As New Hashtable
    If (Not dt Is Nothing And dt.Rows.Count > 0) Then
      ' first find all locations is selected
      MyCommon.QueryStr = "select LocationGroupID from OfferLocations with (NoLock) where OfferID=" & IncentiveID & " and Deleted=0;"
      dtLoc = MyCommon.LRT_Select
      For Each row In dtLoc.Rows
        If (LocGroupList <> "") Then LocGroupList += ","
        LocGroupList = LocGroupList + MyCommon.NZ(row.Item("LocationGroupID"), "-1").ToString
        IsAllLocs = MyCommon.NZ(row.Item("LocationGroupID"), -1) = 1
      Next
      If (Not IsAllLocs AndAlso LocGroupList <> "") Then
        ' find all the locations for the given location groups
        MyCommon.QueryStr = "select LocationID from LocGroupItems with (NoLock) where Deleted = 0 " & _
                            "and LocationGroupID in (" & LocGroupList & ");"
        dtLoc = MyCommon.LRT_Select
        For Each row In dtLoc.Rows
          LocationID = MyCommon.NZ(row.Item("LocationID"), "-1")
          If (Not LocTable.ContainsKey(LocationID.ToString)) Then
            LocTable.Add(MyCommon.NZ(row.Item("LocationID"), "-1").ToString, MyCommon.NZ(row.Item("LocationID"), "-1").ToString)
          End If
        Next
        ' remove the location if it doesn't currently exist for the incentive
        For Each row In dt.Rows
          LocationID = MyCommon.NZ(row.Item("LocationID"), "-1")
          If (Not LocTable.ContainsKey(LocationID.ToString)) Then
            row.Delete()
          End If
        Next
      End If
    End If
  End Sub
</script>
<script type="text/javascript">
  collapseBoxes();
  setComponentsColor('<% Sendb(ComponentColor)%>');
</script>
<script type="text/javascript">
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
  <%
  If (StatusMessage <> "") Then
    Send("alert('" & StatusMessage & "');")
  End If
  %>
</script>
<script runat="server">
  
  '-----------------------------------------------------------------------------------------------------------------------------
  Function GetLastDeployValidationMessage(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long) As String
    Dim lastDeployValidationMsg As String = String.Empty
    Dim dt As DataTable
    MyCommon.QueryStr = "Select LastDeployValidationMessage From CPE_Incentives " & _
                          " Where IncentiveId=@OfferId"
    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If Not dt Is Nothing And dt.Rows.Count = 1 Then
      lastDeployValidationMsg = dt.Rows(0)(0).ToString()
    End If
    Return lastDeployValidationMsg
  End Function
  
  Sub SetLastDeployValidationMessage(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long, ByVal Message As String)
    MyCommon.QueryStr = "Update CPE_Incentives " & _
                       "  Set LastDeployValidationMessage=@Message " & _
                       "  where IncentiveId=@OfferId"
    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    MyCommon.DBParameters.Add("@Message", SqlDbType.NVarChar).Value = Message
    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
  End Sub

  Function IsDeployableOffer(ByRef Logix As Copient.LogixInc, ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByVal ROID As Integer, ByRef ErrorMsg As String) As Boolean
    Dim bDeployable As Boolean = False
    
    ErrorMsg = ""
    bDeployable = MeetsDeploymentReqs(MyCommon, OfferID)
    
    If bDeployable Then
      bDeployable = MeetsTemplateRequirements(MyCommon, ROID)
      If bDeployable Then
        bDeployable = MeetsTieredReqs(MyCommon, OfferID)
        If Not bDeployable Then
          ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.tier-setup-invalid", LanguageID)
        End If
      Else
        ErrorMsg = Copient.PhraseLib.Lookup("offer-sum.required-incomplete", LanguageID)
      End If
      If MyCommon.Fetch_SystemOption(131) = "1" AndAlso MyCommon.Fetch_SystemOption(67) = "0" Then
        bDeployable = MeetsLockOutRequirement(Logix, MyCommon, OfferID)
        If (Not bDeployable) Then
          ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.deployalertforlockout", LanguageID)
        End If
      End If
    Else
      ErrorMsg = Copient.PhraseLib.Lookup("cpeoffer-sum.deployalert", LanguageID)
    End If
    
    If GetCgiValue("deploytransreqskip") <> "1" Then
      If bDeployable Then bDeployable = MeetsTranslationRequirements(MyCommon, OfferID, ROID, ErrorMsg)
    End If
    
    Return bDeployable
  End Function
  
  '-----------------------------------------------------------------------------------------------------------------------------
  
  Function MeetsDeploymentReqs(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer) As Boolean
    Dim bMeetsReqs As Boolean = False
    
    ' The user wants to deploy, so do a quick check for at least one assigned offer location and terminal,
    ' and ensure that there are no unassigned tier values
    MyCommon.QueryStr = "dbo.pa_CPE_IsOfferDeployable"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@IsDeployable", SqlDbType.Bit).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    bMeetsReqs = MyCommon.LRTsp.Parameters("@IsDeployable").Value
    
    Return bMeetsReqs
  End Function
  
  '-----------------------------------------------------------------------------------------------------------------------------
  
  Function MeetsTieredReqs(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Long) As Boolean
    Dim bMeetsReqs As Boolean = False
    
    ' ensure that multi-tiered offers have a tierable condition if there is a tierable reward assigned to the offer.
    MyCommon.QueryStr = "dbo.pa_ValidateOfferTiers"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@IsDeployable", SqlDbType.Bit).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    bMeetsReqs = MyCommon.LRTsp.Parameters("@IsDeployable").Value
    
    Return bMeetsReqs
  End Function
  Function MeetsLockOutRequirement(ByRef Logix As Copient.LogixInc, ByRef MyCommon As Copient.CommonInc, ByVal OfferId As Integer) As Boolean
    
    Dim bMeetsLockoutReq As Boolean = True
    Dim FolderId As Long
    Dim LockOutDays As Integer
    Dim FolderStartDate As Date
    Dim dt As DataTable
    Dim BannerID As Integer = 0
    
    If Logix.UserRoles.DeployOffersPastLockoutDate = False Then
    
      MyCommon.QueryStr = "SELECT BannerID FROM BannerOffers WHERE OfferID=" & OfferId
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        BannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
      End If
      
      MyCommon.QueryStr = "select folderid from folderitems with (nolock) where linkid =" & OfferId
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        FolderId = dt.Rows(0).Item("folderid")
      End If
    
      MyCommon.QueryStr = "SELECT bt.Lockoutdays FROM BannerThemes bt INNER JOIN  FolderThemes ft ON ft.ThemeID = bt.ThemeID " & _
                        " AND bt.BannerID = " & BannerID & " INNER JOIN Folders fo ON ft.FolderID=fo.FolderID " & _
                        " WHERE fo.FolderID = " & FolderId & ""
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        LockOutDays = MyCommon.NZ(dt.Rows(0).Item("lockoutdays"), 0)
        MyCommon.QueryStr = "select startdate from folders with (NoLock) where FolderId=" & FolderId
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
          If Not IsDBNull(dt.Rows(0).Item("startdate")) Then
            FolderStartDate = dt.Rows(0).Item("startdate")
            If FolderStartDate <= Date.Now.AddDays(LockOutDays) Then
              bMeetsLockoutReq = False
            End If
          Else
            bMeetsLockoutReq = True
          End If
        End If
      End If
    End If
    Return bMeetsLockoutReq
  End Function
  '-----------------------------------------------------------------------------------------------------------------------------
  Function MeetsTemplateRequirements(ByRef MyCommon As Copient.CommonInc, ByVal ROID As Integer) As Boolean
    Dim dt As DataTable
    
    MyCommon.QueryStr = "select 'CG' as GroupType, CustomerGroupID as GroupID from CPE_IncentiveCustomerGroups with (NoLock) " & _
                        "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and CustomerGroupID is null " & _
                        "union " & _
                        "select 'PG' as GroupType, ProductGroupID as GroupID from CPE_IncentiveProductGroups with (NoLock) " & _
                            "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and ProductGroupID = -1 " & _
                        "union " & _
                        "select 'PP' as GroupType,ProgramID as GroupID from CPE_IncentivePointsGroups with (NoLock) " & _
                        "where RewardOptionID = " & ROID & " and Deleted=0 and RequiredFromTemplate=1 and ProgramId is null; "
    dt = MyCommon.LRT_Select
    
    Return (dt.Rows.Count = 0)
  End Function
  
  '-----------------------------------------------------------------------------------------------------------------------------
  
  Function MeetsTranslationRequirements(ByRef Common As Copient.CommonInc, ByVal OfferID As Long, ByVal ROID As Long, ByRef ErrorMsg As String) As Boolean
    Dim ReturnVal As Boolean = True
    Dim ErrorTerm As String = ""
    Dim TranslatedErrorTerm As String = ""
    Dim DeployType As String = ""

    ErrorMsg = ""
    If Common.Fetch_SystemOption(124) = "1" Then  'see if multi-language is enabled
      ErrorTerm = CheckForTranslationDeployError(Common, ROID)
      If Not (ErrorTerm = "") Then
        'if this section weren't in place, then you could get around the warning message by clicking defer deploy and then clicking deploy, without specifically addressing the warning
        If GetCgiValue("deferdeploy") <> "" Then
          DeployType = "deferdeploy"
        Else
          DeployType = "deploy"
        End If
        TranslatedErrorTerm = Copient.PhraseLib.DecodeEmbededTokens(ErrorTerm, LanguageID)
        ErrorMsg = Copient.PhraseLib.Detokenize("term.ReqTransFailed", LanguageID, TranslatedErrorTerm) & _
          " <input type=""submit"" class=""regular"" id=""" & DeployType & """ name=""" & DeployType & """ value=""" & Copient.PhraseLib.Lookup("term.yes", LanguageID) & """ onclick=""document.getElementById('deploytransreqskip').value='1';"" />" & _
          "<input type=""hidden"" id=""deploytransreqskip"" name=""deploytransreqskip"" value="""" />"
        ReturnVal = False
      End If
    End If  'multi-language enabled
    
    Return ReturnVal
    
  End Function

  '-----------------------------------------------------------------------------------------------------------------------------
  
  Function CheckForTranslationDeployError(ByRef Common As Copient.CommonInc, ByVal ROID As Long) As String
    Dim ReturnVal As String
    Common.QueryStr = "dbo.pa_ReqTranslationCheck"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
    Common.LRTsp.Parameters.Add("@ErrorTerm", SqlDbType.VarChar, 2000).Direction = ParameterDirection.Output
    Common.LRTsp.ExecuteNonQuery()
    ReturnVal = Common.LRTsp.Parameters("@ErrorTerm").Value
    Common.Close_LRTsp()
    Return ReturnVal
  End Function
  '-----------------------------------------------------------------------------------------------------------------------------
  
  Function GetAssociatedProductGroupIDs(ByVal ROID As Long, ByRef MyCommon As Copient.CommonInc) As String
    Dim PGList As String = "-1"
    Dim dt As DataTable
    Dim PGIDs(-1) As String
    Dim i As Integer
    
    MyCommon.QueryStr = "select ProductGroupID from CPE_IncentiveProductGroups with (NoLock) where RewardOptionID = " & ROID & _
                        " union " & _
                        "select DiscountedProductGroupID as ProductGroupID from CPE_Discounts as DISC with (NoLock) " & _
                        "inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID = DISC.DiscountID and DEL.DeliverableTypeID=2 " & _
                        "where(DISC.Deleted = 0 And DISC.DiscountedProductGroupID > 0 And DEL.Deleted = 0 And DEL.RewardOptionID = " & ROID & ") " & _
                        " union " & _
                        "select ExcludedProductGroupID as ProductGroupID from CPE_Discounts as DISC with (NoLock) " & _
                        "inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID = DISC.DiscountID and DEL.DeliverableTypeID=2 " & _
                        "where(DISC.Deleted = 0 And ExcludedProductGroupID > 0 And DEL.Deleted = 0 And DEL.RewardOptionID = " & ROID & ") " & _
                        "order by ProductGroupID;"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      ReDim PGIDs(dt.Rows.Count - 1)
      For i = 0 To PGIDs.GetUpperBound(0)
        PGIDs(i) = MyCommon.NZ(dt.Rows(0).Item("ProductGroupID"), "0")
      Next
      PGList = String.Join(",", PGIDs)
    End If
    
    Return PGList
  End Function
  
  '-----------------------------------------------------------------------------------------------------------------------------
  
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
  
  '-----------------------------------------------------------------------------------------------------------------------------

  Sub Send_Preference_Info(ByRef Common As Copient.CommonInc, ByVal IncentivePrefsID As Integer, ByVal ROID As Long, ByVal TierLevels As Integer)
    Dim dt, dt2 As DataTable
    Dim PreferenceID As Long = 0
    Dim ComboText As String = ""
    Dim i As Integer = 0
    Dim CellCount As Integer = 0

    
    ' find all the tiers in this preference condition
    Common.QueryStr = "select CIPT.IncentivePrefTiersID, CIPT.TierLevel, CIPT.ValueComboTypeID, CIP.PreferenceID " & _
                      "from CPE_IncentivePrefTiers as CIPT with (NoLock) " & _
                      "inner join CPE_IncentivePrefs as CIP with (NoLock) on CIP.IncentivePrefsID = CIPT.IncentivePrefsID " & _
                      "where CIPT.IncentivePrefsID=" & IncentivePrefsID & " and CIP.RewardOptionID=" & ROID & " " & _
                      "order by CIPT.TierLevel;"
    dt = Common.LRT_Select
    
    If dt.Rows.Count > 0 Then
      For Each row As DataRow In dt.Rows
        PreferenceID = Common.NZ(row.Item("PreferenceID"), 0)
        ComboText = IIf(Common.NZ(row.Item("ValueComboTypeID"), 1) = 1, "term.and", "term.or")
        ComboText = Copient.PhraseLib.Lookup(ComboText, LanguageID)
      
        CellCount += 1
        If CellCount > TierLevels Then Exit For

        Send("<br />")
        If TierLevels > 1 Then
          Send(Copient.PhraseLib.Lookup("term.tier", LanguageID) & CellCount & ": ")
        End If
        
        ' find all the tier values
        Common.QueryStr = "select IPTV.PKID, IPTV.Value, IPTV.DateOperatorTypeID, " & _
                          "  case when POT.PhraseID is null then POT.Description" & _
                          "  else Convert(nvarchar(200), PT.Phrase) end as OperatorText " & _
                          "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                          "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                          "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID & " " & _
                          "where IPTV.IncentivePrefTiersID=" & Common.NZ(row.Item("IncentivePrefTiersID"), 0)
        dt2 = Common.LRT_Select
        For i = 0 To dt2.Rows.Count - 1
          If Common.NZ(dt2.Rows(i).Item("DateOperatorTypeID"), 0) > 0 Then
            Send(Get_Date_Display_Text(Common, dt2.Rows(i).Item("PKID")))
          Else
            Send(Common.NZ(dt2.Rows(i).Item("OperatorText"), "") & " " & Get_Preference_Value(Common, PreferenceID, Common.NZ(dt2.Rows(i).Item("Value"), "")))
          End If

          If i < dt2.Rows.Count - 1 Then
            Send(" <i>" & ComboText.ToLower & "</i> ")
          End If
        Next
      Next
      
      ' account for any tiers that don't have saved information due to increasing the tiers on an existing offer
      For i = CellCount To (TierLevels - 1)
        If i > 0 AndAlso i <= (TierLevels - 1) Then Send("<br />")
        Send(Copient.PhraseLib.Lookup("term.tier", LanguageID) & (i + 1) & ": ")
        Send(Copient.PhraseLib.Lookup("term.undefined", LanguageID))
      Next
      
    End If
      
  End Sub
  
  '-----------------------------------------------------------------------------------------------------------------------------
  
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
  
  '-----------------------------------------------------------------------------------------------------------------------------
  
  Function Get_Date_Display_Text(ByRef Common As Copient.CommonInc, ByVal TierValuePKID As Integer) As String
    Dim DisplayText As String = ""
    Dim dt As DataTable
    Dim ValueModifier As String = ""
    Dim Offset As Integer
    
    Common.QueryStr = "select IPTV.Value, IPTV.ValueModifier, IPTV.ValueTypeID, POT.PhraseID as OperatorPhraseID, " & _
                      "PDOT.PhraseID as DateOpPhraseID " & _
                      "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                      "inner join CPE_PrefDateOperatorTypes as PDOT with (NoLock) on PDOT.PrefDateOperatorTypeID = IPTV.DateOperatorTypeID " & _
                      "where PKID=" & TierValuePKID & ";"
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
    End If
    
    Return DisplayText
  End Function
  
  '-----------------------------------------------------------------------------------------------------------------------------

  Sub GenerateConfirmationBox(ByVal OfferID As Long, ByVal CollisionThreshold As Integer)
    Send("<div id=""confirming"" style=""display:none;"">")
    Send("  <div id=""confirmingwrap"" style=""width:420px;"">")
    Send("    <div class=""box"" id=""confirmingbox"" style=""height:380px;"">")
    Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.ProductCollisions", LanguageID) & "</span></h2>")
    'Send("      <input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" onclick=""javascript:document.getElementById('confirming').style.display='none';"" />")
    Send("      <p>" & Copient.PhraseLib.Lookup("CPEoffer-gen.CollisionConfirm", LanguageID) & "<p>")
    Send("      <form id=""confirmationform"" name=""confirmationform"" action=""#"">")
    Send("        <p style=""text-align:center;"">")
    Send("          <input type=""hidden"" name=""OfferID"" value=""" & OfferID & """ />")
    Send("          <input type=""button"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.RunDetection", LanguageID) & """ onclick=""javascript:loadCollisions();"" />")
    Send("          <input type=""submit"" class=""regular"" id=""confirmingDeploy"" name=""deploy"" value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ />")
    Send("          <input type=""submit"" class=""regular"" id=""confirmingDeferDeploy"" name=""deferdeploy"" value=""" & Copient.PhraseLib.Lookup("CPEoffer-gen.SkipDetection", LanguageID) & """ />")
    Send("          <input type=""button"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""javascript:document.getElementById('confirming').style.display='none';"" />")
    Send("        </p>")
    Send("      </form>")
    Send("    </div>")
    Send("  </div>")
    Send("</div>")
  End Sub
  
  '-----------------------------------------------------------------------------------------------------------------------------

  Sub GenerateCollisionsBox(ByVal OfferID As Long, ByVal CollisionThreshold As Integer)
    Send("<div id=""loading"" style=""display:none;"">")
    Send("  <div id=""loadingwrap"" style=""width:420px;"">")
    Send("    <div class=""box"" id=""loadingbox"" style=""height:380px;"">")
    Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.ProductCollisions", LanguageID) & "</span></h2>")
    'Send("      <input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" onclick=""javascript:document.getElementById('loading').style.display='none';"" />")
    Send("      <div id=""collisionsContent"" style=""height:325px;overflow-y:auto;width:100%;"">")
    Send("        <p>" & Copient.PhraseLib.Lookup("CPEoffer-gen.FindingCollisions", LanguageID) & " " & Copient.PhraseLib.Lookup("CPEoffer-gen.DeployAnyway", LanguageID) & "<p>")
    Send("        <p style=""text-align:center;padding-top:80px;""><img src=""../images/loadingAnimation.gif"" alt=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.loading", LanguageID) & """ /></p>")
    Send("      </div>")
    Send("      <form id=""collisionform"" name=""collisionform"" action=""#"">")
    Send("        <p style=""text-align:center;"">")
    Send("          <input type=""hidden"" name=""OfferID"" value=""" & OfferID & """ />")
    Send("          <input type=""submit"" class=""regular"" id=""collisionDeploy"" name=""deploy"" value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """ />")
    Send("          <input type=""submit"" class=""regular"" id=""collisionDeferDeploy"" name=""deferdeploy"" value=""" & Copient.PhraseLib.Lookup("term.deploy", LanguageID) & """ />")
    Send("          <input type=""button"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onclick=""javascript:document.getElementById('loading').style.display='none';"" />")
    Send("        </p>")
    Send("      </form>")
    Send("    </div>")
    Send("  </div>")
    Send("</div>")
  End Sub
  
  '-----------------------------------------------------------------------------------------------------------------------------
  
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
  
  GenerateConfirmationBox(OfferID, CollisionThreshold)
  GenerateCollisionsBox(OfferID, CollisionThreshold)
  Send_BodyEnd()
done:
  If MyCommon.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    If MyCommon.PMRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_PrefManRT()
  End If
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
