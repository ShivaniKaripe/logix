<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: web-offer-sum.aspx 
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
  Dim t As Integer = 0
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
  Dim DaysDiff As Integer = 0
  Dim rowCount As Integer = 0
  Dim counter As Integer = 0
  Dim ErrorPhrase As String = ""
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
  Dim TenderExcludedAmt As Object
  Dim Popup As Boolean = False
  Dim TierLevels As Integer = 1
  Dim FolderNames As String = ""
  Dim EngineID As Integer = 3
  Dim EnginePhraseID As Integer = 0
  Dim EngineSubTypeID As Integer = 0
  Dim EngineSubTypePhraseID As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "web-offer-sum.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)
  
  MyCommon.QueryStr = "select I.EngineID, I.EngineSubTypeID, R.RewardOptionID, R.TierLevels " & _
                      "from CPE_Incentives as I with (NoLock) " & _
                      "left join CPE_RewardOptions as R with (NoLock) on R.IncentiveID=I.IncentiveID " & _
                      "where I.IncentiveID=" & OfferID & " and R.TouchResponse=0;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
    roid = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
    TierLevels = MyCommon.NZ(rst.Rows(0).Item("TierLevels"), 1)
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
                     "&FolderID=" & MyCommon.NZ(row.Item("FolderID"), "0") & "');"">" & MyCommon.NZ(row.Item("FolderName"), "") & "</a>"
    Next
  Else
    FolderNames = "<a href=""javascript:openPopup('/logix/folder-browse.aspx?OfferID=" & OfferID & "');"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</a>"
  End If
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("offer-new.aspx")
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
            Response.AddHeader("Location", "web-offer-gen.aspx?OfferID=" & OfferID)
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
            Response.AddHeader("Location", "web-offer-gen.aspx?OfferID=" & OfferID)
            GoTo done
        Catch ex As Exception
            If ex.Message = "error.couldnot-processoffers" Then
                infoMessage = Copient.PhraseLib.Lookup(ex.Message, LanguageID)
            Else
                infoMessage = ex.Message
            End If
        End Try
  ElseIf (Request.QueryString("deploy") <> "") Then
    IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorPhrase)
    If (IsDeployable) Then
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "web-offer-sum.aspx?OfferID=" & OfferID)
      GoTo done
    Else
      infoMessage = Copient.PhraseLib.Lookup(ErrorPhrase, LanguageID)
    End If
  ElseIf (Request.QueryString("sendoutbound") <> "") Then
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set CRMEngineUpdateLevel=CRMEngineUpdateLevel+1 where IncentiveID=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-sendtocrm", LanguageID))
  ElseIf (Request.QueryString("delete") <> "") Then
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,Deleted=1,UpdateLevel=UpdateLevel+1 where IncentiveID=" & OfferID
    MyCommon.LRT_Execute()
    ' Mark the shadow table offer as deleted as well.
    MyCommon.QueryStr = "update CPE_ST_Incentives with (RowLock) set Deleted=1, UpdateLevel = (select UpdateLevel from CPE_Incentives with (NoLock)where IncentiveID=" & OfferID & ") where IncentiveID=" & OfferID
    MyCommon.LRT_Execute()
    
    'remove the banners assigned to this offer
    If (BannersEnabled) Then
      MyCommon.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = " & OfferID
      MyCommon.LRT_Execute()
    End If
    
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
      'infoMessage = MyExport.GetErrorMsg
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
    IsDeployable = IsDeployableOffer(MyCommon, OfferID, roid, ErrorPhrase)
    If (IsDeployable) Then
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set DeployDeferred=1, UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), LastUpdate=getdate() where IncentiveID=" & OfferID
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deferdeploy", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "web-offer-sum.aspx?OfferID=" & OfferID)
      GoTo done
    Else
      infoMessage = Copient.PhraseLib.Lookup(ErrorPhrase, LanguageID)
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
                Else
                    ' Udapte Level must be increment ,on every single process . 
                    MyCommon.QueryStr = "select LastUpdateLevel from PromoEngineUpdateLevels with (NoLock) where LinkID=" & OfferID & " and EngineID= " & EngineID & " and ItemType=1;"
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                      NewUpdateLevel = MyCommon.NZ(rst.Rows(0).Item("LastUpdateLevel"), 0)
                    Else
                      NewUpdateLevel = 0
                    End If
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
            MyCommon.QueryStr = "dbo.pc_Copy_CPE_Offer"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@SourceOfferID", SqlDbType.BigInt).Value = Request.QueryString("OfferID")
            MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
            MyCommon.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = EngineSubTypeID
            MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.BigInt).Value = AdminUserID
            MyCommon.LRTsp.Parameters.Add("@NewOfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            OfferID = MyCommon.LRTsp.Parameters("@NewOfferID").Value
            MyCommon.Close_LRTsp()
    
            If (OfferID > 0) Then
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-copy", LanguageID))
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "web-offer-sum.aspx?OfferID=" & OfferID)
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
                        " PromoClassID, CRMEngineID, Priority, StartDate, EndDate, EveryDOW, EligibilityStartDate, " & _
                        " EligibilityEndDate, TestingStartDate, TestingEndDate, P1DistQtyLimit, P1DistTimeType, " & _
                        " P1DistPeriod, P2DistQtyLimit, P2DistTimeType, P2DistPeriod, P3DistQtyLimit, " & _
                        " P3DistTimeType, P3DistPeriod, EnableImpressRpt, EnableRedeemRpt, CPE.CreatedDate, CPE.LastUpdate, CPE.Deleted, ExportToEDW, " & _
                        " CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt, CRMRestricted, StatusFlag, DeployDeferred, " & _
                        " OC.OfferCategoryID, OC.Description as CategoryName, CPEIP.PhraseID, PT.Phrase as PriorityPhrase, " & _
                        " AU1.FirstName + ' ' + AU1.LastName as CreatedBy, AU2.FirstName + ' ' + AU2.LastName as LastUpdatedBy, " & _
                        " CPE.EngineSubTypeID " & _
                        " from CPE_Incentives as CPE with (NoLock) " & _
                        " left join OfferIDs as OID with (NoLock) on OID.OfferID=CPE.IncentiveID " & _
                        " left join PromoEngines as PE with (NoLock) on PE.EngineID=OID.EngineID " & _
                        " left join PromoEngineSubTypes as PEST with (NoLock) on PEST.PromoEngineID=OID.EngineID and PEST.SubTypeID=OID.EngineSubTypeID " & _
                        " left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID " & _
                        " left join CPE_IncentivePriorities as CPEIP with (NoLock) on CPEIP.PriorityID = IsNull(Priority,3) " & _
                        " left join PhraseText as PT with (NoLock) on PT.PhraseID=CPEIP.PhraseID " & _
                        " left join AdminUsers as AU1 with (NoLock) on AU1.AdminUserID = CPE.CreatedByAdminID " & _
                        " left join AdminUsers as AU2 with (NoLock) on AU2.AdminUserID = CPE.LastUpdatedByAdminID " & _
                        " where IncentiveID=" & Request.QueryString("OfferID") & " and CPE.Deleted=0 and PT.LanguageID=" & LanguageID & ";"
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
      TempDateTime = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
      'TempDateTime = TempDateTime.AddDays(1)
      If TempDateTime < Today() Then
        Expired = True
      End If
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 3)
      EnginePhraseID = MyCommon.NZ(rst.Rows(0).Item("EnginePhraseID"), 0)
      EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
      EngineSubTypePhraseID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypePhraseID"), 0)
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
    
  End If
  
  ShowCRM = (MyCommon.Fetch_SystemOption(25) <> "0")
  
  If (IsTemplate) Then
    ActiveSubTab = 27
    IsTemplateVal = "IsTemplate"
  Else
    ActiveSubTab = 27
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
  
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
%>
<script type="text/javascript" language="javascript">
  var divElems = new Array("generalbody", "periodbody", "limitsbody","deploymentbody", 
                           "locationbody", "notificationbody", "conditionbody", "rewardbody", "validationbody");
  var divVals  = new Array(1, 2, 4, 8, 16, 32, 64, 128, 256);
  var divImages = new Array("imgGeneral", "imgPeriod", "imgLimits", "imgDeployment",  
                            "imgLocations", "imgNotifications", "imgConditions", "imgRewards", "imgValidation");
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

<form id="mainform" name="mainform" action="#">
  <input type="hidden" name="OfferID" id="OfferID" value="<% Sendb(offerid) %>" />
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
    ShowActionButton = (Logix.UserRoles.CreateTemplate And Not IsTemplate) OrElse (Logix.UserRoles.CRUDOfferFromTemplate And IsTemplate) _
                        OrElse ((Logix.UserRoles.DeleteOffer AndAlso Not IsTemplate) Or (Logix.UserRoles.DeleteTemplates AndAlso IsTemplate)) _
                        OrElse (Logix.UserRoles.SendOffersToCRM And Not IsTemplate And ShowCRM) OrElse (MyCommon.Fetch_SystemOption(73) <> "")
    If (Not LinksDisabled OrElse IsTemplate) Then
      If (ShowActionButton) Then
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
            Send_DeferDeploy()
          End If
        End If
        If (((Logix.UserRoles.DeployNonTemplateOffers AndAlso Not FromTemplate) OrElse (Logix.UserRoles.DeployTemplateOffers AndAlso FromTemplate)) And Not IsTemplate) Then
          If (Expired And MyCommon.Fetch_CPE_SystemOption(80) = "1") Then
          Else
            Send_Deploy()
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
    %>
    <div id="column1">
      <div class="box" id="general">
        <%  
          If (LinksDisabled) Then
            Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</span></h2>")
          Else
            Send("<h2 style=""float:left;""><span><a href=""web-offer-gen.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.general", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("generalbody", "imgGeneral", Copient.PhraseLib.Lookup("term.general", LanguageID), True)
          Send("<div id=""generalbody"">")
          Send("  <table summary=""" & Copient.PhraseLib.Lookup("term.general", LanguageID) & """ cellpadding=""0"" cellspacing=""0"">")
          Send("    <tr>")
          Send("      <td style=""width:130px;""><b>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":</b></td>")
          Send("      <td>" & MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0) & "</td>")
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
          <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>" cellpadding="0" cellspacing="0">
            <tr>
              <td style="width: 85px;">
                <b><% Sendb(Copient.PhraseLib.Lookup("term.production", LanguageID))%>:</b>
              </td>
              <td>
                <%
                  LongDate = MyCommon.NZ(rst.Rows(0).Item("StartDate"), "1/1/1900")
                  If LongDate > "1/1/1900" Then Sendb(Logix.ToShortDateString(LongDate, MyCommon)) Else Sendb("?")
                  Sendb(" - ")
                  LongDate = MyCommon.NZ(rst.Rows(0).Item("EndDate"), "1/1/1900")
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
          </table>
        </div>
      </div>
      
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
            LongDate = MyCommon.NZ(rst.Rows(0).Item("CPEOARptDate"), "1/1/1900")
            If LongDate > "1/1/1900" Then
              DaysDiff = DateDiff(DateInterval.Day, rst.Rows(0).Item("CPEOARptDate"), DateTime.Today)
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
            If LongDate > "1/1/1900" Then
              DaysDiff = DateDiff(DateInterval.Day, rst.Rows(0).Item("CPEOADeploySuccessDate"), DateTime.Today)
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
            <%Sendb(Copient.PhraseLib.Lookup("term.laststatus", LanguageID))%>:
          </h3>
          <%Send("&nbsp;" & MyCommon.SplitNonSpacedString(MyCommon.NZ(rst.Rows(0).Item("CPEOADeployRpt"), ""), 25))%>
          <br />
          <hr class="hidden" />
        </div>
      </div>
      <% End If%>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="offernotifications">
        <%
          If (LinksDisabled) Then
            Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.notifications", LanguageID) & "</span></h2>")
          Else
            Send("<h2 style=""float:left;""><span><a href=""web-offer-not.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.notifications", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("notificationbody", "imgNotifications", "Notifications", True)
        %>
        <div id="notificationbody">
          <%
            ' Printed message notifications
            MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID " & _
                                "from CPE_Deliverables as D with (NoLock) " & _
                                "inner join PrintedMessages PM with (NoLock) on D.OutputID=PM.MessageID " & _
                                "inner join PrintedMessageTiers as PMT with (NoLock) on PM.MessageID=PMT.MessageID " & _
                                "where D.Deleted=0 and D.RewardOptionPhase=1 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel=1;"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Dim Details As StringBuilder
              Send("<h3>" & Copient.PhraseLib.Lookup("term.printedmessages", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Details = New StringBuilder(200)
                Details.Append(ReplaceTags(MyCommon.NZ(row.Item("BodyText"), "")))
                If (Details.ToString().Length > 80) Then
                  Details = Details.Remove(77, (Details.Length - 77))
                  Details.Append("...")
                End If
                Details.Replace(vbCrLf, "<br />")
                Send("<li>""" & MyCommon.SplitNonSpacedString(Details.ToString, 25) & """</li>")
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
              Send("<h3>" & Copient.PhraseLib.Lookup("term.cashiermessages", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Send("<li>""" & MyCommon.NZ(row.Item("Line1"), "") & "<br />" & MyCommon.NZ(row.Item("Line2"), "") & """</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
            End If
            
            ' Graphics Notifications
            MyCommon.QueryStr = "select OSA.Name as GraphicName, OSA.OnScreenAdID as AdID, OSA.Width as Width, OSA.Height as Height, OSA.GraphicSize as Size, OSA.ImageType as Type, " & _
                                "D.DeliverableID, D.ScreenCellID as CellID, OSA.Width, OSA.Height, OSA.ImageType, SC.Name as CellName, SL.Name as LayoutName " & _
                                "from OnScreenAds as OSA with (NoLock) Inner Join CPE_Deliverables as D with (NoLock) on OSA.OnScreenAdID=D.OutputID and D.RewardOptionID=" & roid & " " & _
                                "and OSA.Deleted=0 and D.Deleted=0 and D.DeliverableTypeID=1 and D.RewardOptionPhase=1 " & _
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
            
            'Accumulation Notifications
            MyCommon.QueryStr = "select PM.MessageID, PM.MessageTypeID, PMT.BodyText, D.DeliverableID " & _
                                "from CPE_deliverables D with (NoLock) inner join PrintedMessages PM with (NoLock) on D.OutputID = PM.MessageID " & _
                                "inner join PrintedMessageTiers PMT with (NoLock) on PM.MessageID = PMT.MessageID " & _
                                "where D.Deleted = 0 and D.RewardOptionPhase=2 and D.RewardOptionID=" & roid & " and D.DeliverableTypeID=4 and PMT.TierLevel = 1;"
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              counter = counter + 1
              Dim Details As StringBuilder
              Send("<h3>" & Copient.PhraseLib.Lookup("term.accumulationmessage", LanguageID) & "</h3>")
              Send("<ul class=""condensed"">")
              For Each row In rst.Rows
                Details = New StringBuilder(200)
                Details.Append(ReplaceTags(MyCommon.NZ(row.Item("BodyText"), "")))
                If (Details.ToString().Length > 80) Then
                  Details = Details.Remove(77, (Details.Length - 77))
                  Details.Append("...")
                End If
                Details.Replace(vbCrLf, "<br />")
                Send("<li>""" & MyCommon.SplitNonSpacedString(Details.ToString, 25) & """</li>")
              Next
              Send("</ul>")
            End If
            
            If (counter = 0) Then
              Send("<h3>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</h3>")
            End If
            counter = 0
          %>
          <hr class="hidden" />
        </div>
      </div>
      <div class="box" id="offerconditions">
        <%  
          If (LinksDisabled) Then
            Send("<h2 style=""float:left;""><span>" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & "</span></h2>")
          Else
            Send("<h2 style=""float:left;""><span><a href=""web-offer-con.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.conditions", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("conditionbody", "imgConditions", Copient.PhraseLib.Lookup("term.conditions", LanguageID), True)
        %>
        <div id="conditionbody">
          <%
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
              For Each row In rst.Rows
                Sendb("<li>")
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
                i = i + 1
                MyCommon.QueryStr = "select count(*) as GCount from GroupMembership with (NoLock) where CustomerGroupID = " & MyCommon.NZ(row.Item("CustomerGroupID"), "-1") & " And Deleted = 0"
                rst2 = MyCommon.LXS_Select()
                For Each row2 In rst2.Rows
                  If Not MyCommon.NZ(row.Item("AnyCardholder"), False) AndAlso Not MyCommon.NZ(row.Item("AnyCustomer"), False) AndAlso Not MyCommon.NZ(row.Item("NewCardholders"), False) AndAlso Not MyCommon.NZ(row.Item("AnyCAMCardholder"), False) Then
                    Sendb(" (" & row2.Item("GCount") & ") ")
                  ElseIf (IsDBNull(row.Item("CustomerGroupID")) AndAlso MyCommon.NZ(row.Item("RequiredFromTemplate"), False)) Then
                    Sendb("<span class=""red"">* ")
                    Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))
                    Sendb(" " & Copient.PhraseLib.Lookup("term.by", LanguageID))
                    Sendb(" " & Copient.PhraseLib.Lookup("term.template", LanguageID))
                    Sendb("</span>")
                  Else
                    Sendb(" ")
                  End If
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
                    End If
                  Next
                Next
              End If
              Send("</li>")
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
                  Sendb("<span class=""red"">* ")
                  Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))
                  Sendb(" " & Copient.PhraseLib.Lookup("term.by", LanguageID))
                  Sendb(" " & Copient.PhraseLib.Lookup("term.template", LanguageID))
                  Sendb("</span>")
                Else
                  Sendb(MyCommon.SplitNonSpacedString(MyCommon.NZ(row.Item("ProgramName"), ""), 25))
                End If
                Send("</li>")
              Next
              Send("</ul>")
              Send("<br class=""half"" />")
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
            Send("<h2 style=""float:left;""><span><a href=""web-offer-rew.aspx?OfferID=" & OfferID & """>" & Copient.PhraseLib.Lookup("term.rewards", LanguageID) & "</a></span></h2>")
          End If
          Send_BoxResizer("rewardbody", "imgRewards", Copient.PhraseLib.Lookup("term.rewards", LanguageID), True)
        %>
        <div id="rewardbody">
          <%
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

<script type="text/javascript">
  collapseBoxes();
  setComponentsColor('<% Sendb(ComponentColor) %>');
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
  Function IsDeployableOffer(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByVal ROID As Integer, ByRef ErrorPhrase As String) As Boolean
    Dim bDeployable As Boolean = False
    
    ErrorPhrase = ""
    bDeployable = MeetsDeploymentReqs(MyCommon, OfferID)
    
    If Not bDeployable Then
      ErrorPhrase = "web-offer.deployalert"
    End If
    
    Return bDeployable
  End Function
  
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
  Logix = Nothing
  MyCommon = Nothing
%>
