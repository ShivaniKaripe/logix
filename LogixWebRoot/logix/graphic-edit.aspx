<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<%
  ' *****************************************************************************
  ' * FILENAME: graphic-edit.aspx 
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
  Dim row As System.Data.DataRow
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dstPoints As System.Data.DataTable
  Dim dstAds As System.Data.DataTable
  Dim dstName As System.Data.DataTable
  Dim dstAssociated As System.Data.DataTable
  Dim dst As DataTable
  Dim pgPromoVarID As String
  Dim pgTotalPoints As Long
  Dim pgCreated As String
  Dim pgUpdated As String
  Dim ProgramName As String
  Dim PromoVarID As String
  Dim ProgramID As Long
  Dim rowCount As Integer
  Dim assocName As String
  Dim assocID As String
  Dim l_adID As Int32 = 0
  Dim longDate As New DateTime
  Dim longDateString As String
  Dim adName As String = ""
  Dim adDescription As String = ""
  Dim adWidth As Integer = 0
  Dim adHeight As Integer = 0
  Dim adDimension As String = ""
  Dim adLastUpload As String = ""
  Dim adDisplayDuration As String = ""
  Dim adGraphicSize As String = ""
  Dim adImageType As String = ""
  Dim adStoreResponse As String = ""
  Dim adResponseChecked As String = ""
  Dim adBackgroundImg As String = ""
  Dim adClientFileName As String = ""
  Dim inUse As Boolean = False
  Dim AllowDelete As Boolean = True
  Dim Touchable As Boolean = False
  Dim ModifyTouch As Boolean = True
  Dim Redeployed As Boolean = False
  Dim File As HttpPostedFile
  Dim tnFile As HttpPostedFile
  Dim ImageType As String = "-1" 'unsupported image type
  Dim GraphicFileName As String = ""
  Dim ClientFileName As String = ""
  Dim strQuery As String = ""
  Dim TouchpointsEditable As Boolean = True
  Dim MD5sum As String = "NOMD5"
  Dim sizeOfData As Integer
  'Dim graphicBuf() As Byte = Nothing
  Dim FileData() As Byte
  Dim ReadCount As Integer
  Dim ShowActionButton As Boolean = False
  Dim OfferCtr As Integer = 0
  Dim DeletingOffer As Boolean = False
  Dim IE6ScrollFix As String = ""
  Dim i As Integer
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "graphic-edit.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Dim DEFAULT_GRAPHIC_PATH As String = MyCommon.Get_Install_Path
  
  If (Request.QueryString("infoMessage") <> "") Then
    infoMessage = Request.QueryString("infoMessage")
  End If
  
  If Request.QueryString("LargeFile") = "true" Then
    infoMessage = Copient.PhraseLib.Lookup("error.UploadTooLarge", LanguageID)
  End If
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("graphic-edit.aspx")
  End If
  
  If (Request.QueryString("save") <> "" OrElse Request.QueryString("unsavedChanges") = "true") Then
    ' somebody clicked save
    Dim Name As String = ""
    Dim Description As String
    Dim DisplayDuration As Integer
    Dim NumRecs As Long
    Dim StoreResponse As String
    l_adID = MyCommon.Extract_Decimal(Request.QueryString("OnScreenAdID"), MyCommon.GetAdminUser.Culture)
    ' check for an entered name; do not allow creation if no name specified
    Name = MyCommon.Parse_Quotes(Left(Logix.TrimAll(Request.QueryString("name")), 255))
    Description = MyCommon.Parse_Quotes(Trim(Request.QueryString("desc")))
    DisplayDuration = MyCommon.Extract_Decimal(Trim(Request.QueryString("duration")), MyCommon.GetAdminUser.Culture)
    If DisplayDuration < 2 Then DisplayDuration = 2
    If DisplayDuration > 600 Then DisplayDuration = 600
    If Request.QueryString("storeresponse") = "on" Then
      StoreResponse = "1"
    Else
      StoreResponse = "0"
    End If
    If (Name = "") Then
      infoMessage = Copient.PhraseLib.Lookup("graphics.noname", LanguageID)
    Else
      'see if the Name they specified is already in use
      NumRecs = 0
      MyCommon.QueryStr = "SELECT count(*) AS NumRecs FROM OnScreenAds with (NoLock) WHERE Name='" & Name & "' AND not(OnScreenAdID=" & l_adID & ") AND Deleted=0;"
      dstName = MyCommon.LRT_Select
      If Not (dstName.Rows.Count = 0) Then
        NumRecs = MyCommon.NZ(dstName.Rows(0).Item("NumRecs"), 0)
      End If
      If NumRecs = 0 Then
        MyCommon.QueryStr = "update OnScreenAds with (RowLock) set Name=N'" & Name & "', Description=N'" & Description & "', DisplayDuration=" & DisplayDuration & _
                            ", CPEStatusFlag = CASE WHEN CPEStatusFlag=2 THEN 2 ELSE 1 END " & _
                            ", UEStatusFlag = CASE WHEN UEStatusFlag=2 THEN 2 ELSE 1 END " & _
                            ", StoreResponse=" & StoreResponse & ", LastUpdate=getdate(), UpdateLevel=UpdateLevel+1 where OnScreenAdID=" & l_adID & ";"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "update TouchAreas with (RowLock) set LastUpdate=getdate() where OnScreenAdID=" & l_adID & " and Deleted=0;"
        MyCommon.LRT_Execute()
        If l_adID <> 0 Then
          MyCommon.Activity_Log(12, l_adID, AdminUserID, Copient.PhraseLib.Lookup("history.graphic-edit", LanguageID))
        End If
      Else
        adName = Name
        adDescription = Description
        adDisplayDuration = DisplayDuration
        adStoreResponse = StoreResponse
        infoMessage = Copient.PhraseLib.Lookup("graphics.duplicatename", LanguageID)
      End If
    End If
  End If
  
  ' any GET parms inbound?
  If (Request.QueryString("Delete") <> "") Then
    DeletingOffer = True
    l_adID = MyCommon.Extract_Decimal(Request.QueryString("OnScreenAdID"), MyCommon.GetAdminUser.Culture)
    MyCommon.QueryStr = "select I.IncentiveID as OfferCt from CPE_Deliverables as D with (NoLock) " & _
                        "Inner Join CPE_RewardOptions as RO with (NoLock) on D.RewardOptionID=RO.RewardOptionID and D.OutputID=" & l_adID & _
                        "  and DeliverableTypeID=1 and RO.Deleted=0 and D.Deleted=0 " & _
                        "Inner Join CPE_Incentives as I with (NoLock) on RO.IncentiveID=I.IncentiveID and I.Deleted=0;"
    dst = MyCommon.LRT_Select
    If (dst.Rows.Count = 0) Then
      ' check that there are no deployed offers that use this graphic
      MyCommon.QueryStr = "dbo.pa_AssociatedOffers_ST"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@LinkType", SqlDbType.Int).Value = 5
      MyCommon.LRTsp.Parameters.Add("@LinkID", SqlDbType.Int).Value = l_adID
      dst = MyCommon.LRTsp_select
      MyCommon.Close_LRTsp()
      If (dst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("term.inusedeployment", LanguageID) & " : ("
        For OfferCtr = 0 To dst.Rows.Count - 1
          infoMessage &= MyCommon.NZ(dst.Rows(OfferCtr).Item("IncentiveID"), "")
        Next
        infoMessage &= ")"
      Else
        MyCommon.QueryStr = "update OnScreenAds with (RowLock) set Deleted=1, CPEStatusFlag=1, UEStatusFlag=1, UpdateLevel=UpdateLevel+1 where OnScreenAdID=" & l_adID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(12, l_adID, AdminUserID, Copient.PhraseLib.Lookup("history.graphic-delete", LanguageID))
        RemoveGraphics(l_adID)
        'remove any touch areas from this graphic
        MyCommon.QueryStr = "update TouchAreas with (RowLock) set Deleted=1 where OnScreenAdID=" & l_adID
        MyCommon.LRT_Execute()
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "graphic-list.aspx")
        GoTo done
      End If
    Else
      infoMessage = Copient.PhraseLib.Lookup("graphics.inuse", LanguageID)
    End If
  ElseIf (Request.QueryString("OnScreenAdID") = "new") Then
    ' add a record
    Dim Name As String = ""
    Dim Description As String
    Dim NumRecs As Long
    Dim AdID As Long
    ' check for an entered name; do not allow creation if no name specified
    Name = Left(Logix.TrimAll(Request.QueryString("name")), 255)
    Description = Trim(Request.QueryString("desc"))
    If (Name = "") Then
      infoMessage = Copient.PhraseLib.Lookup("graphics.noname", LanguageID)
    Else
      'see if the Name they specified is already in use
      NumRecs = 0
      MyCommon.QueryStr = "select count(*) as NumRecs from OnScreenAds with (NoLock) where Deleted=0 and name=@Name;"
      MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = Name
      dstName = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
      If Not (dstName.Rows.Count = 0) Then
        NumRecs = MyCommon.NZ(dstName.Rows(0).Item("NumRecs"), 0)
      End If
      If NumRecs = 0 Then
        MyCommon.QueryStr = "insert into OnScreenAds with (RowLock) (Name, Description, DisplayDuration, StoreResponse, Deleted, LastUpdate, CPEStatusFlag, UEStatusFlag) values (@Name,@Description, 2, 0, 0 , getdate(), 1, 1)"
        MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = Name
        MyCommon.DBParameters.Add("@Description", SqlDbType.NVarChar).Value = Description
        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
        AdID = 0
        MyCommon.QueryStr = "select OnScreenAdID from OnScreenAds with (NoLock) where Deleted=0 and name=@Name;"
        MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar).Value = Name
        dstName = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If Not (dstName.Rows.Count = 0) Then
          AdID = MyCommon.NZ(dstName.Rows(0).Item("OnScreenAdID"), 0)
        End If
        If AdID = 0 Then
          infoMessage = Copient.PhraseLib.Lookup("graphics.error", LanguageID)
        Else
          MyCommon.Activity_Log(12, AdID, AdminUserID, Copient.PhraseLib.Lookup("history.graphic-create", LanguageID))
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "graphic-edit.aspx?OnScreenAdID=" & AdID)
          GoTo done
        End If
      Else
        infoMessage = Copient.PhraseLib.Lookup("graphics.duplicatename", LanguageID)
      End If
    End If
  ElseIf (Request.QueryString("add") <> "") Then
    ' add a touchpoint
    Dim TouchName As String
    Dim X As Integer
    Dim Y As Integer
    Dim Width As Integer
    Dim Height As Integer
    l_adID = MyCommon.Extract_Decimal(Request.QueryString("OnScreenAdID"), MyCommon.GetAdminUser.Culture)
    TouchName = Left(Logix.TrimAll(MyCommon.Strip_Quotes(Request.QueryString("txtAreaName"))), 200)
    X = MyCommon.Extract_Decimal(Request.QueryString("txtXPos"), MyCommon.GetAdminUser.Culture)
    Y = MyCommon.Extract_Decimal(Request.QueryString("txtYPos"), MyCommon.GetAdminUser.Culture)
    Width = MyCommon.Extract_Decimal(Request.QueryString("txtWidth"), MyCommon.GetAdminUser.Culture)
    Height = MyCommon.Extract_Decimal(Request.QueryString("txtHeight"), MyCommon.GetAdminUser.Culture)
    If TouchName = "" Then
    Else
      MyCommon.QueryStr = "insert into TouchAreas with (RowLock) (OnScreenAdID, Name, X, Y, Width, Height, Deleted, LastUpdate) values (" & l_adID & ", N'" & TouchName & "', " & X & ", " & Y & ", " & Width & ", " & Height & ", 0, getdate());"
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update OnScreenAds with (RowLock) set LastUpdate=getdate(), UpdateLevel=UpdateLevel+1, CPEStatusFlag=CASE WHEN CPEStatusFlag=2 THEN 2 ELSE 1 END, UEStatusFlag=CASE WHEN UEStatusFlag=2 THEN 2 ELSE 1 END where OnScreenAdID=" & l_adID
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(12, l_adID, AdminUserID, Copient.PhraseLib.Lookup("history.graphic-addtouchpoint", LanguageID))
    End If
  ElseIf (Request.QueryString("deleteTP") <> "") Then
    Dim AreaID As Long
    Dim AdID As Long
    Dim PrevTouchable As Boolean
    AreaID = MyCommon.Extract_Decimal(Request.QueryString("areaId"), MyCommon.GetAdminUser.Culture)
    AdID = MyCommon.Extract_Decimal(Request.QueryString("adId"), MyCommon.GetAdminUser.Culture)
    Touchable = False
    PrevTouchable = False
    MyCommon.QueryStr = "select count(*) as NumRecs from TouchAreas with (NoLock) where OnScreenAdID=" & AdID & " and Deleted=0;"
    dstPoints = MyCommon.LRT_Select()
    
    If (dstPoints.Rows.Count > 0) Then
      If MyCommon.NZ(dstPoints.Rows(0).Item("NumRecs"), 0) > 0 Then PrevTouchable = True
    End If
    MyCommon.QueryStr = "update TouchAreas with (RowLock) set Deleted=1, LastUpdate=getdate() where AreaID=" & AreaID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "select count(*) as NumRecs from TouchAreas with (NoLock) where OnScreenAdID=" & AdID & " and Deleted=0;"
    dstPoints = MyCommon.LRT_Select()
    If (dstPoints.Rows.Count > 0) Then
      If MyCommon.NZ(dstPoints.Rows(0).Item("NumRecs"), 0) > 0 Then Touchable = True
    End If
    If Touchable = False And PrevTouchable = True Then
      MyCommon.QueryStr = "update OnScreenAds with (RowLock) set StoreResponse=0, LastUpdate=getdate(), UpdateLevel=UpdateLevel+1 where OnScreenAdID=" & AdID & " and Deleted=0;"
      MyCommon.LRT_Execute()
    End If
    MyCommon.QueryStr = "update onscreenads with (RowLock) set LastUpdate=getdate(), UpdateLevel=UpdateLevel+1, " & _
                        "  CPEStatusFlag = CASE WHEN CPEStatusFlag=2 THEN 2 ELSE 1 END, UEStatusFlag = CASE WHEN UEStatusFlag=2 THEN 2 ELSE 1 END where OnScreenAdID=" & AdID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(12, AdID, AdminUserID, Copient.PhraseLib.Lookup("history.graphics-deletetouchpoint", LanguageID))
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "graphic-edit.aspx?OnScreenAdID=" & AdID)
    GoTo done
  ElseIf (Request.QueryString("mode") = "accept") Then
    l_adID = MyCommon.Extract_Decimal(Request.QueryString("adId"), MyCommon.GetAdminUser.Culture)
    If Request.Files.Count >= 1 Then
      Dim GraphicPath As String = ""
      Dim img As System.Drawing.Image
      Dim format As System.Drawing.Imaging.ImageFormat
      Dim GraphicSize As Integer, width As Integer, height As Integer
      Dim imgExt As String = "jpg" ' default to jpg for extension
      Dim imgStream As MemoryStream
      File = Request.Files.Get(0)
      If InStr(UCase(File.ContentType), "IMAGE", CompareMethod.Text) > 0 Then
        GraphicPath = MyCommon.Fetch_SystemOption(47)
        If (GraphicPath = "") Then
          GraphicPath = DEFAULT_GRAPHIC_PATH
        End If
        If Not (Right(GraphicPath, 1) = "\") Then
          GraphicPath = GraphicPath & "\"
        End If
        img = System.Drawing.Image.FromStream(File.InputStream)
        ClientFileName = File.FileName
        GraphicSize = File.ContentLength
        width = img.PhysicalDimension.Width
        height = img.PhysicalDimension.Height
        format = img.RawFormat
        If (format.Equals(System.Drawing.Imaging.ImageFormat.Jpeg)) Then
          ImageType = "1"
          imgExt = "jpg"
        ElseIf (format.Equals(System.Drawing.Imaging.ImageFormat.Gif)) Then
          ImageType = "2"
          imgExt = "gif"
        Else
          infoMessage = File.FileName & ". " & Copient.PhraseLib.Lookup("graphics.unsupportedformat", LanguageID)
        End If
        If (ImageType <> "-1") Then
          GraphicFileName = GraphicPath & l_adID & "img." & imgExt
          File.SaveAs(GraphicFileName)
          MakeThumbnail(GraphicPath & l_adID & "img_tn.", img, imgExt)
          File.InputStream.Seek(0, SeekOrigin.Begin)
          ' Calculate the MD5 Sum
          ReDim FileData(Request.Files(0).ContentLength - 1)
          ReadCount = Request.Files(0).InputStream.Read(FileData, 0, Request.Files(0).ContentLength)
          'uncomment to view raw data
          'Send(Encoding.Default.GetString(FileData))
          MD5sum = MyCommon.MD5(Encoding.Default.GetString(FileData))
          FileData = Nothing
          MyCommon.QueryStr = "Update OnScreenAds with (RowLock) set GraphicSize=" & GraphicSize & ", ClientFileName='" & MyCommon.Parse_Quotes(ClientFileName) & _
                              "', Width=" & width & ", Height=" & height & ", ImageType=" & ImageType & ", LastUpload=getdate(), CPEStatusFlag=2, UEStatusFlag=2, " & _
                              "MD5sum='" & MD5sum & "', UpdateLevel=UpdateLevel+1 where OnScreenAdID=" & l_adID & ";"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(12, l_adID, AdminUserID, Copient.PhraseLib.Lookup("history.graphic-upload", LanguageID))
        End If
      Else
        infoMessage = Copient.PhraseLib.Lookup("graphics.nothingselected", LanguageID)
      End If
      MyCommon.QueryStr = "select Name, Description, DisplayDuration, StoreResponse from OnscreenAds with (NoLock) where OnScreenAdID=" & l_adID & ";"
      dst = MyCommon.LRT_Select
      If dst.Rows.Count > 0 Then
        adName = MyCommon.NZ(dst.Rows(0).Item("Name"), "")
        adDescription = MyCommon.NZ(dst.Rows(0).Item("Description"), "")
        adDisplayDuration = MyCommon.NZ(dst.Rows(0).Item("DisplayDuration"), "")
        adStoreResponse = MyCommon.NZ(dst.Rows(0).Item("StoreResponse"), "")
      End If
      strQuery = "graphic-edit.aspx?OnScreenAdID=" & l_adID & "&name=" & adName & "&desc=" & adDescription & "&displayduration=" & adDisplayDuration & "&storeresponse=" & adStoreResponse
      If (infoMessage <> "") Then
        strQuery = strQuery & "&infoMessage=" & infoMessage
      End If
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", strQuery)
      GoTo done
    End If
  ElseIf (Request.QueryString("redeploy") <> "") Then
    l_adID = MyCommon.Extract_Decimal(Request.QueryString("adId"), MyCommon.GetAdminUser.Culture)
    MyCommon.QueryStr = "update OnScreenAds with (RowLock) set CPEStatusFlag=2, UEStatusFlag=2, UpdateLevel=UpdateLevel+1 where OnScreenAdID=" & l_adID
    MyCommon.LRT_Execute()
    infoMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
    Redeployed = True
  ElseIf (Request.QueryString("OnScreenAdID") <> "") Then
    ' simple edit/search mode
    l_adID = MyCommon.NZ(Request.QueryString("OnScreenAdID"), "0")
  Else
    ' no ad id passed ... what now ?
    l_adID = 0
  End If
  
  ' grab this graphic
  MyCommon.QueryStr = "select OSA.Name, OSA.Description, OSA.DisplayDuration, OSA.GraphicSize, OSA.Width, OSA.Height, OSA.ImageType, " & _
                      "OSA.StoreResponse, OSA.LastUpload, SC.BackgroundImg, OSA.ClientFileName from OnScreenAds as OSA with (NoLock) " & _
                      "left join ScreenCells as SC with (NoLock) on SC.BackgroundImg=OSA.OnScreenAdID and SC.Deleted=0 " & _
                      "where OSA.Deleted=0 and OSA.OnScreenAdID=" & l_adID & ";"
  dstAds = MyCommon.LRT_Select
  If (dstAds.Rows.Count > 0) Then
    'Assign values to the page variable
    If (infoMessage = "" OrElse Redeployed OrElse DeletingOffer) Then
      adName = MyCommon.NZ(dstAds.Rows(0).Item("Name"), "")
      adDescription = MyCommon.NZ(dstAds.Rows(0).Item("Description"), "")
      adDisplayDuration = MyCommon.NZ(dstAds.Rows(0).Item("DisplayDuration"), 2)
      adStoreResponse = MyCommon.NZ(dstAds.Rows(0).Item("StoreResponse"), "")
      If (adStoreResponse = True) Then
        adResponseChecked = " checked=""checked"""
      End If
    End If
    adGraphicSize = MyCommon.NZ(dstAds.Rows(0).Item("GraphicSize"), 0)
    adWidth = MyCommon.NZ(dstAds.Rows(0).Item("Width"), 0)
    adHeight = MyCommon.NZ(dstAds.Rows(0).Item("Height"), 0)
    adDimension = adWidth & " x" & adHeight
    adImageType = MyCommon.NZ(dstAds.Rows(0).Item("ImageType"), "")
    If (Not IsDBNull(dstAds.Rows(0).Item("LastUpload"))) Then
      adLastUpload = Logix.ToLongDateTimeString(dstAds.Rows(0).Item("LastUpload"), MyCommon)
    Else
      adLastUpload = Copient.PhraseLib.Lookup("term.never", LanguageID)
    End If
    adBackgroundImg = MyCommon.NZ(dstAds.Rows(0).Item("BackgroundImg"), "")
    adClientFileName = MyCommon.NZ(dstAds.Rows(0).Item("ClientFileName"), "")
    'Check if Touchareas are already set up
    MyCommon.QueryStr = "select count(*) as NumRecs from TouchAreas with (NoLock) where OnScreenAdID=" & l_adID & " and Deleted=0;"
    dstPoints = MyCommon.LRT_Select
    If (dstPoints.Rows.Count > 0) Then
      If (MyCommon.NZ(dstPoints.Rows(0).Item("NumRecs"), 0) > 0) Then
        Touchable = True
      End If
    End If
    'Check if Modify Touch is allowed
    MyCommon.QueryStr = "select count(*) as NumRecs " & _
                        "from CPE_Deliverables as D with (NoLock) Inner Join CPE_RewardOptions as RO with (NoLock) on D.RewardOptionID=RO.RewardOptionID and D.OutputID=" & l_adID & " and DeliverableTypeID=1 and RO.Deleted=0 and D.Deleted=0 " & _
                        "Inner Join CPE_Incentives as I with (NoLock) on RO.IncentiveID=I.IncentiveID and I.Deleted=0 " & _
                        "where I.EndDate>=getdate();"
    dstPoints = MyCommon.LRT_Select
    If (dstPoints.Rows.Count > 0) Then
      If (MyCommon.NZ(dstPoints.Rows(0).Item("NumRecs"), 0) > 0) Then
        ModifyTouch = False
      End If
    End If
  ElseIf (Request.QueryString("new") <> "New") And (l_adID > 0) Then
    MyCommon.QueryStr = "select Name from OnScreenAds with (NoLock) where OnScreenAdID=" & l_adID & ";"
    dstAds = MyCommon.LRT_Select
    If dstAds.Rows.Count > 0 Then
      adName = MyCommon.NZ(dstAds.Rows(0).Item("Name"), "")
    End If
    Send_HeadBegin("term.graphic", , l_adID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Logos()
    Send_Tabs(Logix, 6)
    Send_Subtabs(Logix, 61, 3, , l_adID)
    Send("")
    Send("<div id=""intro"">")
    Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.graphic", LanguageID) & " #" & l_adID & IIf(adName <> "", ": " & adName, "") & "</h1>")
    Send("</div>")
    Send("<div id=""main"">")
    Send("  <div id=""infobar"" class=""red-background"">")
    Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
    Send("  </div>")
    Send("</div>")
    Send_BodyEnd()
    GoTo done
  Else
    pgPromoVarID = 0
    l_adID = 0
    adDescription = ""
    pgCreated = ""
    pgUpdated = ""
    adName = Copient.PhraseLib.Lookup("term.newgraphic", LanguageID)
  End If
  
  ' Check if the touchpoints are editable. They are only editable if this graphic is not joined to an offer
  MyCommon.QueryStr = "select count(RO.IncentiveID) as GraphicIncentives from CPE_Deliverables D with (NoLock) " & _
                      "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = D.RewardOptionID " & _
                      "where DeliverableTypeID=1 and OutputID=" & l_adID & " and D.Deleted=0 and RO.Deleted=0;"
  dst = MyCommon.LRT_Select
  If (dst.Rows.Count > 0) Then
    TouchpointsEditable = (dst.Rows(0).Item("GraphicIncentives") = 0)
  End If
  
  Send_HeadBegin("term.graphic", , l_adID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 6)
  Send_Subtabs(Logix, 61, 3, , l_adID)
  
  If (Logix.UserRoles.AccessGraphics = False) Then
    Send_Denied(1, "perm.graphics-access")
    Send_BodyEnd()
    GoTo done
  End If
%>
<script type="text/javascript" language="javascript">
function clearNewTP() {
  var frm = document.mainform;
  
  if (frm != null) {
    frm.txtAreaName.value = ""
    frm.txtXPos.value = "";
    frm.txtYPos.value = "";
    frm.txtWidth.value = "";
    frm.txtHeight.value = "";
  }
}

function deleteTouchPt(areaId, adId) {
  var elemAreaId = document.getElementById("areaId");
  var elemDelTp = document.getElementById("deleteTP");
  
  //document.location = "graphic-edit.aspx?deleteTP=1&areaId=" + areaId + "&adId=" + adId;
  if (elemAreaId != null) { elemAreaId.value = areaId; }
  if (elemDelTp != null)  { elemDelTp.value  = "1"; }
  
  checkForChanges();
  document.mainform.submit();
}

function showUploadPage(adId, adName) {
  document.location = "graphic-upload.aspx?adId=" + adId + "&adName=" + adName;
}

function showImageWindow(imgPath) {
  popW = 700;
  popH = 540;
  siteWindow = window.open("test.html","Popup", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
  siteWindow.document.writeln('<!DOCTYPE html ')
  siteWindow.document.writeln('     PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\"')
  siteWindow.document.writeln('     \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">')
  siteWindow.document.writeln('<html xmlns=\"http://www.w3.org/1999/xhtml\">')
  siteWindow.document.writeln('<head>')
  siteWindow.document.writeln('<title>' + '<% Sendb(Copient.PhraseLib.Lookup("term.logix", LanguageID))%>' + ' > ' + '<% Sendb(Copient.PhraseLib.Lookup("reward.graphic-fullsizedimage", LanguageID))%>' + '<\/title>')
  siteWindow.document.writeln('<\/head>')
  siteWindow.document.writeln('<body>')
  siteWindow.document.writeln('<center><img id=\"imgFull\" src=\"../images/clear.png\" alt=\"\" title=\"\" /><\/center>')
  siteWindow.document.writeln('<\/body>')
  siteWindow.document.writeln('<\/html>')
  var imgSmall = document.getElementById("imgGraphic");
  var imgLarge = siteWindow.document.getElementById("imgFull");
  imgLarge.src = imgSmall.src;
  siteWindow.document.close();
  siteWindow.focus();
}

function checkForChanges() {
  var elemChanged = document.getElementById("unsavedChanges")
  var elemName = document.getElementById("name");
  var elemDesc = document.getElementById("desc");
  var elemDuration = document.getElementById("duration");
  
  if (elemChanged != null && elemName != null && elemName.value != elemName.defaultValue) {
    elemChanged.value = "true";
    return true;
  }
  if (elemChanged != null && elemDesc != null && elemDesc.value != elemDesc.defaultValue) {
    elemChanged.value = "true";
    return true;
  }
  if (elemChanged != null && elemDuration != null && elemDuration.value != elemDuration.defaultValue) {
    elemChanged.value = "true";
    return true;
  }
}
//      function chooseFile() {
//      document.getElementById("filedata").click();
//   }
//   function fileonclick()
//   {
//   var filename=document.getElementById("filedata").value;
//    document.getElementById("lblfileupload").innerText = filename.replace("C:\\fakepath\\", "");
//  }
function openPreviewWin() {
  var url = "graphic-preview.aspx?adId=<% Sendb(l_adID) %>"
  var popW = <% Sendb(adWidth + 20) %>;
  var popH = <% Sendb(adHeight + 50) %>;
  previewWindow = window.open(url,"Popup", "width="+popW+", height="+popH+", top="+calcTop(popH)+", left="+calcLeft(popW)+", toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes");
  previewWindow.focus();
}

function isValidPath() {
  var filePath = document.uploadform.filedata.value
  var retVal = true;
  var agt=navigator.userAgent.toLowerCase();
  var browser = '<% Sendb(Request.Browser.Browser) %>'
  
  if (browser == 'IE') {
    if (filePath.length >=2) {
      if (filePath.charAt(1)!=":") {
        alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidfilepath", LanguageID))%>');
        retVal = false;
      }
    } else {
      alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidfilepath", LanguageID))%>');
      retVal = false;
    }
  }
  return retVal;
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
<%
  ' If no name, description, etc., then pull them off the query string
  If adName = "" And adDescription = "" And adDisplayDuration = "" Then
    adName = Request.QueryString("name")
    adDescription = Request.QueryString("desc")
    adDisplayDuration = Request.QueryString("displayduration")
    adStoreResponse = Request.QueryString("storeresponse")
    If (adStoreResponse = "True") Then
      adResponseChecked = " checked=""checked"""
    End If
  End If
%>

<form action="#" method="get" id="mainform" name="mainform">
  <input type="hidden" id="unsavedChanges" name="unsavedChanges" value="false" />
  <input type="hidden" id="deleteTP" name="deleteTP" value="" />
  <input type="hidden" id="areaId" name="areaId" value="" />
  <input type="hidden" id="adId" name="adId" value="<% Sendb(l_adID) %>" />
  <div id="intro">
    <h1 id="title">
      <% 
        If l_adID = 0 Then
          Sendb(Copient.PhraseLib.Lookup("term.newgraphic", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.graphic", LanguageID) & " #" & l_adID & ": ")
          Sendb(MyCommon.TruncateString(adName, 40))
        End If
      %>
    </h1>
    <div id="controls">
      <%
        If (l_adID = 0) Then
          If (Logix.UserRoles.CreateGraphics) Then
            Send_Save()
          End If
        Else
          ShowActionButton = (Logix.UserRoles.CreateGraphics) OrElse (Logix.UserRoles.EditGraphics) OrElse (Logix.UserRoles.DeleteGraphics)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.DeleteGraphics) Then
              Send_Delete()
            End If
            If (Logix.UserRoles.CreateGraphics) Then
              Send_New()
            End If
            If (Logix.UserRoles.EditGraphics) Then
              Send_ReDeploy()
            End If
            If (Logix.UserRoles.EditGraphics) Then
              Send_Save()
            End If
            If (Logix.UserRoles.EditGraphics) AndAlso (l_adID > 0) Then
              Send_Upload()
            End If
            If Request.Browser.Type = "IE6" Then
              Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:121px;""></iframe>")
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(11, l_adID, AdminUserID)
            End If
          End If
        End If
      %>
    </div>
  </div>
  <%
    If Request.Browser.Type = "IE6" Then
      IE6ScrollFix = " onscroll=""javascript:document.getElementById('uploader').style.display='none';document.getElementById('actionsmenu').style.visibility='hidden';"""
    End If
  %>
  <div id="main"<% Sendb(IE6ScrollFix) %>>
    <%If (infoMessage <> "" AndAlso Redeployed) Then
        Send("<div id=""infobar"" class=""green-background"">" & infoMessage & "</div>")
      ElseIf (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column1">
      <%  If (dstAds.Rows.Count > 0 And l_adID <> 0) Then%>
      <div class="box" id="identity">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <input type="hidden" id="OnScreenAdID" name="OnScreenAdID" value="<% sendb(l_adID) %>" />
        <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <% If (adName Is Nothing) Then adName = ""%>
        <input type="text" class="longest" id="name" name="name" maxlength="100" value="<% Sendb(adName.Replace("""", "&quot;")) %>" /><br />
        <label for="desc"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" id="desc" name="desc" cols="48" rows="3"><% Sendb(adDescription)%></textarea><br />
        <br class="half" />
        <%
          MyCommon.QueryStr = "select ActivityDate from ActivityLog with (NoLock) where ActivityTypeID='12' and LinkID='" & l_adID & "' order by ActivityDate asc;"
          dst = MyCommon.LRT_Select
          sizeOfData = dst.Rows.Count
          Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))
          If (sizeOfData > 0) Then
            Send(" " & Logix.ToLongDateTimeString(dst.Rows(0).Item("ActivityDate"), MyCommon))
          Else
            Send(": " & Copient.PhraseLib.Lookup("term.unknown", LanguageID))
          End If
          Send("<br />")
          Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))
          If (sizeOfData > 0) Then
            Send(" " & Logix.ToLongDateTimeString(dst.Rows(sizeOfData - 1).Item("ActivityDate"), MyCommon))
          Else
            Send(": " & Copient.PhraseLib.Lookup("term.unknown", LanguageID))
          End If
          Send("<br />")
          Sendb(Copient.PhraseLib.Lookup("term.uploaded", LanguageID) & " ")
          Send(adLastUpload & "<br />")
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="displayduration">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.displayduration", LanguageID))%>
          </span>
        </h2>
        <input type="text" class="shortest" id="duration" name="duration" maxlength="5" value="<% Sendb(adDisplayDuration)%>" />
        <label for="duration"><%Sendb(Copient.PhraseLib.Lookup("graphics.minimumduration", LanguageID))%></label>
        <br />
        <hr class="hidden" />
      </div>
      <div class="box" id="offers">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <%  Sendb(Copient.PhraseLib.Lookup("graphics.usedbyoffers", LanguageID))%>
        <div class="boxscroll">
          <%
            MyCommon.QueryStr = "select I.IncentiveID as OfferID, I.IncentiveName as Name,buy.ExternalBuyerId as BuyerID " & _
                                "from CPE_Deliverables as D with (NoLock) Inner Join CPE_RewardOptions as RO with (NoLock) on D.RewardOptionID=RO.RewardOptionID and D.OutputID=" & l_adID & " and DeliverableTypeID=1 and RO.Deleted=0 and D.Deleted=0 " & _
                                "Inner Join CPE_Incentives as I with (NoLock) on RO.IncentiveID=I.IncentiveID and I.Deleted=0" & _
                                 "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId;"
            dstAssociated = MyCommon.LRT_Select
            rowCount = dstAssociated.Rows.Count
            If rowCount > 0 Then
              For Each row In dstAssociated.Rows
                       If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
                Else
                assocName = MyCommon.NZ(row.Item("Name"),Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                End If
                'assocName = row.Item("Name")
                assocID = row.Item("OfferID")

                If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                  Sendb(" <a href=""offer-redirect.aspx?OfferID=" & assocID & """>" & assocName & "</a><br />")
                Else
                  Sendb(assocName & "<br />")
                End If
              Next
            Else
              Send("  " & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
      <div class="box" id="validation">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID))%>
          </span>
        </h2>
        <%
          Dim dtValid As DataTable
          Dim rowOK(), rowWatches(), rowWarnings() As DataRow
          Dim objTemp As Object
          Dim GraceHours As Integer
          Dim GraceCount As Double
          
          objTemp = MyCommon.Fetch_CPE_SystemOption(41)
          If Not (Integer.TryParse(objTemp.ToString, GraceHours)) Then
            GraceHours = 4
          End If
          
          objTemp = MyCommon.Fetch_CPE_SystemOption(42)
          If Not (Double.TryParse(objTemp.ToString, GraceCount)) Then
            GraceCount = 0.1D
          End If
          
          MyCommon.QueryStr = "dbo.pa_ValidationReport_Graphic"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@OnScreenAdID", SqlDbType.Int).Value = l_adID
          MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
          MyCommon.LRTsp.Parameters.Add("@GraceCount", SqlDbType.Decimal, 2).Value = GraceCount
          
          dtValid = MyCommon.LRTsp_select()
          
          rowOK = dtValid.Select("Status=0", "LocationName")
          rowWatches = dtValid.Select("Status=1", "LocationName")
          rowWarnings = dtValid.Select("Status=2", "LocationName")
          
          Send("<a id=""validLink"" href=""javascript:openPopup('validation-report.aspx?type=gr&amp;id=" & l_adID & "&amp;level=0');"">")
          Send(Copient.PhraseLib.Lookup("term.valid", LanguageID) & " " & Copient.PhraseLib.Lookup("term.locations", LanguageID) & " (" & rowOK.Length & ")</a><br />")
          Send("<a id=""watchLink"" href=""javascript:openPopup('validation-report.aspx?type=gr&amp;id=" & l_adID & "&amp;level=1');"">")
          Send(Copient.PhraseLib.Lookup("term.watch", LanguageID) & " " & Copient.PhraseLib.Lookup("term.locations", LanguageID) & " (" & rowWatches.Length & ")</a><br />")
          Send("<a id=""warningLink"" href=""javascript:openPopup('validation-report.aspx?type=gr&amp;id=" & l_adID & "&amp;level=2');"">")
          Send(Copient.PhraseLib.Lookup("term.warning", LanguageID) & " " & Copient.PhraseLib.Lookup("term.locations", LanguageID) & " (" & rowWarnings.Length & ")</a><br />")
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <%
        Dim GraphicPath As String = DEFAULT_GRAPHIC_PATH
        Dim imgExt As String = "1"
        Dim imgDisplayWidth As Integer = 350, imgDisplayHeight As Integer = 208
        Dim tempWidth As Integer, tempHeight As Integer
        Dim GraphicResized As Boolean = False
        
        MyCommon.QueryStr = "select OptionValue from SystemOptions with (nolock) where OptionID = 47;"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          GraphicPath = dst.Rows(0).Item("OptionValue")
          If (GraphicPath.Trim().Length = 0) Then
            GraphicPath = DEFAULT_GRAPHIC_PATH
          End If
          If Not (Right(GraphicPath, 1) = "\") Then
            GraphicPath = GraphicPath & "\"
          End If
        End If
        
        If (adImageType = "1") Then
          imgExt = "jpg"
        ElseIf (adImageType = "2") Then
          imgExt = "gif"
        End If
        If (adGraphicSize > 0) Then
          GraphicPath = GraphicPath & CStr(l_adID) & "img." & imgExt
        End If
        
        If (System.IO.File.Exists(GraphicPath)) Then
          If (adHeight > imgDisplayHeight OrElse adWidth > imgDisplayWidth) Then
            tempHeight = adHeight
            tempWidth = adWidth
            While (tempHeight > imgDisplayHeight OrElse tempWidth > imgDisplayWidth)
              tempHeight /= 1.1
              tempWidth /= 1.1
            End While
            imgDisplayHeight = tempHeight
            imgDisplayWidth = tempWidth
            GraphicResized = True
          Else
            imgDisplayHeight = adHeight
            imgDisplayWidth = adWidth
          End If
        End If
      %>
      <div class="box" id="image">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.image", LanguageID))%>
          </span>
        </h2>
        <% If (adGraphicSize > 0) Then%>
        <center>
          <a href="javascript:showImageWindow('<% Sendb(GraphicPath) %>');">
            <img id="imgGraphic" src="graphic-display-img.aspx?path=<% Sendb(GraphicPath) %>&amp;lang=<% Sendb(LanguageID)%>" alt="<% Sendb(Copient.PhraseLib.Lookup("term.image", LanguageID))%>" title="<% Sendb(Copient.PhraseLib.Lookup("term.image", LanguageID))%>" width="<% Sendb(imgDisplayWidth) %>" height="<% Sendb(imgDisplayHeight) %>" /><br />
          </a>
        </center>
        <br class="half" />
        <% If (GraphicResized) Then%>
        <center>
          <a href="javascript:showImageWindow('<% Sendb(GraphicPath) %>');">(<% Sendb(Copient.PhraseLib.Lookup("graphics.seefullsize", LanguageID))%>)</a><br />
        </center>
        <% End If%>
        <br class="half" />
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.details", LanguageID))%>">
          <tr>
            <td>
              <% Sendb(Copient.PhraseLib.Lookup("term.uploaded", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.from", LanguageID), VbStrConv.Lowercase) & ":")%>
            </td>
            <td colspan="2">
              <% Sendb(MyCommon.SplitNonSpacedString(adClientFileName, 25))%>
            </td>
          </tr>
          <tr>
            <td>
              <%Sendb(Copient.PhraseLib.Lookup("term.imagetype", LanguageID) & ":")%>
            </td>
            <td>
              <%  
                If (imgExt = "jpg") Then
                  Sendb(Copient.PhraseLib.Lookup("term.jpeg", LanguageID))
                ElseIf (imgExt = "gif") Then
                  Sendb(Copient.PhraseLib.Lookup("term.gif", LanguageID))
                End If
              %>
            </td>
            <td>
            </td>
          </tr>
          <tr>
            <td>
              <%Sendb(Copient.PhraseLib.Lookup("term.width", LanguageID) & ":")%>
            </td>
            <td>
              <% Sendb(adWidth)%>
              &nbsp;<%Sendb(Copient.PhraseLib.Lookup("term.pixels", LanguageID))%></td>
            <td>
            </td>
          </tr>
          <tr>
            <td>
              <%Sendb(Copient.PhraseLib.Lookup("term.height", LanguageID) & ":")%>
            </td>
            <td>
              <% Sendb(adHeight)%>
              &nbsp;<%Sendb(Copient.PhraseLib.Lookup("term.pixels", LanguageID))%></td>
            <td>
            </td>
          </tr>
        </table>
        <% Else%>
        <center>
          <i><% Sendb(Copient.PhraseLib.Lookup("graphics.notuploaded", LanguageID))%></i>
        </center>
        <% End If%>
        <hr class="hidden" />
      </div>
      <div class="box" id="touchpoints">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.touchpoints", LanguageID))%>
          </span>
        </h2>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.touchpoints", LanguageID))%>">
          <thead>
            <tr>
              <th>
              </th>
              <th class="th-name" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.areaname", LanguageID))%>
              </th>
              <th class="th-xpos" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.xpos", LanguageID))%>
              </th>
              <th class="th-ypos" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.ypos", LanguageID))%>
              </th>
              <th class="th-width" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.width", LanguageID))%>
              </th>
              <th class="th-height" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.height", LanguageID))%>
              </th>
              <th>
              </th>
            </tr>
          </thead>
          <tbody>
            <%
              Dim tpAreaID As String
              Dim tpName As String
              Dim tpX As Integer
              Dim tpY As Integer
              Dim tpWidth As Integer
              Dim tpHeight As Integer
              
              MyCommon.QueryStr = "select AreaID, Name, X, Y, Width, Height from TouchAreas with (NoLock) where OnScreenAdID=" & l_adID & " and Deleted=0 order by Name;"
              dstPoints = MyCommon.LRT_Select()
              
              rowCount = dstPoints.Rows.Count
              If rowCount > 0 Then
                For Each row In dstPoints.Rows
                  tpName = MyCommon.NZ(row.Item("Name"), "")
                  tpAreaID = MyCommon.NZ(row.Item("AreaID"), "")
                  tpX = MyCommon.NZ(row.Item("X"), 0)
                  tpY = MyCommon.NZ(row.Item("Y"), 0)
                  tpWidth = MyCommon.NZ(row.Item("Width"), 0)
                  tpHeight = MyCommon.NZ(row.Item("Height"), 0)
                  
                  Send("<tr>")
                  If (Logix.UserRoles.EditGraphics) Then
                    Sendb("    <td><input type=""button"" class=""ex"" name=""ex"" value=""X"" title=""")
                    Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))
                    Sendb(""" onclick=""javascript:deleteTouchPt(" & tpAreaID & ", " & l_adID & ");""")
                    Sendb(IIf(TouchpointsEditable, "", " disabled=""disabled"""))
                    Sendb(" />")
                    Send("</td>")
                  Else
                    Send("<td>&nbsp;</td>")
                  End If
                  Send("    <td>" & MyCommon.SplitNonSpacedString(tpName, 15) & "</td>")
                  Send("    <td>" & tpX & "</td>")
                  Send("    <td>" & tpY & "</td>")
                  Send("    <td>" & tpWidth & "</td>")
                  Send("    <td>" & tpHeight & "</td>")
                  Send("    <td></td>")
                  Send("</tr>")
                Next
              Else
                Send("<tr>")
                Send("    <td></td>")
                Send("    <td colspan=""5"">" & Copient.PhraseLib.Lookup("graphics.NoTouchpoints", LanguageID) & "</td>")
                Send("    <td></td>")
                Send("</tr>")
              End If
            %>
            <% If (TouchpointsEditable) Then%>
            <tr>
              <td><input type="button" class="ex" value="X" name="ex" title="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>" onclick="javascript:clearNewTP();" /></td>
              <td><input type="text" id="txtAreaName" name="txtAreaName" value="" maxlength="100" size="10" /></td>
              <td><input type="text" id="txtXPos" name="txtXPos" value="" maxlength="5" style="width: 35px;" /></td>
              <td><input type="text" id="txtYPos" name="txtYPos" value="" maxlength="5" style="width: 35px;" /></td>
              <td><input type="text" id="txtWidth" name="txtWidth" value="" maxlength="5" style="width: 35px;" /></td>
              <td><input type="text" id="txtHeight" name="txtHeight" value="" maxlength="5" style="width: 35px;" /></td>
              <% If (adGraphicSize > 0) Then%>
              <td><a class="hidden" href="graphic-mapping.aspx?adId=<% Sendb(l_adID)%>">►</a><a href="javascript:openPopup('graphic-mapping.aspx?adId=<% Sendb(l_adID)%>');"><% Sendb(Copient.PhraseLib.Lookup("term.map", LanguageID))%></a></td>
              <% Else%>
              <td><span style="visibility: hidden;"><% Sendb(Copient.PhraseLib.Lookup("term.map", LanguageID))%></span></td>
              <% End If%>
            </tr>
            <% Else%>
            <tr>
              <td>
              </td>
              <td colspan="5" class="darkred">
                <small><% Sendb(Copient.PhraseLib.Lookup("graphics.TouchpointsUneditable", LanguageID))%></small></td>
              <td>
              </td>
            </tr>
            <% End If%>
          </tbody>
        </table>
        <% If (Logix.UserRoles.EditGraphics AndAlso TouchpointsEditable) Then%>
        <input type="submit" class="regular" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID))%>" />
        <% End If%>
        <% If (Touchable) Then%>
        <input type="button" class="regular" id="preview" name="preview" onclick="javascript:openPreviewWin();" value="<% Sendb(Copient.PhraseLib.Lookup("term.preview", LanguageID))%>" /><br />
        <% End If%>
        <% If (Touchable) Then%>
        <input type="checkbox" id="storeresponse" name="storeresponse" title="<% Sendb(Copient.PhraseLib.Lookup("graphics.captureresponse", LanguageID)) %>"<% sendb(adresponsechecked) %> />
        <label for="storeresponse">
          <% Sendb(Copient.PhraseLib.Lookup("term.captureresponse", LanguageID))%>
          -
          <% Sendb(Copient.PhraseLib.Lookup("graphics.captureresponse", LanguageID))%>
        </label>
        <% End If%>
        <span style="line-height: 0.1;">&nbsp;</span><br />
        <hr class="hidden" />
      </div>
      <div class="box" id="cells">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.screencells", LanguageID))%>
          </span>
        </h2>
        <% Sendb(Copient.PhraseLib.Lookup("graphics.compatible", LanguageID))%>
        <br />
        <div class="boxscroll" style="height: 75px;">
          <%
            Dim dstScreenCells As DataTable = Nothing
            Dim RowCt As Integer
                        
            MyCommon.QueryStr = "select SL.Name as LayoutName, SC.Name as CellName, SL.LayoutID, SC.CellID " & _
                            "from ScreenLayouts as SL with (NoLock) Inner Join ScreenCells as SC with (NoLock) on SC.LayoutID=SL.LayoutID and SL.Deleted=0 and SC.Deleted=0 " & _
                            "where SC.ContentsID=1 and SC.Width=" & adWidth & " and SC.Height=" & adHeight & " order by SL.Name, SC.Name;"
            dstScreenCells = MyCommon.LRT_Select
            If (dstScreenCells.Rows.Count > 0) Then
              Send("<table border=""0"" cellpadding=""2"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.cells", LanguageID) & """>")
              Send("<tr>")
              Send("  <th scope=""col"">" & Copient.PhraseLib.Lookup("term.layout", LanguageID) & "</th>")
              Send("  <th></th>")
              Send("  <th scope=""col"">" & Copient.PhraseLib.Lookup("term.cell", LanguageID) & "</th>")
              Send("</tr>")
              RowCt = dstScreenCells.Rows.Count - 1
              For i = 0 To RowCt
                Send("<tr>")
                Send("  <td valign=""top""><a href=""layout-edit.aspx?LayoutID=" & MyCommon.NZ(dstScreenCells.Rows(i).Item("LayoutID"), 0) & """ >" & MyCommon.NZ(dstScreenCells.Rows(i).Item("LayoutName"), "") & "</a></td>")
                Send("  <td></td>")
                Send("  <td valign=""top""><a href=""layout-cell.aspx?LayoutID=" & MyCommon.NZ(dstScreenCells.Rows(i).Item("LayoutID"), 0) & "&amp;CellID=" & MyCommon.NZ(dstScreenCells.Rows(i).Item("CellID"), 0) & """ >" & MyCommon.NZ(dstScreenCells.Rows(i).Item("CellName"), "") & "</a></td>")
                Send("</tr>")
              Next
              Send("</table>")
            Else
              Send(Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
      
      <%ElseIf (Request.QueryString("New") <> "") Or (Request.QueryString("OnScreenAdID") = "new") Or (Request.QueryString("OnScreenAdID") = "0") Or (Request.QueryString("OnScreenAdID") = "") Then%>
      <div class="box" id="newIdentity">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <%If (l_adID <> 0) Then
            Send(Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & l_adID)
            Send("<br />")
          End If
        %>
        <input type="hidden" id="OnScreenAdID" name="OnScreenAdID" value="new" />
        <br class="half" />
        <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID) & ":")%>
        <br />
        <input type="text" class="longest" id="name" name="name" maxlength="100" />
        <br />
        <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID) & ":")%>
        <br />
        <textarea class="longest" id="desc" name="desc" cols="48" rows="3"></textarea><br />
        <hr class="hidden" />
      </div>
      <%
      Else
      %>
      <div class="box" id="identity">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <%
          Sendb("ID " & l_adID & " " & Copient.PhraseLib.Lookup("term.notfound", LanguageID))
        End If
        %>
        <hr class="hidden" />
      </div>
      <br clear="all" />
    </div>
</form>

<div id="uploader" style="display: none;">
  <div id="uploadwrap">
    <div class="box" id="uploadbox">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("term.upload", LanguageID))%>
        </span>
      </h2>
      <form action="graphic-edit.aspx?mode=accept&amp;adName=<% Sendb(Server.UrlEncode(adName))%>&amp;adId=<% Sendb(l_adID)%>" id="uploadform" name="uploadform" onsubmit="return isValidPath();" method="post" enctype="multipart/form-data">
        <%
          Sendb("<input type=""button"" class=""ex"" id=""uploaderclose"" name=""uploaderclose"" value=""X"" style=""float:right;position:relative;top:-27px;"" ")
          Sendb("onclick=""javascript:document.getElementById('uploader').style.display='none';"" />")
        %>
        <% Send(Copient.PhraseLib.Lookup("graphics.uploadinstructions1", LanguageID) & " <b>'" & MyCommon.SplitNonSpacedString(adName, 25) & "'</b> " & Copient.PhraseLib.Lookup("graphics.uploadinstructions2", LanguageID))%>
        <br />
        <br class="half" />
        <%
          If (Logix.UserRoles.EditCustomerGroups) Then
            Send("     <input type=""file"" id=""filedata"" name=""filedata"" accept=""image/*"" size=""20"" />")
        '           Send("<div id=""divfile"" style=""height:0px;overflow:hidden"">")
        'Send("<input type=""file"" id=""filedata"" name=""fileInput"" onchange=""fileonclick()"" accept=""image/*"" size=""20"" />")
        'Send("</div>")
        'Send("<button type=""button"" onclick=""chooseFile();"">"&Copient.PhraseLib.Lookup("term.browse", LanguageID) &"</button>")
        'Send("<label id=""lblfileupload"" name=""lblfileupload"">"&Copient.PhraseLib.Lookup("term.nofilesselected", LanguageID)&"</label>")
            Send("     <input type=""submit"" class=""regular"" id=""uploadfile"" name=""uploadfile"" value=""" & Copient.PhraseLib.Lookup("term.upload", LanguageID) & """ />")
            Send("     <br />")
          End If
        %>
      </form>
      <hr class="hidden" />
    </div>
  </div>
</div>

<script runat="server">
  Const THUMBNAIL_SIZE As Integer = 120
  
  Private Sub MakeThumbnail(ByVal filename As String, ByVal image As Drawing.Image, ByVal imgExt As String)
    Dim width, height As Integer
    ' Get the source bitmap.
    Dim bm_source As New Bitmap(image)
    
    'Resize while maintaining aspect
    If (image.Height > THUMBNAIL_SIZE OrElse image.Width > THUMBNAIL_SIZE) Then
      Dim tempHeight As Integer = image.Height
      Dim tempWidth As Integer = image.Width
      While (tempHeight > THUMBNAIL_SIZE OrElse tempWidth > THUMBNAIL_SIZE)
        tempHeight /= 1.1
        tempWidth /= 1.1
      End While
      height = tempHeight
      width = tempWidth
    Else
      height = image.Height
      width = image.Width
    End If
    
    ' Make a bitmap for the result.
    Dim bm_dest As New Bitmap(width, height)
    
    ' Make a Graphics object for the result Bitmap.
    Dim gr_dest As Graphics = Graphics.FromImage(bm_dest)
    
    ' Copy the source image into the destination bitmap.
    gr_dest.DrawImage(bm_source, 0, 0, _
        bm_dest.Width + 1, _
        bm_dest.Height + 1)
        
    If (UCase(imgExt) = "GIF") Then
      bm_dest.Save(filename & imgExt, ImageFormat.Gif)
    Else
      bm_dest.Save(filename & imgExt, ImageFormat.Jpeg)
    End If
  End Sub
  
  Private Sub RemoveGraphics(ByVal AdId As Long)
    Dim GraphicPath As String
    Dim MyCommon As New Copient.CommonInc
    
    MyCommon.Open_LogixRT()
    
    Dim DEFAULT_GRAPHIC_PATH As String = MyCommon.Get_Install_Path
    
    GraphicPath = MyCommon.Fetch_SystemOption(47)
    If (GraphicPath = "") Then
      GraphicPath = DEFAULT_GRAPHIC_PATH
    End If
    If Not (Right(GraphicPath, 1) = "\") Then
      GraphicPath = GraphicPath & "\"
    End If
    
    If (File.Exists(GraphicPath & AdId & "img.jpg")) Then File.Delete(GraphicPath & AdId & "img.jpg")
    If (File.Exists(GraphicPath & AdId & "img_tn.jpg")) Then File.Delete(GraphicPath & AdId & "img_tn.jpg")
    If (File.Exists(GraphicPath & AdId & "img.gif")) Then File.Delete(GraphicPath & AdId & "img.gif")
    If (File.Exists(GraphicPath & AdId & "img_tn.gif")) Then File.Delete(GraphicPath & AdId & "img_tn.gif")
    
    MyCommon.Close_LogixRT()
  End Sub
</script>

<script type="text/javascript">
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (l_adID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(11, l_adID, AdminUserID)
    End If
  End If
  Send_BodyEnd("mainform", "name")
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
