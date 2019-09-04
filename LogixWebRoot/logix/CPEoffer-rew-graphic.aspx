<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-rew-graphic.aspx 
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
  Dim rst2 As DataTable
  Dim row As DataRow
  Dim IsTemplate As Boolean = False
  Dim OfferID As Long
  Dim Name As String = ""
  Dim RewardID As String
  Dim Phase As Integer
  Dim CellID As String
  Dim OptionID As String
  Dim NumRecs As Integer
  Dim Width As Integer
  Dim Height As Integer
  Dim GraphicPath As String = ""
  Dim PreviewGraphicPath As String = ""
  Dim jsPath As String
  Dim ImageType As String
  Dim Shaded As String = "shaded"
  Dim DeliverableID As Integer
  Dim NewOptionName As String
  Dim NewOptionID As String = ""
  Dim OnScreenAdName As String = ""
  Dim AdID As String = ""
  Dim bPreviewOnly As Boolean
  Dim PreviewDisplayVal As String
  Dim PreviewClose As String
  Dim QryStr As String = ""
  Dim PhaseTitle As String = ""
  Dim CloseAfterSave As Boolean = False
  Dim ShowAllLayouts As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim EngineID As Integer = 2
  Dim BannersEnabled As Boolean = True
  Dim ValidTiers As Boolean = True
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-rew-graphic.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  OptionID = MyCommon.Extract_Val(Request.QueryString("roid"))
  CellID = MyCommon.Extract_Val(Request.QueryString("cellselect"))
  DeliverableID = MyCommon.Extract_Val(Request.QueryString("DeliverableID"))
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  
  Phase = MyCommon.Extract_Val(Request.QueryString("phase"))
  If (Phase = 0) Then Phase = MyCommon.Extract_Val(Request.Form("Phase"))
  If (Phase = 0) Then Phase = 3
  Select Case Phase
    Case 1 ' Notification
      PhaseTitle = "term.graphicnotification"
    Case 2 ' Accumulation
      PhaseTitle = ""
    Case 3 ' Reward
      PhaseTitle = "term.graphicreward"
    Case Else
      PhaseTitle = "term.graphicreward"
  End Select
  
  AdID = MyCommon.Extract_Val(Request.QueryString("ad"))
  bPreviewOnly = IIf(Request.QueryString("preview") <> "", True, False)
  If (bPreviewOnly) Then
    PreviewDisplayVal = "block;"
    PreviewClose = "window.close();"
    PreviewGraphicPath = "graphic-display-img.aspx?path="
    PreviewGraphicPath += LoadGraphicPath(AdID)
    PreviewGraphicPath = PreviewGraphicPath.Replace("\", "\\")
    If (Request.QueryString("imagetype") <> "") Then
      ImageType = IIf(Request.QueryString("imagetype") = "2", "gif", "jpg")
      PreviewGraphicPath = PreviewGraphicPath.Replace("_tn.jpg", "_tn." & ImageType)
      PreviewGraphicPath = PreviewGraphicPath.Replace("_tn." & ImageType, "." & ImageType)
    End If
  Else
    PreviewDisplayVal = "none;"
    PreviewClose = "closePreview();"
    PreviewGraphicPath = ""
  End If
  
  If (Request.QueryString("infoMessage") <> "") Then
    infoMessage = Request.QueryString("infoMessage")
  ElseIf (Request.QueryString("saved") <> "") Then
    infoMessage = Copient.PhraseLib.Lookup("reward.graphic-added", LanguageID) & " " & OfferID
  End If
  
  If (Request.QueryString("save") <> "" AndAlso Request.QueryString("additem") <> "") Then
    'Get the name of the graphic .. gonna need it later
    MyCommon.QueryStr = "select Name, ImageType from OnScreenAds with (NoLock) where OnScreenAdID=" & Request.QueryString("additem") & " and Deleted=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      OnScreenAdName = MyCommon.NZ(rst.Rows(0).Item("Name"), "No Graphic Name")
      ImageType = MyCommon.NZ(rst.Rows(0).Item("ImageType"), "")
    End If
    
    ' Get the ROID
    If (OptionID = 0) Then
      MyCommon.QueryStr = "select RewardOptionID, isnull(HHEnable, 0) as HHEnable from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        OptionID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
      End If
    End If
    
    'Add the OnScreenAd to Deliverables
    MyCommon.QueryStr = "dbo.pa_CPE_AddGraphic"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@ROID", SqlDbType.Int, 4).Value = OptionID
    MyCommon.LRTsp.Parameters.Add("@OutputID", SqlDbType.Int, 4).Value = Request.QueryString("additem")
    MyCommon.LRTsp.Parameters.Add("@CellID", SqlDbType.Int, 4).Value = CellID
    MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
    MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.BigInt, 8).Direction = ParameterDirection.Output
    
    MyCommon.LRTsp.ExecuteNonQuery()
    DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
    
    MyCommon.Close_LRTsp()
    
    'See if the selected Graphic has any touchable areas and create new Reward Options based on those areas
    MyCommon.QueryStr = "select AreaID, Name from TouchAreas as TA with (NoLock) where TA.OnScreenAdID=" & Request.QueryString("additem") & " and TA.Deleted=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        'Create the New Reward Option
        NewOptionName = OnScreenAdName & " - " & MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("reward.graphic-noareaname", LanguageID))
        
        MyCommon.QueryStr = "dbo.pa_CPE_AddGraphicOption"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OptionName", SqlDbType.NVarChar, 255).Value = NewOptionName
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int, 4).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@NewOptionID", SqlDbType.BigInt, 8).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        NewOptionID = MyCommon.LRTsp.Parameters("@NewOptionID").Value
        MyCommon.Close_LRTsp()
        
        'Create a new DeliverableROIDs record
        MyCommon.QueryStr = "insert into CPE_DeliverableROIDs with (RowLock) (DeliverableID, AreaID, RewardOptionID, IncentiveID, Deleted, LastUpdate) values (" & DeliverableID & ", " & MyCommon.NZ(row.Item("AreaID"), 0) & ", " & NewOptionID & ", " & OfferID & ", 0, getdate());"
        MyCommon.LRT_Execute()
      Next
    End If
    
    If (EngineID = 5) Then
      MyCommon.QueryStr = "update CPE_deliverables with (RowLock) set Priority= (select Max(IsNull(Priority,0)) + 1 from CPE_Deliverables where RewardOptionID=" & OptionID & " and Deleted=0) where DeliverableID=" & DeliverableID
      MyCommon.LRT_Execute()
    End If
    
    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID & ";"
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("CPE_Reward.createdgraphic", LanguageID))
    
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    If Not (CloseAfterSave) Then
      Response.Status = "301 Moved Permanently"
      QryStr = "?OfferID=" & OfferID & "&ad=" & Request.QueryString("additem") & "&imagetype=" & ImageType & "&preview=1&saved=1&Phase=" & Phase & "&EngineID=" & EngineID
      Response.AddHeader("Location", "CPEoffer-rew-graphic.aspx" & QryStr)
    Else
      Send("<script type=""text/javascript"" language=""javascript"">")
      Send("    window.close();")
      Send("</script>")
      Response.End()
    End If
  End If
  
    Send_HeadBegin("term.offer", PhaseTitle, OfferID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript" language="javascript">
  function reloadForNewCell() {
      document.mainform.submit()
  }
  
  function showFullSize(path) {
      var elem = document.getElementById("imagepreview");
      var elemList = document.getElementById("selector");
      var elemCell = document.getElementById("cell");
      var btnSave = document.getElementById("save");
      var elemClose = document.getElementById("btnClose");
      
      if (elem != null) {
          elem.style.display = "block";
          var elemImg = document.getElementById("fsImage");
          if (elemImg != null) {
              elemImg.src = "graphic-display-img.aspx?path=" + escape(path);
          }
          if (elemList != null) { elemList.style.display="none"; }
          if (elemCell != null) { elemCell.style.display="none"; }
          if (btnSave != null) { btnSave.style.display="none";  }
          if (elemClose != null) { elemClose.style.display="block"; }
      }
  }
  
  function closePreview() {
      var elem = document.getElementById("imagepreview");
      var elemList = document.getElementById("selector");
      var elemCell = document.getElementById("cell");
      var btnSave = document.getElementById("save");
      var elemClose = document.getElementById("btnClose");
      
      if (elem != null) { elem.style.display = "none"; }
      if (elemClose != null) { elemClose.style.display="none"; }
      if (elemList != null) { elemList.style.display="block"; }
      if (elemCell != null) { elemCell.style.display="block"; }
      if (btnSave != null) { btnSave.style.display="block";  }
  }
</script>
<%
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (Phase = 3) Then
    If (EngineID = Copient.CommonInc.InstalledEngines.CAM) Then
      Send("  opener.location = 'CAM/CAM-offer-rew.aspx?OfferID=" & OfferID & "'; ")
    ElseIf (EngineID = Copient.CommonInc.InstalledEngines.Website) Then
      Send("  opener.location = 'web-offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Else
      Send("  opener.location = 'CPEoffer-rew.aspx?OfferID=" & OfferID & "'; ")
    End If
  ElseIf (Phase = 1) Then
    If (EngineID = Copient.CommonInc.InstalledEngines.Website) Then
      Send("  opener.location = 'web-offer-not.aspx?OfferID=" & OfferID & "'; ")
    Else
      Send("  opener.location = 'offer-channels.aspx?OfferID=" & OfferID & "'; ")
    End If
  End If
  Send("} ")
  Send("</script>")
  Send_HeadEnd()
  
  If (IsTemplate) Then
    Send_BodyBegin(12)
  Else
    Send_BodyBegin(2)
  End If
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send_Denied(2, "perm.offers-access-templates")
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function ChangeParentDocument() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="DeliverableID" name="DeliverableID" value="<% sendb(DeliverableID) %>" />
    <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" id="RewardID" name="RewardID" value="<% sendb(RewardID) %>" />
    <input type="hidden" id="roid" name="roid" value="<% sendb(OptionID) %>" />
    <input type="hidden" id="Phase" name="Phase" value="<% Sendb(Phase)%>" />
    <input type="hidden" id="EngineID" name="EngineID" value="<%Sendb(EngineID) %>" />
    <%If (isTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup(PhaseTitle, LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup(PhaseTitle, LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <%If (Not bPreviewOnly) Then
          Sendb("<input type=""button"" class=""regular"" id=""btnClose"" name=""btnClose"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID))
          Send(""" style=""display:none;"" onclick=""javascript:" & PreviewClose & """ />")
          If DeliverableID = 0 Then
            Send_Save(" onclick=""this.style.visibility='hidden';""")
          Else
            Send_Save()
          End If
        Else
          Sendb("<input type=""button"" class=""regular"" id=""btnClose"" name=""btnClose"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID))
          Send(""" onclick=""javascript:" & PreviewClose & """ />")
        End If
      %>
    </div>
  </div>
  <div id="main" style="overflow: auto;">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <div class="box" id="cell" style="display: <% Sendb(IIf(bPreviewOnly, "none", "block"))%>;">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.cell", LanguageID))%>
          </span>
        </h2>
        <%
          ' show all layouts when offer is a website or email offer
          If (EngineID = 3 OrElse EngineID = 5) Then
            ShowAllLayouts = True
            NumRecs = 1
          Else
            'send the list of available Screen Cells for this promotion
            NumRecs = 0
            'MyCommon.QueryStr = "select count(*) as NumRecs from OfferTerminals as OT where OfferID=" & OfferID & " and OT.Excluded=0;"
            MyCommon.QueryStr = "select ot.TerminalTypeID, tt.AnyTerminal from offerterminals ot with (NoLock) " & _
                                "inner join terminaltypes tt with (NoLock) on ot.terminaltypeid = tt.terminaltypeid  " & _
                                "where offerid=" & OfferID & " and tt.deleted=0 and tt.EngineID=2 and ot.excluded=0;"
            rst = MyCommon.LRT_Select()
            NumRecs = rst.Rows.Count
            For Each row In rst.Rows
              ShowAllLayouts = ShowAllLayouts OrElse MyCommon.NZ(row.Item("AnyTerminal"), False)
            Next
          End If
          
          If NumRecs > 0 And Not ShowAllLayouts Then
            'query for ScreenLayouts/Cells that are linked to in store locations that are used with this promotion
            MyCommon.QueryStr = "select Distinct SL.Name as LayoutName, SC.Name as CellName, SC.CellID " & _
                                "from OfferTerminals as OT with (NoLock) Inner Join TerminalTypes as TT with (NoLock) on OT.TerminalTypeID=TT.TerminalTypeID and OT.Excluded=0 " & _
                                "Inner Join ScreenLayouts as SL with (NoLock) on TT.LayoutID=SL.LayoutID " & _
                                "Inner Join ScreenCells as SC with (NoLock) on SL.LayoutID=SC.LayoutID and SC.Deleted=0 and SC.ContentsID=1 " & _
                                "Where OfferID=" & OfferID & ";"
          ElseIf NumRecs > 0 And ShowAllLayouts Then
            'query for all ScreenLayouts/Cells that are linked to ANY in store location since this promtion is available at all locations
            'MyCommon.QueryStr = "select Distinct SL.Name as LayoutName, SC.Name as CellName, SC.CellID " & _
            '                    "from TerminalTypes as TT with (NoLock) Inner Join ScreenLayouts as SL with (NoLock) on TT.LayoutID=SL.LayoutID and SL.Deleted=0 " & _
            '                    "Inner Join ScreenCells as SC with (NoLock) on SL.LayoutID=SC.LayoutID and SC.Deleted=0 and SC.ContentsID=1;"
            MyCommon.QueryStr = "select Distinct SL.Name as LayoutName, SC.Name as CellName, SC.CellID " & _
                                "from ScreenLayouts as SL with (NoLock) " & _
                                "Inner Join ScreenCells as SC with (NoLock) on SL.LayoutID=SC.LayoutID " & _
                                "and SC.Deleted=0 and SL.Deleted=0 and SC.ContentsID=1;"

          End If
          
          If (NumRecs > 0) Then
            rst = MyCommon.LRT_Select()
            If (rst.Rows.Count > 0) Then
              Send("<select class=""longer"" id=""cellselect"" name=""cellselect"" onchange=""javascript:reloadForNewCell();"">")
              Send("<option value=""0"">(" & Copient.PhraseLib.Lookup("reward.graphic-selectcell", LanguageID) & ")</option>")
              For Each row In rst.Rows
                Sendb("<option value=""" & MyCommon.NZ(row.Item("CellID"), 0) & """")
                If (rst.Rows.Count = 1) Then
                  Sendb(" selected=""selected""")
                  CellID = MyCommon.NZ(row.Item("CellID"), "0")
                ElseIf MyCommon.NZ(row.Item("CellID"), 0) = CellID Then
                  Sendb(" selected=""selected""")
                End If
                Send(">" & MyCommon.NZ(row.Item("LayoutName"), "") & " - " & MyCommon.NZ(row.Item("CellName"), "") & "</option>")
              Next
              Send("</select><br />")
            End If
          Else
            Send(Copient.PhraseLib.Lookup("cpeoffer-rew-graphic-noscreenlayout", LanguageID))
          End If
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="selector" style="display: <% Sendb(IIf(bPreviewOnly, "none", "block"))%>;">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.graphics", LanguageID))%>
          </span>
        </h2>
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.image", LanguageID))%>">
          <tr>
            <th class="th-select" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>
            </th>
            <th class="th-image" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.image", LanguageID))%>
            </th>
            <th class="th-id" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
            </th>
            <th class="th-name" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
            </th>
            <th class="th-dimensions" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.dimensions", LanguageID))%>
            </th>
            <th class="th-format" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.format", LanguageID))%>
            </th>
            <th class="th-touchpoints" scope="col">
              <% Sendb(Copient.PhraseLib.Lookup("term.touchpoints", LanguageID))%>
            </th>
          </tr>
          <tbody>
            <%
              MyCommon.QueryStr = "select SL.Name as LayoutName, SC.Name as CellName, SC.Width, SC.Height from ScreenLayouts as SL with (NoLock) Inner Join ScreenCells as SC with (NoLock) on SL.LayoutID=SC.LayoutID and SC.Deleted=0 and SL.Deleted=0 where SC.CellID=" & CellID & ";"
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                Width = MyCommon.NZ(rst.Rows(0).Item("Width"), 1)
                Height = MyCommon.NZ(rst.Rows(0).Item("Height"), 1)
              End If
              'show the selected graphic, if applicable, at the top of the list
              'show the list of available On-Screen Ads
              MyCommon.QueryStr = "select OSA.Name as Name, OSA.OnScreenAdID as PKID, OSA.ImageType, " & _
                                  "(select Count(AreaID) as TPCount from TouchAreas with (NoLock) where Deleted = 0 and OnScreenAdID=OSA.OnScreenAdID) as Touchpoints " & _
                                  "from OnScreenAds OSA with (NoLock) " & _
                                  "where Deleted=0 and Width=" & Width & " and Height=" & Height & " " & _
                                  "and not OSA.OnScreenAdID in (select OutputID from CPE_Deliverables with (NoLock) where Deleted=0 and DeliverableTypeID=1 and RewardOptionID in (select RewardOptionID from CPE_RewardOptions with (NoLock) where RewardOptionID = " & OptionID & " or IncentiveID=" & OfferID & ")) order by OSA.Name;"
              'Send(MyCommon.QueryStr)
              rst = MyCommon.LRT_Select
              If (rst.Rows.Count > 0) Then
                For Each row In rst.Rows
                  ImageType = IIf(MyCommon.NZ(row.Item("ImageType"), 1) = 1, "jpg", "gif")
                  GraphicPath = LoadGraphicPath(MyCommon.NZ(row.Item("PKID"), 0))
                  jsPath = GraphicPath.Replace("\", "\\")
                  jsPath = jsPath.Replace("_tn.jpg", "." & ImageType)
                  GraphicPath = GraphicPath.Replace("_tn.jpg", "_tn." & ImageType)
                  Send("<tr class=""" & Shaded & """>")
                  Send("<td align=""center""><input type=""radio"" id=""additem" & MyCommon.NZ(row.Item("PKID"), "0") & """ name=""additem"" value=""" & MyCommon.NZ(row.Item("PKID"), "0") & """ /></td>")
                  Send("<td><a href=""javascript:showFullSize('" & jsPath & "');""><img id=""imgGraphic" & MyCommon.NZ(row.Item("PKID"), "0") & """ src=""graphic-display-img.aspx?path=" & GraphicPath & "&lang=" & LanguageID & """ alt=""" & Copient.PhraseLib.Lookup("reward.graphic-clicktoview", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("reward.graphic-clicktoview", LanguageID) & """ /></a></td>")
                  Send("<td>" & MyCommon.NZ(row.Item("PKID"), "&nbsp;") & "</td>")
                  Send("<td>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</td>")
                  Send("<td>" & Width & "&nbsp;x&nbsp;" & Height & "</td>")
                  Send("<td>" & ImageType.ToUpper & "</td>")
                  Send("<td>" & MyCommon.NZ(row.Item("Touchpoints"), "0") & "</td>")
                  Send("</tr>")
                  Shaded = IIf(Shaded = "shaded", "", "shaded")
                Next
                GraphicPath = ""
              ElseIf CellID > 0 Then
                Send("<tr>")
                Send("<td colspan=""7""><center><i>" & Copient.PhraseLib.Lookup("reward.graphic-nographics", LanguageID) & "</i></center></td>")
                Send("</tr>")
              Else
                Send("<tr>")
                Send("<td colspan=""7""></td>")
                Send("</tr>")
              End If
            %>
          </tbody>
        </table>
        <hr class="hidden" />
      </div>
      <div class="box" id="imagepreview" style="display: <% Sendb(IIf(bPreviewOnly, "block", "none"))%>;">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.imagepreview", LanguageID))%>
          </span>
        </h2>
        <%
          If (Not bPreviewOnly) Then PreviewGraphicPath = GraphicPath
          If (PreviewGraphicPath = "") Then
            Sendb("<img src=""#"" id=""fsImage"" alt=""" & Copient.PhraseLib.Lookup("reward.graphic-fullsizedimage", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("reward.graphic-fullsizedimage", LanguageID) & """ />")
          Else
            Sendb("<img src=""" & PreviewGraphicPath & """ id=""fsImage"" alt=""" & Copient.PhraseLib.Lookup("reward.graphic-fullsizedimage", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("reward.graphic-fullsizedimage", LanguageID) & """ />")
          End If
        %>
      </div>
    </div>
  </div>
</form>

<script runat="server">
  Public Function LoadGraphicPath(ByVal adId As Long) As String
    Const DEFAULT_GRAPHIC_PATH As String = "C:\"
    Dim MyCommon As New Copient.CommonInc
    Dim dst As DataTable
    Dim GraphicPath As String = DEFAULT_GRAPHIC_PATH
    Dim imgExt As String = "jpg"
    Try
      MyCommon.Open_LogixRT()
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
    Catch ex As Exception
      ' do nothing
    Finally
      MyCommon.Close_LogixRT()
    End Try
    Return GraphicPath & CStr(adId) & "img_tn.jpg"
  End Function
</script>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>
<%
done:
  MyCommon.Close_LogixRT()
  If (bPreviewOnly) Then
    Send_BodyEnd()
  Else
    Send_BodyEnd("mainform", "cellselect")
  End If
  Logix = Nothing
  MyCommon = Nothing
%>
