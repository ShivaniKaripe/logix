<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-rew-bins.aspx 
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
  Dim drsTender() As DataRow
  Dim drTender As DataRow
  Dim drNone As DataRow
  Dim row As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim RewardID As String
  Dim NumRanges As Integer
  Dim LinkID As Long
  Dim RewardAmountTypeID As Integer
  Dim TriggerQty As Integer
  Dim ApplyToLimit As Integer
  Dim DoNotItemDistribute As Boolean
  Dim TransactionLevelSelected As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim maxNumRanges As Integer = 30
  Dim sUp As String
  Dim sDown As String
  Dim sDelete As String
  Dim DistPeriod As Integer
  Dim UseSpecialPricing As Boolean
  Dim SPRepeatAtOccur As Integer
  Dim Tiered As Integer
  Dim SponsorID As Integer
  Dim PromoteToTransLevel As Boolean
  Dim RewardLimit As Integer
  Dim RewardLimitTypeID As Integer
  Dim Disallow_Edit As Boolean = True
  Dim FromTemplate As Boolean
  Dim IsTemplate As Boolean = False
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim l As Integer
  Const sBinColDelim As String = ";"
  Const sBinRowDelim As String = ":"
  Dim sXml As String
  Dim sCellName As String
  Dim sRow As String
  Dim sCellValue As String
  Dim sCurrentBegin As String
  Dim sPreviousEnd As String
  Dim sRows() As String
  Dim sCols() As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
 
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  OfferID = Request.QueryString("OfferID")
  RewardID = Request.QueryString("RewardID")
  NumRanges = Request.QueryString("NumRanges")
  
  MyCommon.AppName = "offer-rew-bins.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = rst.Rows(0).Item("Name")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
  End If
  
  ' Get list of tenders
  MyCommon.QueryStr = "select ExtTendercode, Description from TenderTypes with (NoLock) order by Description"
  rst = MyCommon.LRT_Select
  drNone = rst.NewRow()
  drNone.Item(0) = ""
  drNone.Item(1) = "None"
  rst.Rows.Add(drNone)
  drsTender = rst.Select("len(ExtTenderCode) = 0 or (len(ExtTenderCode)= 4 and substring(ExtTenderCode,3,2) <> '00')")
  If drsTender.Length = 1 Then
    infoMessage = Copient.PhraseLib.Lookup("offer-rew-bins.NoValidTenders", LanguageID)
  End If
  Response.Expires = 0
  Send_HeadBegin("term.offer", "term.xmlbinrangesreward", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
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
  
  ' we need to determine our linkid for updates and tiered
  MyCommon.QueryStr = "select LinkID,Tiered,SponsorID,PromoteToTransLevel,RewardDistPeriod,RewardLimit,RewardLimitTypeID,TriggerQty,RewardAmountTypeID, " & _
   "UseSpecialPricing, SPRepeatAtOccur,ApplyToLimit,DoNotItemDistribute from OfferRewards with (NoLock) where RewardID=" & RewardID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    DistPeriod = MyCommon.NZ(row.Item("RewardDistPeriod"), 0)
    LinkID = row.Item("LinkID")
    Tiered = MyCommon.NZ(row.Item("Tiered"), 0)
    SponsorID = MyCommon.NZ(row.Item("SponsorID"), 0)
    PromoteToTransLevel = MyCommon.NZ(row.Item("PromoteToTransLevel"), 0)
    RewardLimit = MyCommon.NZ(row.Item("RewardLimit"), 0)
    RewardLimitTypeID = MyCommon.NZ(row.Item("RewardLimitTypeID"), 2)
    RewardAmountTypeID = MyCommon.NZ(row.Item("RewardAmountTypeID"), 1)
    TriggerQty = MyCommon.NZ(row.Item("TriggerQty"), 1)
    ApplyToLimit = MyCommon.NZ(row.Item("ApplyToLimit"), 1)
    UseSpecialPricing = MyCommon.NZ(row.Item("UseSpecialPricing"), 0)
    SPRepeatAtOccur = MyCommon.NZ(row.Item("SPRepeatAtOccur"), 1)
    DoNotItemDistribute = row.Item("DoNotItemDistribute")
  Next
  
  MyCommon.QueryStr = "select OFR.RewardID,Tiered,O.Numtiers,O.OfferID,XT.TierLevel,XT.XmlText from OfferRewards as OFR with (NoLock) " & _
                      "left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID left join RewardXmlTiers as XT with (NoLock) on OFR.RewardID=XT.RewardID where OFR.RewardID=" & RewardID
  rst = MyCommon.LRT_Select()
  sXml = MyCommon.NZ(rst.Rows(0).Item("XmlText"), sBinColDelim & sBinColDelim)
  If (Request.QueryString("save") <> "") Then
    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
      End If
      MyCommon.QueryStr = "update OfferRewards with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
      " where RewardID=" & RewardID
      MyCommon.LRT_Execute()
    End If
    ' NO tiers for bin Ranges
    If (Tiered = 0) Then
      sXml = ""
      sPreviousEnd = "0"
      sCurrentBegin = "0"
      infoMessage = ""
      For i = 1 To NumRanges
        For j = 1 To 3
          sCellName = "R" & i & "C" & j
          sCellValue = Request.QueryString(sCellName)
          If j = 1 Then
            If (sCellValue = "" Or (Not IsNumeric(sCellValue))) Then
              If infoMessage = "" Then
                infoMessage = Copient.PhraseLib.Lookup("term.xmlnumeric", LanguageID) & " " & Copient.PhraseLib.Lookup("term.xmlbegin", LanguageID)
              End If
              sCurrentBegin = "0"
            ElseIf (Int64.Parse(sPreviousEnd) >= Int64.Parse(sCellValue)) Then
              If infoMessage = "" Then
                infoMessage = Copient.PhraseLib.Lookup("term.xmlend", LanguageID) & " (" & sPreviousEnd & ") " & Copient.PhraseLib.Lookup("term.xmlgreater", LanguageID) & " " & Copient.PhraseLib.Lookup("term.xmlbegin", LanguageID) & " (" & sCellValue & ")"
              End If
              sCurrentBegin = "0"
            Else
              sCurrentBegin = sCellValue
            End If
          ElseIf j = 2 Then
            If (sCellValue = "" Or (Not IsNumeric(sCellValue))) Then
              If infoMessage = "" Then
                infoMessage = Copient.PhraseLib.Lookup("term.xmlnumeric", LanguageID) & " " & Copient.PhraseLib.Lookup("term.xmlend", LanguageID)
              End If
              sPreviousEnd = "0"
            ElseIf (Int64.Parse(sCurrentBegin) >= Int64.Parse(sCellValue)) Then
              If infoMessage = "" Then
                infoMessage = Copient.PhraseLib.Lookup("term.xmlbegin", LanguageID) & " (" & sCurrentBegin & ") " & Copient.PhraseLib.Lookup("term.xmlgreater", LanguageID) & " " & Copient.PhraseLib.Lookup("term.xmlend", LanguageID) & " (" & sCellValue & ")"
              End If
              sPreviousEnd = "0"
            Else
              sPreviousEnd = sCellValue
            End If
          ElseIf j = 3 Then
            If sCellValue = "" And infoMessage = "" Then
              infoMessage = Copient.PhraseLib.Lookup("term.xmlvalid", LanguageID) & " " & Copient.PhraseLib.Lookup("term.tender", LanguageID)
            End If
          End If
          sXml += sCellValue
          If j <> 3 Then
            sXml += sBinColDelim
          End If
        Next
        If i <> NumRanges Then
          sXml += sBinRowDelim
        End If
      Next
      If infoMessage = "" Then
        MyCommon.QueryStr = "dbo.pt_XmlTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@RewardID", SqlDbType.BigInt).Value = RewardID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@XmlText", SqlDbType.NVarChar, 4000).Value = sXml
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()

        MyCommon.QueryStr = "update OfferRewards with (RowLock) set TCRMAStatusFlag=3,CMOAStatusFlag=2 where RewardID=" & RewardID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.xmlbinranges", LanguageID))
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
        MyCommon.LRT_Execute()
        
        CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")

      End If
    End If
  Else
    For k = 1 To NumRanges
      sUp = "^" & k
      sDown = "v" & k
      sDelete = "x" & k
      If (Request.QueryString(sUp) <> "") Then
        sXml = ""
        For i = 1 To NumRanges
          If i = k Then
            sXml += sBinColDelim & sBinColDelim & sBinRowDelim
          End If
          For j = 1 To 3
            sCellName = "R" & i & "C" & j
            sCellValue = Request.QueryString(sCellName)
            sXml += sCellValue
            If j <> 3 Then
              sXml += sBinColDelim
            End If
          Next
          If i <> NumRanges Then
            sXml += sBinRowDelim
          End If
        Next
        Exit For
      ElseIf (Request.QueryString(sDown) <> "") Then
        sXml = ""
        For i = 1 To NumRanges
          For j = 1 To 3
            sCellName = "R" & i & "C" & j
            sCellValue = Request.QueryString(sCellName)
            sXml += sCellValue
            If j <> 3 Then
              sXml += sBinColDelim
            End If
          Next
          If i = k Then
            sXml += sBinRowDelim & sBinColDelim & sBinColDelim
          End If
          If i <> NumRanges Then
            sXml += sBinRowDelim
          End If
        Next
        Exit For
      ElseIf (Request.QueryString(sDelete) <> "") Then
        sXml = ""
        l = 0
        For i = 1 To NumRanges
          If i <> k Then
            l += 1
            For j = 1 To 3
              sCellName = "R" & i & "C" & j
              sCellValue = Request.QueryString(sCellName)
              sXml += sCellValue
              If j <> 3 Then
                sXml += sBinColDelim
              End If
            Next
            If l <> NumRanges - 1 Then
              sXml += sBinRowDelim
            End If
          End If
        Next
        If sXml = "" Then
          sXml = sBinColDelim & sBinColDelim
        End If
        Exit For
      End If
    Next
  End If
  
  If (IsTemplate Or FromTemplate) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Disallow_Edit from OfferRewards with (NoLock) where RewardID=" & RewardID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      For Each row In rst.Rows
        ' ok there are some rows for the template
        Disallow_Edit = MyCommon.NZ(row.Item("Disallow_Edit"), True)
      Next
    End If
  End If
  
  Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("    opener.location = 'offer-rew.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("  }")
  Send("</script>")
%>
<form action="offer-rew-bins.aspx" id="mainform" name="mainform">
  <div id="intro">
    <input type="hidden" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" name="Name" value="<% sendb(Name) %>" />
    <input type="hidden" name="RewardID" value="<% sendb(RewardID) %>" />
    <input type="hidden" name="IsTemplate" value="<% 
    if(istemplate)then 
    sendb("IsTemplate")
    else 
    sendb("Not") 
    end if
%>" />
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & Copient.PhraseLib.Lookup("term.xmlbinrangesreward", LanguageID))%>
    </h1>
    <%If (IsTemplate) Then%>
    <span class="temp2">
      <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
      <label for="temp-employees"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
    </span>
    <%End If%>
    <div id="controls">
      <% If Not (IsTemplate) Then
           If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then Send_Save()
         Else
           If (Logix.UserRoles.EditTemplates) Then Send_Save()
         End If    
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column">
      <div class="box" id="groups">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.xmlbinranges", LanguageID))%>
          </span>
        </h2>
        <table id="bins" summary="<% Sendb(Copient.PhraseLib.Lookup("term.xmlbinranges", LanguageID))%>">
          <thead>
            <tr>
              <th align="left" class="th-reorder" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.insert", LanguageID))%>
              </th>
              <th align="left" class="th-del" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID))%>
              </th>
              <th align="left" class="th-longid" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.xmlbegin", LanguageID))%>
              </th>
              <th align="left" class="th-longid" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.xmlend", LanguageID))%>
              </th>
              <th align="left" class="th-tender" scope="col">
                <% Sendb(Copient.PhraseLib.Lookup("term.tender", LanguageID))%>
              </th>
            </tr>
          </thead>
          <tbody>
            <%
              sRows = sXml.Split(sBinRowDelim)
              i = 0
              NumRanges = sRows.Length
              For Each sRow In sRows
                i = i + 1
                sUp = "^" & i
                sDown = "v" & i
                sDelete = "x" & i
                sCols = sRow.Split(sBinColDelim)
            %>
            <tr class="shaded">
              <td>
                <%
                  Sendb("<input class=""up"" id=""up"" name=""" & sUp & """ title=""" & Copient.PhraseLib.Lookup("term.up", LanguageID) & """ ")
                  If NumRanges < maxNumRanges Then
                    Sendb(" type=""submit"" value=""▲"" />")
                  Else
                    Sendb(" type=""submit"" value=""▲"" disabled=""disabled"" />")
                  End If
                  Sendb("<input class=""down"" id=""down"" name=""" & sDown & """ title=""" & Copient.PhraseLib.Lookup("term.down", LanguageID) & """ ")
                  If NumRanges < maxNumRanges Then
                    Send(" type=""submit"" value=""▼"" />")
                  Else
                    Send(" type=""submit"" value=""▼"" disabled=""disabled"" />")
                  End If
                %>
              </td>
              <td>
                <% Sendb("<input class=""ex"" type=""submit"" value=""X"" name=""" & sDelete & """ title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ />")%>
              </td>
              <%
                j = 0
                For Each sCellValue In sCols
                  j = j + 1
                  sCellName = "R" & i & "C" & j
                  If j = 3 Then
                    Send("<td>")
                    Send("<select class=""long"" id=""" & sCellName & """ name= """ & sCellName & """>")
                    For Each drTender In drsTender
                      If sCellValue = drTender.Item(0) Then
                        Sendb("<option value=""" & drTender.Item(0) & """ selected=""selected"">" & drTender.Item(1) & "</option>")
                      Else
                        Sendb("<option value=""" & drTender.Item(0) & """>" & drTender.Item(1) & "</option>")
                      End If
                    Next
                    Send("</select>")
                    Send("</td>")
                  Else
                    Send("<td>")
                    Send("<input class=""mediumshort"" id=""" & sCellName & """ name=""" & sCellName & """ type=""text"" value=""" & sCellValue & """ maxlength=""12"" />")
                    Send("</td>")
                  End If
                Next
              %>
            </tr>
            <% Next%>
          </tbody>
        </table>
        <% Send("<input type=""hidden"" name=""NumRanges"" value=""" & NumRanges & """ />")%>
        <br class="half" />
        <br />
      </div>
    </div>
  </div>
  
<script type="text/javascript">
  <% If (CloseAfterSave) Then %>
    <% If (infoMessage = "") Then %>
        window.close();
    <% End If %>
  <% End If %>
</script>
  
</form>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "R1C1")
  Logix = Nothing
  MyCommon = Nothing
%>
