<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-con-tender.aspx 
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
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim tmpString As String
  Dim NumTiers As Integer
  Dim Tiered As Integer
  Dim Disallow_Edit As Boolean = True
  Dim bUseTemplateLocks As Boolean
  Dim IsTemplate As Boolean = False
  Dim CloseAfterSave As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim bDisallowEditValue As Boolean = False
  Dim bDisallowEditRewards As Boolean = False
  Dim bDisallowEditPp As Boolean = False
  Dim sDisabled As String
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-con-tender.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Tiered = False
  OfferID = Request.QueryString("OfferID")
  ConditionID = Request.QueryString("ConditionID")
  NumTiers = Request.QueryString("NumTiers")
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  ' fetch the name
  MyCommon.QueryStr = "Select Name,IsTemplate,FromTemplate from Offers with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Name = rst.Rows(0).Item("Name")
    IsTemplate = MyCommon.NZ(rst.Rows(0).Item("IsTemplate"), False)
    If IsTemplate Then
      bUseTemplateLocks = False
    Else
      bUseTemplateLocks = MyCommon.NZ(rst.Rows(0).Item("FromTemplate"), False)
    End If
  End If
  
  If (IsTemplate Or bUseTemplateLocks) Then
    ' lets dig the permissions if its a template
    MyCommon.QueryStr = "select Tiered,Disallow_Edit,DisallowEdit1,DisallowEdit2,DisallowEdit3,DisallowEdit4 from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("Disallow_Edit"), True)
      bDisallowEditPp = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit1"), False)
      bDisallowEditValue = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit2"), False)
      bDisallowEditRewards = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit3"), False)
      Tiered = MyCommon.NZ(rst.Rows(0).Item("Tiered"), False)
      If Tiered Then
        bDisallowEditRewards = True
      End If
      If bUseTemplateLocks Then
        If Disallow_Edit Then
          bDisallowEditPp = True
          bDisallowEditValue = True
          bDisallowEditRewards = True
        Else
          Disallow_Edit = bDisallowEditPp And bDisallowEditValue And bDisallowEditRewards
        End If
      End If
    End If
  End If

  Send_HeadBegin("term.offer", "term.tendercondition", OfferID)
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
  
  If (Request.QueryString("save") <> "") Then
    
    If Not (bUseTemplateLocks And bDisallowEditPp) Then
      ' delete all the existing tender types for this condition
      MyCommon.QueryStr = "delete from ConditionTenderTypes with (RowLock) where ConditionID=" & ConditionID
      MyCommon.LRT_Execute()
      ' lets check those tender inputs
      MyCommon.QueryStr = "select TenderTypeID,Description from TenderTypes with (NoLock)"
      rst = MyCommon.LRT_Select
      For Each row In rst.Rows
        tmpString = "t" & row.Item("TenderTypeID")
        If (Request.QueryString(tmpString) = "on") Then
          MyCommon.QueryStr = "insert into ConditionTenderTypes with (RowLock) (ConditionID,TenderTypeID) values(" & ConditionID & " ," & row.Item("TenderTypeID") & ")"
          MyCommon.LRT_Execute()
        End If
      Next
    End If
    
    If Not (bUseTemplateLocks And bDisallowEditRewards) Then
      If (Request.QueryString("granted") <> "") Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set GrantTypeID=" & Request.QueryString("granted") & " where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If
    End If
    
    If Not (bUseTemplateLocks And bDisallowEditValue) Then
      If (Request.QueryString("valuetype") <> "") Then
        MyCommon.QueryStr = "update OfferConditions with (RowLock) set QtyUnitType=2 where ConditionID=" & ConditionID & ";"
        MyCommon.LRT_Execute()
      End If
      If (Request.QueryString("tier0") <> "" And Request.QueryString("Tiered") = "False") Then
        ' delete the current tier ammounts
        MyCommon.QueryStr = "delete from ConditionTiers with (RowLock) where ConditionID=" & ConditionID
        MyCommon.LRT_Execute()
        'dbo.pt_ConditionTiers_Update @ConditionID bigint, @TierLevel int,@AmtRequired decimal(12,3)
        MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
        MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = 0
        MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier0"))
        If (MyCommon.Extract_Val(Request.QueryString("tier0")) < 0) Then
          infoMessage = Copient.PhraseLib.Lookup("condition.badvalue", LanguageID)
        End If
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      ElseIf (Request.QueryString("Tiered") = "True") Then
        ' delete the current tier ammounts
        MyCommon.QueryStr = "delete from ConditionTiers with (RowLock) where ConditionID=" & ConditionID
        MyCommon.LRT_Execute()
        Dim x As Integer
        For x = 1 To NumTiers
          'Response.Write("->@ConditionID: " & ConditionID & " @TierLevel: " & x & " @AmtRequired: " & Request.QueryString("tier" & x) & "<br />")
          MyCommon.QueryStr = "dbo.pt_ConditionTiers_Update"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = MyCommon.Extract_Val(ConditionID)
          MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = x
          MyCommon.LRTsp.Parameters.Add("@AmtRequired", SqlDbType.Decimal).Value = MyCommon.Extract_Val(Request.QueryString("tier" & x))
          If (x > 1) And Int((MyCommon.Extract_Val(Request.QueryString("tier" & x))) < Int(MyCommon.Extract_Val(Request.QueryString("tier" & (x - 1))))) Then
            infoMessage = Copient.PhraseLib.Lookup("condition.tiervalues", LanguageID)
          ElseIf (MyCommon.Extract_Val(Request.QueryString("tier" & x)) < 0) Then
            infoMessage = Copient.PhraseLib.Lookup("condition.badvalue", LanguageID)
          End If
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        Next
      End If
    End If

    If (Request.QueryString("IsTemplate") = "IsTemplate") Then
      ' time to update the status bits for the templates
      Dim form_Disallow_Edit As Integer = 0
      Dim iDisallowEditPp As Integer = 0
      Dim iDisallowEditValue As Integer = 0
      Dim iDisallowEditRewards As Integer = 0
      
      Disallow_Edit = False
      bDisallowEditValue = False
      bDisallowEditRewards = False
      bDisallowEditPp = False

      If (Request.QueryString("Disallow_Edit") = "on") Then
        form_Disallow_Edit = 1
        Disallow_Edit = True
      End If
      If (Request.QueryString("DisallowEditPp") = "on") Then
        iDisallowEditPp = 1
        bDisallowEditPp = True
      End If
      If (Request.QueryString("DisallowEditValue") = "on") Then
        iDisallowEditValue = 1
        bDisallowEditValue = True
      End If
      If (Request.QueryString("DisallowEditRewards") = "on") Then
        iDisallowEditRewards = 1
        bDisallowEditRewards = True
      End If

      MyCommon.QueryStr = "update OfferConditions with (RowLock) set Disallow_Edit=" & form_Disallow_Edit & _
      ",DisallowEdit1=" & iDisallowEditPp & _
      ",DisallowEdit2=" & iDisallowEditValue & _
      ",DisallowEdit3=" & iDisallowEditRewards & _
      " where ConditionID=" & ConditionID
      MyCommon.LRT_Execute()
    End If
    
    MyCommon.QueryStr = "update OfferConditions with (RowLock) set TCRMAStatusFlag=2,CRMAStatusFlag=2 where ConditionID=" & ConditionID
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where offerid=" & OfferID
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-tender", LanguageID))
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
  End If
  
  Send("<script type=""text/javascript"">")
    Send("function ChangeParentDocument() { ")
    Send("  if (opener != null) {")
    Send("    var newlocation = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
  Send("    opener.location = 'offer-con.aspx?OfferID=" & OfferID & "'; ")
    Send("    } ")
    Send("  }")
    Send("}")
  Send("</script>")
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <input type="hidden" id="OfferID" name="OfferID" value="<% sendb(OfferID) %>" />
    <input type="hidden" id="ConditionID" name="ConditionID" value="<% sendb(ConditionID) %>" />
    <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
            if(istemplate)then 
            sendb("IsTemplate")
            else 
            sendb("Not") 
            end if
        %>" />
    <%If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.tendercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.tendercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="temp-employees" name="Disallow_Edit"
          <% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="temp-employees">
          <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
        </label>
      </span>
      <% End If%>
      <% If Not (IsTemplate) Then
           If (Logix.UserRoles.EditOffer And Not (bUseTemplateLocks And Disallow_Edit)) Then Send_Save()
         Else
           If (Logix.UserRoles.EditTemplates) Then Send_Save()
         End If    
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="types">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.types", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditPp1" name="DisallowEditPp" <% if(bDisallowEditPp)then send(" checked=""checked""") %> />
            <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
          </span>
          <% ElseIf (bUseTemplateLocks And bDisallowEditPp) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="DisallowEditPp2" name="DisallowEditPp" disabled="disabled" checked="checked" />
            <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
          </span>
          <% End If%>
        </h2>
        <%
          Dim checkBox As String
          MyCommon.QueryStr = " select TT.TenderTypeID, TT.Description, isnull(CTT.PKID, 0) as InUse from TenderTypes as TT with (NoLock) Left Join ConditionTenderTypes as CTT with (NoLock) on TT.TenderTypeID=CTT.TenderTypeID and CTT.ConditionID=" & ConditionID & " order by Description"
          rst = MyCommon.LRT_Select
          If (bUseTemplateLocks And bDisallowEditPp) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
          For Each row In rst.Rows
            If (row.Item("InUse") <> 0) Then checkBox = " checked=""checked""" Else checkBox = ""
            Send("<input class=""checkbox"" id=""t" & row.Item("TenderTypeID") & """ name=""t" & row.Item("TenderTypeID") & """ type=""checkbox""" & checkBox & sDisabled & " /><label for=""t" & row.Item("TenderTypeID") & """>" & MyCommon.SplitNonSpacedString(row.Item("Description"), 25) & "</label><br />")
            checkBox = ""
          Next
        %>
        <hr class="hidden" />
      </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <div class="box" id="value">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.value", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate or (bUseTemplateLocks and bDisallowEditValue)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditValue" name="DisallowEditValue"<% if(bDisallowEditValue)then send(" checked=""checked""") %> <% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %>/>
          <label for="temp-Tiers"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
        </span>
        <br class="printonly" />
        <% End If%>
        <label for="tier0">
          <% Sendb(Copient.PhraseLib.Lookup("condition.valueneeded", LanguageID))%>
        </label>
        <br />
        <%
          MyCommon.QueryStr = "select LinkID,Tiered,O.Numtiers,QtyUnitType,O.OfferID,CT.TierLevel,CT.AmtRequired from OfferConditions as OC with (NoLock) left join Offers as O with (NoLock) on O.OfferID=OC.OfferID left  join ConditionTiers as CT with (NoLock) on OC.ConditionID=CT.ConditionID where OC.ConditionID=" & ConditionID
          rst = MyCommon.LRT_Select()
          Dim q As Integer
          q = 1
          If (bUseTemplateLocks And bDisallowEditValue) Then
            sDisabled = " disabled=""disabled"""
          Else
            sDisabled = ""
          End If
          For Each row In rst.Rows
            If (row.Item("Tiered") = 0) Then
              Send("<input id=""tier0"" name=""tier0"" type=""text"" size=""6"" maxlength=""9"" value=""" & row.Item("AmtRequired") & """" & sDisabled & " /><br />")
            Else
              Send("<label for=""tier" & q & """><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & q & ":</b></label>")
              Send("<input id=""tier" & q & """ name=""tier" & q & """ type=""text"" size=""6"" value=""" & row.Item("AmtRequired") & """" & sDisabled & " /><br />")
              Tiered = True
            End If
            q = q + 1
          Next
          Send("<input type=""hidden"" name=""NumTiers"" value=""" & MyCommon.NZ(row.Item("NumTiers"), "") & """ />")
          Send("<input type=""hidden"" name=""Tiered"" value=""" & MyCommon.NZ(row.Item("Tiered"), "") & """ />")
          MyCommon.QueryStr = "select LinkID,ExcludedID,PointsRedeemInstant,MinOrderItemsOnly,GrantTypeID,DoNotItemDistribute from OfferConditions with (NoLock) where ConditionID=" & ConditionID & ";"
          rst = MyCommon.LRT_Select
          For Each row In rst.Rows
          Next
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="grants" <% if(tiered) then sendb("style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.rewards", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate Or IsTemplate Or (bUseTemplateLocks And bDisallowEditRewards)) Then%>
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="DisallowEditRewards" name="DisallowEditRewards"
            <% if(bDisallowEditRewards)then send(" checked=""checked""") %><% If(bUseTemplateLocks) Then sendb(" disabled=""disabled""") %> />
          <label for="temp-Tiers">
            <% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%>
          </label>
        </span>
        <br class="printonly" />
        <% End If%>
        <% Sendb(Copient.PhraseLib.Lookup("condition.rewardsgranted", LanguageID))%>
        <br />
        <input class="radio" id="eachtime" name="granted" value="3" type="radio" <% if(MyCommon.NZ(row.item("granttypeid"), 0)=3)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="eachtime">
          <% Sendb(Copient.PhraseLib.Lookup("condition.eachtime", LanguageID))%>
        </label>
        <br />
        <input class="radio" id="equalto" name="granted" value="1" type="radio" <% if(MyCommon.NZ(row.item("granttypeid"),0)=1)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="equalto">
          <% Sendb(Copient.PhraseLib.Lookup("condition.equalto", LanguageID))%>
        </label>
        <br />
        <input class="radio" id="greaterthan" name="granted" value="2" type="radio" <% if(row.item("granttypeid")=2)then sendb(" checked=""checked""") %><% If(bUseTemplateLocks and bDisallowEditRewards) Then sendb(" disabled=""disabled""") %> />
        <label for="greaterthan">
          <% Sendb(Copient.PhraseLib.Lookup("condition.greaterthan", LanguageID))%>
        </label>
      </div>
    </div>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% End If %>
</script>

<%
done:
  MyCommon.Close_LogixRT()
  If Tiered Then
    Send_BodyEnd("mainform", "tier1")
  Else
    Send_BodyEnd("mainform", "tier0")
  End If
  MyCommon = Nothing
  Logix = Nothing
%>
