<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: PLU-detail.aspx 
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
  Dim PLU As Decimal = 0
  Dim PLUString As String = ""
  Dim LastUpdate As Date
  Dim RangeBegin As Decimal = 0
  Dim RangeEnd As Decimal = 0
  Dim RangeLocked As Boolean = True
  Dim MultipleOffers As Boolean = False
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  Dim row As DataRow
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "PLU-detail.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  PLU = CDec(MyCommon.Extract_Val(Request.QueryString("PLU")))
  PLUString = Request.QueryString("PLU")
  RangeBegin = CDec(MyCommon.Fetch_SystemOption(198))
  RangeEnd = CDec(MyCommon.Fetch_SystemOption(199))
  RangeLocked = MyCommon.Fetch_CPE_SystemOption(95)
  MultipleOffers = MyCommon.Fetch_CPE_SystemOption(94)
  
  Send_HeadBegin("term.triggercode", , PLUString)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 4)
  
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(1, "perm.admin-configuration")
    GoTo done
  End If
  
  MyCommon.QueryStr = "select LastUpdate from CPE_IncentivePLUs with (NoLock) where PLU='" & PLUString & "' order by LastUpdate DESC;"
  rst = MyCommon.LRT_Select()
  If (rst.Rows.Count > 0) Then
    LastUpdate = MyCommon.NZ(rst.Rows(0).Item("LastUpdate"), "1/1/2000")
  End If
%>
<form action="#" id="mainform" name="mainform">
  <input type="hidden" id="PLU" name="PLU" value="<% Sendb(PLU) %>" />
  <div id="intro">
    <%
      Sendb("<h1 id=""title"">")
      Sendb(Copient.PhraseLib.Lookup("term.triggercode", LanguageID) & " " & PLUString)
      Send("</h1>")
    %>
    <div id="controls">
      <%
        Send("<div class=""actionsmenu"" id=""actionsmenu"">")
        Send("</div>")
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
      <div class="box" id="offers">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <div class="boxscroll">
          <%
            If (Request.QueryString("PLU") <> "") Then
              MyCommon.QueryStr = "select distinct CIP.IncentivePLUID, CIP.RewardOptionID, CIP.PLU, RO.IncentiveID as OfferID, " & _
                                  " I.IncentiveName as Name, I.EligibilityEndDate as ProdEndDate from CPE_IncentivePLUs as CIP with (NoLock) " & _
                                  " inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=CIP.RewardOptionID " & _
                                  " inner join CPE_Incentives as I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                  " where PLU='" & PLUString & "' and RO.Deleted=0 and I.Deleted=0 " & _
                                  " order by IncentiveName;"
              rst = MyCommon.LRT_Select
              If rst.Rows.Count > 0 Then
                For Each row In rst.Rows
                  If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                    Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</a>")
                  Else
                    Sendb(MyCommon.NZ(row.Item("Name"), ""))
                  End If
                  If (MyCommon.NZ(row.Item("ProdEndDate"), Today) < Today) Then
                    Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                  End If
                  Send("<br />")
                Next
              Else
                Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
              End If
            Else
              Send(Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>

    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
    </div>
    <br clear="all" />
  </div>
</form>

<script type="text/javascript">
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  MyCommon = Nothing
  Logix = Nothing
%>
