<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-sharedselect.aspx 
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
  Dim OfferID As Long
  Dim OfferName As String
  Dim rst As DataTable
  Dim row As DataRow
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-sharedselect.aspx"
  ' open database connection
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = Request.QueryString("OfferID")
  OfferName = Request.QueryString("Name")
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  MyCommon.QueryStr = "select ExtOfferID,OfferID,Name,DistPeriod,DistPeriodLimit from offers with (NoLock) where SharedLimitID=0 and visible=1 and deleted=0 and not OfferID=" & OfferID & " order by Name"
  rst = MyCommon.LRT_Select()
  
  Send_HeadBegin("term.offer", "term.select", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
function ChangeParentDocument(url) { 
    opener.location = url; 
    window.close()
    }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.AccessOffers = False) Then
    Send_Denied(2, "perm.offers-access")
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
<div id="intro">
  <h1 id="title">
    <%  Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferID & " " & Copient.PhraseLib.Lookup("term.sharedlimits", LanguageID))%>
  </h1>
</div>
<div id="main">
  <form action="#" id="mainform" name="mainform">
    <% Sendb(Copient.PhraseLib.Lookup("offer-sharedselect.main", LanguageID))%>
    <br class="half" />
    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID)) %>" style="width: 100%;">
      <thead>
        <tr>
          <th align="left" class="th-xid" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID))%>
          </th>
          <th align="left" class="th-id" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </th>
          <th align="left" class="th-name" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </th>
          <th align="left" class="th-limits" scope="col">
            <% Sendb(Copient.PhraseLib.Lookup("term.limits", LanguageID))%>
          </th>
        </tr>
      </thead>
      <tbody>
        <%
          For Each row In rst.Rows
            Send("<tr class=""" & Shaded & """>")
            Send("  <td>" & row.Item("ExtOfferID") & "</td>")
            Send("  <td>" & row.Item("OfferID") & "</td>")
            Send("  <td><a href=""JavaScript:ChangeParentDocument('offer-gen.aspx?mode=addShared&amp;OfferID=" & OfferID & "&amp;ID=" & row.Item("OfferID") & "')""  >" & row.Item("Name") & Copient.PhraseLib.Lookup("offer-sharedselect.SampleOffer", LanguageID) & "</a></td>")
            Sendb("  <td>" & row.Item("distPeriodLimit"))
            Sendb(" " & Copient.PhraseLib.Lookup("term.per", LanguageID) & " ")
            If Not IsDBNull(row.Item("distPeriod")) Then
              If (row.Item("distPeriod") > 0) Then
                Sendb(row.Item("distPeriod") & " " & Copient.PhraseLib.Lookup("term.days", LanguageID))
              Else
                Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID))
              End If
            End If
            Send("  </td>")
            Send("</tr>")
            If Shaded = "shaded" Then
              Shaded = ""
            Else
              Shaded = "shaded"
            End If
          Next
        %>
      </tbody>
    </table>
  </form>
</div>
<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
