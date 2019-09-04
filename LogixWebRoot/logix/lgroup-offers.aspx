<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: lgroup-offers.aspx 
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
    Dim dt As DataTable
    Dim row As DataRow
    Dim Name As String = ""
    Dim LocationGroupID As Integer
    Dim OfferCount As Integer = 0
    Dim i As Integer
    Dim OfferName As String = ""
    Dim OfferID As Integer
    Dim EngineID As Integer
    Dim EngineName As String = "CPE"
    Dim DeployStatus As Integer
    Dim OffersToAdd() As String
    Dim OffersToRemove() As String
    Dim OfferList As New StringBuilder()
    Dim bIsErrorMsg As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249) = "1", True, False)



    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "lgroup-offers.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)

    LocationGroupID = Request.QueryString("LocationGroupID")
    MyCommon.QueryStr = "select LG.Name, LG.EngineID, PE.Description as EngineName from LocationGroups LG with (NoLock) " & _
                        "inner join PromoEngines PE with (NoLock) on PE.EngineID=LG.EngineID " & _
                        "where LocationGroupID=" & LocationGroupID & ";"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
        Name = dt.Rows(0).Item("Name")
        EngineID = dt.Rows(0).Item("EngineID")
        EngineName = dt.Rows(0).Item("EngineName")
    End If

    DeployStatus = MyCommon.Extract_Val(Request.QueryString("deployType"))
    If DeployStatus = 0 Then DeployStatus = 1

    If (Request.QueryString("addAvailable") <> "") Then
        OffersToAdd = Request.QueryString.GetValues("availableOffers")
        If (Not OffersToAdd Is Nothing AndAlso OffersToAdd.Length > 0) Then
            For i = 0 To OffersToAdd.GetUpperBound(0)
                OfferID = OffersToAdd(i)
                MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = LocationGroupID
                MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
                MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
                MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                MyCommon.LRTsp.ExecuteNonQuery()
                MyCommon.Close_LRTsp()
                MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferId=" & OfferID & " and LocationGroupID=" & LocationGroupID
                MyCommon.LRT_Execute()
                If EngineID = 2 OrElse EngineID = 9 Then
                    If (DeployStatus = 2 AndAlso MeetsDeploymentReqs(MyCommon, OfferID, EngineID)) Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(),LastDeployValidationMessage='term.validationsuccessful' where IncentiveID=@OfferID"
                        'MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate() where IncentiveID=" & OfferID
                    Else
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1 where IncentiveID=@OfferID"
                    End If
                    MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                Else
                    If (DeployStatus = 2) Then
                        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1 where OfferID=@OfferID"
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                        'MyCommon.LRT_Execute()
                        MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                    Else
                        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1 where offerid=@OfferID"
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                    End If
                End If
                ' MyCommon.LRT_Execute()
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
                MyCommon.Activity_Log(11, LocationGroupID, AdminUserID, Copient.PhraseLib.Lookup("lgroup-edit.addoffers", LanguageID) & ": " & OfferID)
            Next
        End If
    ElseIf (Request.QueryString("removeSelected") <> "") Then
        OffersToRemove = Request.QueryString.GetValues("selectedOffers")
        If (Not OffersToRemove Is Nothing AndAlso OffersToRemove.Length > 0) Then
            For i = 0 To OffersToRemove.GetUpperBound(0)
                OfferList.Append(OffersToRemove(i))
                OfferID = OffersToRemove(i)
                If (i < OffersToRemove.GetUpperBound(0)) Then OfferList.Append(", ")

                MyCommon.QueryStr = "Update OfferLocations with (RowLock) set Deleted=1, LastUpdate=getdate(), StatusFlag=2, TCRMAStatusFlag=3 where " & _
                                    "LocationGroupID = " & LocationGroupID & " and OfferID=" & OfferID & ";"
                MyCommon.LRT_Execute()
                If EngineID = 2 OrElse EngineID = 9 Then
                    If (DeployStatus = 2 AndAlso MeetsDeploymentReqs(MyCommon, OfferID, EngineID)) Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate() where IncentiveID=" & OfferID & ";"
                    Else
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1 where IncentiveID=" & OfferID & ";"
                    End If
                Else
                    If (DeployStatus = 2) Then
                        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1 where OfferID=" & OfferID & ";"
                    Else
                        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1 where offerid=" & OfferID & ";"
                    End If
                End If
                MyCommon.LRT_Execute()
            Next

            MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-deploy", LanguageID))
            MyCommon.Activity_Log(11, LocationGroupID, AdminUserID, Copient.PhraseLib.Lookup("lgroup-edit.removeoffers", LanguageID) & ": " & OfferList.ToString)
        End If
    End If

    Send_HeadBegin("term.storegroup", , LocationGroupID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    Send_BodyBegin(2)

    If (Logix.UserRoles.AccessStoreGroups = False) Then
        Send_Denied(2, "perm.lgroup-access")
        GoTo done
    End If

    Send("<script type=""text/javascript"">")
    Send("  var bSkipUnload = true;")
    Send("")
    Send("  function ChangeParentDocument() { ")
    Send("    if (opener != null && !opener.closed && bSkipUnload != true) {")
    Send("      opener.location = 'lgroup-edit.aspx?LocationGroupID=" & LocationGroupID & "'; ")
    Send("    }")
    Send("  } ")
    Send("")
    Send("  var refresh = true")
    Send("")
    Send("  window.onunload = refreshParent;")
    Send("")
    Send("  function refreshParent() { ")
    Send("    if (refresh) {")
    Send("      window.opener.location.reload(); ")
    Send("    }")
    Send("  }")
    Send("")
    Send("  function noRefresh() {")
    Send("    refresh = false;")
    Send("  }")
    Send("")
    Send("</script>")
%>
<form action="lgroup-offers.aspx" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.storegroup", LanguageID) & " #" & LocationGroupID & ": " & MyCommon.TruncateString(Name, 35))%>
    </h1>
    <div id="controls">
      <input type="hidden" id="OfferID" name="LocationGroupID" value="<% sendb(LocationGroupID) %>" />
      <input type="hidden" id="Name" name="Name" value="<% sendb(Name) %>" />
    </div>
  </div>
  <div id="main">
    <%If (infoMessage <> "" And bIsErrorMsg) Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      ElseIf (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column">
      <div class="box" id="offers">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>
          </span>
        </h2>
        <b><% Sendb(Copient.PhraseLib.Lookup("term.deployment", LanguageID))%> &nbsp <% Sendb(Copient.PhraseLib.Lookup("term.options", LanguageID))%>:</b><br />
        <input type="radio" id="deployType2" name="deployType" value="2" checked="checked" />
        <label for="deployType2"><% Sendb(Copient.PhraseLib.Lookup("lgroup-offers.RedeployOffers", LanguageID))%></label>
        <br />
        <input type="radio" id="deployType1" name="deployType" value="1" />
        <label for="deployType1"><% Sendb(Copient.PhraseLib.Lookup("lgroup-offers.MarkOffersAsModified", LanguageID))%></label>
        <br />
        <br />
        <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.offers", LanguageID))%>">
          <tr>
            <td>
              <b><% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID))%>:</b>
            </td>
            <td>
            </td>
            <td>
              <b><% Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID))%>:</b>
            </td>
          </tr>
          <tr>
            <td>
              <select class="longer" id="availableOffers" name="availableOffers" multiple="multiple" size="10">
                <% 
                    MyCommon.QueryStr = "select OfferID, Name, ProdEndDate from AllOffersListView with (NoLock) " &
                                        "where Deleted=0 and IsTemplate=0 and EngineID in (" & EngineID & IIf(EngineID = 2, ",6", "") & ") and StatusFlag not in (11, 12) "
                    If (bEnableRestrictedAccessToUEOfferBuilder) Then  MyCommon.QueryStr  &= GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"") & " "
                    MyCommon.QueryStr &= " and OfferID Not IN (" &
                                        "  select O.OfferID as OfferID from OfferLocations as OL with (NoLock) " &
                                        "  left join Offers as O with (NoLock) on O.OfferID=OL.OfferID " &
                                        "  where O.Deleted=0 and OL.Deleted=0 and OL.LocationGroupID=" & LocationGroupID &
                                        "   union " &
                                        "  select I.IncentiveID as OfferID from OfferLocations OL with (NoLock) " &
                                        "  inner join CPE_Incentives I with (NoLock) on OL.OfferID=I.IncentiveID " &
                                        "  inner join OfferIDs OI with (NoLock) on OI.OfferID=I.IncentiveID " &
                                        "  where OL.LocationGroupID=" & LocationGroupID & " and OL.Deleted=0 and I.Deleted=0 and I.StatusFlag <> 11 and I.StatusFlag <> 12 "
                    If (bEnableRestrictedAccessToUEOfferBuilder) Then  MyCommon.QueryStr  &= GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"I") & " "
                    MyCommon.QueryStr &=" ) order by Name ASC;"

                    dt = MyCommon.LRT_Select
                    OfferCount = dt.Rows.Count
                    If (OfferCount > 0) Then
                        For Each row In dt.Rows
                            OfferName = MyCommon.NZ(row.Item("Name"), "")
                            OfferName = IIf(OfferName.Trim = "", "(" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & row.Item("OfferID") & ")", OfferName)
                            If (MyCommon.NZ(row.Item("ProdEndDate"), Today) < Today) Then OfferName += " (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")"
                            Send("<option value=""" & row.Item("OfferID") & """>" & row.Item("OfferID") & ": " & OfferName & "</option>")
                        Next
                    End If
                %>
              </select>
            </td>
            <td align="center">
              <input type="submit" class="arrowadd" id="addAvailable" name="addAvailable" value="&#187;" /><br />
              <input type="submit" class="arrowrem" id="removeSelected" name="removeSelected" value="&#171;" /><br />
              <br />
            </td>
            <td>
              <select class="longer" id="selectedOffers" name="selectedOffers" multiple="multiple" size="10">
                <% 
                    MyCommon.QueryStr = "select 1 as EngineID, O.Name as Name,O.OfferID as OfferID,O.ProdEndDate from OfferLocations as OL with (NoLock) left join Offers as O with (NoLock) on O.OfferID=OL.OfferID " &
                                  " where O.Deleted=0 and OL.Deleted=0 and OL.LocationGroupID=" & LocationGroupID &
                                  " union " &
                                  "select OI.EngineID, I.IncentiveName as Name, I.IncentiveID as OfferID, I.EndDate as ProdEndDate from OfferLocations OL with (NoLock) " &
                                  "inner join CPE_Incentives I with (NoLock) on OL.OfferID = I.IncentiveID " &
                                  "inner join OfferIDs OI with (NoLock) on OI.OfferID = I.IncentiveID " &
                                  "where OL.LocationGroupID=" & LocationGroupID & " and OL.Deleted=0 and I.Deleted=0 and I.StatusFlag <> 11 and I.StatusFlag <> 12"
                    If (bEnableRestrictedAccessToUEOfferBuilder) Then MyCommon.QueryStr &=GetRestrictedAccessToUEBuilderQuery(MyCommon,Logix,"I") & " "
                    MyCommon.QueryStr &= " order by Name;"

                    dt = MyCommon.LRT_Select
                    OfferCount = dt.Rows.Count
                    If (OfferCount > 0) Then
                        For Each row In dt.Rows
                            OfferName = MyCommon.NZ(row.Item("Name"), "")
                            OfferName = IIf(OfferName.Trim = "", Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & row.Item("OfferID"), OfferName)
                            Send("<option value=""" & row.Item("OfferID") & """>" & row.Item("OfferID") & ": " & OfferName & "</option>")
                        Next
                    End If
                %>
              </select>
            </td>
          </tr>
        </table>
        <hr class="hidden" />
      </div>
    </div>
  </div>
</form>
<script runat="server">
    Function MeetsDeploymentReqs(ByRef MyCommon As Copient.CommonInc, ByVal OfferID As Integer, ByVal EngineID As Integer) As Boolean
        Dim bMeetsReqs As Boolean = False
        ' The user wants to deploy, so do a quick check for at least one assigned offer location and terminal,
        ' and ensure that there are no unassigned tier values
        If EngineID = 2 Then
            MyCommon.QueryStr = "dbo.pa_CPE_IsOfferDeployable"
        ElseIf EngineID = 9 Then
            MyCommon.QueryStr = "dbo.pa_UE_IsOfferDeployable"
        End If
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@IsDeployable", SqlDbType.Bit).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        bMeetsReqs = MyCommon.LRTsp.Parameters("@IsDeployable").Value
        Return bMeetsReqs
    End Function
</script>
<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd()
  Logix = Nothing
  MyCommon = Nothing
%>
