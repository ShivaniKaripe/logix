<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEaccum-adjust.aspx 
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
  
%>
<script runat="server">
  Dim MyCommon As New Copient.CommonInc
</script>
<%
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
    Dim Logix As New Copient.LogixInc
  Dim CustUpdateRA As New Copient.CustomerUpdate(MyCommon)
  Dim rst2 As DataTable
  Dim CustomerExtID As String = ""
  Dim OfferID As Long
  Dim KeyCt As Integer = 0
  Dim InfoMsgCode As Integer = 0
  Dim OfferDesc As String = ""
  Dim OfferName As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim CardPK As Long = 0
  Dim CustomerPK As Long = 0
  Dim SessionID As String = ""
  Dim IsUSAirMiles As Boolean = False
  Dim AdjustPermitted As Boolean = False
  Dim HHPK As Long = 0
  Dim HHCardPK As Long = 0

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Const REWARD_ALREADY_EARNED As Integer = 1
  Const REWARD_NOT_RECEIVED As Integer = 2
  Const ADJUSTMENT_AMT_NOT_SENT As Integer = 3
  Const ADJUSTMENT_IMPROPER_FORMAT As Integer = 4
  Const ERROR_DURING_ADJUST As Integer = 5
  Const ADJUSTMENT_BELOW_ZERO As Integer = 6

  Response.Expires = 0
  MyCommon.AppName = "CPEaccum-adjust.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  CustomerExtID = Request.QueryString("CustomerExtId")
  OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
  CardPK = MyCommon.Extract_Val(Request.QueryString("CardPK"))
  CustomerPK = MyCommon.Extract_Val(Request.QueryString("CustomerPK"))
  
  If Session("SessionID") IsNot Nothing AndAlso Session("SessionID").ToString.Trim <> "" Then
    SessionID = Session("SessionID").ToString.Trim
    CopientNotes &= "SessionID: " & SessionID
  End If

  Send_HeadBegin("term.customer", "term.accumulation", CustomerPK)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
  var linkToHH = false;

  function showDetail(row, btn) {
    var elemTr = document.getElementById("histdetail" + row);

    if (elemTr != null && btn != null) {
      elemTr.style.display = (btn.value == "+") ? "" : "none";
      btn.value = (btn.value == "+") ? "-" : "+";
    }
  }

  function HandleSwitchToHH() {
    var refreshElem = document.getElementById("RefreshParent");

    linkToHH = true;
    if (refreshElem != null) {
      refreshElem.value = "true";
    }
  }

</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(2)
  
  ' Grab the offer description from the appropriate table
  If (OfferID > 0) Then
    MyCommon.QueryStr = "select IncentiveName as Name, Description, OID.EngineID, OID.EngineSubTypeID from CPE_Incentives as I with (NoLock) " & _
                        "inner join OfferIDs as OID on OID.OfferID=I.IncentiveID " & _
                        "where I.IncentiveID=" & OfferID & " " & _
                        " union " & _
                        "select Name, Description, OID.EngineID, OID.EngineSubTypeID from Offers as O with (NoLock) " & _
                        "inner join OfferIDs as OID on OID.OfferID=O.OfferID " & _
                        "where O.OfferID=" & OfferID & ";"
    rst2 = MyCommon.LRT_Select
    If rst2.Rows.Count > 0 Then
      OfferDesc = MyCommon.NZ(rst2.Rows(0).Item("Description"), "")
      OfferName = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
      If (MyCommon.NZ(rst2.Rows(0).Item("EngineID"), -1) = 2) AndAlso (MyCommon.NZ(rst2.Rows(0).Item("EngineSubTypeID"), -1) = 2) Then
        IsUSAirMiles = True
      End If
    End If
  End If
  
  If (Logix.UserRoles.EditAccumBalances) AndAlso (IsUSAirMiles = False) Then
    AdjustPermitted = True
  ElseIf (Logix.UserRoles.EditAirmilesAccumBalances) AndAlso (IsUSAirMiles = True) Then
    AdjustPermitted = True
  End If
  
  If (Logix.UserRoles.AccessAccumBalances = False) AndAlso (IsUSAirMiles = False) Then
    Send_Denied(2, "perm.customers-accumbalaccess")
    GoTo done
  ElseIf (Logix.UserRoles.AccessAirmilesAccumBalances = False) AndAlso (IsUSAirMiles = True) Then
    Send_Denied(2, "perm.customers-accumbalaccess-airmiles")
    GoTo done
  End If
  
  If (Request.QueryString("mode") = "addaccum") Then
    Add_Accumulation(AdminUserID, SessionID)
  End If
  
  If (Request.QueryString("infoMsgCode") <> "") Then
    InfoMsgCode = MyCommon.Extract_Val(Request.QueryString("infoMsgCode"))
    If (InfoMsgCode = REWARD_ALREADY_EARNED) Then
      InfoMessage = Copient.PhraseLib.Lookup("CPE_accum-adjust-rwdearned", LanguageID)
    ElseIf (InfoMsgCode = REWARD_NOT_RECEIVED) Then
      InfoMessage = Copient.PhraseLib.Lookup("CPE_accum-adj-not-rcvd", LanguageID)
    ElseIf (InfoMsgCode = ADJUSTMENT_AMT_NOT_SENT) Then
      InfoMessage = Copient.PhraseLib.Lookup("CPE_accum-adj-noadjamt", LanguageID)
    ElseIf (InfoMsgCode = ADJUSTMENT_IMPROPER_FORMAT) Then
      InfoMessage = Copient.PhraseLib.Lookup("CPE_accum-adj-improperformat", LanguageID)
      InfoMessage &= " (" & Copient.PhraseLib.Lookup("term.expectedformat", LanguageID) & ": "
      Select Case MyCommon.Extract_Val(Request.QueryString("unittype"))
        Case 1
          InfoMessage &= Copient.PhraseLib.Lookup("term.positiveintegers", LanguageID)
        Case 2
          InfoMessage &= Copient.PhraseLib.Lookup("CPEaccum-adjust.twodecimal", LanguageID)
        Case 3
          InfoMessage &= Copient.PhraseLib.Lookup("CPEaccum-adjust.threedecimal", LanguageID)
        Case Else
          InfoMessage &= Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      End Select
      InfoMessage &= ")"
    ElseIf (InfoMsgCode = ERROR_DURING_ADJUST) Then
      InfoMessage &= Copient.PhraseLib.Lookup("CPEaccum-adjust.error", LanguageID)
    ElseIf (InfoMsgCode = ADJUSTMENT_BELOW_ZERO) Then
      InfoMessage &= Copient.PhraseLib.Lookup("CPE_accum-adjust-subzero", LanguageID)
    End If
  End If
  If IsAccumulationAdjustmentAllowed(OfferID, CustomerPK, HHPK, HHCardPK) = False Then
    InfoMessage = Copient.PhraseLib.Lookup("customer-inquiry.accum-hh-adjust-note", LanguageID)
    AdjustPermitted = False
  End If
  
  Send("<form method=""get"" action=""CPEaccum-adjust.aspx"" id=""mainform"" name=""mainform"">")
  Send("<div id=""intro"">")
  Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & Request.QueryString("OfferID") & ": " & MyCommon.TruncateString(OfferName, 40) & "</h1>")
  Send("  <div id=""controls"">")
  If (AdjustPermitted) Then
    Send("<input type=""submit"" accesskey=""s"" class=""regular"" id=""save"" name=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ />")
  End If
  Send("  </div>")
  Send("</div>")
  Send("<div id=""main"">")
  If (InfoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & InfoMessage & "</div>")
  Send("  <div id=""column"">")
  If (HHPK > 0 AndAlso HHCardPK > 0) Then
    Send("<br /><a href=""CPEaccum-adjust.aspx?")
    KeyCt = Request.QueryString.Keys.Count
    For i = 0 To KeyCt - 1
      If (Request.QueryString.Keys(i) = "CustomerPK") Then
        Send("CustomerPK=" & HHPK)
      ElseIf (Request.QueryString.Keys(i) = "CardPK") Then
        If HHCardPK > 0 Then
          Send("CardPK=" & HHCardPK)
        Else
          Send(Request.QueryString.Keys(i) & "=" & Request.QueryString.Item(i))
        End If
      Else
        Send(Request.QueryString.Keys(i) & "=" & Request.QueryString.Item(i))
      End If
      If (i < KeyCt - 1) Then Send("&")
    Next
    Send(""" onclick=""javascript:HandleSwitchToHH();"" >")
    Send(Copient.PhraseLib.Lookup("sv.hh-adjust-linktext", LanguageID) & "</a><br/><br/>")
  End If
  If (OfferDesc <> "") Then
    Send("    <p>" & MyCommon.SplitNonSpacedString(OfferDesc, 50) & "</p>")
  End If

  HandleOfferAccumulation(CustomerExtID, CustomerPK, OfferID, AdjustPermitted)
  Send("  </div>")
  Send("</div>")
  Send("</form>")
%>
<script runat="server">
    Dim Logix As New Copient.LogixInc
    Dim MyCryptLib As New Copient.CryptLib
  Dim CustUpdateRA As New Copient.CustomerUpdate(MyCommon)
  
  Const REWARD_ALREADY_EARNED As Integer = 1
  Const REWARD_NOT_RECEIVED As Integer = 2
  Const ADJUSTMENT_AMT_NOT_SENT As Integer = 3
  Const ADJUSTMENT_IMPROPER_FORMAT As Integer = 4
  Const ERROR_DURING_ADJUST As Integer = 5
  Const ADJUSTMENT_BELOW_ZERO As Integer = 6
  
  Function IsAccumulationAdjustmentAllowed(ByVal OfferID As Long, ByVal CustomerPK As Long, ByRef HHPK As Long, ByRef HHCardPK As Long) As Boolean
    Dim dt As DataTable
    Dim HHEnable As Boolean = False
    Dim Result = True
    If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
    If (MyCommon.LXSadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixXS()
    MyCommon.QueryStr = "SELECT HHEnable from CPE_RewardOptions WHERE IncentiveID = @IncentiveID"
    MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
    If dt.Rows.Count > 0 Then HHEnable = MyCommon.NZ(dt.Rows(0).Item("HHEnable"), False)
    
    If HHEnable Then
      MyCommon.QueryStr = "Select HHPK from Customers where CustomerPK = @CustomerPK and HHPK <> 0"
      MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
      dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
      Result = (dt.Rows.Count = 0)
      
      If (Result = False) Then
        HHPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
        MyCommon.QueryStr = "Select CardPK from CardIDs where CustomerPK = @CustomerPK and CardTypeID = @CardTypeID"
        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = HHPK
        MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = Copient.commonShared.CardTypes.HOUSEHOLD
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        If dt.Rows.Count > 0 Then
          HHCardPK = MyCommon.NZ(dt.Rows(0).Item("CardPK"), 0)
        End If
      End If
    End If
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    Return Result
  End Function
  
  Sub HandleOfferAccumulation(ByVal CustomerExtID As String, ByVal CustomerPK As Long, ByVal OfferID As String, ByVal AdjustPermitted As Boolean)
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim rst3 As DataTable
    Dim dt As DataTable
    Dim dt2 As DataTable
    Dim row As DataRow
    Dim row1 As DataRow
    Dim EngineID As Integer
    Dim AccumProgram As Boolean = False
    Dim PointsProgram As Boolean = False
    Dim RewardOptionID As Long
    Dim HHEnable As Boolean
    Dim HHPrimaryID As Long
    Dim UnitType As Integer  '1=Items  2=Dollars
    Dim UpdateAccum As Boolean = False
    Dim TotalAccum As Decimal
    Dim CurrentAccum As Decimal
    Dim DistCardNum As String = ""
    Dim LocationName As String = ""
    'Dim CustomerPK As Long
    Dim AccumAdj As String = ""
    Dim Overthreshold As Boolean = False
    Dim QtyForIncentive As Double = 0
    Dim PendingRdm As Integer = 0
    Dim j As Integer = 0
    Dim OfferIsAirMiles As Boolean = False
    Dim DeletedAODRecord As Boolean = False
    Dim AODRecord As Boolean = False
    Dim DeletedDisplayRecord As Boolean = False
    Dim DisplayAmount As Single = 0
    Dim DisplayFormat As String = ""
    Dim DisplayCount As Integer = 0
    Dim TimesDisplayed As Integer = 0
    Dim AODOnly As Boolean = False
    Dim AODDeletedManual As Boolean = False
    Dim RunOnce As Boolean = False
    Dim SecondRun As Boolean = False
    Dim PresentedCustomerID As String = ""
    Dim ResolvedCustomerID As String = ""
    
    
    If (MyCommon.LRTadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixRT()
    If (MyCommon.LXSadoConn.State <> ConnectionState.Open) Then MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
    MyCommon.QueryStr = "select EngineID, EngineSubTypeID from OfferIDs with (NoLock) where OfferID = " & OfferID
    rst = MyCommon.LRT_Select
    
    If (rst.Rows.Count > 0) Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
      If (EngineID = 2 AndAlso MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), -1) = 2) Then OfferIsAirMiles = True
    End If
    
    If OfferIsAirMiles Then
      UpdateAccum = Logix.UserRoles.EditAirmilesAccumBalances
    Else
      UpdateAccum = Logix.UserRoles.EditAccumBalances
    End If
    
    ' At this time, only CPE engine offers allow Accumulation adjustments
    If (EngineID = 2) Then
      MyCommon.QueryStr = "select IPG.AccumMin, RO.RewardOptionID, RO.HHEnable, IPG.QtyUnitType, IPG.QtyForIncentive " & _
                          "from CPE_IncentiveProductGroups as IPG with (NoLock) Inner Join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID and IPG.Deleted=0 and IPG.ExcludedProducts=0 and RO.Deleted=0 " & _
                          "where RO.IncentiveID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
          AccumProgram = True
        End If
        RewardOptionID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), 0)
        HHEnable = MyCommon.NZ(rst.Rows(0).Item("HHEnable"), False)
        UnitType = MyCommon.NZ(rst.Rows(0).Item("QtyUnitType"), 2)
        QtyForIncentive = MyCommon.NZ(rst.Rows(0).Item("QtyForIncentive"), 0)
      End If
      
      If HHEnable Then
        MyCommon.QueryStr = "select HHPK from Customers with (NoLock) where CustomerPK in " & _
                            " (select CustomerPK from CardIDs with (NoLock) where ExtCardID = '" & MyCryptLib.SQL_StringEncrypt(CustomerExtID) & "')"
        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
          HHPrimaryID = MyCommon.NZ(rst.Rows(0).Item("HHPK"), 0)
        End If
      End If
      
      If AccumProgram Then
        'Query for the CustomerPK from the External ID
        If (CustomerPK = 0) Then
          MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerExtID) & "';"
          rst = MyCommon.LXS_Select()
          If (rst.Rows.Count > 0) Then
            CustomerPK = rst.Rows(0).Item("CustomerPK")
          End If
        End If
                
        'only make this available to administrators who have access to adjust accumulation amounts
        If UpdateAccum Then
          Send("")
          
          Send("    <div class=""box"" id=""accum""" & IIf(AdjustPermitted, "", " style=""display:none;""") & ">")
          Send("      <h2><span>" & Copient.PhraseLib.Lookup("term.accumadjustment", LanguageID) & "</span></h2>")
          Send("      <input type=""hidden"" id=""mode""            name=""mode""           value=""addaccum"" />")
          Send("      <input type=""hidden"" id=""userid""          name=""userid""         value=""" & CustomerPK & """ />")
          Send("      <input type=""hidden"" id=""CustomerPK""      name=""CustomerPK""     value=""" & CustomerPK & """ />")
          Send("      <input type=""hidden"" id=""roid""            name=""roid""           value=""" & RewardOptionID & """ />")
          Send("      <input type=""hidden"" id=""incentiveid""     name=""incentiveid""    value=""" & OfferID & """ />")
          Send("      <input type=""hidden"" id=""OfferID""         name=""OfferID""        value=""" & OfferID & """ />")
          Send("      <input type=""hidden"" id=""cardnumber""      name=""cardnumber""     value=""" & CustomerExtID & """ />")
          Send("      <input type=""hidden"" id=""CustomerExtId""   name=""CustomerExtId""  value=""" & CustomerExtID & """ />")
          Send("      <input type=""hidden"" id=""noback""          name=""noback""         value=""1"" />")
          Send("      <input type=""hidden"" id=""RefreshParent"" name=""RefreshParent"" value=""false"" />  ")
          AccumAdj = Request.QueryString("accumadj")
          If (UnitType = 1) Or (UnitType = 3) Then
            Send("      <label for=""accumadj"">" & Copient.PhraseLib.Lookup("CPE_accum-adj-amt", LanguageID) & "</label>: <input type=""text"" id=""accumadj"" name=""accumadj"" maxlength=""9"" size=""10"" value=""" & AccumAdj & """ />")
          Else
            Send("      <label for=""accumadj"">" & Copient.PhraseLib.Lookup("CPE_accum-adj-amt", LanguageID) & "</label>: $<input type=""text"" id=""accumadj"" name=""accumadj"" maxlength=""9"" size=""10"" value=""" & AccumAdj & """ />")
          End If
          Send("    </div>")
          Send("<hr class=""hidden"" />")
          Send("")
        End If
        MyCommon.QueryStr = "select 'ND' as AdjustType, Col1 as LocalID, Col2 as ServerSerial, Col6 as QtyAdjust, Col7 as PriceAdjust, Convert(datetime, IsNull(Col8, getdate())) as AdjustDate, Replayed, ReplayedDate " & _
                            "from CPE_UploadTemp_RA_ND with (NoLock) " & _
                            "where Col3 = '" & RewardOptionID & "' and (Col4 = '" & CustomerPK & "' or Col5='" & CustomerPK & "') " & _
                            "union " & _
                            "select 'N' as AdjustType, Col1 as LocalID, Col2 as ServerSerial, Col6 as QtyAdjust, Col7 as PriceAdjust, Convert(datetime, IsNull(Col8, getdate())) as AdjustDate, Replayed, ReplayedDate " & _
                            "from CPE_UploadTemp_RA_N with (NoLock) where Col3 = '" & RewardOptionID & "' and (Col4 = '" & CustomerPK & "' or Col5='" & CustomerPK & "') " & _
                            "order by AdjustType desc, AdjustDate desc;"
        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
          Send("<div class=""box"" id=""pending"">")
          Send("<h2><span>" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & "</span></h2>")
          Send("<span style=""float:right; font-size:9px; position:relative; top: -22px;""><a href=""CPEaccum-adjust.aspx?OfferID=" & OfferID & "&OfferName=&CustomerPK=" & CustomerPK & "&CustomerExtId=" & CustomerExtID & """>" & Copient.PhraseLib.Lookup("term.refresh", LanguageID) & "</a></span>")
          Send("<table border=""0"" cellpadding=""2"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.pending", LanguageID) & """>")
          Send(" <thead>")
          Send("  <tr>")
          Send("   <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
          Send("   <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</th>")
          Send("   <th class=""th-longid"" scope=""col"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & "</th>")
          Send("   <th class=""th-adjustment"" scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & "</th>")
          'If HHEnable Then
          '  Send("   <th class=""th-cardholder"" scope=""col"" id=""cardnumber"">" & Copient.PhraseLib.Lookup("term.cardnumber", LanguageID) & "</th>")
          'End If
          Send("  </tr>")
          Send(" </thead>")
          Send(" <tbody>")
          For Each row In rst.Rows
            Send("  <tr>")
            If IsDBNull(row.Item("AdjustDate")) Then
              Send("    <td>&nbsp;</td>")
            Else
              Send("    <td>" & Logix.ToShortDateTimeString(row.Item("AdjustDate"), MyCommon) & "</td>")
            End If
            If MyCommon.NZ(row.Item("ServerSerial"), 0) = -9 Then
              Send("    <td>" & Copient.PhraseLib.Lookup("term.logix-manual-entry", LanguageID) & "</td>")
            Else
              Send("    <td>" & MyCommon.NZ(row.Item("ServerSerial"), "&nbsp;") & "</td>")
            End If
            If (MyCommon.NZ(row.Item("AdjustType"), "N") = "ND") Then
              Send("    <td>" & Copient.PhraseLib.Lookup("CPE_accum-adjust-QtyNext", LanguageID) & "</td>")
            Else
              Send("    <td>" & Copient.PhraseLib.Lookup("CPE_accum-adjust-QtyCarried", LanguageID) & "</td>")
            End If
            If (UnitType = 1) Or (UnitType = 3) Then
              Send("    <td style=""text-align:right;"">" & MyCommon.NZ(row.Item("QtyAdjust"), "0") & "</td>")
            Else
              Send("    <td style=""text-align:right;"">" & FormatCurrency(MyCommon.NZ(row.Item("PriceAdjust"), "0.00")) & "</td>")
            End If
            'If HHEnable Then
            '  Send("   <td style=""text-align:center;"">" & DistCardNum & "</td>")
            'End If
            Send("  </tr>")
          Next
          Send(" </tbody>")
          Send("</table>")
          Send("</div>")
        End If
        
        If HHEnable Then
          MyCommon.QueryStr = "select RA.AccumulationDate, RA.LastServerID, RA.ServerSerial, RA.CustomerPK, RA.PurchCustomerPK, RA.TotalPrice, RA.QtyPurchased, RA.Deleted, RA.LocationID, " & _
                              "PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate, RA.LastUpdate, RA.POSTimeStamp, 'RA' as SourceTable " & _
                              "from CPE_RewardAccumulation as RA with (NOLOCK) where (RA.CustomerPK=" & CustomerPK & " or RA.CustomerPK=" & HHPrimaryID & ") and RA.RewardOptionID=" & RewardOptionID & " and RA.Deleted=0 " & _
                              "union all " & _
                              "select RA2.AccumulationDate, RA2.LastServerID, RA2.ServerSerial, RA2.CustomerPK, RA2.PurchCustomerPK, RA2.TotalPrice, RA2.QtyPurchased, RA2.Deleted, RA2.LocationID, " & _
                              "PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, 0 As Replayed, ReplayedDate, RA2.LastUpdate, RA2.POSTimeStamp, 'Archive' as SourceTable " & _
                              "from CPE_RA_Archive as RA2 with (NOLOCK) where (RA2.CustomerPK=" & CustomerPK & " or RA2.CustomerPK=" & HHPrimaryID & ") and RA2.RewardOptionID=" & RewardOptionID & " " & _
                              " order by AccumulationDate;"
          
        Else
          MyCommon.QueryStr = "select RA.AccumulationDate, RA.LastServerID, RA.ServerSerial, RA.CustomerPK, RA.PurchCustomerPK, RA.TotalPrice, RA.QtyPurchased, RA.Deleted, RA.LocationID, " & _
                              "PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, Replayed, ReplayedDate, RA.LastUpdate, RA.POSTimeStamp, 'RA' as SourceTable " & _
                              "from CPE_RewardAccumulation as RA with (NOLOCK) where RA.CustomerPK=" & CustomerPK & " and RA.RewardOptionID=" & RewardOptionID & " and RA.Deleted=0 " & _
                              "union all " & _
                              "select RA2.AccumulationDate, RA2.LastServerID, RA2.ServerSerial, RA2.CustomerPK, RA2.PurchCustomerPK, RA2.TotalPrice, RA2.QtyPurchased, RA2.Deleted, RA2.LocationID, " & _
                              "PresentedCustomerID, PresentedCardTypeID, ResolvedCustomerID, HHID, 0 as Replayed, ReplayedDate, RA2.LastUpdate, RA2.POSTimeStamp, 'Archive' as SourceTable " & _
                              "from CPE_RA_Archive as RA2 with (NOLOCK) where RA2.CustomerPK=" & CustomerPK & " and RA2.RewardOptionID=" & RewardOptionID & " " & _
                              " order by AccumulationDate;"
        End If
        rst = MyCommon.LXS_Select
        
        Send("<div class=""box"" id=""history"">")
        Send("<h2><span>" & Copient.PhraseLib.Lookup("term.history", LanguageID) & "</span></h2>")
        Send("<div id=""listing"" class=""boxscroll"" style=""height:350px;width:630px;"">")
        If (rst.Rows.Count > 0) Then
          Send("<table border=""0"" cellpadding=""2"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.history", LanguageID) & """>")
          Send(" <thead>")
          Send("  <tr>")
          Send("   <th scope=""col"" style=""width:30px;"">&nbsp;</th>")
          Send("   <th class=""th-datetime"" scope=""col"">" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
          Send("   <th class="""" scope=""col"" style=""text-align:center;width:20px;"" title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """>" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</th>")
          Send("   <th class=""th-name"" scope=""col"">" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</th>")
          Send("   <th class=""th-adjustment"" scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & "</th>")
          'If HHEnable Then
          '  Send("   <th class=""th-cardholder"" scope=""col"" id=""cardnumber"">" & Copient.PhraseLib.Lookup("term.cardnumber", LanguageID) & "</th>")
          'End If
          Send("  </tr>")
          Send(" </thead>")
          Send(" <tbody>")
          TotalAccum = 0
          CurrentAccum = 0
          j = 0
          For Each row In rst.Rows
            'j += 1
            If HHEnable Then
              'DistCardNum = "0"
              'MyCommon.QueryStr = "select ClientUserID1 from Users where UserID=" & MyCommon.NZ(rst.Rows(0).Item("PurchUserID"), 0) & ";"
              'rst2 = MyCommon.LXS_Select
              'If Not (rst2.Rows.Count > 0) Then
              '    DistCardNum = MyCommon.NZ(rst2.Rows(0).Item("ClientUserID1"), "0")
              'End If
            End If
            
            'Resetting variables for each record
            LocationName = ""
            AODRecord = False
            DeletedDisplayRecord = False
            DisplayCount = 0
            TimesDisplayed = 0
            AODOnly = False
            AODDeletedManual = False
            RunOnce = False
            SecondRun = False
            
            'If the record is a delete record in the Archive table and has been altered/created by the AOD, then it needs to be displayed twice.
            'The first time it is displayed it will be shown as it was originally received and the second time it will be displayed as its is currently.
            If MyCommon.NZ(row.Item("SourceTable"), "") = "Archive" AndAlso MyCommon.NZ(row.Item("Deleted"), False) AndAlso MyCommon.NZ(row.Item("LastServerID"), 0) = -10 Then
              DisplayCount = 2
              DeletedDisplayRecord = True
            Else
              DisplayCount = 1
            End If
                               
            If MyCommon.NZ(row.Item("ServerSerial"), 0) = -10 AndAlso MyCommon.NZ(row.Item("LastServerID"), 0) = -10 Then AODOnly = True 'If the record has a server serial of -10 and LastServerID of -10 then it is completely from AOD
            If MyCommon.NZ(row.Item("ServerSerial"), 0) = -9 AndAlso DeletedDisplayRecord = True Then AODDeletedManual = True 'If the record is from a Logix Manual Entry and has been altered by the AOD
            
            While TimesDisplayed < DisplayCount
              j += 1 'Increase the count for the table id
              DeletedAODRecord = False
              LocationName = ""
              If (TimesDisplayed = 0 AndAlso DisplayCount = 1) Then RunOnce = True
              If (TimesDisplayed = 1 AndAlso DisplayCount = 2) Then SecondRun = True
              'If this is an AOD record, must be displayed twice, and this is the second time of displaying it or is only to be displayed once or is entirely an AOD record
              If (MyCommon.NZ(row.Item("LastServerID"), 0) = -10 AndAlso ((SecondRun = True OrElse RunOnce = True) OrElse AODOnly)) Then
                AODRecord = True
                LocationName = "<span>" & Copient.PhraseLib.Lookup("term.autoofferdist", LanguageID) & "</span>"
                If MyCommon.NZ(row.Item("Deleted"), False) AndAlso (TimesDisplayed = 1 AndAlso DisplayCount = 2) Then
                  DeletedAODRecord = True
                End If
                'If this is an AOD record, must be displayed twice, and this is the second time of displaying it or is only to be displayed once or is a Logix Manual Entry record changed by AOD
              ElseIf (MyCommon.NZ(row.Item("ServerSerial"), 0) = -9 AndAlso ((SecondRun = True OrElse RunOnce = True) OrElse AODDeletedManual)) Then
                LocationName = "<span>" & Copient.PhraseLib.Lookup("term.logix-manual-entry", LanguageID) & "</span>"
              Else
                If MyCommon.NZ(row.Item("LocationID"), 0) > 0 Then
                  MyCommon.QueryStr = "select ExtLocationCode from Locations where LocationID=" & row.Item("LocationID") & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count > 0 Then
                    LocationName = "<span alt=""" & MyCommon.NZ(row.Item("LocationID"), "") & """ title=""" & MyCommon.NZ(row.Item("LocationID"), "") & """>" & MyCommon.NZ(rst2.Rows(0).Item("ExtLocationCode"), "") & "</span>"
                  Else
                    LocationName = "<span alt=""" & MyCommon.NZ(row.Item("LocationID"), "") & """ title=""" & MyCommon.NZ(row.Item("LocationID"), "") & """>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</span>"
                  End If
                End If
              End If
              Send("  <tr id=""hist" & j & """" & IIf(DeletedAODRecord, " style=""background-color:#ffff99;""", "") & ">")
              Send("    <td><input class=""ex more"" type=""button"" value=""+"" name=""ex"" title=""" & Copient.PhraseLib.Lookup("term.details", LanguageID) & """ onclick=""javascript:showDetail(" & j & ", this);"" /></td>")
              Sendb("   <td nowrap align=""left"">")
              'If the record is a deleted AOD record and needs to be displayed twice then the LastUpdate column is used to display the date of the AOD display and the AccumulatioDate is used to display the date of the original accumulation record.
              If (DeletedDisplayRecord AndAlso (TimesDisplayed = 1 AndAlso DisplayCount = 2)) Then  'If this is a deleted AOD record that needs to be displayed twice and this is the second time displaying it then use the LastUpdate column
                If MyCommon.NZ(row.Item("LastUpdate"), "1/1/1980") = "1/1/1980" Then
                  Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                Else
                  Sendb(Logix.ToShortDateTimeString(MyCommon.NZ(row.Item("LastUpdate"), New Date(1900, 1, 1)), MyCommon))
                End If
              Else  'Else use the normal AccumulationDate
                Dim ShowPOSTimeStamp As Boolean = IIf(MyCommon.Fetch_CPE_SystemOption(131) = "1", True, False)
                If ShowPOSTimeStamp AndAlso MyCommon.NZ(row.Item("POSTimeStamp"), "1/1/1980") <> "1/1/1980" Then
                  'If MyCommon.NZ(row.Item("POSTimeStamp"), "1/1/1980") = "1/1/1980" Then
                  'Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                  'Else
                  Sendb(Logix.ToShortDateTimeString(MyCommon.NZ(row.Item("POSTimeStamp"), New Date(1900, 1, 1)), MyCommon))
                  'End If
                Else
                  If MyCommon.NZ(row.Item("AccumulationDate"), "1/1/1980") = "1/1/1980" Then
                    Sendb(Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                  Else
                    Sendb(Logix.ToShortDateTimeString(MyCommon.NZ(row.Item("AccumulationDate"), New Date(1900, 1, 1)), MyCommon))
                  End If
                End If
              End If
              Send("</td>")
              Send("   <td style=""text-align:center;"">" & IIf(MyCommon.NZ(row.Item("Replayed"), False), "<span title=""" & Copient.PhraseLib.Lookup("term.replayed", LanguageID) & """ style=""color:#dd0000;cursor:default;font-size:10px;font-weight:bold;"">" & Left(Copient.PhraseLib.Lookup("term.replayed", LanguageID), 1) & "</span>", "") & "</td>")
              Send("   <td align=""left"">" & LocationName & "</td>")
              Sendb("   <td align=""right""")
              'If the record is used to be redeemed for an offer then it should be displayed in red.
              If MyCommon.NZ(row.Item("Deleted"), False) AndAlso DeletedAODRecord = False AndAlso DeletedDisplayRecord = False Then
                Sendb(" style=""color:#ff0000;""")
              End If
              Sendb(">")
              If UnitType = 1 Then
                DisplayAmount = Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 0)
                If AODRecord Then
                  Send(FormatAODRecords(UnitType, DisplayAmount, DeletedAODRecord) & "</td>")
                Else
                  Send(Format(Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 0), "###,##0") & "</td>")
                End If
                If DeletedAODRecord = False AndAlso DeletedDisplayRecord = False Then TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 0)
                If MyCommon.NZ(row.Item("Deleted"), False) = False Then CurrentAccum = CurrentAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 0)
              ElseIf UnitType = 2 Then
                DisplayAmount = Math.Round(MyCommon.NZ(row.Item("TotalPrice"), 0), 2)
                If AODRecord Then
                  Send(FormatAODRecords(UnitType, DisplayAmount, DeletedAODRecord) & "</td>")
                Else
                  If DisplayAmount < 0 Then 'If the amount for a dollar accumulation is negative, change this so it can be displayed properly.
                    DisplayAmount = DisplayAmount * -1
                    Send("- $" & Format(DisplayAmount, "###,##0.00") & "</td>")
                  Else
                    Send("$" & Format(Math.Round(MyCommon.NZ(row.Item("TotalPrice"), 0), 2), "###,##0.00") & "</td>")
                  End If
                End If
                If DeletedAODRecord = False AndAlso DeletedDisplayRecord = False Then TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("TotalPrice"), 0), 2)
                If MyCommon.NZ(row.Item("Deleted"), False) = False Then CurrentAccum = CurrentAccum + Math.Round(MyCommon.NZ(row.Item("TotalPrice"), 0), 2)
              ElseIf UnitType = 3 Then
                DisplayAmount = Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 3)
                If AODRecord Then
                  Send(FormatAODRecords(UnitType, DisplayAmount, DeletedAODRecord) & "</td>")
                Else
                  Send(Format(Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 3), "###,##0.000") & "</td>")
                End If
                If DeletedAODRecord = False AndAlso DeletedDisplayRecord = False Then TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 3)
                If MyCommon.NZ(row.Item("Deleted"), False) = False Then CurrentAccum = CurrentAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 3)
              End If
              'If HHEnable Then
              '  Send("   <td style=""text-align:center;"">" & DistCardNum & "</td>")
              'End If
              Send("  </tr>")
              PresentedCustomerID = MyCommon.NZ(row.Item("PresentedCustomerID"), "Unknown")
              ResolvedCustomerID = MyCommon.NZ(row.Item("ResolvedCustomerID"), "Unknown")
              If (ResolvedCustomerID = "0" OrElse ResolvedCustomerID = "Unknown") Then
                MyCommon.QueryStr = "select ExtCardID from CardIDs with (NoLock) where CustomerPK=" & MyCommon.NZ(row.Item("CustomerPK"), 0) & ";"
                dt2 = MyCommon.LXS_Select
                If dt2.Rows.Count > 0 Then
                  ResolvedCustomerID = IIf(MyCryptLib.SQL_StringDecrypt(dt2.Rows(0).Item("ExtCardID").ToString())= "", "Unknown", MyCryptLib.SQL_StringDecrypt(dt2.Rows(0).Item("ExtCardID").ToString()))
                End If
              End If
              Send("  <tr id=""histdetail" & j & """ style=""display:none;color:#777777;"">")
              Send("    <td></td>")
              Send("    <td colspan=""4"">")
              Send("      " & Copient.PhraseLib.Lookup("term.presented", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & PresentedCustomerID & " &nbsp;|&nbsp; ")
              Send("      " & Copient.PhraseLib.Lookup("term.resolved", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & ResolvedCustomerID & " &nbsp;|&nbsp; ")
              Send("      " & Copient.PhraseLib.Lookup("term.household", LanguageID) & " " & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & MyCommon.NZ(row.Item("HHID"), "Unknown"))
              Send("    </td>")
              Send("  </tr>")
              TimesDisplayed += 1
            End While
          Next
          Send("")
          Send("  <tr>")
          Send("   <td colspan=""4""></td>")
          Send("   <td style=""text-align:right;""><img src=""/images/blackdot.png"" style=""width:60px;height:1px;"" alt="""" /></td>")
          Send("  </tr>")
          Send("")
          Send("  <tr class=""shaded"">")
          Send("   <td></td>")
          Send("   <td colspan=""3"" align=""left"">" & Copient.PhraseLib.Lookup("CPE_accum-adj-total", LanguageID) & ":</td>")
          Sendb("   <td align=""right"">")
          If UnitType = 1 Then
            Sendb(Format(TotalAccum, "###,##0"))
          ElseIf UnitType = 2 Then
            Sendb("$" & Format(TotalAccum, "###,##0.00"))
          ElseIf UnitType = 3 Then
            Sendb(Format(TotalAccum, "###,##0.000"))
          End If
          Send("</td>")
          Send("  </tr>")
          Send("  <tr class=""shadedmid"">")
          Send("   <td></td>")
          Send("   <td colspan=""3"" align=""left"">" & Copient.PhraseLib.Lookup("CPE_accum-adj-current", LanguageID) & ":</td>")
          Sendb("   <td align=""right"">")
          If UnitType = 1 Then
            Sendb(Format(CurrentAccum, "###,##0"))
          ElseIf UnitType = 2 Then
            Sendb("$" & Format(CurrentAccum, "###,##0.00"))
          ElseIf UnitType = 3 Then
            Sendb(Format(CurrentAccum, "###,##0.000"))
          End If
          Sendb("<input type=""hidden"" name=""CurrentAccum"" value=""" & CurrentAccum & """ />")
          Send("</td>")
          Send("  </tr>")
          
          MyCommon.QueryStr = "select RewardAccumulationID, Overthreshold from CPE_RewardAccumulation with (NoLock)" & _
                              "where CustomerPK=" & CustomerPK & " and RewardOptionID=" & RewardOptionID & " order by RewardAccumulationID desc;"
          dt = MyCommon.LXS_Select()
          If dt.Rows.Count > 0 Then
            Overthreshold = MyCommon.NZ(dt.Rows(0).Item("Overthreshold"), 0)
          End If
          PendingRdm = Math.Floor(CurrentAccum / QtyForIncentive)
          If Overthreshold AndAlso PendingRdm > 0 Then
            'List the number of pending redemptions
            Send("  <tr class=""shaded"">")
            If PendingRdm = 1 Then
              Send("   <td colspan=""5"" align=""left"" style=""color:GREEN;"">" & PendingRdm & " " & Copient.PhraseLib.Lookup("term.pending", LanguageID) & " " & Copient.PhraseLib.Lookup("term.redemption", LanguageID) & "</td>")
            Else
              Send("   <td colspan=""5"" align=""left"" style=""color:GREEN;"">" & PendingRdm & " " & Copient.PhraseLib.Lookup("term.pending", LanguageID) & " " & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & "</td>")
            End If
            Send("  </tr>")
          End If
          Send(" </tbody>")
          Send("</table>")
        Else
          Send(Copient.PhraseLib.Lookup("CPE_accum-adj-noaccum", LanguageID))
        End If
        Send("</div>")
        Send("</div>")
      End If
    End If
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
  End Sub
  
  Sub Add_Accumulation(ByVal AdminID As String, ByVal SessionID As String)
    Dim RewardOptionID As Long
    Dim IncentiveID As Long
    Dim IncentiveName As String = ""
    Dim UserID As Long
    Dim UnitType As Integer  '1=Items  2=Dollars
    Dim QtyForIncentive As Long
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim row As DataRow
    Dim RewardEarned As Boolean
    Dim AccumAdj As Double
    Dim TotalPrice As Double
    Dim QtyPurchased As Double
    Dim PrevTotalPrice As Double
    Dim PrevQtyPurchased As Double
    Dim ExtraTotalPrice As Double
    Dim ExtraQtyPurchased As Double
    Dim CurrentAccum As Double
    Dim EDiscount As Boolean
    Dim ErrorMsg As String = ""
    Dim NumRecs As Long
    Dim UserGroupID As Long
    Dim LogStr As String = ""
    Dim HHEnable As Boolean
    Dim RewardDistDatesCt As Integer = 0
    Dim DistributionDate As Date
    Dim QryStr As String = ""
    Dim CustomerExtId As String = ""
    Dim P3DistQtyLimit As Integer = 0
    Dim AwardedUnits As Integer = 0
    Dim i As Integer = 0
    Dim CustomerPK As Long
    Dim SvAdjust As New Copient.StoredValue
    Dim AccumLocalID, AccumLocalID2 As Long
    Dim Fields As New Copient.CommonInc.ActivityLogFields
    Dim AssocLinks(-1) As Copient.CommonInc.ActivityLink
    

    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    
    'If Not (UpdateAccum) Then
    '    Deny_Update_Access(AdminID)
    'End If
    
      
    RewardOptionID = MyCommon.Extract_Val(Request.QueryString("roid"))
    UserID = MyCommon.Extract_Val(Request.QueryString("userid"))
    CustomerExtId = Request.QueryString("cardnumber")
    CustomerPK = Request.QueryString("CustomerPK")
    CurrentAccum = Math.Round(CDbl(MyCommon.Extract_Val(Request.QueryString("CurrentAccum"))), 3)
    
    HHEnable = False
    'get information about the accumulation promotion
    MyCommon.QueryStr = "select IPG.AccumMin, IPG.QtyUnitType, IPG.QtyForIncentive, RO.IncentiveID, RO.HHEnable, I.IncentiveID, I.IncentiveName " & _
               "from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
               "Inner Join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID and IPG.Deleted=0 and RO.Deleted=0 and IPG.ExcludedProducts=0 " & _
               "Inner Join CPE_Incentives as I with (NoLock) on RO.IncentiveID=I.IncentiveID " & _
               "where RO.RewardOptionID=" & RewardOptionID & " and RO.TouchResponse=0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      UnitType = MyCommon.NZ(rst.Rows(0).Item("QtyUnitType"), 2)
      IncentiveID = MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0)
      QtyForIncentive = MyCommon.NZ(rst.Rows(0).Item("QtyForIncentive"), 0)
      HHEnable = MyCommon.NZ(rst.Rows(0).Item("HHEnable"), False)
      IncentiveName = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
      IncentiveID = MyCommon.NZ(rst.Rows(0).Item("IncentiveID"), 0)
    End If
    
    If Not IsNumeric(Request.QueryString("accumadj")) Then
      Response.Status = "301 Moved Permanently"
      QryStr = "?infoMsgCode=" & ADJUSTMENT_IMPROPER_FORMAT & "&unittype=" & UnitType & "&accumadj=" & AccumAdj & "&OfferID=" & IncentiveID & "&CustomerExtId=" & CustomerExtId & "&CustomerPK=" & CustomerPK
      Response.AddHeader("Location", "CPEaccum-adjust.aspx" & QryStr)
      Exit Sub
    End If
    
    AccumAdj = Math.Round(CDbl(MyCommon.Extract_Val(Request.QueryString("accumadj"))), 3)
    ' check if the adjustment amount entry matches the unit type
    If (Not IsProperFormat(UnitType, AccumAdj)) Then
      Response.Status = "301 Moved Permanently"
      QryStr = "?infoMsgCode=" & ADJUSTMENT_IMPROPER_FORMAT & "&unittype=" & UnitType & "&accumadj=" & AccumAdj & "&OfferID=" & IncentiveID & "&CustomerExtId=" & CustomerExtId & "&CustomerPK=" & CustomerPK
      Response.AddHeader("Location", "CPEaccum-adjust.aspx" & QryStr)
      Exit Sub
    End If
    
    If AccumAdj = 0 Then
      Response.Status = "301 Moved Permanently"
      QryStr = "?infoMsgCode=" & ADJUSTMENT_AMT_NOT_SENT & "&accumadj=" & AccumAdj & "&OfferID=" & IncentiveID & "&CustomerExtId=" & CustomerExtId & "&CustomerPK=" & CustomerPK
      Response.AddHeader("Location", "CPEaccum-adjust.aspx" & QryStr)
      Exit Sub
    End If
    
    If CurrentAccum + AccumAdj < 0 Then
      Response.Status = "301 Moved Permanently"
      QryStr = "?infoMsgCode=" & ADJUSTMENT_BELOW_ZERO & "&accumadj=" & AccumAdj & "&OfferID=" & IncentiveID & "&CustomerExtId=" & CustomerExtId & "&CustomerPK=" & CustomerPK
      Response.AddHeader("Location", "CPEaccum-adjust.aspx" & QryStr)
      Exit Sub
    End If
    
    If HHEnable Then  'lookup the primaryid for the card
      MyCommon.QueryStr = "select isnull(HHPK, 0) as HHPrimaryID from Customers with (NoLock) where CustomerPK=" & UserID & ";"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        If rst.Rows(0).Item("HHPrimaryID") <> 0 Then
          UserID = rst.Rows(0).Item("HHPrimaryID")
        End If
      End If
    End If
    
    'see if the distribution limits have been blown
    RewardEarned = False
    'MyCommon.QueryStr = "select RD.DistributionID, RD.RewardOptionID " & _
    '           "from CPE_RewardDistribution as RD Inner Join CPE_RewardOptions as RO on RO.RewardOptionID=RD.RewardOptionID and RO.Deleted=0 and RD.Deleted=0 inner join CPE_Incentives as I on RO.IncentiveID=I.IncentiveID and I.Deleted=0 " & _
    '           "Where RD.RewardOptionID=" & RewardOptionID & " and RD.Phase=3 and RD.UserID=" & UserID & " and ((I.P3DistQtyLimit>0) and(    (DistributionDate>= dateadd(hour, (I.P3DistPeriod * -1), getdate()) and I.P3DistTimeType=2) or (I.P3DistTimeType=1 and I.P3DistPeriod>0 and DistributionDate>= dateadd(day, ((datediff(day, I.startdate, getdate())/I.P3DistPeriod)*I.P3DistPeriod-1), I.StartDate) and DistributionDate<=dateadd(day, ((datediff(day, I.startdate, getdate())/I.P3DistPeriod)*I.P3DistPeriod+I.P3DistPeriod), I.StartDate)) ) );"
              
    ' first check if this reward is unlimited
    MyCommon.QueryStr = "select P3DistQtyLimit from CPE_Incentives I with (NoLock) " & _
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.IncentiveId = I.IncentiveId " & _
                        "where I.Deleted = 0 and RO.Deleted = 0 and RO.RewardOptionID= " & RewardOptionID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      P3DistQtyLimit = MyCommon.NZ(rst.Rows(0).Item("P3DistQtyLimit"), 0)
    End If
    
    If (P3DistQtyLimit > 0) Then
      ' then check if any reward distribution records exist for this roid
      MyCommon.QueryStr = "select RD.DistributionID, DistributionDate, IncentiveID from CPE_RewardDistribution as RD with (NoLock) where deleted=0 and RD.RewardOptionID=" & RewardOptionID & " and RD.Phase=3 and RD.CustomerPK=" & UserID & ";"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count = 0) Then
        RewardEarned = False
      Else
        ' if any reward distribution records exist then check to see if they meet the limits in the incentives table
        For Each row In rst.Rows
          DistributionDate = MyCommon.NZ(row.Item("DistributionDate"), Nothing)
          If (Not DistributionDate = Nothing) Then
            MyCommon.QueryStr = "select IncentiveID from CPE_Incentives I with (NoLock) " & _
                                "where deleted=0 and incentiveid = " & MyCommon.NZ(row.Item("IncentiveID"), -1) & " and ( ( ('" & DistributionDate.ToString & "'>= dateadd(hour, (I.P3DistPeriod * -1), getdate()) and I.P3DistTimeType=2) or (I.P3DistTimeType=1 and I.P3DistPeriod>0 and '" & DistributionDate.ToString & "'>= dateadd(day, ((datediff(day, I.startdate, getdate())/I.P3DistPeriod)*I.P3DistPeriod-1), I.StartDate) and '" & DistributionDate.ToString & "'<=dateadd(day, ((datediff(day, I.startdate, getdate())/I.P3DistPeriod)*I.P3DistPeriod+I.P3DistPeriod), I.StartDate)) ) );"
            rst2 = MyCommon.LRT_Select
            If (rst2.Rows.Count > 0) Then
              RewardDistDatesCt += rst2.Rows.Count
              If (RewardDistDatesCt >= P3DistQtyLimit) Then
                RewardEarned = True
                Exit For
              End If
            End If
          End If
        Next
        ' take it to the limit...one more time
        AwardedUnits = IIf((P3DistQtyLimit - RewardDistDatesCt) < 0, 0, (P3DistQtyLimit - RewardDistDatesCt))
      End If
    End If
    
    'Send("AwardedUnits: " & AwardedUnits)
    'Response.End()
    If RewardEarned Then
      Response.Status = "301 Moved Permanently"
      QryStr = "?infoMsgCode=" & REWARD_ALREADY_EARNED & "&accumadj=" & AccumAdj & "&OfferID=" & IncentiveID & "&CustomerExtId=" & CustomerExtId
      Response.AddHeader("Location", "CPEaccum-adjust.aspx" & QryStr)
      Exit Sub
    Else
      'fetch the accumulation totals for this ROID
      QtyPurchased = 0
      TotalPrice = 0
      MyCommon.QueryStr = "select sum(RA.TotalPrice) as TotalPrice, sum(RA.QtyPurchased) as QtyPurchased " & _
                          "from CPE_RewardAccumulation as RA with (NoLock) " & _
                          "where RA.CustomerPK=" & UserID & " and RA.RewardOptionID=" & RewardOptionID & " and Deleted=0;"
      rst = MyCommon.LXS_Select
      If Not (rst.Rows.Count = 0) Then
        PrevTotalPrice = Math.Round(MyCommon.NZ(rst.Rows(0).Item("TotalPrice"), 0), 2)
        PrevQtyPurchased = MyCommon.NZ(rst.Rows(0).Item("QtyPurchased"), 0)
      End If
      ExtraQtyPurchased = 0
      ExtraTotalPrice = 0
      If (UnitType = 1) Or (UnitType = 3) Then
        TotalPrice = 0
        
        ' find out how many units to award
        AwardedUnits = CInt(Math.Floor(CDec((PrevQtyPurchased + AccumAdj) / QtyForIncentive)))
        
        If (AwardedUnits > 0 And UnitType = 3) Then
          RewardEarned = True
          QtyPurchased = AccumAdj
          ExtraQtyPurchased = (PrevQtyPurchased + AccumAdj) - (AwardedUnits * QtyForIncentive)
          If ExtraQtyPurchased < 0 Then ExtraQtyPurchased = 0
        ElseIf (AwardedUnits > 0 And UnitType = 1) Then
          RewardEarned = True
          QtyPurchased = AccumAdj
          TotalPrice = 0
        Else
          QtyPurchased = AccumAdj
        End If
        'If (PrevQtyPurchased + AccumAdj) >= QtyForIncentive Then
        '  RewardEarned = True
        '  QtyPurchased = QtyForIncentive - PrevQtyPurchased 'calculates the amount needed to reach the threshold
        '  ExtraQtyPurchased = AccumAdj - QtyPurchased  'calculates the amount carried forward after the threshold was crossed
        '  If ExtraQtyPurchased < 0 Then ExtraQtyPurchased = 0
        'Else
        '  QtyPurchased = AccumAdj
        'End If
        
        'if the adjustment amount would have taken the accumulation total below zero
        'then change the adjustment amount so that it would take to total to zero (not below)
        If (PrevQtyPurchased + AccumAdj < 0) Then
          AccumAdj = AccumAdj + (0 - (PrevQtyPurchased + AccumAdj))
          QtyPurchased = AccumAdj
        End If
        If UnitType = 1 Then LogStr = Copient.PhraseLib.Lookup("CPE_accum-adj-by", LanguageID) & " " & AccumAdj & " " & Copient.PhraseLib.Lookup("CPE_accum-adj-items", LanguageID) & " '" & IncentiveName & "'"
        If UnitType = 3 Then LogStr = Copient.PhraseLib.Lookup("CPE_accum-adj-by", LanguageID) & " " & AccumAdj & " " & Copient.PhraseLib.Lookup("CPE_accum-adj-lbsgals", LanguageID) & " '" & IncentiveName & "'"
      End If
      If UnitType = 2 Then
        QtyPurchased = 0
        
        ' find out how many units to award
        AwardedUnits = CInt(Math.Floor(CDec((PrevTotalPrice + AccumAdj) / QtyForIncentive)))
        
        TotalPrice = AccumAdj
        If (AwardedUnits > 0) Then
          RewardEarned = True
        End If
        
        'If ((PrevTotalPrice + AccumAdj) >= QtyForIncentive) Then
        '  RewardEarned = True
        '  TotalPrice = QtyForIncentive - PrevTotalPrice 'calculates the amount needed to reach the threshold
        '  ExtraTotalPrice = AccumAdj - TotalPrice 'calculates the amount carried forward after the threshold was crossed
        'Else
        '  TotalPrice = AccumAdj
        'End If
        
        'if the adjustment amount would have taken the accumulation total below zero
        'then change the adjustment amount so that it would take to total to zero (not below)
        If (PrevTotalPrice + AccumAdj < 0) Then
          AccumAdj = AccumAdj + (0 - (PrevTotalPrice + AccumAdj))
          TotalPrice = AccumAdj
        End If
        LogStr = Copient.PhraseLib.Lookup("CPE_accum-adj-by", LanguageID) & " $" & AccumAdj & " " & Copient.PhraseLib.Lookup("CPE_accum-adj-underoffer", LanguageID) & " " & IncentiveName
      End If
      
      If RewardEarned Then
        'see if this promotion has an EDiscount or Printed Message Deliverable
        MyCommon.QueryStr = "select DeliverableID from CPE_Deliverables with (NoLock) where RewardOptionID=" & RewardOptionID & " and DeliverableTypeID in (1,2,4,9,10) and RewardOptionPhase=3 and Deleted=0;"
        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
          EDiscount = True
        End If
        
        'add the accumulation and mark it as being over the threshold
        AccumLocalID2 = CustUpdateRA.GetAccumAdjustLocalID(True)
        Try
          MyCommon.QueryStr = "begin transaction;"
          MyCommon.LXS_Execute()
          ' insert the extra portion over and above the accumulation threshold as a new record
          ' pa_CPE_TU_InsertData_RA_N
          MyCommon.QueryStr = "insert into CPE_UploadTemp_RA_N with (RowLock) " & _
                              "  (TableNum, Operation, Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, ServerSerial, LocationID, WaitingAck, Col11, Col12, Col13, Col14, Col15, POSTimeStamp ) " & _
                              "values " & _
                              "  (2, 1, " & AccumLocalID2 & ", -9, " & RewardOptionID & ", " & UserID & ", " & UserID & ", " & QtyPurchased & ", " & TotalPrice & ", getdate(), -9, -9, 0, 1, 0, 0, 0, 0, GETDATE() );"
          MyCommon.LXS_Execute()
          MyCommon.QueryStr = "commit transaction;"
          MyCommon.LXS_Execute()
        Catch
          MyCommon.QueryStr = "rollback transaction;"
          MyCommon.LXS_Execute()

          Response.Status = "301 Moved Permanently"
          QryStr = "?infoMsgCode=" & ERROR_DURING_ADJUST & "&accumadj=" & AccumAdj & "&OfferID=" & IncentiveID & "&CustomerExtId=" & CustomerExtId
          Response.AddHeader("Location", "CPEaccum-adjust.aspx" & QryStr)
          Exit Sub
        End Try
        
        ' log the addition of the offer and any associated offers
        Fields.ActivityTypeID = 25
        Fields.ActivitySubTypeID = 14
        Fields.LinkID = UserID
        Fields.AdminUserID = AdminID
        Fields.Description = LogStr
        Fields.LinkID2 = IncentiveID
        Fields.SessionID = SessionID
        Fields.ActivityValue = AccumAdj.ToString
        Fields.PreAdjustBalance = CurrentAccum
        Fields.Adjustment = New Decimal(AccumAdj)
        Fields.PostAdjustBalance = Decimal.Add(Fields.PreAdjustBalance, Fields.Adjustment)
        
        ReDim AssocLinks(0)
        AssocLinks(0).LinkID = IncentiveID
        AssocLinks(0).LinkTypeID = 1
        Fields.AssociatedLinks = AssocLinks
        
        MyCommon.Activity_Log3(Fields)
        
        'create the DistributionHistory record
        'MyCommon.QueryStr = "insert into DistributionHistory (IncentiveID, RewardOptionID, Phase, UserID, Cashier, DistributionDate, LastUpdate, WaitingACK, LastServerID) values (" & IncentiveID & ", " & RewardOptionID & ", 3, " & UserID & ", 'A" & AdminID & "', getdate(), getdate(), 0, -9);"
        'YHouseDBS.Execute(QueryStr)
        
        'see if there is a GroupMembership deliverable - and give it out
        UserGroupID = 0
        MyCommon.QueryStr = "select OutputID from CPE_Deliverables with (NoLock) where RewardOptionID=" & RewardOptionID & " and DeliverableTypeID=5 and Deleted=0;"
        rst = MyCommon.LRT_Select
        If Not (rst.Rows.Count = 0) Then
          UserGroupID = MyCommon.NZ(rst.Rows(0).Item("OutputID"), 0)
        End If
        If Not (UserGroupID = 0) Then
          NumRecs = 0
          'make sure the user isn't already a member of the group
          MyCommon.QueryStr = "select count(*) as NumRecs from GroupMembership with (NoLock) where CustomerGroupID=" & UserGroupID & " and CustomerPK=" & UserID & " and Deleted=0;"
          rst = MyCommon.LXS_Select
          If Not (rst.Rows.Count = 0) Then
            NumRecs = MyCommon.NZ(rst.Rows(0).Item("NumRecs"), 0)
          End If
          If NumRecs = 0 Then
            ' add customer to group and mark status flag for Traffic Cop to propagate for Cross-Shopping
            MyCommon.QueryStr = "insert into GroupMembership with (RowLock) (CustomerPK, CustomerGroupID, Manual, Deleted, LastUpdate, CPEStatusFlag, UEStatusFlag) values (" & UserID & ", " & UserGroupID & ", 1, 0, getdate(), 1, 1);"
            MyCommon.LXS_Execute()
          End If
        End If
        
      Else 'Reward Not Earned
        AccumLocalID = CustUpdateRA.GetAccumAdjustLocalID(True)
        
        If (AccumLocalID > 0) Then
          ' write the record to the upload temp table for the TranUpdateAgent RA-N to process
          ' pa_CPE_TU_InsertData_RA_N
          MyCommon.QueryStr = "insert into CPE_UploadTemp_RA_N with (RowLock) " & _
                              "  (TableNum, Operation, Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, ServerSerial, LocationID, WaitingAck, Col11, Col12, Col13, Col14, Col15, POSTimeStamp) " & _
                              "values " & _
                              "  (2, 1, " & AccumLocalID & ", -9, " & RewardOptionID & ", " & UserID & ", " & UserID & ", " & QtyPurchased & ", " & TotalPrice & ", getdate(), -9, -9, 0, 0, 0, 0, 0, 0, GETDATE() );"
          MyCommon.LXS_Execute()
        Else
          Response.Status = "301 Moved Permanently"
          QryStr = "?infoMsgCode=" & ERROR_DURING_ADJUST & "&accumadj=" & AccumAdj & "&OfferID=" & IncentiveID & "&CustomerExtId=" & CustomerExtId
          Response.AddHeader("Location", "CPEaccum-adjust.aspx" & QryStr)
          Exit Sub
        End If
        
        ' log the addition of the offer and any associated offers
        Fields.ActivityTypeID = 25
        Fields.ActivitySubTypeID = 14
        Fields.LinkID = UserID
        Fields.AdminUserID = AdminID
        Fields.Description = LogStr
        Fields.LinkID2 = IncentiveID
        Fields.SessionID = SessionID
        Fields.ActivityValue = AccumAdj.ToString
        Fields.PreAdjustBalance = CurrentAccum
        Fields.Adjustment = New Decimal(AccumAdj)
        Fields.PostAdjustBalance = Decimal.Add(Fields.PreAdjustBalance, Fields.Adjustment)
        
        ReDim AssocLinks(0)
        AssocLinks(0).LinkID = IncentiveID
        AssocLinks(0).LinkTypeID = 1
        Fields.AssociatedLinks = AssocLinks
        
        MyCommon.Activity_Log3(Fields)
      End If
    End If
  End Sub
  
  Function IsProperFormat(ByVal UnitType As Integer, ByVal AdjAmt As Double) As Boolean
    Dim FormatOk As Boolean = True
    Dim StrAdjAmt As String = AdjAmt.ToString()
    Dim DecPtPos, CharAfterDec As Integer
    
    DecPtPos = StrAdjAmt.IndexOf(".")
    If (DecPtPos > -1) Then
      CharAfterDec = (StrAdjAmt.Length - (DecPtPos + 1))
    End If

    Select Case UnitType
      Case 1 ' ###,##0
        FormatOk = (DecPtPos = -1)
      Case 2 ' ###,##0.00
        FormatOk = (DecPtPos = -1) OrElse (CharAfterDec >= 0 AndAlso CharAfterDec <= 2)
      Case 3 ' ###,##0.000
        FormatOk = (DecPtPos = -1) OrElse (CharAfterDec >= 0 AndAlso CharAfterDec <= 3)
      Case Else
        FormatOk = True
    End Select
    
    Return FormatOk
  End Function
  
    
  Function FormatAODRecords(ByVal UnitType As Integer, ByVal DisplayAmount As Single, ByVal DeletedAOD As Boolean) As String
    Dim DisplayRecord As String = ""
    
    If UnitType = 1 Then
      If DeletedAOD Then
        If DisplayAmount < 0 Then
          DisplayAmount = DisplayAmount * -1
          DisplayRecord = "+" & Format(DisplayAmount, "###,##0")
        Else
          DisplayRecord = "-" & Format(DisplayAmount, "###,##0")
        End If
      Else
        If DisplayAmount < 0 Then
          'DisplayAmount = DisplayAmount * -1
          DisplayRecord = Format(DisplayAmount, "###,##0")
        Else
          DisplayRecord = "+" & Format(DisplayAmount, "###,##0")
        End If
      End If
    ElseIf UnitType = 2 Then
      If DeletedAOD Then
        If DisplayAmount < 0 Then
          DisplayAmount = DisplayAmount * -1
          DisplayRecord = "+ $" & Format(DisplayAmount, "###,##0.00")
        Else
          DisplayRecord = "- $" & Format(DisplayAmount, "###,##0.00")
        End If
      Else
        If DisplayAmount < 0 Then
          DisplayAmount = DisplayAmount * -1
          DisplayRecord = "- $" & Format(DisplayAmount, "###,##0.00")
        Else
          DisplayRecord = "+ $" & Format(DisplayAmount, "###,##0.00")
        End If
      End If
    ElseIf UnitType = 3 Then
      If DeletedAOD Then
        If DisplayAmount < 0 Then
          DisplayAmount = DisplayAmount * -1
          DisplayRecord = "+" & Format(DisplayAmount, "###,##0.000")
        Else
          DisplayRecord = "-" & Format(DisplayAmount, "###,##0.000")
        End If
      Else
        If DisplayAmount < 0 Then
          'DisplayAmount = DisplayAmount * -1
          DisplayRecord = Format(DisplayAmount, "###,##0.000")
        Else
          DisplayRecord = "+" & Format(DisplayAmount, "###,##0.000")
        End If
      End If
    End If
    
    Return DisplayRecord
  End Function
  
</script>
<script type="text/javascript">
  function ChangeParentDocument() {
    var refreshElem = document.getElementById("RefreshParent");
    if (opener != null && !opener.closed) {
      if (refreshElem != null && refreshElem.value == 'true') {
        if (linkToHH) {
          opener.location = 'customer-offers.aspx?CustPK=<%Sendb(HHPK)%><%Sendb(IIf(HHCardPK > 0, "&CardPK=" & HHCardPK, ""))%>';
        }
        else {
          opener.location = 'customer-offers.aspx?CustPK=<%Sendb(CustomerPK)%><%Sendb(IIf(CardPK > 0, "&CardPK=" & CardPK, ""))%>';
        }
      }
    }
  }
</script>
<%
done:
  Send_BodyEnd("mainform", "accumadj")
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
