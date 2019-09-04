<%@ Page Language="vb" Debug="true" CodeFile="../LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  Dim CopientFileName As String = "XMLFeeds.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim AdminUserID As Long
  Dim Category As String = "0"
  Dim Description As String = ""
  Dim CategoryDesc As String = ""
  Dim OfferName As String = ""
  Dim TempDate As Date
  Dim CardPK As Long
  Dim CustPK As Long
  
  MyCommon.AppName = "CAM-Feeds.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)

  If (LanguageID = 0) Then
    LanguageID = MyCommon.Extract_Val(Request.QueryString("LanguageID"))
  End If
  
 
  If (Request.QueryString("CAMOfferTransactions") <> "") Then
    'Send("CustPK : " & Request.Form("CustPK") & "  OfferID: " & Request.Form("OfferID"))
    CAMOfferTransactions(Long.Parse(Request.Form("CustPK")), Request.Form("ExtCustID"), _
                         Long.Parse(Request.Form("OfferID")), Logix)
  ElseIf (Request.QueryString("CAMTransactionOffers") <> "") Then
    CAMTransactionOffers(Long.Parse(Request.QueryString("CustPK")), Long.Parse(Request.QueryString("OfferID")))
  ElseIf (Request.QueryString("CAMProgramTransactions") <> "") Then
    CAMProgramTransactions(Long.Parse(Request.Form("CustPK")), Request.Form("ExtCustID"), _
                           Long.Parse(Request.Form("ProgramID")), Logix)
  ElseIf (Request.QueryString("CAMOfferSelection") <> "") Then
    If Not Date.TryParse(Request.Form("TransDate"), TempDate) Then
      TempDate = Date.Now
    End If
    If Request.Form("CustPK") <> "" Then
      CustPK = MyCommon.Extract_Val(Request.Form("CustPK"))
      CardPK = MyCommon.Extract_Val(Request.Form("CardPK"))
      CAMOfferSelection(CustPK, CardPK, Request.Form("LogixTransNum"), _
                        Request.Form("TransNum"), TempDate, Request.Form("SearchText"), Logix)
    End If
  Else
    Send("<b>" & Copient.PhraseLib.Lookup("term.noarguments", LanguageID) & "</b>")
  End If
        
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>

<script runat="server">
  Public DefaultLanguageID As Integer = 1
  Public MyCommon As New Copient.CommonInc
  
  ' retrieve all of a customer's transactions for a given offer
  Sub CAMOfferTransactions(ByVal CustomerPK As Long, ByVal ExtCustomerID As String, ByVal OfferID As Long, ByRef Logix As Copient.LogixInc)
    Dim bParsed, IsPtsOffer As Boolean
    Dim dt, dt2 As DataTable
    Dim row As DataRow
    Dim DisabledPtsAdj As String = ""
    Dim TransNumber As String = ""
    Dim LogixTransNum As String = ""
    Dim TransDate As New Date(1980, 1, 1)
    Dim TermNum As String = ""
    Dim TransNum As String = ""
    Dim ExtLocationCode As String = ""
    Dim RedemptionCt As Integer = 0
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    Try
      MyCommon.Open_LogixWH()
      MyCommon.Open_LogixRT()
      MyCommon.Open_LogixXS()
      
      If (Request.QueryString("Lang") <> "") Then
        bParsed = Integer.TryParse(Request.QueryString("lang"), LanguageID)
        If (Not bParsed) Then LanguageID = 1
      End If
      
      MyCommon.QueryStr = "dbo.pa_CustomerOfferHasPointsProgram"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@HasPointsProgram", SqlDbType.Bit).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      IsPtsOffer = MyCommon.LRTsp.Parameters("@HasPointsProgram").Value
      MyCommon.Close_LRTsp()
      
      MyCommon.QueryStr = "select CustomerPrimaryExtId, Max(TransDate) as TransactionDate, ExtLocationCode, sum(RedemptionAmount) as RedemptionAmount, " & _
                          "sum(RedemptionCount) as RedemptionCount, TerminalNum, TransNum, LogixTransNum, count(*) as DetailRecords " & _
                          "from TransRedemptionView with (NoLock) where CustomerTypeID=2 and CustomerPrimaryExtId='" & ExtCustomerID & "' and OfferID=" & OfferID & " " & _
                          "Group by CustomerPrimaryExtId, TransNum, TerminalNum, ExtLocationCode, LogixTransNum " & _
                          "order by TransactionDate desc;"
      dt = MyCommon.LWH_Select
      Send("  <table id=""tblTrans" & OfferID & """ style=""margin-left:15px;width:90%;background-color:#dddddd;"" cellpadding=""0"" cellspacing=""0"">")
      If dt.Rows.Count > 0 Then
        Send("    <tr style=""background-color:#cccccc;""><td colspan=""6"" style=""text-align:center;border-bottom:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.transactions", LanguageID) & "</b></td></tr>")
        Send("    <tr style=""background-color:#b8b8b8;"">")
        Send("      <td style=""text-align:center;border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.redeemed", LanguageID) & "</b></td>")
        Send("    </tr>")
        For Each row In dt.Rows
          LogixTransNum = MyCommon.NZ(row.Item("LogixTransNum"), "").ToString.Trim
          RedemptionCt = MyCommon.NZ(row.Item("RedemptionCount"), 0)
          
          ' override transaction data with the trans history table as it is the table of record
          MyCommon.QueryStr = "select CustomerPrimaryExtId, TransDate as TransactionDate, " & _
                              "       ExtLocationCode, TerminalNum, POSTransNum " & _
                              "from TransHistory with (NoLock) " & _
                              "where LogixTransNum = '" & LogixTransNum & "';"
          dt2 = MyCommon.LWH_Select
          If dt2.Rows.Count > 0 Then
            TransDate = MyCommon.NZ(dt2.Rows(0).Item("TransactionDate"), New Date(1980, 1, 1))
            ExtLocationCode = MyCommon.NZ(dt2.Rows(0).Item("ExtLocationCode"), "")
            TermNum = MyCommon.NZ(dt2.Rows(0).Item("TerminalNum"), "")
            TransNum = MyCommon.NZ(dt2.Rows(0).Item("POSTransNum"), "")
          Else
            TransDate = MyCommon.NZ(row.Item("TransactionDate"), New Date(1980, 1, 1))
            ExtLocationCode = MyCommon.NZ(row.Item("ExtLocationCode"), "")
            TermNum = MyCommon.NZ(row.Item("TerminalNum"), "")
            TransNum = MyCommon.NZ(row.Item("TransNum"), "")
          End If
          
          Send("    <tr>")
          If (IsPtsOffer) Then
            If (Logix.UserRoles.AccessPointsBalances = False) Then
              DisabledPtsAdj = " disabled=""disabled"""
            Else
              DisabledPtsAdj = ""
            End If
            Sendb("   <td style=""text-align:center;""><input type=""button"" class=""adjust"" id=""ptsAdj" & OfferID & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " title=""" & Copient.PhraseLib.Lookup("customer-inquiry.adjust-for-transaction", LanguageID) & """ ")
            Send("onClick=""javascript:openPopup('CAM-point-adjust.aspx?LogixTransNum=" & LogixTransNum & "&OfferID=" & OfferID & "&CustomerPK=" & CustomerPK & "&CustomerExtId=" & ExtCustomerID & "');"" /><span style=""margin-left:8px;""></span></td>")
          End If
          Send("      <td>" & TransDate.ToString & "</td>")
          Send("      <td>" & ExtLocationCode & "</td>")
          Send("      <td>" & TermNum & "</td>")
          Send("      <td style=""word-break:break-all"">" & TransNum & "</td>")
          Send("      <td>" & RedemptionCt & "</td>")
          Send("    </tr>")
        Next
      Else
        Send("    <tr style=""background-color:#cccccc;""><td colspan=""6"" style=""text-align:center;border-bottom:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("customer-inquiry.nohistory", LanguageID) & "</b></td></tr>")
      End If
      Send("  </table>")
    Catch ex As Exception
      Send("<td colspan=""6"">")
      Send(ex.ToString)
      Send("</td>")
    Finally
      MyCommon.Close_LogixWH()
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
    End Try
    
  End Sub
  
  ' retrieve all of the offers redeemed within a customer's transaction
  Sub CAMTransactionOffers(ByVal CustomerPK As Long, ByVal OfferID As Long)
    Dim rst As DataTable
    Dim row As DataRow
    Dim LanguageID As Integer = 1
    Dim bParsed As Boolean
    Dim TotalRedeemCt As Integer = 0
    Dim TotalRedeemAmt As Double = 0.0
    Dim OfferDesc As String = ""
    Dim OfferName As String = ""
    Dim CustomerExtID As String = ""
        Dim CustName As String = ""
        Dim MyCryptLib As New Copient.CryptLib

    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    Try
      MyCommon.Open_LogixWH()
      MyCommon.Open_LogixRT()
      MyCommon.Open_LogixXS()
      
      If (Request.QueryString("Lang") <> "") Then
        bParsed = Integer.TryParse(Request.QueryString("lang"), LanguageID)
        If (Not bParsed) Then LanguageID = 1
      End If
    
      ' get the customers card number from the unique identifier
      MyCommon.QueryStr = "select PrimaryExtID, FirstName, LastName from Customers with (NoLock) where CustomerPK=" & CustomerPK
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
                CustomerExtID = MyCommon.NZ(MyCryptLib.SQL_StringDecrypt(rst.Rows(0).Item("PrimaryExtID")), "")
        CustName = MyCommon.NZ(rst.Rows(0).Item("FirstName"), "") & " " & MyCommon.NZ(rst.Rows(0).Item("LastName"), "")
      End If
    
      ' get the offers description for display purposes
      MyCommon.QueryStr = "select IncentiveName, Description from CPE_Incentives with (NoLock) where IncentiveID = " & OfferID
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        OfferName = MyCommon.NZ(rst.Rows(0).Item("IncentiveName"), "")
        OfferDesc = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
      End If
    
      Send("<br class=""half"" />")
      Send("<table style=""width:95%;font-size:10pt;color:#333333;border:solid 1px #333333;"" cellpadding=""0"" cellspacing=""0"" summary=""" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & """>")
      Send("    <tr style=""background-color:#dddddd;"">")
      Send("      <td><b>" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & ":</b></td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & CustomerExtID & " </td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ": " & CustName & " </td>")
      Send("    </tr>")
      Send("    <tr style=""background-color:#eeeeee;"">")
      Send("      <td><b>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</b>:</td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ": " & OfferID & " </td>")
      Send("      <td>" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ": " & OfferName & " </td>")
      Send("    </tr>")
      Send("    <tr style=""background-color:#eeeeee;"">")
      Send("      <td></td>")
      Send("      <td colspan=""2"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ": " & OfferDesc & "</td>")
      Send("    </tr>")
      Send("</table>")
      Send("<br class=""half"" />")
      Send("<table style=""width:95%;"" summary=""" & Copient.PhraseLib.Lookup("term.list", LanguageID) & """>")
      Send("    <thead>")
      Send("    <tr>")
      Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & "</th>")
      Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</th>")
      Send("        <th scope=""col"">" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:center;"">" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.amount", LanguageID) & "</th>")
      Send("        <th scope=""col"" style=""text-align:right;"">" & Copient.PhraseLib.Lookup("term.redemptions", LanguageID) & "</th>")
      Send("    </tr>")
      Send("    </thead>")
      Send("    <tbody>")
    
      MyCommon.QueryStr = "select ExtLocationCode, RedemptionCount, RedemptionAmount, TransDate, TerminalNum, TransNum from TransRedemptionView with (NoLock) " & _
                          "where OfferID = " & OfferID & " and CustomerTypeID=2 and CustomerPrimaryExtID = '" & CustomerExtID & "' order by TransDate;"
      rst = MyCommon.LWH_Select
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          Send("    <tr>")
          Sendb("       <td><input type=""button"" class=""adjust"" id=""ptsAdj" & MyCommon.NZ(row.Item("OfferID"), "") & """ name=""ptsAdj"" value=""P"" title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
          Send("onClick=""javascript:openPopup('CAM-point-adjust.aspx?OfferID=" & MyCommon.NZ(row.Item("OfferID"), "") & "&CustomerPK=" & CustomerPK & "');"" /></td>")
          Send("        <td>" & MyCommon.NZ(row.Item("TransDate"), "") & "</td>")
          Send("        <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), "") & "</td>")
          Send("        <td style=""text-align:center;"">" & MyCommon.NZ(row.Item("TerminalNum"), "") & "</td>")
          Send("        <td style=""text-align:center;word-break:break-all"">" & MyCommon.NZ(row.Item("TransNum"), "") & "</td>")
          Send("        <td style=""text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionAmount"), "") & "</td>")
          Send("        <td style=""text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionCount"), "") & "</td>")
          Send("    </tr>")
          TotalRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
          TotalRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
        Next
        Send("    <tr style=""height:15px;"">")
        Send("        <td colspan=""5""></td>")
        Send("        <td><hr></td>")
        Send("        <td><hr></td>")
        Send("    </tr>")
        Send("    <tr>")
        Send("        <td colspan=""3"">" & Copient.PhraseLib.Lookup("CPE_accum-adj-transamt", LanguageID) & ": " & rst.Rows.Count & "</td>")
        Send("        <td colspan=""2""></td>")
        Send("        <td style=""text-align:right;"">" & Format(TotalRedeemAmt, "$ #,###,##0.000") & "</td>")
        Send("        <td style=""text-align:right;"">" & TotalRedeemCt & "</td>")
        Send("    </tr>")
      Else
        Send("    <tr>")
        Send("        <td colspan=""7"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("CPE_accum-adj-notranshistory", LanguageID) & "</i></td>")
        Send("    </tr>")
      End If
    
      Send("</tbody>")
      Send("</table>")
      Send("<br /><br />")
      Send("<center>")
      Send("<input type=""button"" id=""btnNewTrans"" name=""btnNewTrans"" value=""" & Copient.PhraseLib.Lookup("term.newtransaction", LanguageID) & """ />")
      Send("<br />")
      Send("<small>" & Format(DateTime.Now, "dd MMM yyyy, HH:mm:ss") & "</small>")
      Send("</center>")
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixWH()
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
    End Try
    
  End Sub

  ' retrieve all of a customer's transactions for a given points program
  Sub CAMProgramTransactions(ByVal CustomerPK As Long, ByVal ExtCustomerID As String, ByVal ProgramID As Long, ByRef Logix As Copient.LogixInc)
    Dim bParsed As Boolean
    Dim dt, dt2, dtOffers As DataTable
    Dim row As DataRow
    Dim DisabledPtsAdj As String = ""
    Dim TransNumber As String = ""
    Dim OfferIDList As String = ""
    Dim OfferID As Long
    Dim LogixTransNum As String = ""
    Dim TransDate As New Date(1980, 1, 1)
    Dim TermNum As String = ""
    Dim TransNum As String = ""
    Dim ExtLocationCode As String = ""
    Dim RedemptionCt As Integer = 0
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    Try
      MyCommon.Open_LogixWH()
      MyCommon.Open_LogixRT()
      MyCommon.Open_LogixXS()
      
      If (Request.QueryString("Lang") <> "") Then
        bParsed = Integer.TryParse(Request.QueryString("lang"), LanguageID)
        If (Not bParsed) Then LanguageID = 1
      End If
      
      ' find all the offers where this points program is used
      MyCommon.QueryStr = "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_IncentivePointsGroups IPG with (NoLock) " & _
                          "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                          "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                          "INNER JOIN PointsPrograms PP with (NoLock) on IPG.ProgramID = PP.ProgramID " & _
                          "WHERE IPG.ProgramID = " & ProgramID & " and IPG.Deleted = 0 and RO.Deleted=0 and I.Deleted=0 and PP.Deleted=0 " & _
                          "UNION " & _
                          "SELECT DISTINCT I.IncentiveName as Name, I.IncentiveID as OfferID from CPE_DeliverablePoints DP " & _
                          "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=DP.RewardOptionID " & _
                          "INNER JOIN CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                          "WHERE DP.Deleted=0 and RO.Deleted=0 and I.Deleted=0 and ProgramID=" & ProgramID
      dtOffers = MyCommon.LRT_Select

      If dtOffers.Rows.Count > 0 Then
        For Each row In dtOffers.Rows
          If OfferIDList <> "" Then OfferIDList &= ","
          OfferIDList &= MyCommon.NZ(row.Item("OfferID"), -1)
        Next
      End If
      
      MyCommon.QueryStr = "select CustomerPrimaryExtId, Max(TransDate) as TransactionDate, ExtLocationCode, sum(RedemptionAmount) as RedemptionAmount, " & _
              "sum(RedemptionCount) as RedemptionCount, TerminalNum, TransNum, LogixTransNum, OfferID, count(*) as DetailRecords " & _
              "from TransRedemptionView with (NoLock) where CustomerTypeID=2 and CustomerPrimaryExtId ='" & ExtCustomerID & "' and OfferID in (" & OfferIDList & ") " & _
              "Group by CustomerPrimaryExtId, TransNum, TerminalNum, ExtLocationCode, LogixTransNum, OfferID " & _
              "order by TransactionDate desc;"

      dt = MyCommon.LWH_Select
      Send("  <table id=""tblTrans" & ProgramID & """ style=""margin-left:15px;width:90%;background-color:#dddddd;"" cellpadding=""0"" cellspacing=""0"">")
      If dt.Rows.Count > 0 Then
        Send("    <tr style=""background-color:#cccccc;""><td colspan=""7"" style=""text-align:center;border-bottom:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.transactions", LanguageID) & "</b></td></tr>")
        Send("    <tr style=""background-color:#b8b8b8;"">")
        Send("      <td style=""text-align:center;border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.store", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.offerid", LanguageID) & "</b></td>")
        Send("      <td style=""border:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("term.redeemed", LanguageID) & "</b></td>")
        Send("    </tr>")
        For Each row In dt.Rows
          OfferID = MyCommon.NZ(row.Item("OfferID"), -1)
          LogixTransNum = MyCommon.NZ(row.Item("LogixTransNum"), "").ToString.Trim
          RedemptionCt = MyCommon.NZ(row.Item("RedemptionCount"), 0)
          
          ' override transaction data with the trans history table as it is the table of record
          MyCommon.QueryStr = "select CustomerPrimaryExtId, TransDate as TransactionDate, " & _
                              "       ExtLocationCode, TerminalNum, POSTransNum " & _
                              "from TransHistory with (NoLock) " & _
                              "where LogixTransNum = '" & LogixTransNum & "';"
          dt2 = MyCommon.LWH_Select
          If dt2.Rows.Count > 0 Then
            TransDate = MyCommon.NZ(dt2.Rows(0).Item("TransactionDate"), New Date(1980, 1, 1))
            ExtLocationCode = MyCommon.NZ(dt2.Rows(0).Item("ExtLocationCode"), "")
            TermNum = MyCommon.NZ(dt2.Rows(0).Item("TerminalNum"), "")
            TransNum = MyCommon.NZ(dt2.Rows(0).Item("POSTransNum"), "")
          Else
            TransDate = MyCommon.NZ(row.Item("TransactionDate"), New Date(1980, 1, 1))
            ExtLocationCode = MyCommon.NZ(row.Item("ExtLocationCode"), "")
            TermNum = MyCommon.NZ(row.Item("TerminalNum"), "")
            TransNum = MyCommon.NZ(row.Item("TransNum"), "")
          End If
          Send("    <tr>")
          If (Logix.UserRoles.AccessPointsBalances = False) Then
            DisabledPtsAdj = " disabled=""disabled"""
          Else
            DisabledPtsAdj = ""
          End If
          Sendb("   <td style=""text-align:center;""><input type=""button"" class=""adjust"" id=""ptsAdj" & OfferID & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " title=""" & Copient.PhraseLib.Lookup("customer-inquiry.adjust-for-transaction", LanguageID) & """ ")
          Send("onClick=""javascript:openPopup('CAM-point-adjust-program.aspx?LogixTransNum=" & LogixTransNum & "&ProgramID=" & ProgramID & "&CustomerPK=" & CustomerPK & "&CustomerExtId=" & ExtCustomerID & "');"" /><span style=""margin-left:8px;""></span></td>")
          Send("      <td>" & TransDate.ToString & "</td>")
          Send("      <td>" & ExtLocationCode & "</td>")
          Send("      <td>" & TermNum & "</td>")
          Send("      <td style=""word-break:break-all"">" & TransNum & "</td>")
          Send("      <td>" & MyCommon.NZ(row.Item("OfferID"), "").ToString & "</td>")
          Send("      <td>" & RedemptionCt & "</td>")
          Send("    </tr>")
        Next
      Else
        Send("    <tr style=""background-color:#cccccc;""><td colspan=""7"" style=""text-align:center;border-bottom:solid 1px #c6c6c6;""><b>" & Copient.PhraseLib.Lookup("customer-inquiry.nohistory", LanguageID) & "</b></td></tr>")
      End If
      Send("  </table>")
      
    Catch ex As Exception
      Send("<td colspan=""6"">")
      Send(ex.ToString)
      Send("</td>")
    Finally
      MyCommon.Close_LogixWH()
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
    End Try
    
  End Sub

  ' retrieve all the CAM offers 
  Sub CAMOfferSelection(ByVal CustomerPK As Long, ByVal CardPK As Long, ByVal LogixTransNum As String, ByVal TransNum As String, _
                        ByVal TransDate As Date, ByVal SearchText As String, ByRef Logix As Copient.LogixInc)
    Dim dt, dtOffers, dtPrograms As DataTable
    Dim row As DataRow
    Dim sortedRows() As DataRow
    Dim cgXML As String = ""
    Dim AllCAMCardholdersID As Long = 0
    Dim OfferList As String = ""
    Dim SearchFilter As String = ""
    Dim StatusTable As Hashtable
    Dim OfferStatusCode As Copient.LogixInc.STATUS_FLAGS
    Dim OfferStatus As String = ""
    Dim DisabledPtsAdj As String = ""
    Dim ProgramID As Long = 0
    Dim SearchID As Long = 0
    Dim Name As String = ""
    Dim Description As String = ""
    Dim RowCounter As Integer = 0
    Dim IsExpired As Boolean = False
    
    Try
      MyCommon.Open_LogixRT()
      MyCommon.Open_LogixXS()
      
      If SearchText Is Nothing Then SearchText = ""
      SearchText = SearchText.Trim
      SearchID = MyCommon.Extract_Val(SearchText)
      
      ' find the CardPK if one is not sent
      MyCommon.QueryStr = "select Top 1 CardPK from CardIDs with (NoLock) where CardtypeID=2 and CustomerPK=" & CustomerPK
      dt = MyCommon.LXS_Select
      If dt.Rows.Count > 0 Then
        CardPK = MyCommon.NZ(dt.Rows(0).Item("CardPK"), 0)
      End If
      
      ' find all the offers for which a customer is eligible, then determine which of those offers have points programs
      cgXML = "<customergroups>"
      MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1;"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0) Then
        AllCAMCardholdersID = MyCommon.NZ(dt.Rows(0).Item("CustomerGroupID"), 0)
        cgXML &= "<id>" & AllCAMCardholdersID & "</id>"
      End If

      MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0"
      dt = MyCommon.LXS_Select()

      If dt.Rows.Count > 0 Then
        For Each row In dt.Rows
          cgXML &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "0") & "</id>"
        Next
      End If
      cgXML &= "</customergroups>"
            
      MyCommon.QueryStr = "dbo.pa_CAM_CustomerTranOffers"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = cgXML
      MyCommon.LRTsp.Parameters.Add("@IsEmployee", SqlDbType.Bit).Value = 0
      MyCommon.LRTsp.Parameters.Add("@Favorite", SqlDbType.Bit).Value = 0
      MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Value = 1
      dtOffers = MyCommon.LRTsp_select
      MyCommon.Close_LRTsp()

      ' draw the title bar and the close button
      Send("<h2 style=""width:100%;height:21px;color:#ffffff;background-color:#000080;"">&nbsp;" & Copient.PhraseLib.Lookup("cam.select-for-adjust", LanguageID) & "</h2>")
      Sendb("<span style=""position:relative;top:-20px;float:right;color:#ffffff;font-size:11pt;padding-right:2px;cursor:pointer;cursor:hand;font-weight:bold;"" ")
      Send("       onclick=""closeDialog();"" title=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """>X</span>")
      Send("<br class=""half"" />")
      
      Send("<table style=""width:480px;"">")
      Send("  <tr>")
      Send("    <td style=""font-size:8pt;width:300px;"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":" & IIf(TransNum = "", Copient.PhraseLib.Lookup("term.new", LanguageID), TransNum & "&nbsp;&nbsp;(<a href=""javascript:showSelectDialog('', '', '');"">" & Copient.PhraseLib.Lookup("term.new", LanguageID) & "</a>)"))
      Send("      <br />" & Copient.PhraseLib.Lookup("term.date", LanguageID) & ":" & TransDate.ToString("dd MMM yyyy, HH:mm:ss"))
      Send("    </td>")
      Send("    <td style=""text-align:right;"">")
      Send("      <input type=""text"" id=""txtSearch"" name=""txtSearch"" class=""short"" value=""" & SearchText.Trim & """  onkeydown=""handleSearchKeyDown(event, '" & LogixTransNum & "','" & TransNum & "','" & TransDate.ToString("dd MMM yyyy, HH:mm:ss") & "');"" />")
      Send("      <input type=""button"" name=""btnSearch"" id=""btnSearch"" value=""" & Copient.PhraseLib.Lookup("term.search", LanguageID) & """ onclick=""searchFromDialog('" & LogixTransNum & "','" & TransNum & "','" & TransDate.ToString("dd MMM yyyy, HH:mm:ss") & "');"" />")
      Send("    </td>")
      Send("  </tr>")
      Send("</table>")
      'Send("<br class=""half"" />")
      
      ' draw the offer results
      Send("<h2 style=""margin-left:5px;width:475px;background-color:#c8c8c8;"">" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & "</h2>")
      Send("<table style=""margin-left:10px;width:470px;"">")
      Send("  <thead>")
      Send("    <tr>")
      Send("      <th style=""width:40px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & "</th>")
      Send("      <th style=""width:70px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
      Send("      <th style=""width:100px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
      Send("      <th>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & "</th>")
      Send("    </tr>")
      Send("   </thead>")
      Send("   <tbody>")
      
      ' restrict the adjustable programs to those offer for which the customer is eligible
      If dtOffers.Rows.Count > 0 Then
        sortedRows = dtOffers.Select("", "Name")
        
        ' use the transaction date for determining the status of the offer 
        StatusTable = LoadOfferStatuses(sortedRows, TransDate, MyCommon, Logix)

        ' get only the offers that were active or expired on the transaction date
        For Each row In sortedRows
          OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
          OfferStatus = StatusTable.Item(MyCommon.NZ(row.Item("OfferID"), "0").ToString)
          If (OfferStatus IsNot Nothing) Then
            OfferStatusCode = OfferStatus
          End If
          
          IsExpired = (OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_EXPIRED)
          If (IsExpired OrElse OfferStatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) Then
            Name = MyCommon.NZ(row.Item("Name"), "").ToString.ToUpper
            Description = MyCommon.NZ(row.Item("Description"), "").ToString.ToUpper
            If (SearchText = "" OrElse MyCommon.NZ(row.Item("OfferID"), -1) = SearchID OrElse Name.IndexOf(SearchText.ToUpper) > -1 OrElse Description.IndexOf(SearchText.ToUpper) > -1) Then
              RowCounter += 1
              Send("  <tr style=""background-color:#" & IIf((RowCounter Mod 2) = 0, "e8e8e8", "f0f0f0") & """>")
              IIf(Logix.UserRoles.AccessPointsBalances = False, DisabledPtsAdj = " disabled=""disabled""", DisabledPtsAdj = "")
              Sendb("    <td style=""text-align:center;""><input type=""button"" class=""adjust"" id=""ptsAdj" & MyCommon.NZ(row.Item("OfferID"), "") & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
              Send("onClick=""javascript:openPopup('CAM-point-adjust.aspx?LogixTransNum=" & LogixTransNum & "&OfferID=" & MyCommon.NZ(row.Item("OfferID"), "") & "&CustomerPK=" & CustomerPK & IIf(CardPK > 0, "&CardPK=" & CardPK, "") & "');"" />")
              Send("    </td>")
              Send("    <td>" & MyCommon.NZ(row.Item("OfferID"), 0) & "</td>")
              Send("    <td>" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & IIf(IsExpired, " (<i>" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & "</i>)", "") & "</td>")
              Send("    <td>" & MyCommon.NZ(row.Item("Description"), "&nbsp;") & "</td>")
              Send("  </tr>")
            End If
            
            If OfferList <> "" Then OfferList &= ","
            OfferList &= MyCommon.NZ(row.Item("OfferID"), -1)
          End If
        Next

        If RowCounter = 0 Then
          Send("  <tr style=""background-color:#f0f0f0" & """>")
          Send("    <td colspan=""4"" style=""text-align:center;""><i>" & Copient.PhraseLib.Lookup("phierarchy.noresultsfound", LanguageID) & "</i></td>")
          Send("  </tr>")
        End If

        If OfferList.Trim = "" Then OfferList = "-1"
        SearchFilter &= " and INC.IncentiveID IN (" & OfferList & ")"
        If SearchText <> "" Then
          SearchFilter &= " and (PP.ProgramID=" & SearchID & " or PP.ProgramName like N'%" & MyCommon.Parse_Quotes(SearchText) & "%' or PP.Description like N'%" & MyCommon.Parse_Quotes(SearchText) & "%') "
        End If
      End If
      Send("    </tbody>")
      Send("  </table>")
      Send("  <br />")
      
      ' draw the points programs results
      Send("<h2 style=""margin-left:5px;width:475px;background-color:#c8c8c8;"">" & Copient.PhraseLib.Lookup("term.pointsprograms", LanguageID) & "</h2>")
      Send("<table style=""margin-left:10px;width:470px;"">")
      Send("  <thead>")
      Send("    <tr>")
      Send("      <th style=""width:40px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & "</th>")
      Send("      <th style=""width:70px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
      Send("      <th style=""width:100px;"" scope=""col"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & "</th>")
      Send("      <th>" & Copient.PhraseLib.Lookup("term.description", LanguageID) & "</th>")
      Send("    </tr>")
      Send("   </thead>")
      Send("   <tbody>")
      
      MyCommon.QueryStr = "select distinct PP.ProgramID, PP.ProgramName, PP.Description from PointsPrograms AS PP with (NoLock) " & _
                            "inner join CPE_IncentivePointsGroups AS IPG with (NoLock) on IPG.ProgramID = PP.ProgramID and IPG.Deleted=0 " & _
                            "inner join CPE_RewardOptions AS RO with (NoLock) on RO.RewardOptionID = IPG.RewardOptionID and RO.Deleted=0 " & _
                            "inner join CPE_Incentives AS INC with (NoLock) on INC.IncentiveID = RO.IncentiveID and INC.Deleted=0 and INC.EngineID=6 " & _
                            "where PP.Deleted=0 " & SearchFilter & " " & _
                            "union " & _
                            "select distinct PP.ProgramID, PP.ProgramName, PP.Description from PointsPrograms AS PP with (NoLock) " & _
                            "inner join CPE_DeliverablePoints AS DPT with (NoLock) on DPT.ProgramID = PP.ProgramID and DPT.Deleted=0 " & _
                            "inner join CPE_Deliverables as DEL with (NoLock) on DEL.OutputID = DPT.PKID and DEL.DeliverableTypeID=8 and DEL.Deleted=0 " & _
                            "inner join CPE_RewardOptions AS RO with (NoLock) on RO.RewardOptionID = DEL.RewardOptionID and RO.Deleted=0 " & _
                            "inner join CPE_Incentives AS INC with (NoLock) on INC.IncentiveID = RO.IncentiveID and INC.Deleted=0 and INC.EngineID=6 " & _
                            "where PP.Deleted=0 " & SearchFilter & " " & _
                            "order by PP.ProgramName; "
      dtPrograms = MyCommon.LRT_Select
      If dtPrograms.Rows.Count > 0 Then
        RowCounter = 0
        For Each row In dtPrograms.Rows
          RowCounter += 1
          ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
          Send("  <tr style=""background-color:#" & IIf((RowCounter Mod 2) = 0, "e8e8e8", "f0f0f0") & """>")
          IIf(Logix.UserRoles.AccessPointsBalances = False, DisabledPtsAdj = " disabled=""disabled""", DisabledPtsAdj = "")
          Sendb("   <td style=""text-align:center;""><input type=""button"" class=""adjust"" id=""ptsAdj" & ProgramID & """ name=""ptsAdj"" value=""P""" & DisabledPtsAdj & " title=""" & Copient.PhraseLib.Lookup("term.adjust", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.points", LanguageID), VbStrConv.Lowercase) & """ ")
          Send("        onClick=""javascript:openPopup('CAM-point-adjust-program.aspx?LogixTransNum=" & LogixTransNum & "&ProgramID=" & ProgramID & "&CustomerPK=" & CustomerPK & "');"" /><span style=""margin-left:8px;""></span></td>")
          Send("    </td>")
          Send("    <td>" & ProgramID & "</td>")
          Send("    <td>" & MyCommon.NZ(row.Item("ProgramName"), "&nbsp;") & "</td>")
          Send("    <td>" & MyCommon.NZ(row.Item("Description"), "&nbsp;") & "</td>")
          Send("  </tr>")
        Next
      Else
        Send("  <tr style=""background-color:#f0f0f0" & """>")
        Send("    <td colspan=""4"" style=""text-align:center;""><i>" & Copient.PhraseLib.Lookup("phierarchy.noresultsfound", LanguageID) & "</i></td>")
        Send("  </tr>")
      End If
      Send("  </tbody>")
      Send(" </table>")
      Send(" <br />")
      Send(" <center>")
      Send("  <input type=""button"" name=""btnClose"" id=""btnClose"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """ onclick=""closeDialog();"" />")
      Send(" </center>")
      
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon.Close_LogixXS()
    End Try
      
  End Sub
      
  Private Function LoadOfferStatuses(ByVal rows() As DataRow, ByVal StatusDate As Date, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc) As Hashtable
    Dim Statuses As New Hashtable(200)
    Dim i, ct, activeCt As Integer
    Dim OfferList() As String = Nothing
    
    ct = rows.Length
    If (ct > 0) Then
      ReDim OfferList(ct - 1)
      For i = 0 To ct - 1
        If rows(i).RowState <> DataRowState.Deleted Then
          OfferList(i) = MyCommon.NZ(rows(i).Item("OfferID"), "0")
          activeCt += 1
        End If
      Next
      ' trim the offer list array to remove the empty elements caused by the filtered-out rows
      If (activeCt >= 1) Then
        ReDim Preserve OfferList(activeCt - 1)
        Statuses = Logix.GetStatusForOffers(OfferList, LanguageID, StatusDate)
      End If
    End If
    
    Return Statuses
  End Function
  
  </script>
