<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS" %>


<script runat="server">
  Public MyCommon As New Copient.CommonInc
  Dim CopientFileName As String = "PrintedMessageFeeds.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
    Dim SystemCacheData As ICacheData
  Dim Logix As New Copient.LogixInc
    Dim restrictRewardforRPOS As Boolean = False
  Dim EngineID As Integer = 0
  Dim OfferID As Integer = 0
  Dim Phase As Integer = 0
  Dim PrinterTypeID As Integer = 0
  
  '-------------------------------------------------------------------------------------------------------------
  
  'this function checks to see if the offer that uses this printed message required a customer group (any customer group besides Any Customer)
  Function isOfferTargeted(ByVal EngineID As Integer, ByVal OfferID As Long) As Boolean
    Dim ReturnVal As Boolean = True
    Dim dst As DataTable
    
    'currently this lookup only needs to be run for the UE promo engine
    If EngineID = 9 Then
      'query to see if AnyCustomer is the only conditional non-excluded customer group associated with this offer 
      MyCommon.QueryStr = "select isnull(CG.AnyCustomer, 0) as AnyCustomer " & _
                          "from CPE_IncentiveCustomerGroups as ICG with (NoLock) Inner Join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID and ICG.Deleted=0 and RO.deleted=0 " & _
                          "Inner Join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                          "where RO.IncentiveID=" & OfferID & " and ICG.ExcludedUsers = 0;"
      dst = MyCommon.LRT_Select
      If dst.Rows.Count = 1 Then
        If dst.Rows(0).Item("AnyCustomer") = True Then ReturnVal = False
      End If
    End If
    
    Return ReturnVal
    
  End Function
  
  '-------------------------------------------------------------------------------------------------------------
  
  Sub GetMarkupTags(ByVal EngineID As Integer, ByVal OfferID As Long, ByVal Phase As Integer, ByVal PrinterTypeID As Integer)
    Dim rst As DataTable
    Dim row As DataRow
    Dim DisabledAttribute As String = ""
    Dim CleanID As String = ""
    Dim TextAreaName As String = ""
    Dim ScorecardWorthy As Boolean = False
    Dim CentralRendered As Boolean = False
    Dim NumParams As String = 0
    Dim Param1Phrase As String = ""
    Dim Param2Phrase As String = ""
    Dim Param3Phrase As String = ""
    Dim Param4Phrase As String = ""
    Dim iRewardTypeId As Integer
    Dim sLifetimePointsId As String = ""
    Dim TargetedOffer As Boolean = True
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    CMS.AMS.CurrentRequest.Resolver.AppName = "UEpmsg-feeds.aspx"
        SystemCacheData = CurrentRequest.Resolver.Resolve(Of ICacheData)()
        restrictRewardforRPOS = (SystemCacheData.GetSystemOption_UE_ByOptionId(234) = "1")
    If Phase = 3 Then
      MyCommon.QueryStr = "select Priority from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        If MyCommon.NZ(rst.Rows(0).Item("Priority"), 0) >= 900 Then
          ScorecardWorthy = True
        End If
      End If
    End If
    
        If EngineID = 9 Then
            ScorecardWorthy = True
        End If
        
    Try
      If EngineID = 0 Or EngineID = 1 Then
        iRewardTypeId = 3
        sLifetimePointsId = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(37), "")
      Else
        iRewardTypeId = 4
      End If
      
      TargetedOffer = isOfferTargeted(EngineID, OfferID)
      
      MyCommon.QueryStr = "dbo.pa_Printed_Message_Tags"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
      MyCommon.LRTsp.Parameters.Add("@PrinterTypeID", SqlDbType.Int).Value = PrinterTypeID
      MyCommon.LRTsp.Parameters.Add("@RewardTypeID", SqlDbType.Int).Value = iRewardTypeId
      rst = MyCommon.LRTsp_select
      If (rst.Rows.Count > 0) Then
        For Each row In rst.Rows
          CentralRendered = IIf(row.Item("CentralRendered"), True, False)
          NumParams = MyCommon.NZ(row.Item("NumParams"), 0)
          DisabledAttribute = IIf(MyCommon.NZ(row.Item("PrinterTypeID"), -1) = -1, " disabled=""disabled""", "")
          DisabledAttribute = IIf(PrinterTypeID = 999, "", DisabledAttribute)
          If Not (TargetedOffer) And row.Item("CustomerRequired") = True Then DisabledAttribute = " disabled=""disabled"""
          CleanID = MyCommon.NZ(row.Item("ButtonText"), "")
          CleanID = CleanID.Replace("#", "AMT")
          CleanID = CleanID.Replace("$", "DOL")
          CleanID = CleanID.Replace("/", "Off")
          If (CleanID = "NETDOL") Or (CleanID = "INITIALDOL") Or (CleanID = "EARNEDDOL") Or (CleanID = "REDEEMEDDOL") Then
            Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 1, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
          ElseIf (CleanID = "PTBAL") Or (CleanID = "PTVAL") Or (CleanID = "NETAMT") Or (CleanID = "INITIALAMT") Or (CleanID = "EARNEDAMT") Or (CleanID = "REDEEMEDAMT") Then
            If EngineID = 9 AndAlso TargetedOffer = False Then
              DisabledAttribute = IIf(CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IPointsProgramService)().IsAnyCustomerPointProgramExist(), String.Empty, "disabled=""disabled""")
            End If
            Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 2, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
          ElseIf (CleanID = "SVVAL") Or (CleanID = "SVBAL") Or (CleanID = "SVBALEXP") Or (CleanID = "SVVALEXP") Or (CleanID = "SVLIMIT") Or (CleanID = "SVREDEEM") Then
            If EngineID = 9 AndAlso TargetedOffer = False Then
              DisabledAttribute = IIf(CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IStoredValueProgramService)().IsAnyCustomerSVProgramExist(), String.Empty, "disabled=""disabled""")
            End If
            Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 3, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
          ElseIf (CleanID = "LIFETIMEAMT") Then
            If sLifetimePointsId <> "" Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 8, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            End If
          ElseIf (CleanID = "SCORECARD") Then
            If ScorecardWorthy Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 4, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            End If
          ElseIf (CleanID = "SVSCORECARD") Then
            If ScorecardWorthy Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 6, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            End If
          ElseIf (CleanID = "DSCORECARD") Then
            If ScorecardWorthy Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 7, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            End If
          ElseIf (CleanID = "SVRATIO") Or (CleanID = "SVSCRATIO") Then
            Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 5, this.value);"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
          ElseIf (CentralRendered) AndAlso (NumParams > 0) Then
            If MyCommon.NZ(row.Item("Param1PhraseID"), 0) > 0 Then
              Param1Phrase = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("Param1PhraseID"), 0), LanguageID)
              Param1Phrase = Replace(Param1Phrase, "'", "\'")
            End If
            If MyCommon.NZ(row.Item("Param2PhraseID"), 0) > 0 Then
              Param2Phrase = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("Param2PhraseID"), 0), LanguageID)
              Param2Phrase = Replace(Param2Phrase, "'", "\'")
            End If
            If MyCommon.NZ(row.Item("Param3PhraseID"), 0) > 0 Then
              Param3Phrase = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("Param3PhraseID"), 0), LanguageID)
              Param3Phrase = Replace(Param3Phrase, "'", "\'")
            End If
            If MyCommon.NZ(row.Item("Param4PhraseID"), 0) > 0 Then
              Param4Phrase = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("Param4PhraseID"), 0), LanguageID)
              Param4Phrase = Replace(Param4Phrase, "'", "\'")
            End If
            If (NumParams = 1) Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 11, this.value, '" & Param1Phrase & "');"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            ElseIf (NumParams = 2) Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 21, this.value, '" & Param1Phrase & "', '" & Param2Phrase & "');"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            ElseIf (NumParams = 3) Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 31, this.value, '" & Param1Phrase & "', '" & Param2Phrase & "', '" & Param3Phrase & "');"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            ElseIf (NumParams = 4) Then
              Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 41, this.value, '" & Param1Phrase & "', '" & Param2Phrase & "', '" & Param3Phrase & "', '" & Param4Phrase & "');"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
            End If
                    ElseIf (CleanID = "UPCA" OrElse CleanID = "EAN13" OrElse CleanID = "CODE128" OrElse CleanID = "UPCB" OrElse CleanID = "/BARCODE" OrElse CleanID = "BARCODE" OrElse CleanID = "CODE39") Then
                        If Not restrictRewardforRPOS Then
                            Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""edInsert" & (StrConv(CleanID, VbStrConv.ProperCase)) & "('" & TextAreaName & "');"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
                        End If
          Else
            TextAreaName = IIf(EngineID = Copient.CommonInc.InstalledEngines.CPE, "t1_text", "tier0")
            If (TextAreaName = "tier0") Then
              If (MyCommon.Extract_Val(Request.QueryString("NumTiers")) > 0) Then
                TextAreaName = "tier1"
              End If
            End If
            Sendb("<input type=""button"" id=""ed_" & (StrConv(CleanID, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""edInsert" & (StrConv(CleanID, VbStrConv.ProperCase)) & "('" & TextAreaName & "');"" value=""" & (StrConv(MyCommon.NZ(row.Item("ButtonText"), ""), VbStrConv.ProperCase)) & """ />")
          End If
        Next
      End If
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
</script>


<%
    '-----------------------------------------------------------------------------------------
    'Execution starts here ... 

    MyCommon.AppName = "PrintedMessageFeeds.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)


    EngineID = MyCommon.Extract_Val(Request.QueryString("EngineID"))
    OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
    CheckIfValidOffer(MyCommon, OfferID)
    Phase = MyCommon.Extract_Val(Request.QueryString("Phase"))
    PrinterTypeID = MyCommon.Extract_Val(Request.QueryString("PrinterTypeID"))

    If (Request.QueryString("Mode") = "MarkupTags") Then
        Response.Expires = 0
        Response.Clear()
        Response.ContentType = "text/html"
        GetMarkupTags(EngineID, OfferID, Phase, PrinterTypeID)
    Else
        Send("<b>" & Copient.PhraseLib.Lookup("feeds.noarguments", LanguageID) & "!</b>")
        Send(Request.RawUrl)
    End If
    Response.Flush()
    Response.End()

%>