<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<script runat="server">

  Public CopientFileName As String = "CashierMessageFeeds.aspx"
  Public CopientFileVersion As String = "7.3.1.138972"
  Public CopientProject As String = "Copient Logix"
  Public CopientNotes As String = ""
  Public MyCommon As New Copient.CommonInc
  Public Logix As New Copient.LogixInc

    
  Sub GetMarkupTags(ByVal EngineID As Integer, ByVal PrinterTypeID As Integer)
    Dim rst As DataTable
    Dim row As DataRow
    Dim DisabledAttribute As String = ""
    Dim cleanid As String = ""
    Dim TextAreaName As String = ""
    Dim OfferID As Long = MyCommon.NZ(Request.QueryString("OfferID"), 0)
    Dim isAnyCustomer = False
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    CMS.AMS.CurrentRequest.Resolver.AppName = "CashierMessageFeeds.aspx"
    
    Try
      MyCommon.QueryStr = "dbo.pa_Cashier_Message_Tags"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
      MyCommon.LRTsp.Parameters.Add("@RewardTypeID", SqlDbType.Int).Value = IIf(EngineID = 1, 4, 9)
      rst = MyCommon.LRTsp_select
      If (rst.Rows.Count > 0) Then
        isAnyCustomer = CPEOffer_Has_AnyCustomer(MyCommon, OfferID)
        If OfferID > 0 AndAlso isAnyCustomer AndAlso EngineID <> 9 Then
          DisabledAttribute = "disabled=""disabled"""
        End If
        For Each row In rst.Rows
          cleanid = row.Item("Tag")
          cleanid = cleanid.Replace("#", "Amt")
          cleanid = cleanid.Replace("$", "Dol")
          cleanid = cleanid.Replace("/", "Off")
          If (cleanid = "SVBAL") Or (cleanid = "SVVALNET") Then
            If EngineID = 9 AndAlso isAnyCustomer Then
              DisabledAttribute = IIf(CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IStoredValueProgramService)().IsAnyCustomerSVProgramExist(), String.Empty, "disabled=""disabled""")
            End If
            Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 1, this.value);"" value=""" & (StrConv(row.Item("Tag"), VbStrConv.ProperCase)) & """ />")
          ElseIf (cleanid = "PTBAL") Then
            If EngineID = 9 AndAlso isAnyCustomer Then
              DisabledAttribute = IIf(CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IPointsProgramService)().IsAnyCustomerPointProgramExist(), String.Empty, "disabled=""disabled""")
            End If
            Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""javascript:showDialogSpan(true, 2, this.value);"" value=""" & (StrConv(row.Item("Tag"), VbStrConv.ProperCase)) & """ />")
          Else
            TextAreaName = IIf(EngineID = Copient.CommonInc.InstalledEngines.CPE, "t1_text", "tier0")
            If (TextAreaName = "tier0") Then
              If (MyCommon.Extract_Val(Request.QueryString("NumTiers")) > 0) Then
                TextAreaName = "tier1"
              End If
            End If
            Sendb("<input type=""button"" id=""ed_" & (StrConv(cleanid, VbStrConv.Lowercase)) & """ class=""ed_button""" & DisabledAttribute & " title=""" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & """ onclick=""edInsert" & (StrConv(cleanid, VbStrConv.ProperCase)) & "('" & TextAreaName & "');"" value=""" & (StrConv(row.Item("Tag"), VbStrConv.ProperCase)) & """ />")
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
  
  Response.Expires = 0
  MyCommon.AppName = "CashierMessageFeeds.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Request.QueryString("Mode") = "MarkupTags") Then
    Response.Clear()
    Response.ContentType = "text/html"
    GetMarkupTags(MyCommon.Extract_Val(Request.QueryString("EngineID")), MyCommon.Extract_Val(Request.QueryString("PrinterTypeID")))
  Else
    Send("<b>" & Copient.PhraseLib.Lookup("feeds.noarguments", LanguageID) & "!</b>")
    Send(Request.RawUrl)
  End If
  Response.Flush()
  Response.End()
%>
