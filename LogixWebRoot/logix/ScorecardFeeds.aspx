<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: ScorecardFeeds.aspx 
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
  ' *W
  ' * PURPOSE : 
  ' *
  ' * NOTES   : 
  ' *
  ' * Version : 7.3.1.138972 
  ' *
  ' *****************************************************************************

  Dim CopientFileName As String = "ScorecardFeeds.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  
  MyCommon.AppName = "ScorecardFeeds.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (LanguageID = 0) Then
    LanguageID = MyCommon.Extract_Val(Request.QueryString("LanguageID"))
  End If
  
  If (Request.QueryString("ScorecardFieldsForReward") <> "") Then
    ScorecardFieldsForReward(MyCommon, MyCommon.Extract_Val(Request.QueryString("EngineID")), MyCommon.Extract_Val(Request.QueryString("OfferID")), MyCommon.Extract_Val(Request.QueryString("DeliverableID")), MyCommon.Extract_Val(Request.QueryString("ProgramID")), MyCommon.Extract_Val(Request.QueryString("ScorecardTypeID")))
  Else
    Send("<b>" & Copient.PhraseLib.Lookup("term.noarguments", LanguageID) & "</b>")
  End If
  
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>

<script runat="server">
  Public DefaultLanguageID
  
  Sub ScorecardFieldsForReward(ByRef MyCommon As Copient.CommonInc, ByVal EngineID As Long, ByVal OfferID As Long, ByVal DeliverableID As Long, ByVal ProgramID As Long, ByVal ScorecardTypeID As Integer)
    Dim rst As DataTable
    Dim row As DataRow
    Dim ScorecardIDAtProgram As Integer = 0
    Dim ScorecardDescAtProgram As String = ""
    Dim ScorecardBoldAtProgram As Boolean = False
    Dim ScorecardIsNull As Boolean = True
    Dim ScorecardID As Integer = 0
    Dim ScorecardDesc As String = ""
    Dim DefaultExists As Boolean = False
    Dim PKID As Integer = 0
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    Try
      MyCommon.Open_LogixRT()
      
      ' First get the program-level scorecard details (which applies only the points and stored value scorecards)
      If (ScorecardTypeID = 1 OrElse ScorecardTypeID = 2) Then
        If ScorecardTypeID = 1 Then
          MyCommon.QueryStr = "select ScorecardID, ScorecardDesc, ScorecardBold from PointsPrograms with (NoLock) " & _
                              "where ProgramID=" & ProgramID & ";"
        Else
          MyCommon.QueryStr = "select ScorecardID, ScorecardDesc, ScorecardBold from StoredValuePrograms with (NoLock) " & _
                              "where SVProgramID=" & ProgramID & ";"
        End If
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          ScorecardIDAtProgram = MyCommon.NZ(rst.Rows(0).Item("ScorecardID"), 0)
          ScorecardDescAtProgram = MyCommon.NZ(rst.Rows(0).Item("ScorecardDesc"), "")
          ScorecardBoldAtProgram = MyCommon.NZ(rst.Rows(0).Item("ScorecardBold"), False)
        End If
      End If
      
      ' Next get the reward-level scorecard details (which applies to all scorecards, including discounts)
      If ScorecardTypeID = 1 Then
        MyCommon.QueryStr = "select PKID, ScorecardID, ScorecardDesc, ScorecardBold from CPE_DeliverablePoints with (NoLock) " & _
                            "where DeliverableID=" & DeliverableID & ";"
      ElseIf ScorecardTypeID = 2 Then
        MyCommon.QueryStr = "select PKID, ScorecardID, ScorecardDesc, ScorecardBold from CPE_DeliverableStoredValue with (NoLock) " & _
                            "where DeliverableID=" & DeliverableID & ";"
      ElseIf ScorecardTypeID = 3 Then
        MyCommon.QueryStr = "select DiscountID as PKID, ScorecardID, ScorecardDesc, ScorecardBold from CPE_Discounts with (NoLock) " & _
                            "where DeliverableID=" & DeliverableID & ";"
      End If
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        If IsDBNull(rst.Rows(0).Item("ScorecardID")) Then
          ScorecardIsNull = True
        Else
          ScorecardIsNull = False
          ScorecardID = MyCommon.NZ(rst.Rows(0).Item("ScorecardID"), 0)
        End If
        ScorecardDesc = MyCommon.NZ(rst.Rows(0).Item("ScorecardDesc"), "")
        ScorecardBoldAtProgram = MyCommon.NZ(rst.Rows(0).Item("ScorecardBold"), False)
        PKID = MyCommon.NZ(rst.Rows(0).Item("PKID"), 0)
      End If
      
      ' Also, check to see if a default scorecard has been declared
      MyCommon.QueryStr = "select ScorecardID from Scorecards with (NoLock) where Deleted=0 " & _
                          " and EngineID=" & EngineID & " and ScorecardTypeID=" & ScorecardTypeID & " and DefaultForEngine=1;"
      rst = MyCommon.LRT_Select
      If rst.Rows.Count > 0 Then
        DefaultExists = True
      End If
      
      ' Finally, draw the inputs table, which may vary based on the stuff above
      Send("<table summary=""" & Copient.PhraseLib.Lookup("term.scorecard", LanguageID) & """>")
      If ProgramID > 0 Then
        Send("  <tr>")
        Send("    <td style=""width:70px;"">")
        Send("      <label for=""ScorecardID"">" & Copient.PhraseLib.Lookup("CPEoffer-rew-discount.IncludeOnScorecard", LanguageID) & ":</label>")
        Send("    </td>")
        Send("    <td>")
        Send("      <select class=""medium"" id=""ScorecardID"" name=""ScorecardID""" & IIf(ScorecardIDAtProgram = 0, "", " disabled=""disabled""") & " onchange=""toggleScorecardText();"">")
        MyCommon.QueryStr = "select ScorecardID, Description, EngineID, DefaultForEngine from Scorecards " & _
                            "where ScorecardTypeID=" & ScorecardTypeID & " and Deleted=0 and EngineID=" & EngineID & ";"
        rst = MyCommon.LRT_Select
        If ScorecardIDAtProgram = 0 Then
          Send("        <option value=""0"">" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</option>")
          If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("ScorecardID"), 0) = ScorecardID) Then
                Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """ selected=""selected"">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
              ElseIf (MyCommon.NZ(row.Item("DefaultForEngine"), False) = True) AndAlso (MyCommon.NZ(row.Item("EngineID"), -1) = EngineID) AndAlso (ScorecardIsNull = True) Then
                Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """ selected=""selected"">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
              Else
                Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
              End If
            Next
          End If
        Else
          If rst.Rows.Count > 0 Then
            For Each row In rst.Rows
              Send("        <option value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """" & IIf(MyCommon.NZ(row.Item("ScorecardID"), 0) = ScorecardIDAtProgram, " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
            Next
          End If
        End If
        Send("      </select>")
        Send("      <input type=""hidden"" id=""ScorecardID"" name=""ScorecardID""" & IIf(ScorecardIDAtProgram = 0, " disabled=""disabled""", "") & " value=""" & MyCommon.NZ(row.Item("ScorecardID"), 0) & """>")
        Send("    </td>")
        Send("  </tr>")
        Send("  <tr id=""ScorecardDescLine""" & IIf((ScorecardIDAtProgram = 0 AndAlso DefaultExists = False) OrElse (ScorecardID = 0 AndAlso ScorecardIsNull = False AndAlso ScorecardIDAtProgram = 0), " style=""display:none;""", "") & ">")
        Send("    <td>")
        Send("      <label for=""ScorecardDesc"">" & Copient.PhraseLib.Lookup("term.scorecardtext", LanguageID) & ":</label>")
        Send("    </td>")
        Send("    <td>")
        'Send("      <input type=""text"" class=""medium"" id=""ScorecardDesc"" name=""ScorecardDesc"" maxlength=""31""" & IIf(ScorecardIDAtProgram = 0, "", " disabled=""disabled""") & " value=""" & IIf(ScorecardIDAtProgram = 0, ScorecardDesc.Replace("""", "&quot;"), ScorecardDescAtProgram.Replace("""", "&quot;")) & """ />")
        'Multilanguage inputs:
        Dim Localization As Copient.Localization
        Dim MLI As New Copient.Localization.MultiLanguageRec
        Localization = New Copient.Localization(MyCommon)
        If ScorecardTypeID = 1 Then 'points
          MLI.MLTableName = "CPE_DeliverablePointsTranslations"
          MLI.MLIdentifierName = "DeliverablePointsID"
          MLI.StandardTableName = "CPE_DeliverablePoints"
        ElseIf ScorecardTypeID = 2 Then 'stored value
          MLI.MLTableName = "CPE_DeliverableSVTranslations"
          MLI.MLIdentifierName = "DeliverableSVID"
          MLI.StandardTableName = "CPE_DeliverableStoredValue"
        End If
        MLI.ItemID = PKID
        MLI.StandardIdentifierName = "PKID"
        MLI.MLColumnName = "ScorecardDesc"
        MLI.StandardValue = IIf(ScorecardIDAtProgram = 0, ScorecardDesc.Replace("""", "&quot;"), ScorecardDescAtProgram.Replace("""", "&quot;"))
        MLI.InputName = "ScorecardDesc"
        MLI.InputID = "ScorecardDesc"
        MLI.InputType = "text"
        MLI.MaxLength = 31
        MLI.CSSStyle = "width:204px;"
        MLI.Disabled = IIf(ScorecardIDAtProgram = 0, False, True)
        Send(Localization.SendTranslationInputs(MyCommon, MLI, Request.QueryString, 2))
        Send("      <input type=""hidden"" id=""ScorecardDesc"" name=""ScorecardDesc""" & IIf(ScorecardIDAtProgram = 0, " disabled=""disabled""", "") & " value=""" & IIf(ScorecardIDAtProgram = 0, ScorecardDesc.Replace("""", "&quot;"), ScorecardDescAtProgram.Replace("""", "&quot;")) & """ />")
        Send("      <input type=""hidden"" id=""ScorecardDesc"" name=""ScorecardDesc""" & IIf(ScorecardIDAtProgram = 0, " disabled=""disabled""", "") & " value=""" & IIf(ScorecardIDAtProgram = 0, ScorecardDesc.Replace("""", "&quot;"), ScorecardDescAtProgram.Replace("""", "&quot;")) & """ />")
        Send("    </td>")
        Send("  </tr>")
        
        ' Commenting out the bolding inputs.  If we want to restore the ability of users to control this, restore this line (and the inputs for the stored procedures)
        'Send("  <tr>")
        'Send("    <td><label for=""ScorecardBold"">Bold on scorecard:</label></td>")
        'If Not ScorecardBoldAtProgram Then
        '  Send("    <td><input type=""checkbox"" id=""ScorecardBold"" name=""ScorecardBold""" & IIf(ScorecardBold, " checked=""checked""", "") & " /></td>")
        'Else
        '  Send("    <td><input type=""checkbox"" id=""ScorecardBold"" name=""ScorecardBold"" checked=""checked"" disabled=""disabled"" /></td>")
        'End If
        'Send("  </tr>")
      Else
        
      End If
      Send("</table>")
    Catch ex As Exception
      
    Finally
      
    End Try
  End Sub
</script>
