<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="System.Collections.Generic" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  Dim CopientFileName As String = "preference-feeds.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Integer
  Dim Common As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim PreferenceID As Long
  Dim Mode As String = ""
  Dim RetMsg As String = ""
  Dim Success As Boolean = False
  Dim DisallowEdit As Boolean = False

  Response.Expires = 0
  Response.Cache.SetCacheability(HttpCacheability.NoCache)

  Common.AppName = "preference-feeds.aspx"
  Common.Open_LogixRT()
  If Not (Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER)) Then
    Send(Copient.PhraseLib.Lookup("preference-feeds.EPMNotEnabled", LanguageID))
    Response.End()
  End If
  Common.Open_PrefManRT()
  AdminUserID = Verify_AdminUser(Common, Logix)
  
  DisallowEdit = (Common.Extract_Val(GetCgiValue("edit")) = -1)
  Long.TryParse(GetCgiValue("prefId"), PreferenceID)
  Mode = GetCgiValue("mode")
  
  Select Case Mode
    Case "loadPrefConValues"
      Response.Clear()
      Response.ContentType = "text/html"
      Send_Pref_Con_Values(Common, PreferenceID, DisallowEdit)
    Case "loadCMPrefConValues"
      Response.Clear()
      Response.ContentType = "text/html"
      Send_CMPref_Con_Values(Common, PreferenceID, DisallowEdit)
    Case "savePrefCon"
      Success = Save_Pref_Con(Common, PreferenceID, RetMsg)
      Send_Save_Pref_Con_Response(Success, RetMsg)
    Case "saveCMPrefCon"
      Success = Save_CMPref_Con(Common, RetMsg)
      Send_Save_Pref_Con_Response(Success, RetMsg)
    Case Else
      Send("<b>" & Copient.PhraseLib.Lookup("feeds.noarguments", LanguageID) & "!</b>")
      Send(Request.RawUrl)
  End Select

  Response.Flush()
  Response.End()

  Common.Close_LogixRT()
  Common.Close_PrefManRT()
  Common = Nothing
  
%>

<script runat="server">
  Structure TierData
    Dim PKID As Integer
    Dim Level As Integer
    Dim ValueComboID As Integer
    Dim Values As List(Of TierValue)
  End Structure

  Structure TierValue
    Dim OperatorTypeID As Integer
    Dim Value As String
    Dim DisplayText As String
    Dim ValueTypeID As Integer
    Dim DateOperatorTypeID As Integer
    Dim ValueModifier As String
    Dim DaysBefore As Integer
    Dim DaysAfter As Integer
  End Structure

  
  Enum ValueCombo
    And_Type = 1
    Or_Type = 2
  End Enum

  Enum OperatorTypes
    Equals = 1
    NotEquals = 2
    GreaterThan = 3
    LessThan = 4
  End Enum

  Enum DataTypes
    ListType = 1
    NumericRange = 2
    AlphaNumeric = 3
    Numeric = 4
    BooleanType = 5
    Theme = 6
    DateType = 7
    Likert = 8
  End Enum
  
  Const ANNIVERSARY_DATE_OP As Integer = 2
  
  Sub Send_Pref_Con_Values(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal DisallowEdit As Boolean)
    Dim TierLevel, TierCount As Integer
    Dim IncentivePrefsID As Integer = 0
    Dim Tiers As List(Of TierData)
    Dim OfferID As Long
    Dim DisabledAttribute As String = IIf(DisallowEdit, " disabled=""disabled""", "")

    ' note: javascript functions called by these form fields are found in prefentry.js file.
    
    Try
      Integer.TryParse(GetCgiValue("incentiveprefsid"), IncentivePrefsID)
      Long.TryParse(GetCgiValue("offerid"), OfferID)
      
      TierCount = Get_Tier_Levels(Common, OfferID)
      If TierCount <= 0 Then TierCount = 1
      Tiers = Load_Tiers(Common, PreferenceID, IncentivePrefsID, TierCount)
      
      For TierLevel = 1 To TierCount
        Send("<div id=""value_tier" & TierLevel & """>")
        Send_Tier_Header(TierCount, TierLevel)
        Send("  <table style=""width:98%;"">")
        Send_Date_Operator_Types(Common, PreferenceID, TierLevel, DisabledAttribute, 2)
        Send("    <tr>")
        Send("      <td style=""width:60px;"">" & Copient.PhraseLib.Lookup("term.operation", LanguageID) & ":</td>")
        Send("      <td>")
        Send("        <select id=""optype_tier" & TierLevel & """ name=""optype_tier" & TierLevel & """ class=""short""" & DisabledAttribute & ">")
        Send_Operator_Types(Common, PreferenceID, Tiers, TierLevel, 2)
        Send("        </select>")
        Send("      </td>")
        Send("      <td style=""text-align:right;"">")
        Send("        <input type=""button"" id=""addvalue_tier" & TierLevel & """ name=""addvalue_tier" & TierLevel & """ value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""addValueEntry(" & TierLevel & ");""" & DisabledAttribute & " />")
        Send("      </td>")
        Send("    </tr>")
        Send_Value_Types(Common, PreferenceID, TierLevel, DisabledAttribute, 2)
        Send("    <tr id=""trVal_tier" & TierLevel & """>")
        Send("      <td>" & Copient.PhraseLib.Lookup("term.value", LanguageID) & ":</td>")
        Send("      <td colspan=""2"">")
        Send_Value_Entry(Common, PreferenceID, Tiers, TierLevel, DisabledAttribute)
        Send("      </td>")
        Send("    </tr>")
        Send_Value_Modifier(Common, PreferenceID, TierLevel, DisabledAttribute)
        Send("    <tr>")
        Send("      <td colspan=""3"">")
        Send("        <select id=""values_tier" & TierLevel & """ name=""values_tier" & TierLevel & """ class=""longer"" size=""4""" & DisabledAttribute & ">")
        Send_Saved_Values(Common, Tiers, TierLevel)
        Send("        </select>")
        Send("        <input type=""hidden"" id=""allvalues_tier" & TierLevel & """ name=""allvalues_tier" & TierLevel & """ value="""" />")
        Send("      </td>")
        Send("    </tr>")
        Send("    <tr>")
        Send("      <td><input type=""button"" id=""remove_tier" & TierLevel & """ name=""remove_tier" & TierLevel & """ value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ onclick=""removeValueEntry(" & TierLevel & ");""" & DisabledAttribute & " /></td>")
        Send("      <td colspan=""2"" style=""text-align:center;"">")
        Send_Value_Combo(Common, PreferenceID, Tiers, TierLevel, DisabledAttribute)
        Send("      </td>")
        Send("    </tr>")
        Send("  </table>")
        Send("</div>")
      Next

    Catch ex As Exception
      Send(ex.ToString)
    End Try
  End Sub
    
  Sub Send_CMPref_Con_Values(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal DisallowEdit As Boolean)
    Dim TierLevel, TierCount As Integer
    Dim ConditionID As Long = 0
    Dim Tiers As List(Of TierData)
    Dim OfferID As Long
    Dim DisabledAttribute As String = IIf(DisallowEdit, " disabled=""disabled""", "")

    ' note: javascript functions called by these form fields are found in prefentry.js file.
    
    Try
      Integer.TryParse(GetCgiValue("conditionid"), ConditionID)
      Long.TryParse(GetCgiValue("offerid"), OfferID)
      
      TierCount = 1
      Tiers = Load_CM_Condition(Common, PreferenceID, ConditionID)
      
      For TierLevel = 1 To TierCount
        Send("<div id=""value_tier" & TierLevel & """>")
        Send("  <table style=""width:98%;"">")
        Send_Date_Operator_Types(Common, PreferenceID, TierLevel, DisabledAttribute, 0)
        Send("    <tr>")
        Send("      <td style=""width:60px;"">" & Copient.PhraseLib.Lookup("term.operation", LanguageID) & ":</td>")
        Send("      <td>")
        Send("        <select id=""optype_tier" & TierLevel & """ name=""optype_tier" & TierLevel & """ class=""medium"" onchange=""refreshDateValueBox(" & TierLevel & ");""" & DisabledAttribute & ">")
        Send_Operator_Types(Common, PreferenceID, Tiers, TierLevel, 0)
        Send("        </select>")
        Send("      </td>")
        Send("      <td style=""text-align:right;"">")
        Send("        <input type=""button"" id=""addvalue_tier" & TierLevel & """ name=""addvalue_tier" & TierLevel & """ value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & """ onclick=""addValueEntry(" & TierLevel & ");""" & DisabledAttribute & " />")
        Send("      </td>")
        Send("    </tr>")
        Send_Value_Types(Common, PreferenceID, TierLevel, DisabledAttribute, 0)
        Send("    <tr id=""trVal_tier" & TierLevel & """>")
        Send("      <td>" & Copient.PhraseLib.Lookup("term.value", LanguageID) & ":</td>")
        Send("      <td colspan=""2"">")
        Send_Value_Entry(Common, PreferenceID, Tiers, TierLevel, DisabledAttribute)
        Send("      </td>")
        Send("    </tr>")
        Send_Value_Modifier(Common, PreferenceID, TierLevel, DisabledAttribute)
        Send_Range(Common, PreferenceID, TierLevel, DisabledAttribute)
        Send("    <tr>")
        Send("      <td colspan=""3"">")
        Send("        <select id=""values_tier" & TierLevel & """ name=""values_tier" & TierLevel & """ class=""longer"" size=""4""" & DisabledAttribute & ">")
        Send_Saved_Values(Common, Tiers, TierLevel)
        Send("        </select>")
        Send("        <input type=""hidden"" id=""allvalues_tier" & TierLevel & """ name=""allvalues_tier" & TierLevel & """ value="""" />")
        Send("      </td>")
        Send("    </tr>")
        Send("    <tr>")
        Send("      <td><input type=""button"" id=""remove_tier" & TierLevel & """ name=""remove_tier" & TierLevel & """ value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ onclick=""removeValueEntry(" & TierLevel & ");""" & DisabledAttribute & " /></td>")
        Send("      <td colspan=""2"" style=""text-align:center;"">")
        Send_Value_Combo(Common, PreferenceID, Tiers, TierLevel, DisabledAttribute)
        Send("      </td>")
        Send("    </tr>")
        Send("  </table>")
        Send("</div>")
      Next

    Catch ex As Exception
      Send(ex.ToString)
    End Try
  End Sub

  Sub Send_Tier_Header(ByVal TierCount As Integer, ByVal TierLevel As Integer)
    If TierCount > 1 Then
      If TierLevel > 1 Then
        Send("<hr />")
      End If
      Send("<b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & TierLevel & "</b>")
    End If
  End Sub
  
  Sub Send_Date_Operator_Types(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal TierLevel As Integer, ByVal DisabledAttribute As String, ByVal EngineID As Integer)
    Dim dt As DataTable
    Dim DataTypeID As Integer
    
    DataTypeID = Get_Pref_Data_Type(Common, PreferenceID)

    ' only show the date operator dropdown for date preferences 
    If DataTypeID = 7 Then
      Send("    <tr>")
      Send("      <td style=""width:60px;"">" & Copient.PhraseLib.Lookup("term.dateoperator", LanguageID) & ":</td>")
      Send("      <td colspan=""2"">")
      Send("        <select id=""dateoptype_tier" & TierLevel & """ name=""dateoptype_tier" & TierLevel & """ class=""medium""  onchange=""" & IIf(EngineID = 0, "refreshDateValueBox(", "handleDateOpTypeChange(") & TierLevel & ")""" & DisabledAttribute & ">")

      Common.QueryStr = "select PrefDateOperatorTypeID, " & _
                        "  case when PT.PhraseID  is null then PDOT.Description " & _
                        "  else convert(nvarchar(50), PT.Phrase) end as PhraseText " & _
                        "from CPE_PrefDateOperatorTypes as PDOT with (NoLock) " & _
                        "left join PhraseText as PT with (NoLock) on PT.PhraseID = PDOT.PhraseID and LanguageID=" & LanguageID & " " & _
                        "where PrefDateOperatorTypeID > 0;"
      dt = Common.LRT_Select
      For Each row As DataRow In dt.Rows
        Send("          <option value=""" & Common.NZ(row.Item("PrefDateOperatorTypeID"), 0) & """>" & Common.NZ(row.Item("PhraseText"), "") & "</option>")
      Next

      Send("        </select>")
      Send("      <td>")
      Send("    </tr>")
    End If

  End Sub
  
  Sub Send_Value_Types(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal TierLevel As Integer, ByVal DisabledAttribute As String, ByVal EngineID As Integer)
    Dim dt As DataTable
    Dim DataTypeID As Integer
    
    DataTypeID = Get_Pref_Data_Type(Common, PreferenceID)

    Select Case DataTypeID
      Case 7 ' only show for the date preference type
        Send("    <tr>")
        Send("      <td style=""width:60px;"">" & Copient.PhraseLib.Lookup("term.datecomparison", LanguageID) & ":</td>")
        Send("      <td colspan=""2"">")
        Send("        <select id=""valtype_tier" & TierLevel & """ name=""valtype_tier" & TierLevel & """ class=""medium"" onchange=""" & IIf(EngineID=0, "refreshDateValueBox(", "handleValueTypeChange(") & TierLevel & ")""" & DisabledAttribute & ">")

        Common.QueryStr = "select PVT.PrefValueTypeID, " & _
                          "  case when PT.PhraseID  is null then PVT.Name " & _
                          "  else convert(nvarchar(50), PT.Phrase) end as PhraseText " & _
                          "from CPE_PrefValueTypes as PVT with (NoLock) " & _
                          "left join PhraseText as PT with (NoLock) on PT.PhraseID = PVT.PhraseID and LanguageID=" & LanguageID & " " & _
                          "where PVT.PrefDataTypeID = 7;"
        dt = Common.LRT_Select
        For Each row As DataRow In dt.Rows
          Send("          <option value=""" & Common.NZ(row.Item("PrefValueTypeID"), 0) & """>" & Common.NZ(row.Item("PhraseText"), "") & "</option>")
        Next

        Send("        </select>")
        Send("      <td>")
        Send("    </tr>")
    End Select

  End Sub
  
  Sub Send_Value_Modifier(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal TierLevel As Integer, ByVal DisabledAttribute As String)
    Dim DataTypeID As Integer
    
    DataTypeID = Get_Pref_Data_Type(Common, PreferenceID)

    Select Case DataTypeID
      Case 7
        Send("    <tr id=""trValMod_tier" & TierLevel & """ style=""display:none;"">")
        Send("      <td style=""width:60px;"">" & Copient.PhraseLib.Lookup("term.dateoffset", LanguageID) & ":</td>")
        Send("      <td colspan=""2"">")
        Send("        <input type=""text"" id=""valmod_tier" & TierLevel & """ name=""valmod_tier" & TierLevel & """ style=""width:98%;"" value=""""" & DisabledAttribute & " />")
        Send("      </td>")
        Send("    </tr>")
    End Select
        
  End Sub
  
  Sub Send_Range(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal TierLevel As Integer, ByVal DisabledAttribute As String)
    Dim DataTypeID As Integer
    
    DataTypeID = Get_Pref_Data_Type(Common, PreferenceID)

    Select Case DataTypeID
      Case 7 ' DATE
        Send("    <tr id=""trRange_tier" & TierLevel & """ style=""display:none;"">")
        Send("      <td style=""width:60px;"">" & Copient.PhraseLib.Lookup("term.days", LanguageID) & ":</td>")
        Send("      <td>" & Copient.PhraseLib.Lookup("term.before", LanguageID))
        Send("        <input type=""text"" id=""daysbefore_tier" & TierLevel & """ name=""daysbefore_tier" & TierLevel & """ style=""width:40px;;"" value=""""" & DisabledAttribute & " />")
        Send("      </td>")
        Send("      <td>" & Copient.PhraseLib.Lookup("term.after", LanguageID))
        Send("        <input type=""text"" id=""daysafter_tier" & TierLevel & """ name=""daysafter_tier" & TierLevel & """ style=""width:40px;;"" value=""""" & DisabledAttribute & " />")
        Send("      </td>")
        Send("    </tr>")
    End Select
        
  End Sub
  
  Sub Send_Operator_Types(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal Tiers As List(Of TierData), ByVal TierLevel As Integer, ByVal EngineID As Integer)
    Dim dt As DataTable
    Dim DataTypeID As Integer = 0
    Dim MultiValued As Boolean = False
    
    DataTypeID = Get_Pref_Data_Type(Common, PreferenceID)
    MultiValued = Is_Pref_MultiValued(Common, PreferenceID)
    
    Common.QueryStr = "select PrefOperatorTypeID, " & _
                      "  case when PT.PhraseID  is null then POT.Description " & _
                      "  else convert(nvarchar(10), PT.Phrase) end as PhraseText " & _
                      "from CPE_PrefOperatorTypes as POT with (NoLock) " & _
                      "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID

    If DataTypeID = DataTypes.ListType Then
      Common.QueryStr &= " where PrefOperatorTypeID in (1,2);"
    ElseIf DataTypeID = DataTypes.BooleanType Then
      Common.QueryStr &= " where PrefOperatorTypeID = 1;"
    ElseIf Not MultiValued AndAlso (Tier_Value_Has_Equals(Tiers, TierLevel) AndAlso Get_Tier_Value_ComboID(Tiers, TierLevel) = ValueCombo.And_Type) Then
      Common.QueryStr &= " where PrefOperatorTypeID not in (1, 5);"
    ElseIf EngineID <> 0 Then
      Common.QueryStr &= " where PrefOperatorTypeID not in(5);"
    End If
    
    dt = Common.LRT_Select
    For Each row As DataRow In dt.Rows
      Send("          <option value=""" & Common.NZ(row.Item("PrefOperatorTypeID"), 0) & """>" & Common.NZ(row.Item("PhraseText"), "") & "</option>")
    Next
    
  End Sub
  
  Function Get_Pref_Data_Type(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long) As Integer
    Dim DataTypeID As Integer = 0
    Dim dt As DataTable
    
    Common.QueryStr = "select DataTypeID from Preferences with (NoLock) where PreferenceID=" & PreferenceID
    dt = Common.PMRT_Select
    If dt.Rows.Count > 0 Then
      DataTypeID = Common.NZ(dt.Rows(0).Item("DataTypeID"), 0)
    End If
    
    Return DataTypeID
  End Function
  
  Function Is_Pref_MultiValued(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long) As Boolean
    Dim MultiValued As Boolean = False
    Dim dt As DataTable
    
    Try
      Common.QueryStr = "select MultiValue from Preferences with (NoLock) where PreferenceID=" & PreferenceID & " and Deleted=0;"
      dt = Common.PMRT_Select
      If dt.Rows.Count > 0 Then
        MultiValued = Common.NZ(dt.Rows(0).Item("MultiValue"), False)
      End If
    Catch ex As Exception
      Common.Write_Log("pref-feeds.txt", ex.ToString, True)
      MultiValued = False
    End Try
    
    Return MultiValued
  End Function
  

  Sub Send_Value_Entry(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal Tiers As List(Of TierData), _
                       ByVal TierLevel As Integer, ByVal DisabledAttribute As String)

    Dim DataTypeID As Integer = 0
    Dim SelTierVals As New List(Of TierValue)
    Dim Tier As TierData
    
    DataTypeID = Get_Pref_Data_Type(Common, PreferenceID)
    Select Case DataTypeID
      Case DataTypes.ListType, DataTypes.BooleanType
        For Each Tier In Tiers
          ' find the correct tier
          If Tier.Level = TierLevel Then
            SelTierVals = Tier.Values
            Exit For
          End If
        Next

        Send("        <select id=""valueentry_tier" & TierLevel & """ name=""valueentry_tier" & TierLevel & """ style=""width:98%;""" & DisabledAttribute & ">")
        If DataTypeID = DataTypes.ListType Then
          Send_Pref_List_Items(Common, PreferenceID, SelTierVals)
        ElseIf DataTypeID = DataTypes.BooleanType Then
          Send_Pref_Boolean_Options(Common, PreferenceID, SelTierVals)
        End If
        Send("        </select>")
      Case DataTypes.NumericRange
        Send("        <input type=""text"" id=""valueentry_tier" & TierLevel & """ name=""valueentry_tier" & TierLevel & """ style=""width:48%;"" value=""""" & DisabledAttribute & " />")
        Send_Numeric_Range_Label(Common, PreferenceID)
      Case DataTypes.Numeric, DataTypes.DateType
        Send("        <input type=""text"" id=""valueentry_tier" & TierLevel & """ name=""valueentry_tier" & TierLevel & """ style=""width:98%;"" value=""""" & DisabledAttribute & " />")
      Case DataTypes.Likert
        Send("        <input type=""text"" id=""valueentry_tier" & TierLevel & """ name=""valueentry_tier" & TierLevel & """ style=""width:48%;"" value=""""" & DisabledAttribute & " />")
        Send_Likert_Label(Common, PreferenceID)
    End Select
  End Sub
  
  Sub Send_Pref_List_Items(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal SelTierVals As List(Of TierValue))
    Dim dt As DataTable
    Dim SelList As String = "''"
    
    For Each tv As TierValue In SelTierVals
      If SelList.Length > 0 Then SelList &= ","
      SelList &= "'" & tv.Value & "'"
    Next
    
    Common.QueryStr = "select PLI.Value, " & _
                      "  case when UPT.PhraseID is null then PLI.Name " & _
                      "  else convert(nvarchar(100), UPT.Phrase) end as PhraseText " & _
                      "from PreferenceListItems as PLI with (NoLock) " & _
                      "left join UserPhraseText as UPT with (NoLock) on UPT.PhraseID = PLI.NamePhraseID and LanguageID=" & LanguageID & " " & _
                      "where PreferenceID=" & PreferenceID & IIf(SelList.Length > 0, " and PLI.Value not in (" & SelList & ") ", " ") & _
                      "order by PhraseText"
    dt = Common.PMRT_Select
    For Each row As DataRow In dt.Rows
      Send("          <option value=""" & Common.NZ(row.Item("Value"), 0) & """>" & Common.NZ(row.Item("PhraseText"), "") & "</option>")
    Next
  End Sub

  Sub Send_Pref_Boolean_Options(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal SelTierVals As List(Of TierValue))
    Dim TrueSelected, FalseSelected As Boolean
    
    For Each tv As TierValue In SelTierVals
      TrueSelected = TrueSelected OrElse (tv.Value = "1" OrElse tv.Value.ToLower = "true")
      FalseSelected = FalseSelected OrElse (tv.Value = "0" OrElse tv.Value.ToLower = "false")
    Next

    If Not TrueSelected Then Send("          <option value=""1"">" & Copient.PhraseLib.Lookup("term.true", LanguageID) & "</option>")
    If Not FalseSelected Then Send("          <option value=""0"">" & Copient.PhraseLib.Lookup("term.false", LanguageID) & "</option>")
  End Sub
  
  Sub Send_Numeric_Range_Label(ByVal Common As Copient.CommonInc, ByVal PreferenceID As Long)
    Dim dt As DataTable
    Dim row As DataRow
    Dim i, endCount As Integer
    Dim MinDec, MaxDec As Decimal
    Dim MinVals As New List(Of Decimal)
    Dim MaxVals As New List(Of Decimal)
    
    Common.QueryStr = "select MinimumValue, MaximumValue from PreferenceNumericRanges with (NoLock) " & _
                      "where PreferenceID=" & PreferenceID & " order by MinimumValue;"
    dt = Common.PMRT_Select
    endCount = dt.Rows.Count - 1
    If endCount > -1 Then
      Sendb("&nbsp;(")
      For i = 0 To endCount
        row = dt.Rows(i)
        MinDec = Common.NZ(row.Item("MinimumValue"), 0D)
        MinVals.Add(MinDec)
        Sendb(MinDec.ToString("0.###") & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & " ")

        MaxDec = Common.NZ(row.Item("MaximumValue"), 0D)
        MaxVals.Add(MaxDec)
        Sendb(MaxDec.ToString("0.###") & IIf(i < endCount, "; ", ""))
      Next
      Send(")")

      i = 0
      For Each d As Decimal In MinVals
        i += 1
        Send("<input type=""hidden"" name=""minvalue"" id=""minvalue" & i & """ value=""" & d.ToString("0.###") & """ />")
      Next

      i = 0
      For Each d As Decimal In MaxVals
        i += 1
        Send("<input type=""hidden"" name=""maxvalue"" id=""maxvalue" & i & """ value=""" & d.ToString("0.###") & """ />")
      Next
    End If
    
  End Sub

  Sub Send_Likert_Label(ByVal Common As Copient.CommonInc, ByVal PreferenceID As Long)
    Dim dt As DataTable
    Dim row As DataRow
    Dim i, endCount As Integer
    Dim MinVals As New List(Of Integer)
    Dim MaxVals As New List(Of Integer)
    
    Common.QueryStr = "select MinValue, MaxValue from PreferenceLikerts with (NoLock) " & _
                      "where PreferenceID=" & PreferenceID & " order by MinValue;"
    dt = Common.PMRT_Select
    endCount = dt.Rows.Count - 1
    If endCount > -1 Then
      Sendb("&nbsp;(")
      For i = 0 To endCount
        row = dt.Rows(i)
        MinVals.Add(CInt(Common.NZ(row.Item("MinValue"), 0)))
        MaxVals.Add(CInt(Common.NZ(row.Item("MaxValue"), 0)))
                    
        Sendb(Common.NZ(row.Item("MinValue"), 0) & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & " " & _
             Common.NZ(row.Item("MaxValue"), 0) & IIf(i < endCount, "; ", ""))
      Next
      Send(")")
      
      i = 0
      For Each v As Integer In MinVals
        i += 1
        Send("<input type=""hidden"" name=""minvalue"" id=""minvalue" & i & """ value=""" & v.ToString() & """ />")
      Next

      i = 0
      For Each v As Integer In MaxVals
        Send("<input type=""hidden"" name=""maxvalue"" id=""maxvalue" & i & """ value=""" & v.ToString() & """ />")
      Next
    End If

  End Sub
  
  Sub Send_Saved_Values(ByRef Common As Copient.CommonInc, ByVal Tiers As List(Of TierData), ByVal TierLevel As Integer)
    Dim Tier As TierData
    Dim Vals As New List(Of TierValue)
    
    If Tiers IsNot Nothing AndAlso Tiers.Count > 0 Then
      For Each Tier In Tiers
        ' find the correct tier to write
        If Tier.Level = TierLevel Then
          Vals = Tier.Values

          ' write each of the tiers saved values
          For Each v As TierValue In Vals
            Sendb("<option value=""" & v.OperatorTypeID & "|" & v.Value & "|" & v.DateOperatorTypeID & "|" & v.ValueTypeID & "|" & v.ValueModifier)
            If v.DateOperatorTypeID = ANNIVERSARY_DATE_OP Then
              Sendb("|" & v.DaysBefore & "|" & v.DaysAfter)
            End If
            Send(""">" & v.DisplayText & "</option>")
          Next
          Exit For
        End If
      Next
    End If
    
  End Sub
  
  Sub Send_Value_Combo(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal Tiers As List(Of TierData), ByVal TierLevel As Integer, ByVal DisabledAttribute As String)
    Dim Tier As TierData
    Dim AndAttributes As String = ""
    Dim OrAttributes As String = ""
    Dim NeedsDefaulted As Boolean = True
    Dim DataTypeID As Integer
    
    DataTypeID = Get_Pref_Data_Type(Common, PreferenceID)
    
    If Tiers IsNot Nothing AndAlso Tiers.Count > 0 Then
      For Each Tier In Tiers
        ' find the correct tier
        If Tier.Level = TierLevel Then
          Select Case Tier.ValueComboID
            Case 1 ' and
              AndAttributes &= " checked=""checked"""
              NeedsDefaulted = False
            Case 2 'or
              OrAttributes &= " checked=""checked"""
              NeedsDefaulted = False
          End Select
          Exit For
        End If
      Next
    End If

    ' when single valued, if equal or not equal is selected then 'And' should be disabled because the condition would never evaluate
    If DataTypeID = DataTypes.BooleanType OrElse (Not Is_Pref_MultiValued(Common, PreferenceID) AndAlso Tier_Value_Has_Equals(Tiers, TierLevel)) Then
      AndAttributes = " disabled=""disabled"""
      'OrAttributes = ""
      'NeedsDefaulted = True
    End If

    If NeedsDefaulted Then
      OrAttributes &= " checked=""checked"""
    End If

    Send("        <input type=""radio"" name=""valcbo_tier" & TierLevel & """ id=""valcboand" & TierLevel & """ onclick=""handleAndClick(" & TierLevel & ");"" value=""1""" & AndAttributes & "" & DisabledAttribute & " />")
    Send("        <label for=""valcboand" & TierLevel & """>" & Copient.PhraseLib.Lookup("term.and", LanguageID) & "</label>")
    Send("        <input type=""radio"" name=""valcbo_tier" & TierLevel & """ id=""valcboor" & TierLevel & """ onclick=""handleOrClick(" & TierLevel & ");"" value=""2""" & OrAttributes & "" & DisabledAttribute & " />")
    Send("        <label for=""valcboor" & TierLevel & """>" & Copient.PhraseLib.Lookup("term.or", LanguageID) & "</label>")

  End Sub

  ' determines if at least one value for the tier uses the "equal" operator
  Function Tier_Value_Has_Equals(ByVal Tiers As List(Of TierData), ByVal TierLevel As Integer) As Boolean
    Dim Used As Boolean = False
    
    If Tiers IsNot Nothing Then
      For Each t As TierData In Tiers
        If t.Level = TierLevel Then
          For Each v As TierValue In t.Values
            If v.OperatorTypeID = OperatorTypes.Equals Then
              Used = True
            End If
          Next
        End If
      Next
    End If
    
    Return Used
  End Function
  
  Function Get_Tier_Value_ComboID(ByVal Tiers As List(Of TierData), ByVal TierLevel As Integer) As Integer
    Dim ComboID As Integer = ValueCombo.Or_Type
    
    If Tiers IsNot Nothing Then
      For Each t As TierData In Tiers
        If t.Level = TierLevel Then
          ComboID = t.ValueComboID
        End If
      Next
    End If
    
    Return ComboID
  End Function
  
  Function Get_Tier_Levels(ByRef Common As Copient.CommonInc, ByVal OfferID As Long) As Integer
    Dim TierLevels As Integer
    Dim dt As DataTable
    
    ' find the number of tiers for this offer
    Common.QueryStr = "select RO.TierLevels from CPE_RewardOptions as RO with (NoLock) " & _
                      "where RO.IncentiveID=" & OfferID & " and TouchResponse=0;"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      TierLevels = Common.NZ(dt.Rows(0).Item("TierLevels"), 0)
    End If

    Return TierLevels
  End Function
  
  
  Function Load_Tiers(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal IncentivePrefsID As Integer, ByVal TierLevels As Integer) As List(Of TierData)
    Dim dt, dt2 As DataTable
    Dim i As Integer
    Dim Tiers As New List(Of TierData)
    Dim Tier As TierData
    Dim TierVal As TierValue
    Dim PrefDataTypeID As Integer
    
    PrefDataTypeID = Get_Pref_Data_Type(Common, PreferenceID)
    
    ' find the tier levels and their associated values
    For i = 1 To TierLevels
      Common.QueryStr = "select IPT.IncentivePrefTiersID, IPT.ValueComboTypeID " & _
                        "from CPE_IncentivePrefTiers as IPT with (NoLock) " & _
                        "inner join CPE_IncentivePrefs as CIP with (NoLock) on CIP.IncentivePrefsID = IPT.IncentivePrefsID " & _
                        "where IPT.IncentivePrefsID=" & IncentivePrefsID & " and IPT.TierLevel=" & i & " and CIP.PreferenceID=" & PreferenceID
      dt = Common.LRT_Select
      If dt.Rows.Count > 0 Then
        Tier = New TierData
        Tier.Values = New List(Of TierValue)
        
        Tier.Level = i
        Tier.PKID = Common.NZ(dt.Rows(0).Item("IncentivePrefTiersID"), 0)
        Tier.ValueComboID = Common.NZ(dt.Rows(0).Item("ValueComboTypeID"), 1)
        
        Common.QueryStr = "select IPTV.PKID, IPTV.OperatorTypeID, IPTV.Value, IPTV.ValueTypeID, IPTV.DateOperatorTypeId, IPTV.ValueModifier, " & _
                          "  case when PT.PhraseID  is null then POT.Description " & _
                          "  else convert(nvarchar(10), PT.Phrase) end as PhraseText " & _
                          "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                          "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                          "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID & _
                          "where IncentivePrefTiersID=" & Tier.PKID
        dt2 = Common.LRT_Select
        For Each row As DataRow In dt2.Rows
          TierVal = New TierValue
          With TierVal
            .OperatorTypeID = Common.NZ(row.Item("OperatorTypeID"), 1)
            .Value = Common.NZ(row.Item("Value"), "")
            .DateOperatorTypeID = Common.NZ(row.Item("DateOperatorTypeID"), 0)
            .ValueTypeID = Common.NZ(row.Item("ValueTypeID"), 0)
            .ValueModifier = Common.NZ(row.Item("ValueModifier"), "")
            .DisplayText = Common.NZ(row.Item("PhraseText"), "")
          End With

          Select Case PrefDataTypeID
            Case DataTypes.ListType
              TierVal.DisplayText &= " " & Get_Pref_List_Item_Name(Common, PreferenceID, TierVal.Value)
            Case DataTypes.Numeric, DataTypes.NumericRange, DataTypes.Likert
              TierVal.DisplayText &= " " & TierVal.Value
            Case DataTypes.BooleanType
              TierVal.DisplayText &= " " & Copient.PhraseLib.Lookup("term." & IIf(TierVal.Value = "1" OrElse TierVal.Value.ToUpper = "TRUE", "true", "false"), LanguageID)
            Case DataTypes.DateType
              TierVal.DisplayText = Get_Date_Display_Text(Common, Common.NZ(row.Item("PKID"), 0), 2)
          End Select
          
          Tier.Values.Add(TierVal)
        Next

        Tiers.Add(Tier)
      Else
        Tiers.Add(New TierData)
      End If
    Next
    
    Return Tiers
  End Function


  Function Load_CM_Condition(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal ConditionID As Integer) As List(Of TierData)
    Dim dt As DataTable
    Dim Tiers As New List(Of TierData)
    Dim Tier As TierData
    Dim TierVal As TierValue
    Dim PrefDataTypeID As Integer
    
    PrefDataTypeID = Get_Pref_Data_Type(Common, PreferenceID)
    
    Tier = New TierData
    Tier.Values = New List(Of TierValue)
        
    Tier.Level = 1
    Common.QueryStr = "select CPV.PKID, CPV.OperatorTypeID, CPV.Value, CPV.ValueTypeID, CPV.ValueComboTypeID, " & _
                      "  CPV.DateOperatorTypeId, CPV.ValueModifier, CPV.DaysBefore, CPV.DaysAfter, " & _
                      "  case when PT.PhraseID  is null then POT.Description " & _
                      "  else convert(nvarchar(10), PT.Phrase) end as PhraseText " & _
                      "from CM_ConditionPreferenceValues as CPV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = CPV.OperatorTypeID " & _
                      "left join PhraseText as PT with (NoLock) on PT.PhraseID = POT.PhraseID and LanguageID=" & LanguageID & _
                      "where ConditionID=" & ConditionID & " and PreferenceID = " & PreferenceID
    dt = Common.LRT_Select
    For Each row As DataRow In dt.Rows
      Tier.ValueComboID = Common.NZ(row.Item("ValueComboTypeID"), 2)
      
      TierVal = New TierValue
      With TierVal
        .OperatorTypeID = Common.NZ(row.Item("OperatorTypeID"), 1)
        .Value = Common.NZ(row.Item("Value"), "")
        .DateOperatorTypeID = Common.NZ(row.Item("DateOperatorTypeID"), 0)
        .ValueTypeID = Common.NZ(row.Item("ValueTypeID"), 0)
        .ValueModifier = Common.NZ(row.Item("ValueModifier"), "")
        .DaysBefore = Common.NZ(row.Item("DaysBefore"), 0)
        .DaysAfter = Common.NZ(row.Item("DaysAfter"), 0)
        .DisplayText = Common.NZ(row.Item("PhraseText"), "")
      End With

      Select Case PrefDataTypeID
        Case DataTypes.ListType
          TierVal.DisplayText &= " " & Get_Pref_List_Item_Name(Common, PreferenceID, TierVal.Value)
        Case DataTypes.Numeric, DataTypes.NumericRange, DataTypes.Likert
          TierVal.DisplayText &= " " & TierVal.Value
        Case DataTypes.BooleanType
          TierVal.DisplayText &= " " & Copient.PhraseLib.Lookup("term." & IIf(TierVal.Value = "1" OrElse TierVal.Value.ToUpper = "TRUE", "true", "false"), LanguageID)
        Case DataTypes.DateType
          TierVal.DisplayText = Get_Date_Display_Text(Common, Common.NZ(row.Item("PKID"), 0), 0)
      End Select
          
      Tier.Values.Add(TierVal)
    Next

    Tiers.Add(Tier)
    
    Return Tiers
  End Function

  Function Get_Date_Display_Text(ByRef Common As Copient.CommonInc, ByVal TierValuePKID As Integer, ByVal EngineID As Integer) As String
    Dim DisplayText As String = ""
    Dim dt As DataTable
    Dim ValueModifier As String = ""
    Dim Offset, DaysBefore, DaysAfter As Integer
    
    If EngineID = 0 Then
      Common.QueryStr = "select CPV.Value, CPV.ValueModifier, CPV.ValueTypeID, CPV.DaysBefore, CPV.DaysAfter, POT.PhraseID as OperatorPhraseID, " & _
                        "PDOT.PhraseID as DateOpPhraseID, CPV.DateOperatorTypeID " & _
                        "from CM_ConditionPreferenceValues as CPV with (NoLock) " & _
                        "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = CPV.OperatorTypeID " & _
                        "inner join CPE_PrefDateOperatorTypes as PDOT with (NoLock) on PDOT.PrefDateOperatorTypeID = CPV.DateOperatorTypeID " & _
                        "where PKID=" & TierValuePKID & ";"
    Else
    Common.QueryStr = "select IPTV.Value, IPTV.ValueModifier, IPTV.ValueTypeID, POT.PhraseID as OperatorPhraseID, " & _
                      "PDOT.PhraseID as DateOpPhraseID " & _
                      "from CPE_IncentivePrefTierValues as IPTV with (NoLock) " & _
                      "inner join CPE_PrefOperatorTypes as POT with (NoLock) on POT.PrefOperatorTypeID = IPTV.OperatorTypeID " & _
                      "inner join CPE_PrefDateOperatorTypes as PDOT with (NoLock) on PDOT.PrefDateOperatorTypeID = IPTV.DateOperatorTypeID " & _
                      "where PKID=" & TierValuePKID & ";"
    End If

    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      DisplayText = Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("DateOpPhraseID"), ""), LanguageID) & " "
      DisplayText &= Copient.PhraseLib.Lookup(Common.NZ(dt.Rows(0).Item("OperatorPhraseID"), ""), LanguageID) & " "
      If Common.NZ(dt.Rows(0).Item("ValueTypeID"), 0) = 1 Then
        DisplayText &= "[" & Copient.PhraseLib.Lookup("term.currentdate", LanguageID).ToLower & "]"
        ValueModifier = Common.NZ(dt.Rows(0).Item("ValueModifier"), "")
        If ValueModifier <> "" AndAlso Integer.TryParse(ValueModifier, Offset) Then
          ValueModifier = " " & IIf(Offset < 0, " - ", " + ") & Math.Abs(Offset)
        End If
        DisplayText &= ValueModifier
      Else
        DisplayText &= " " & Common.NZ(dt.Rows(0).Item("Value"), "")
      End If

      If EngineID = 0 AndAlso Common.NZ(dt.Rows(0).Item("DateOperatorTypeID"), 0) = ANNIVERSARY_DATE_OP Then
        DaysBefore = Common.NZ(dt.Rows(0).Item("DaysBefore"), 0)
        DaysAfter = Common.NZ(dt.Rows(0).Item("DaysAfter"), 0)

        If DaysBefore > 0 AndAlso DaysAfter > 0 Then
          DisplayText &= " (-" & DaysBefore & " / +" & DaysAfter & ")"
        ElseIf DaysBefore > 0 AndAlso DaysAfter = 0 Then
          DisplayText &= " (-" & DaysBefore & ")"
        ElseIf DaysBefore = 0 AndAlso DaysAfter > 0 Then
          DisplayText &= " (+" & DaysAfter & ")"
    End If
      End If
    End If
    
    Return DisplayText
  End Function
  

  Function Get_Pref_List_Item_Name(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByVal ItemValue As String) As String
    Dim Name As String = ""
    Dim dt As DataTable
    
    Common.QueryStr = "select case when UPT.Phrase is null then PLI.Name " & _
                      "       else CONVERT(nvarchar(200), UPT.Phrase) end as PhraseText " & _
                      "from PreferenceListItems as PLI with (NoLock) " & _
                      "left join UserPhraseText as UPT with (NoLock)on UPT.PhraseID = PLI.NamePhraseID " & _
                      "  and UPT.LanguageID=" & LanguageID & " " & _
                      "where PLI.PreferenceID=" & PreferenceID & " and PLI.Value=N'" & ItemValue & "';"
    dt = Common.PMRT_Select
    If dt.Rows.Count > 0 Then
      Name = Common.NZ(dt.Rows(0).Item("PhraseText"), "")
    End If
    
    Return Name
  End Function
  
  Function Save_Pref_Con(ByRef Common As Copient.CommonInc, ByVal PreferenceID As Long, ByRef RetMsg As String) As Boolean
    Dim Saved As Boolean = False
    Dim IncentivePrefsID, IncentivePrefTiersID, ValuePKID As Integer
    Dim ROID, OfferID As Long
    Dim SelectedPrefID, ValueComboID As Integer
    Dim TierLevel As Integer = 1
    Dim TierCount As Integer = 1
    Dim RollbackSave As Boolean = False
    Dim TempStr As String = ""
    Dim TierVals As List(Of TierValue)
    Dim DisallowEdit As Boolean = False
    
    Try
      Integer.TryParse(GetCgiValue("IncentivePrefsID"), IncentivePrefsID)
      Integer.TryParse(Parse_PrefID(GetCgiValue("preferenceid")), SelectedPrefID)
      Integer.TryParse(GetCgiValue("TierCount"), TierCount)
      Long.TryParse(GetCgiValue("ROID"), ROID)
      Long.TryParse(GetCgiValue("OfferID"), OfferID)
      DisallowEdit = (GetCgiValue("Disallow_Edit") = "1")
    
      Common.QueryStr = "BEGIN TRAN T1;"
      Common.LRT_Execute()
      
      ' remove any existing incentive preferences
      If IncentivePrefsID > 0 Then
        Common.QueryStr = "dbo.pt_CPE_IncentivePrefs_Delete"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@IncentivePrefsID", SqlDbType.Int).Value = IncentivePrefsID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
      End If

      Common.QueryStr = "dbo.pt_CPE_IncentivePrefs_Insert"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
      Common.LRTsp.Parameters.Add("@PreferenceID", SqlDbType.Int).Value = SelectedPrefID
      Common.LRTsp.Parameters.Add("@DisallowEdit", SqlDbType.Int).Value = IIf(DisallowEdit, 1, 0)
      Common.LRTsp.Parameters.Add("@IncentivePrefsID", SqlDbType.Int).Direction = ParameterDirection.Output
      Common.LRTsp.ExecuteNonQuery()
      IncentivePrefsID = Common.LRTsp.Parameters("@IncentivePrefsID").Value
      Common.Close_LRTsp()
    
      If IncentivePrefsID > 0 Then
        For TierLevel = 1 To TierCount
          ValueComboID = Common.Extract_Val(GetCgiValue("valcbo_tier" & TierLevel))

          Common.QueryStr = "dbo.pt_CPE_IncentivePrefTiers_Insert"
          Common.Open_LRTsp()
          Common.LRTsp.Parameters.Add("@IncentivePrefsID", SqlDbType.BigInt).Value = IncentivePrefsID
          Common.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int).Value = TierLevel
          Common.LRTsp.Parameters.Add("@ValueComboTypeID", SqlDbType.Int).Value = ValueComboID
          Common.LRTsp.Parameters.Add("@IncentivePrefTiersID", SqlDbType.Int).Direction = ParameterDirection.Output
          Common.LRTsp.ExecuteNonQuery()
          IncentivePrefTiersID = Common.LRTsp.Parameters("@IncentivePrefTiersID").Value
          Common.Close_LRTsp()
      
          If IncentivePrefTiersID > 0 Then
            TierVals = Parse_TierValues("allvalues_tier" & TierLevel)

            
            If TierVals IsNot Nothing Then
              For Each t As TierValue In TierVals
                Common.QueryStr = "dbo.pt_CPE_IncentivePrefTierValues_Insert"
                Common.Open_LRTsp()
                Common.LRTsp.Parameters.Add("@IncentivePrefTiersID", SqlDbType.Int).Value = IncentivePrefTiersID
                Common.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 250).Value = t.Value
                Common.LRTsp.Parameters.Add("@OperatorTypeID", SqlDbType.Int).Value = t.OperatorTypeID
                Common.LRTsp.Parameters.Add("@ValueTypeID", SqlDbType.Int).Value = t.ValueTypeID
                Common.LRTsp.Parameters.Add("@DateOperatorTypeID", SqlDbType.Int).Value = t.DateOperatorTypeID
                Common.LRTsp.Parameters.Add("@ValueModifier", SqlDbType.NVarChar, 250).Value = t.ValueModifier
                Common.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
                Common.LRTsp.ExecuteNonQuery()
                ValuePKID = Common.LRTsp.Parameters("@PKID").Value
                Common.Close_LRTsp()
              Next
            Else
              Throw New Exception(Copient.PhraseLib.Detokenize("preference-feeds.TierValuesError", LanguageID, TierLevel))
            End If
          Else
            Throw New Exception(Copient.PhraseLib.Detokenize("preference-feeds.FailedCreateTier", LanguageID, TierLevel))
          End If

        Next
      Else
        Throw New Exception(Copient.PhraseLib.Lookup("preference-feeds.FailedIncentivePref", LanguageID))
      End If

      Common.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 " & _
                        "where IncentiveID=" & OfferID & ";"
      Common.LRT_Execute()
            ResetOfferApprovalStatus(OfferID)

      Saved = True
      RetMsg = Copient.PhraseLib.Lookup("term.ChangesWereSaved", LanguageID)
            Common.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.con-pref-edit", LanguageID) & ": " & SelectedPrefID)
        Catch ex As Exception
            RollbackSave = True
            RetMsg = Copient.PhraseLib.Lookup("preference-feeds.ErrorEncounteredSave", LanguageID) & " " & ex.ToString
      
    Finally
      Common.QueryStr = IIf(RollbackSave, "ROLLBACK TRAN T1", "COMMIT TRAN T1")
      Common.LRT_Execute()
    End Try
    
    Return Saved
  End Function


  Function Save_CMPref_Con(ByRef Common As Copient.CommonInc, ByRef RetMsg As String) As Boolean
    Dim Saved As Boolean = False
    Dim OfferID, ConditionID, ValuePKID As Long
    Dim SelectedPrefID, ValueComboID As Integer
    Dim TierLevel As Integer = 1
    Dim TierCount As Integer = 1
    Dim RollbackSave As Boolean = False
    Dim TempStr As String = ""
    Dim TierVals As List(Of TierValue)
    Dim DisallowEdit As Boolean = False
    
    Try
      Integer.TryParse(GetCgiValue("conditionid"), ConditionID)
      Integer.TryParse(Parse_PrefID(GetCgiValue("preferenceid")), SelectedPrefID)
      Long.TryParse(GetCgiValue("offerid"), OfferID)
      DisallowEdit = (GetCgiValue("Disallow_Edit") = "1")

      Common.QueryStr = "BEGIN TRAN T1;"
      Common.LRT_Execute()
      
      If ConditionID > 0 Then
        ' first, remove any existing incentive preferences
        Common.QueryStr = "dbo.pt_CM_ConditionPreferenceValues_Delete"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
        
        ' then, add the new values
        TierVals = Parse_TierValues("allvalues_tier" & TierLevel)
        If TierVals IsNot Nothing Then
          For Each t As TierValue In TierVals
            ValueComboID = Common.Extract_Val(GetCgiValue("valcbo_tier" & TierLevel))
            ' days before and days after are only valid for anniversary date operation with the date range operation selected
            If t.DateOperatorTypeID <> 2 OrElse t.OperatorTypeID <> 5 Then
              t.DaysBefore = 0
              t.DaysAfter = 0
            End If
            
            Common.Write_Log("PrefFeeds.txt", "ConditionID = " & ConditionID & " PreferenceID=" & SelectedPrefID, True)
            Common.QueryStr = "dbo.pt_CM_ConditionPreferenceValues_Insert"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@ConditionID", SqlDbType.BigInt).Value = ConditionID
            Common.LRTsp.Parameters.Add("@PreferenceID", SqlDbType.BigInt).Value = SelectedPrefID
            Common.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 250).Value = t.Value
            Common.LRTsp.Parameters.Add("@OperatorTypeID", SqlDbType.Int).Value = t.OperatorTypeID
            Common.LRTsp.Parameters.Add("@ValueTypeID", SqlDbType.Int).Value = t.ValueTypeID
            Common.LRTsp.Parameters.Add("@DateOperatorTypeID", SqlDbType.Int).Value = t.DateOperatorTypeID
            Common.LRTsp.Parameters.Add("@ValueModifier", SqlDbType.NVarChar, 250).Value = t.ValueModifier
            Common.LRTsp.Parameters.Add("@ValueComboTypeID", SqlDbType.Int).Value = ValueComboID
            Common.LRTsp.Parameters.Add("@DaysBefore", SqlDbType.Int).Value = t.DaysBefore
            Common.LRTsp.Parameters.Add("@DaysAfter", SqlDbType.Int).Value = t.DaysAfter
            Common.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
            Common.LRTsp.ExecuteNonQuery()
            ValuePKID = Common.LRTsp.Parameters("@PKID").Value
            Common.Close_LRTsp()
          Next
        Else
          Throw New Exception(Copient.PhraseLib.Detokenize("preference-feeds.TierValuesError", LanguageID, TierLevel))
        End If
        
      End If

      Common.QueryStr = "update Offers with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 " & _
                        "where OfferID=" & OfferID & ";"
      Common.LRT_Execute()

      Saved = True
      RetMsg = Copient.PhraseLib.Lookup("term.ChangesWereSaved", LanguageID)
         
    Catch ex As Exception
      RollbackSave = True
      RetMsg = Copient.PhraseLib.Lookup("preference-feeds.ErrorEncounteredSave", LanguageID) & " " & ex.ToString
      
    Finally
      Common.QueryStr = IIf(RollbackSave, "ROLLBACK TRAN T1", "COMMIT TRAN T1")
      Common.LRT_Execute()
    End Try
    
    Return Saved
  End Function

  Sub Send_Save_Pref_Con_Response(ByVal Saved As Boolean, ByVal RetMsg As String)
    Response.Clear()
    Send(IIf(Saved, "OK", "ERROR"))
    Send("[DATA]:Saved=" & IIf(Saved, 1, 0))
    Send(RetMsg)
  End Sub
  
  Function Parse_PrefID(ByVal str As String) As Integer
    Dim PrefID As Integer = 0
    Dim Tokens As String()
    
    If str IsNot Nothing Then
      Tokens = str.Split("|")
      If Tokens IsNot Nothing Then
        Integer.TryParse(Tokens(0), PrefID)
      End If
    End If
    
    Return PrefID
  End Function

  Function Parse_TierValues(ByVal TokenName As String) As List(Of TierValue)
    Dim TierVals As New List(Of TierValue)
    Dim Tier As New TierValue
    Dim TokenStr As String
    Dim Tokens() As String
    Dim Vals() As String
    
    TokenStr = GetCgiValue(TokenName)
    
    If TokenStr IsNot Nothing Then
      Tokens = TokenStr.Split(",")
      For Each s As String In Tokens
        If s IsNot Nothing Then
          Vals = s.Split("|")
          If Vals.Length >= 5 Then
            Integer.TryParse(Vals(0), Tier.OperatorTypeID)
            Tier.Value = Vals(1)
            Integer.TryParse(Vals(2), Tier.DateOperatorTypeID)
            Integer.TryParse(Vals(3), Tier.ValueTypeID)
            Tier.ValueModifier = Vals(4)
          End If
          
          ' handle days before and after when applicable
          If Vals.Length = 7 Then
            Integer.TryParse(Vals(5), Tier.DaysBefore)
            Integer.TryParse(Vals(6), Tier.DaysAfter)
        End If
          
        End If
        TierVals.Add(Tier)
      Next
    End If
    
    Return TierVals
  End Function
  
</script>
