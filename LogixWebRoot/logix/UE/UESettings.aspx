<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: UEsettings.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2010 - 2013.  All rights reserved by:
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
<script type="text/javascript" language="javascript">

  function onEnter(e) {

    var keynum = 0;
    if (window.event) // IE8 and earlier
      keynum = e.keyCode;
    else if (e.which) // IE9/Firefox/Chrome/Opera/Safari
      keynum = e.which;

    if (keynum < 48 || keynum > 57)
      return false;

    return true;
  }
</script>
<script runat="server">

    Dim Common As New Copient.CommonInc
    Dim UIInc As New Copient.LogixInc
    Dim Handheld As Boolean = False
    Dim MyCryptLib As New Copient.CryptLib()
    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_Settings_List(ByVal InfoMessage As String)

        Dim dst As DataTable
        Dim dst2 As DataTable
        Dim dst3 As DataTable
        Dim SODDdst As DataTable
        Dim DependentCheckdst As DataTable
        Dim DependentID As Integer
        Dim DependentValues As String
        Dim row As DataRow
        Dim row2 As DataRow
        Dim row3 As DataRow
        Dim SODDrow As DataRow
        Dim OptionID As Integer
        Dim OptionValue, TempValue As String
        Dim OpenTagEscape As String = "<>"
        Dim SystemOptionTypeID As String
        Dim Counter As Integer = 0
        Dim DependencyOK As Boolean
        Dim DependencyAND As Boolean
        Dim AnyOK, AllOK As Boolean
        Dim DependentValueArray() As String
        Dim HasDependents As Boolean
        Dim SelectionChange As Boolean
        Dim TempStr As String
        Dim Debug As Boolean = False 'show comments in HTML source for debugging the settings page
        Dim index As Integer
        Dim MaxOptionLength As Integer
        Dim WrapOptionLimit As Integer = 150 'the maximum character width where we will display an option and values on a single line. Beyond this we wrap it on 2 lines.
        Dim OptionValueQuery As String = ""
        Dim OptionPhrase As String = ""
        Dim DisableEdit As String = " disabled=""disabled"""

        If (InfoMessage <> "") Then
            Send("<div id=""statusbar"" class=""green-background"">" & InfoMessage & "</div>")
        End If

        Send("<script type=""text/javascript"">")

        Send("function SubmitSelection () { ")
        Send("document.getElementById('selectionchange').value=""changed"";")
        Send("document.mainform.submit();")
        Send("} ")
        Send("</scr" & "ipt>")

        If Debug Then Send("<!-- selectionchage=""" & GetCgiValue("selectionchange") & """ -->")
        SelectionChange = False
        If Not (GetCgiValue("selectionchange") = "") Then
            'the user clicked changed the value on a control that has dependencies 
            'we need to redisplay the page use the values from the page, rather than values from the database
            SelectionChange = True
        End If
        If Debug Then Send("<!-- SelectionChange=" & SelectionChange & " -->")

        Send("<input type=""hidden"" name=""selectionchange"" id=""selectionchange"" value="""" />")
        If UIInc.UserRoles.ViewHiddenOptions = True Then
            Common.QueryStr = "select SO.OptionName, SO.OptionID, isnull(SO.OptionValue, '') as OptionValue, SO.PhraseID, SO.Visible, isnull(SO.OptionTypeID, 0) as SystemOptionTypeID, isnull(SOT.OptionTypePhraseID, 0) as OptionTypePhraseID, isnull(SOT.OptionTypeName, '') as OptionTypeName, " & _
                              "IsNull(PT.Phrase, ISNULL(PTEng.Phrase,SO.OptionName)) as Phrase, isnull(PT.LanguageID, 1) as LanguageID, isnull(SO.DependencyAND, 1) as DependencyAND, isnull(SOT.UIBoxID, 0) as UIBoxID, " & _
                              "OptionValueQuery, UseDisplayQuery, isnull(DisplayQuery, '') as DisplayQuery " & _
                              "from UE_SystemOptions as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID " & _
                               "Left join PhraseText as PTEng with (NoLock) on SO.PhraseID=PTEng.PhraseID and PTEng.LanguageID=1 " & _
                              "Left Join UE_SystemOptionTypes as SOT on SOT.OptionTypeID = SO.OptionTypeID " & _
                              "order by SO.OptionTypeID, SO.DisplayOrder, SO.OptionName;"
        Else
            Common.QueryStr = "select SO.OptionName, SO.OptionID, isnull(SO.OptionValue, '') as OptionValue, SO.PhraseID, SO.Visible, isnull(SO.OptionTypeID, 0) as SystemOptionTypeID, isnull(SOT.OptionTypePhraseID, 0) as OptionTypePhraseID, isnull(SOT.OptionTypeName, '') as OptionTypeName, " & _
                             "IsNull(PT.Phrase, ISNULL(PTEng.Phrase,SO.OptionName)) as Phrase, isnull(PT.LanguageID, 1) as LanguageID, isnull(SO.DependencyAND, 1) as DependencyAND, isnull(SOT.UIBoxID, 0) as UIBoxID, " & _
                             "OptionValueQuery, UseDisplayQuery, isnull(DisplayQuery, '') as DisplayQuery " & _
                             "from UE_SystemOptions as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID " & _
                               "Left join PhraseText as PTEng with (NoLock) on SO.PhraseID=PTEng.PhraseID and PTEng.LanguageID=1 " & _
                             "Left Join UE_SystemOptionTypes as SOT on SOT.OptionTypeID = SO.OptionTypeID " & _
                             "where(Visible = 1) " & _
                             "order by SO.OptionTypeID, SO.DisplayOrder, SO.OptionName;"
        End If
        Common.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
        dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
        SystemOptionTypeID = 0
        If (dst.Rows.Count > 0) Then
            For Each row In dst.Rows

                OptionID = Common.NZ(row.Item("OptionID"), 0)
                OptionValueQuery = row.Item("OptionValueQuery")

                If SelectionChange Then
                    If OptionID = 163 Then
                        OptionValue = MyCryptLib.SQL_StringEncrypt(GetCgiValue("oid" & OptionID))
                    Else
                        OptionValue = GetCgiValue("oid" & OptionID)
                    End If

                    If Debug Then Send("<!-- Using value from form:" & OptionValue & " -->")
                Else
                    OptionValue = row.Item("OptionValue")
                    If Debug Then Send("<!-- Using value from query:" & OptionValue & " -->")
                End If
                DependencyAND = row.Item("DependencyAND")

                'see if this SystemOption has any display dependencies  
                DependencyOK = True
                AnyOK = False
                AllOK = True
                Common.QueryStr = "select isnull(DependentID, 0) as DependentID, isnull(DependentValues, '') as DependentValues from UE_SO_DisplayDependencies where OptionID=@OptionID;"
                Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                SODDdst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                If SODDdst.Rows.Count > 0 Then
                    For Each SODDrow In SODDdst.Rows
                        DependentID = SODDrow.Item("DependentID")
                        DependentValues = SODDrow.Item("DependentValues")
                        If Not (DependentValues = "") Then
                            DependentValueArray = Split(DependentValues, ",")
                            DependentValues = ""
                            For index = 0 To UBound(DependentValueArray)
                                If Not (DependentValues = "") Then DependentValues = DependentValues & ", "
                                DependentValues = DependentValues & "'" & Trim(DependentValueArray(index)) & "'"
                            Next
                        End If
                        If Not (DependentID = 0) And Not (DependentValues = "") Then
                            Common.QueryStr = "select isnull(OptionValue, '') as OptionValue from UE_SystemOptions where OptionID=@OptionID;"
                            Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = DependentID
                            If Debug Then Send("<!-- Dependent check Query=" & Common.QueryStr & " -->")
                            DependentCheckdst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                            If DependentCheckdst.Rows.Count > 0 Then
                                If SelectionChange Then
                                    TempStr = "'" & GetCgiValue("oid" & DependentID) & "'"
                                Else
                                    TempStr = "'" & DependentCheckdst.Rows(0).Item("OptionValue") & "'"
                                End If
                                If Debug Then Send("<!-- tempstr=" & TempStr & "  DependentValues=" & DependentValues & " -->")
                                If InStr(DependentValues, TempStr, CompareMethod.Text) > 0 Then
                                    AnyOK = True
                                Else
                                    AllOK = False
                                End If
                            End If  'The query returned a row
                        End If  'There is a dependency
                    Next
                    If DependencyAND Then
                        'if display dependencies are AND'ed together, then all of the dependencies must be met
                        DependencyOK = AllOK
                    Else
                        'if display dependencies are OR'ed together, then only only of the dependencies must be met
                        DependencyOK = AnyOK
                    End If

                End If
                If Debug Then Send("<!-- DependencyOK=" & DependencyOK & " -->")

                'check the display query in the UE_SystemOptions table
                If DependencyOK And (row.Item("UseDisplayQuery")) And Not (row.Item("DisplayQuery") = "") Then
                    Common.QueryStr = row.Item("DisplayQuery")
                    If Debug Then Send("<!-- DisplayQuery=" & DependencyOK & " -->")
                    dst2 = Common.LRT_Select
                    If dst2.Rows.Count = 0 Then DependencyOK = False
                    dst2 = Nothing
                End If

                If DependencyOK Then
                    'there aren't any display dependencies, or if there are, they've been met

                    'if we are moving on to a new SystemOptionType, then close the previous display box and open a new one
                    If Common.NZ(row.Item("SystemOptionTypeID"), 0) > Counter Then
                        If Not (Counter = 0) Then
                            Send("</table>")
                            Close_UI_Box()
                        End If
                        SystemOptionTypeID = row.Item("SystemOptionTypeID")
                        Open_UI_Box(row.Item("UIBoxID"), AdminUserID, Common, "", "700px")
                        Send("<br class=""half"" />")
                        Send("<table border=""0"" cellspacing=""2"" summary=""" & Copient.PhraseLib.Lookup("term.uesettings", LanguageID) & " " & row.Item("OptionTypeName") & """>")
                    End If

                    'see how long the description and the longest option value are ... so we can wrap if it's too long
                    MaxOptionLength = 0
                    If Not (OptionValueQuery = "") Then
                        Common.QueryStr = OptionValueQuery
                    Else
                        Common.QueryStr = "select IsNull(PhraseID, 0) as PhraseID, '' As PhraseTerm, isnull(SOV.Description, '') as Description " & _
                                          "from UE_SystemOptionValues as SOV with (NoLock) " & _
                                          "where OptionID=@OptionID " & _
                                          "order by OptionValue;"
                        Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                    End If
                    If Debug Then Send("<!-- OptionValueQuery=" & Common.QueryStr & " -->")
                    dst2 = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                    For Each row2 In dst2.Rows
                        If Not (row2.Item("PhraseID") = 0) Then
                            OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID)
                        Else
                            OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseTerm"), LanguageID)
                        End If
                        If OptionPhrase = "" Then OptionPhrase = row2.Item("Description") 'if there's no phrase result, then use the description from the table
                        If Len(OptionPhrase) > MaxOptionLength Then MaxOptionLength = Len(OptionPhrase)
                    Next

                    Sendb("<tr><td" & IIf(row.Item("Visible"), "", " style=""color:red;""") & ">" & Common.NZ(row.Item("Phrase"), "") & ":&nbsp;")
                    If MaxOptionLength > WrapOptionLimit Then Sendb("<br />&nbsp; &nbsp; &nbsp; &nbsp; ")
                    'see if there are any other options that are dependent on this one (for their visibility)
                    HasDependents = False
                    Common.QueryStr = "select 1 from UE_SO_DisplayDependencies where DependentID=@OptionID;"
                    Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                    dst2 = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dst2.Rows.Count > 0 Then
                        HasDependents = True
                    End If

                    Select Case OptionID
                        Case 116, 117, 118 ' default chargeback departments
                            'only display the selection for default chargeback departments if banners are not enabled
                            'if banners are enabled, then each banner could have different chargeback departments associated with it this messes up the concept of system wide defaults
                            If Common.Fetch_SystemOption(66) = "0" Then  'Banners enabled?
                                Common.QueryStr = "select distinct convert(nvarchar,ChargebackDeptID) as ChargebackDeptID, case when IsNull(ExternalID, '') = '' then Name " & _
                                                  " else ExternalID + ' - ' + Name end as OptionText " & _
                                                  " from ChargeBackDepts with (NoLock) where Deleted=0 and ISNULL(BannerID,0)=0"
                                If OptionID = 118 Then ' basket level
                                    Common.QueryStr &= " and ChargeBackDeptID<>0 "
                                ElseIf OptionID = 117 Then ' dept level
                                    Common.QueryStr &= " and ChargeBackDeptID<>0 and ChargeBackDeptID<>14 "
                                ElseIf OptionID = 116 Then 'item level
                                    Common.QueryStr &= " and ChargeBackDeptID<>10 "
                                End If
                                Common.QueryStr &= " order by OptionText;"
                                dst3 = Common.LRT_Select
                                If dst3.Rows.Count = 0 Then
                                    Send("<input type=""text"" id=""option" & OptionID & """ name=""option" & OptionID & """ value=""" & Common.NZ(row.Item("OptionValue"), "") & """" & IIf(row.Item("Visible"), "", DisableEdit) & " />")
                                Else
                                    Send("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("Visible"), "", DisableEdit) & ">")
                                    For Each row3 In dst3.Rows
                                        If row3.Item("ChargebackDeptID") = Common.NZ(row.Item("OptionValue"), "") Then
                                            Send("      <option value=""" & Common.NZ(row3.Item("ChargebackDeptID"), "") & """ selected=""selected"">" & Common.NZ(row3.Item("OptionText"), "") & "</option>")
                                        Else
                                            Send("      <option value=""" & Common.NZ(row3.Item("ChargebackDeptID"), "") & """>" & Common.NZ(row3.Item("OptionText"), "") & "</option>")
                                        End If
                                    Next
                                    Send("    </select>")
                                End If
                            Else
                                'banners are enabled, so we need to send hidden form values to prevent these from being set to blank in the UE_SystemOptions table
                                Send("<input type=""hidden"" id=""oid" & OptionID & """ & name=""oid" & OptionID & """ value=""" & Common.NZ(row.Item("OptionValue"), "") & """ />")
                            End If  'banners not enabled

                        Case Else
                            If Not (OptionValueQuery = "") Then
                                Common.QueryStr = OptionValueQuery
                            Else
                                Common.QueryStr = "select SOV.OptionValue, isnull(SOV.Description, '') as Description, isnull(SOV.PhraseID, 0) as PhraseID, '' as PhraseTerm " & _
                                                  "from UE_SystemOptionValues as SOV with (NoLock) " & _
                                                  "where OptionID=@OptionID " & _
                                                  "order by OptionValue;"
                            End If
                            Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                            dst2 = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dst2.Rows.Count > 0) Then
                                Sendb("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("visible"), "", DisableEdit))
                                If HasDependents Then
                                    Sendb(" onchange=""javacript: SubmitSelection();""")
                                End If
                                Send(">")
                                For Each row2 In dst2.Rows
                                    TempValue = Common.NZ(row2.Item("OptionValue"), "")
                                    TempValue = TempValue.Replace("<", OpenTagEscape)
                                    Sendb("      <option value=""" & TempValue & """")
                                    If Common.NZ(row2.Item("OptionValue"), "") = OptionValue Then Sendb(" selected=""selected""")
                                    If Not (row2.Item("PhraseID") = 0) Then
                                        OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID)
                                    Else
                                        OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseTerm"), LanguageID)
                                    End If
                                    If OptionPhrase = "" Then OptionPhrase = row2.Item("Description")
                                    Send(">" & OptionPhrase & "</option>")
                                Next
                                Send("    </select>")
                            Else
                                Dim rst2 As DataTable
                                Common.QueryStr = "select PriorityID, Name, PhraseID from UE_Priorities with (NoLock);"
                                rst2 = Common.LRT_Select
                                If (OptionID = 163) Then 'Messaging - RabbitMQ Password.
                                    Send("<input type=""password"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30""" & IIf(row.Item("visible"), "", DisableEdit) & " />")
                                ElseIf (OptionID = 226) Then
                                    Dim upc5_Priority As String = Request.Form("upc5priority")
                                    If (String.IsNullOrWhiteSpace(upc5_Priority) And Not OptionValue = "") Then
                                        upc5_Priority = OptionValue
                                    End If
                                    Send("<select id=""upc5priority"" name=""upc5priority"" " & IIf(row.Item("visible"), "", " disabled=""disabled""") & ">")
                                    For Each row2 In rst2.Rows
                                        Send("<option value=""" & Common.NZ(row2.Item("PriorityID"), 0) & """" & IIf(upc5_Priority = Common.NZ(row2.Item("PriorityID"), 0), " selected=""selected""", "") & ">" &
                                       "" & Copient.PhraseLib.Lookup(Common.NZ(row2.Item("PhraseID"), 0), LanguageID, Common.NZ(row2.Item("Name"), "")) & " </option>")
                                    Next
                                    Send("    </select>")
                                    If (Not String.IsNullOrWhiteSpace(upc5_Priority)) Then
                                        Common.QueryStr = "Update UE_SystemOptions with (RowLock) set OptionValue=@Priority  where Visible=1 and OptionID=@OptionID;"
                                        Common.DBParameters.Add("@Priority", SqlDbType.NVarChar, 255).Value = upc5_Priority
                                        Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                                        Common.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                    End If
                                ElseIf (OptionID = 227) Then
                                    Dim gs1_Priority As String = Request.Form("gs1priority")
                                    If (String.IsNullOrWhiteSpace(gs1_Priority) And Not OptionValue = "") Then
                                        gs1_Priority = OptionValue
                                    End If
                                    Send("<select id=""gs1priority"" name=""gs1priority"" " & IIf(row.Item("visible"), "", " disabled=""disabled""") & ">")
                                    For Each row2 In rst2.Rows
                                        Send("<option value=""" & Common.NZ(row2.Item("PriorityID"), 0) & """" & IIf(gs1_Priority = Common.NZ(row2.Item("PriorityID"), 0), " selected=""selected""", "") & ">" &
                                       "" & Copient.PhraseLib.Lookup(Common.NZ(row2.Item("PhraseID"), 0), LanguageID, Common.NZ(row2.Item("Name"), "")) & " </option>")
                                    Next
                                    Send("    </select>")
                                    If (Not String.IsNullOrWhiteSpace(gs1_Priority)) Then
                                        Common.QueryStr = "Update UE_SystemOptions with (RowLock) set OptionValue=@Priority  where Visible=1 and OptionID=@OptionID;"
                                        Common.DBParameters.Add("@Priority", SqlDbType.NVarChar, 255).Value = gs1_Priority
                                        Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                                        Common.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                    End If
                                ElseIf (OptionID = 158) Then
                                    Send("<input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & OptionValue & """" & IIf(row.Item("visible"), "", DisableEdit) & " onkeypress=""return onEnter(event);"" />")
                                Else
                                    Send("<input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & OptionValue & """" & IIf(row.Item("visible"), "", DisableEdit) & " />")
                                End If
                            End If
                    End Select

                    Send("</td></tr>")
                    Counter = Common.NZ(row.Item("SystemOptionTypeID"), 0)
                Else 'DependencyOK = False - send a hidden form field
                    Send("<tr style=""display:none;""><td><input type=""hidden"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ value=""" & OptionValue & """ /></td></tr>")
                End If 'DependencyOK

            Next
            If Counter > 0 Then
                Send("</table>")
                Close_UI_Box()
            End If
        End If

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Sub Send_Page(ByVal InfoMessage As String)
        Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
        Dim CopientFileVersion As String = "1.0b1.0"
        Dim CopientProject As String = "Preference Manager"
        Dim CopientNotes As String = ""

        Send_HeadBegin("term.uesettings")
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts(New String() {"ajaxSubmit.js"})
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Logos()
        Send_Tabs(UIInc, 8)
        Send_Subtabs(UIInc, 8, 4)
        If UIInc.UserRoles.AccessSystemSettings Then
            Send("<form action=""UESettings.aspx"" id=""mainform"" name=""mainform"" method=""post"">")
            Send("<div id=""intro"">")
            'Send("<div id=""gutter""></div>")
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.uesettings", LanguageID) & "</h1>")
            Send("<div id=""controls"">")
            Send_Save()
            Send("</div>  <!-- controls -->")
            Send("</div>  <!-- intro -->")
            Send("<div id=""main"">")
            Send("  <div id=""column"">")
            Send_Settings_List(InfoMessage)
            Send("</div> <!-- column -->")
            Send("</div> <!-- main -->")
            Send("</form>")
        Else
            Send_Denied(1, Common, , "Access system settings")
        End If

        Send_BodyEnd()

    End Sub

    '-------------------------------------------------------------------------------------------------------------  

    Sub Save_Settings()

        Dim dst As DataTable
        Dim tempstr As String
        Dim OptionObj As Copient.SystemOption = Nothing
        Dim HistoryStr, HistoryPhraseStr As String
        Dim OldOptionValueStr, NewOptionValueStr As String
        Dim InfoMessage As String = ""
        Dim row As DataRow
        Dim OVdst As DataTable

        Common.QueryStr = "select OptionID, OptionName, OptionValue, isnull(PhraseID, 0) as PhraseID from UE_SystemOptions with (NoLock) where Visible=1 order by OptionID;"
        dst = Common.LRT_Select
        If (dst.Rows.Count > 0) Then
            For Each row In dst.Rows
                tempstr = GetCgiValue("oid" & row.Item("OptionID"))
                tempstr = UIInc.TrimAll(tempstr)

                OptionObj = New Copient.SystemOption(row.Item("OptionID"), Common.NZ(row.Item("OptionValue"), ""))

                If (OptionObj.GetOptionID = 163) Then
                    If (tempstr.Length > 0) Then
                        OptionObj.SetNewValue(MyCryptLib.SQL_StringEncrypt(tempstr))
                    Else
                        OptionObj.SetNewValue(OptionObj.GetOldValue)
                    End If
                Else
                    OptionObj.SetNewValue(tempstr)
                End If

                If (OptionObj.GetOptionID = 150) Then
                    If (tempstr.Equals("")) Then
                        OptionObj.SetNewValue(tempstr)
                    Else
                        Dim tmp As Integer
                        If (Not Integer.TryParse(OptionObj.GetNewValue(), tmp)) Then
                            OptionObj.SetNewValue(OptionObj.GetOldValue())
                            If InfoMessage = "" Then InfoMessage = "Please enter time in hours"
                        Else
                            If (OptionObj.GetNewValue() > 0) Then
                                OptionObj.SetNewValue(tempstr)
                            Else
                                OptionObj.SetNewValue(OptionObj.GetOldValue())
                                If InfoMessage = "" Then InfoMessage = "Please enter time in hours"
                            End If
                        End If
                    End If
                End If

                If (OptionObj.GetOptionID = 158) Then
                    Dim tmp As Integer
                    If (Not Integer.TryParse(OptionObj.GetNewValue(), tmp)) Then
                        OptionObj.SetNewValue(2)
                    End If
                    If (OptionObj.GetNewValue() > 10) Then
                        OptionObj.SetNewValue(10)
                    End If
                    If (OptionObj.GetNewValue() < 2) Then
                        OptionObj.SetNewValue(2)
                    End If
                End If
                If (OptionObj.GetOptionID = 217) Then

                    If (Regex.IsMatch(OptionObj.GetNewValue(), "^[A-Za-z0-9,]+$") Or OptionObj.GetNewValue() = "") Then
                        OptionObj.SetNewValue(tempstr)
                    Else
                        InfoMessage = "special characters are not allowed in trackable coupon prefix"
                        OptionObj.SetNewValue(OptionObj.GetOldValue())
                    End If
                End If
                If OptionObj.GetOptionID = 226 Then
                    Dim Priority As String = Request.Form("upc5priority")
                    If Not String.IsNullOrEmpty(Priority) Then
                        OptionObj.SetNewValue(Priority)
                    End If
                End If
                If OptionObj.GetOptionID = 227 Then
                    Dim Priority As String = Request.Form("gs1priority")
                    If Not String.IsNullOrEmpty(Priority) Then
                        OptionObj.SetNewValue(Priority)
                    End If
                End If
                If OptionObj.GetOptionID = 237 AndAlso OptionObj.GetNewValue() = "0" Then
                    If Common.ItemDeptExistsForOffer() Then
                        InfoMessage = Copient.PhraseLib.Lookup("error.proration", LanguageID)
                        OptionObj.SetNewValue(OptionObj.GetOldValue())
                    End If
                End If
                If OptionObj.IsModified Then
                    Common.QueryStr = "Update UE_SystemOptions with (RowLock) set OptionValue=@OptionValue, LastUpdate=getdate() where Visible=1 and OptionID=@OptionID;"
                    Common.DBParameters.Add("@OptionValue", SqlDbType.NVarChar, 255).Value = OptionObj.GetNewValue()
                    Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID()
                    Common.ExecuteNonQuery(Copient.DataBases.LogixRT)

                    If (Common.RowsAffected > 0) Then
                        If InfoMessage = "" Then InfoMessage = Copient.PhraseLib.Lookup("term.changessaved", LanguageID)

                        OldOptionValueStr = OptionObj.GetOldValue
                        Common.QueryStr = "select isnull(PhraseID, 0) as PhraseID from UE_SystemOptionValues where OptionID=@OptionID and OptionValue=@OptionValue;"
                        Common.DBParameters.Add("@OptionValue", SqlDbType.NVarChar, 255).Value = OldOptionValueStr
                        Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID
                        OVdst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                        If OVdst.Rows.Count > 0 Then
                            OldOptionValueStr = "{" & OVdst.Rows(0).Item("PhraseID") & "}"
                        End If

                        NewOptionValueStr = OptionObj.GetNewValue
                        Common.QueryStr = "select isnull(PhraseID, 0) as PhraseID from UE_SystemOptionValues where OptionID=@OptionID and OptionValue=@OptionValue;"
                        Common.DBParameters.Add("@OptionValue", SqlDbType.NVarChar, 255).Value = NewOptionValueStr
                        Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID
                        OVdst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                        If OVdst.Rows.Count > 0 Then
                            NewOptionValueStr = "{" & OVdst.Rows(0).Item("PhraseID") & "}"
                        End If

                        If (OptionObj.GetOptionID = 163) Then
                            HistoryStr = "Edited system settings '" & Common.NZ(row.Item("OptionName"), "") & "'"
                            HistoryPhraseStr = "[history.editsetting] '{" & row.Item("PhraseID") & "}'"
                        Else
                            HistoryStr = "Edited system settings '" & Common.NZ(row.Item("OptionName"), "") & "' from: " & OptionObj.GetOldValue() & " to: " & OptionObj.GetNewValue()
                            HistoryPhraseStr = "[history.editsetting] '{" & row.Item("PhraseID") & "}' [term.from]: " & OldOptionValueStr & " [term.to]: " & NewOptionValueStr
                        End If
                        Common.Activity_Log(56, 0, 0, AdminUserID, HistoryStr, HistoryPhraseStr)
                    End If
                End If
            Next

            ' Refresh cache with new system options
            CMS.AMS.CurrentRequest.Resolver.AppName = Common.AppName
            Dim cacheData As CMS.AMS.Contract.ICacheData = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.ICacheData)()
            cacheData.ClearAllSystemOptionsCache()
            Copient.SystemOptionsCache.RemoveCache(System.Web.HttpContext.Current.Request.Url.Host)
        End If
        Send_Page(InfoMessage)
    End Sub

</script>

<%
  
  '-------------------------------------------------------------------------------------------------------------  
  'Execution begins here ...
  
  Common.AppName = "UESettings.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
  
  AdminUserID = Verify_AdminUser(Common, UIInc)
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  If GetCgiValue("save") <> "" Then
    Save_Settings()
  Else
    Send_Page("")
  End If
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
  UIInc = Nothing

  Response.End()
  
ErrorTrap:
  Response.Write("<pre>" & Common.Error_Processor() & "</pre>")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  Common = Nothing
  UIInc = Nothing
  
%>