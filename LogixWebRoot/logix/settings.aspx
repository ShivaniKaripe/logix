<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%
    ' *****************************************************************************
    ' * FILENAME: settings.aspx 
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
<script type="text/javascript">

    var nVer = navigator.appVersion;
    var nAgt = navigator.userAgent;
    var browserName = navigator.appName;
    var nameOffset, verOffset, ix;

    var browser = navigator.appName;

    // In Opera, the true version is after "Opera" or after "Version"
    if ((verOffset = nAgt.indexOf("Opera")) != -1) {
        browserName = "Opera";
    }
    // In MSIE, the true version is after "MSIE" in userAgent
    else if ((verOffset = nAgt.indexOf("MSIE")) != -1) {
        browserName = "IE";
    }
    // In Chrome, the true version is after "Chrome" 
    else if ((verOffset = nAgt.indexOf("Chrome")) != -1) {
        browserName = "Chrome";
    }
    // In Safari, the true version is after "Safari" or after "Version" 
    else if ((verOffset = nAgt.indexOf("Safari")) != -1) {
        browserName = "Safari";
    }
    // In Firefox, the true version is after "Firefox" 
    else if ((verOffset = nAgt.indexOf("Firefox")) != -1) {
        browserName = "Firefox";
    }
    // In most other browsers, "name/version" is at the end of userAgent 
    else if ((nameOffset = nAgt.lastIndexOf(' ') + 1) <
          (verOffset = nAgt.lastIndexOf('/'))) {
        browserName = nAgt.substring(nameOffset, verOffset);
        fullVersion = nAgt.substring(verOffset + 1);
        if (browserName.toLowerCase() == browserName.toUpperCase()) {
            browserName = navigator.appName;
        }
    }


    if (browserName == "IE") {
        document.attachEvent("onclick", PageClick);
    }
    else {
        document.onclick = function (evt) {
            var target = document.all ? event.srcElement : evt.target;
            if (target.href) {
                if (IsFormChanged(document.mainform)) {
                    var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                    return bConfirm;
                }
            }
        };
    }
    function PageClick(evt) {
        var target = document.all ? event.srcElement : evt.target;

        if (target.href) {
            if (IsFormChanged(document.mainform)) {
                var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                return bConfirm;
            }
        }
    }
   
</script>
<script runat="server">

    Dim Common As New Copient.CommonInc
    Dim UIInc As New Copient.LogixInc
    Dim Handheld As Boolean = False

    '-------------------------------------------------------------------------------------------------------------  



    Sub Send_Settings_List(ByVal InfoMessage As String)

        Dim dst As DataTable
        Dim dst2 As DataTable
        Dim SODDdst As DataTable
        Dim DependentCheckdst As DataTable
        Dim DependentID As Integer
        Dim DependentValues As String
        Dim row As DataRow
        Dim row2 As DataRow
        Dim SODDrow As DataRow
        Dim OptionID As Integer
        Dim OptionValue As String
        Dim TempValue As String
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
        Dim Crypt As New Copient.CryptLib
        Dim IsAccountUnlockEnabled As Boolean = False
        If (InfoMessage <> "") Then
            Dim strParts() As String = InfoMessage.Split(":")
            If strParts(0) = "Error" Or strParts(0) = Copient.PhraseLib.Lookup("term.error", LanguageID) Then
                Send("<div id=""statusbar"" class=""red-background"">" & InfoMessage & "</div>")
            Else
                Send("<div id=""statusbar"" class=""green-background"">" & InfoMessage & "</div>")
            End If
        End If

        Send("<script type=""text/javascript"">")
        Send("$( document ).ready(function() {")
        Send("var value = $('#oid192').val()")
        Send("if(value == '1') {")
        Send("$('#oid191').val('1')")
        Send("$('#oid191').attr('disabled', 'disabled')")
        Send("}")
        Send("});")
        Send("function SubmitSelection () { ")
        Send("document.getElementById('selectionchange').value=""changed"";")
        Send("document.mainform.submit();")
        Send("} ")
        Send("function DisableControl () { ")
        Send("var text=document.getElementById('oid192').value")
        Send("if(text==1) {")
        Send("document.getElementById('oid191').value=1")
        Send("document.getElementById('oid191').disabled=true")
        Send("}")
        Send("else {")
        Send("document.getElementById('oid191').disabled=false")
        Send("}")
        Send("}")
        Send("function AccountLockoutControls () { ")
        Send("var text=document.getElementById('oid334').value")
        Send("if(text==0) {")
        Send("$('#oid335').attr('readonly', true)")
        Send("$('#oid336').attr('readonly', true)")
        Send("$('#oid335').css('background-color' , '#DEDEDE')")
        Send("$('#oid336').css('background-color' , '#DEDEDE')")
        Send("}")
        Send("else {")
        Send("$('#oid335').attr('readonly', false)")
        Send("$('#oid336').attr('readonly', false)")
        Send("$('#oid335').css('background-color' , '')")
        Send("$('#oid336').css('background-color' , '')")
        Send("}")
        Send("}")
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
            Common.QueryStr = "select SO.OptionName, SO.OptionID, isnull(SO.OptionValue, '') as OptionValue, SO.PhraseID, SO.Visible, isnull(SO.OptionTypeID, 0) as SystemOptionTypeID, isnull(SOT.OptionTypePhraseID, 0) as OptionTypePhraseID, isnull(SOT.OptionTypeName, '') as OptionTypeName, " &
                            "IsNull(PT.Phrase, ISNULL(PTEng.Phrase,SO.OptionName)) as Phrase, isnull(PT.LanguageID, 1) as LanguageID, isnull(SO.DependencyAND, 1) as DependencyAND, isnull(SOT.UIBoxID, 0) as UIBoxID, " &
                            "OptionValueQuery, UseDisplayQuery, isnull(DisplayQuery, '') as DisplayQuery " &
                            "from SystemOptions as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID " &
                            "Left join PhraseText as PTEng with (NoLock) on SO.PhraseID=PTEng.PhraseID and PTEng.LanguageID=1 " &
                            "Left Join SystemOptionTypes as SOT on SOT.OptionTypeID = SO.OptionTypeID " &
                            "order by SO.OptionTypeID, SO.DisplayOrder, SO.OptionName;"
        Else
            Common.QueryStr = "select SO.OptionName, SO.OptionID, isnull(SO.OptionValue, '') as OptionValue, SO.PhraseID, SO.Visible, isnull(SO.OptionTypeID, 0) as SystemOptionTypeID, isnull(SOT.OptionTypePhraseID, 0) as OptionTypePhraseID, isnull(SOT.OptionTypeName, '') as OptionTypeName, " &
                              "IsNull(PT.Phrase, ISNULL(PTEng.Phrase,SO.OptionName)) as Phrase, isnull(PT.LanguageID, 1) as LanguageID, isnull(SO.DependencyAND, 1) as DependencyAND, isnull(SOT.UIBoxID, 0) as UIBoxID, " &
                              "OptionValueQuery, UseDisplayQuery, isnull(DisplayQuery, '') as DisplayQuery " &
                              "from SystemOptions as SO with (NoLock) left join PhraseText as PT with (NoLock) on SO.PhraseID=PT.PhraseID and PT.LanguageID=@LanguageID " &
                              "Left join PhraseText as PTEng with (NoLock) on SO.PhraseID=PTEng.PhraseID and PTEng.LanguageID=1 " &
                              "Left Join SystemOptionTypes as SOT on SOT.OptionTypeID = SO.OptionTypeID " &
                              "where(Visible = 1) " &
                              "order by SO.OptionTypeID, SO.DisplayOrder, SO.OptionName;"
        End If
        Common.DBParameters.Add("@LanguageID", SqlDbType.Int).Value = LanguageID
        dst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
        SystemOptionTypeID = 0
        If (dst.Rows.Count > 0) Then
            For Each row In dst.Rows
                OptionID = Common.NZ(row.Item("OptionID"), 0)
                If (ValidateSystemOptionRolePermission(OptionID)) Then
                    OptionValueQuery = row.Item("OptionValueQuery")
                    If SelectionChange Then
                        OptionValue = GetCgiValue("oid" & OptionID)
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
                    Common.QueryStr = "select isnull(DependentID, 0) as DependentID, isnull(DependentValues, '') as DependentValues from SO_DisplayDependencies where OptionID=@OptionID;"
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

                                Common.QueryStr = "select isnull(OptionValue, '') as OptionValue from SystemOptions where OptionID=@OptionID;"
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

                    'check the display query in the SystemOptions table
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
                            Send("<BR class=""half"" />")
                            Send("<table border=0 cellspacing=2 summary=""" & Copient.PhraseLib.Lookup("term.settings", LanguageID) & " " & row.Item("OptionTypeName") & """>")
                        End If

                        'see how long the description and the longest option value are ... so we can wrap if it's too long
                        MaxOptionLength = 0
                        If Not (OptionValueQuery = "") Then
                            Common.QueryStr = OptionValueQuery
                        Else
                            Common.QueryStr = "select IsNull(PhraseID, 0) as PhraseID, '' As PhraseTerm, isnull(SOV.Description, '') as Description " &
                                              "from SystemOptionValues as SOV with (NoLock) " &
                                              "where OptionID=@OptionID " &
                                              "order by OptionValue;"
                            Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                        End If
                        If Debug Then Send("<!-- OptionValueQuery=" & Common.QueryStr & " -->")
                        dst2 = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                        For Each row2 In dst2.Rows
                            If Not (Common.NZ(row2.Item("PhraseID"), 0) = 0) Then
                                OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID)
                            ElseIf Not (Common.NZ(row2.Item("PhraseTerm"), "") = "") Then
                                OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseTerm"), LanguageID)
                            Else
                                OptionPhrase = row2.Item("Description") 'if there's no phrase result, then use the description from the table
                            End If
                            If Len(OptionPhrase) > MaxOptionLength Then MaxOptionLength = Len(OptionPhrase)
                        Next

                        Sendb("<TR><TD" & IIf(row.Item("Visible"), "", " style=""color:red;""") & ">" & Common.NZ(row.Item("Phrase"), "") & ":&nbsp;")
                        If MaxOptionLength > WrapOptionLimit Then Sendb("<BR>&nbsp; &nbsp; &nbsp; &nbsp; ")
                        'see if there are any other options that are dependent on this one (for their visibility)
                        HasDependents = False
                        Common.QueryStr = "select 1 from SO_DisplayDependencies where DependentID=@OptionID;"
                        Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                        dst2 = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                        If dst2.Rows.Count > 0 Then
                            HasDependents = True
                        End If

                        If Not (OptionValueQuery = "") Then
                            Common.QueryStr = OptionValueQuery
                        Else
                            Common.QueryStr = "select isnull(SOV.OptionValue, '') as OptionValue, isnull(SOV.Description, '') as Description, isnull(SOV.PhraseID, 0) as PhraseID, '' as PhraseTerm " &
                                              "from SystemOptionValues as SOV with (NoLock) " &
                                              "where OptionID=@OptionID " &
                                              "order by OptionValue;"
                            Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionID
                        End If
                        dst2 = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dst2.Rows.Count > 0) Then
                            Sendb("    <select id=""oid" & OptionID & """ name=""oid" & OptionID & """" & IIf(row.Item("visible"), "", DisableEdit))
                            If HasDependents Then
                                Sendb(" onchange=""javacript: SubmitSelection();""")
                            End If
                            If OptionID = 192 Then
                                Send(" onchange=""javacript: DisableControl();""")
                            ElseIf OptionID = 334 Then
                                Send(" onchange=""javacript: AccountLockoutControls();""")
                                If row.Item("OptionValue") = "1" Then
                                    IsAccountUnlockEnabled = True
                                End If
                            End If
                            Send(">")
                            For Each row2 In dst2.Rows
                                TempValue = Common.NZ(row2.Item("OptionValue"), "")
                                TempValue = TempValue.Replace("<", OpenTagEscape)
                                Sendb("      <option value=""" & TempValue & """")
                                If row2.Item("OptionValue").ToString = OptionValue Then Sendb(" selected=""selected""")
                                If Not (Common.NZ(row2.Item("PhraseID"), 0) = 0) Then
                                    OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID)
                                ElseIf Not (Common.NZ(row2.Item("PhraseTerm"), "") = "") Then
                                    OptionPhrase = Copient.PhraseLib.Lookup(row2.Item("PhraseTerm"), LanguageID)
                                Else
                                    OptionPhrase = row2.Item("Description")
                                End If

                                Send(">" & OptionPhrase & "</option>")
                            Next
                            Send("    </select>")
                            ' End If
                        ElseIf (OptionID = 232) Then
                            Send("<input type=""password"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30""" & IIf(row.Item("visible"), "", DisableEdit) & " />")
                        ElseIf (OptionID = 299 Or OptionID = 300) Then
                            Send("<input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & OptionValue & """" & IIf(row.Item("visible"), "", DisableEdit) & " maxlength=""2"" />")
                        ElseIf (OptionID = 335 Or OptionID = 336) Then
                            'User Account Lockout Seeting 335 , 336 should be only editable when OptionID 334 is True
                            Dim TempOptionValueUI = GetCgiValue("oid334")
                            TempOptionValueUI = UIInc.TrimAll(TempOptionValueUI)
                            If (IsAccountUnlockEnabled) Then
                                Send("<input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & OptionValue & """" & IIf(row.Item("visible"), "", DisableEdit) & " />")
                            Else
                                Send("<input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & OptionValue & """" & IIf(row.Item("visible"), "", DisableEdit) & " readonly= ""readonly"" style=""background-color: rgb(222, 222, 222)""/>")
                            End If
                        Else
                            Send("<input type=""text"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ size=""30"" value=""" & OptionValue & """" & IIf(row.Item("visible"), "", DisableEdit) & " />")

                        End If
                        Send("</TD></TR>")
                        Counter = Common.NZ(row.Item("SystemOptionTypeID"), 0)

                    Else 'DependencyOK = False - send a hidden form field
                        Send("<input type=""hidden"" id=""oid" & OptionID & """ name=""oid" & OptionID & """ value=""" & OptionValue & """ />")
                    End If 'DependencyOK
                End If
            Next
            If Not (Counter = 0) Then
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

        Send_HeadBegin("term.settings")
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
            Send("<form action=""settings.aspx"" id=""mainform"" name=""mainform"" method=""post"">")
            Send("<div id=""intro"">")
            'Send("<div id=""gutter""></div>")
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.settings", LanguageID) & "</h1>")
            Send("<div id=""controls"">")
            If UIInc.UserRoles.AccessSystemSettings Then
                Send_Save()
            End If
            Send("</div>  <!-- controls -->")

            Send("</div>  <!-- intro -->")
            Send("<div id=""main"">")
            'Send("<div class=""gutter""></div>")
            Send("  <div id=""column"">")
            Send_Settings_List(InfoMessage)
            Send("</div> <!-- column -->")
            Send("</div> <!-- main -->")
            Send("</form>")

        Else
            Send_Denied(1, "perm.admin-settings")
        End If

        Send_BodyEnd()

    End Sub
    '-------------------------------------------------------------------------------------------------------------  


    Function ValidateSystemOptionRolePermission(ByVal SystemOptionID As Integer)
        Dim retValue As Boolean = True
        Select Case SystemOptionID
            Case 334, 335, 336
                'User Account Lockout
                retValue = UIInc.UserRoles.AccessUserAccountLockout
        End Select
        Return retValue
    End Function
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
        Dim Option184Value As String = ""
        Dim intOption184Value As Integer = 0
        Dim Option192Value As String = ""
        Dim Option308Value As String = ""
        Dim intOption308Value As Integer = 0
        Dim intOption192Value As Integer = 0
        Dim Option296Value As String = ""
        Dim intOption296Value As Integer = 0
        Dim OTPValue As String = ""
        Dim intOTPValue As Integer = 0
        Dim intOTPAttemptsValue As Integer = 0
        Dim OTPAttemptsValue As String = String.Empty
        Dim Option232Value As String = String.Empty
        Dim OptionValue As String = ""
        Dim intOptionValue As Integer = 0


        Dim Crypt As New Copient.CryptLib

        Common.QueryStr = "select OptionID, OptionName, OptionValue, isnull(PhraseID, 0) as PhraseID from SystemOptions with (NoLock) where Visible=1 order by OptionID desc;"
        dst = Common.LRT_Select
        If (dst.Rows.Count > 0) Then
            For Each row In dst.Rows
                If (ValidateSystemOptionRolePermission(Integer.Parse(row.Item("OptionID")))) Then
                    tempstr = GetCgiValue("oid" & row.Item("OptionID"))
                    tempstr = UIInc.TrimAll(tempstr)
                    If (row.Item("OptionID") = "335") Then
                        Dim a As Boolean = False
                    End If
                    OptionObj = New Copient.SystemOption(row.Item("OptionID"), Common.NZ(row.Item("OptionValue"), ""))
                    OptionObj.SetNewValue(tempstr)

                    If OptionObj.IsModified Then
                        If OptionObj.GetOptionID() = 42 Then '' Check whether Logging path is valid and accessible
                            Try
                                If Not String.IsNullOrWhiteSpace(OptionObj.GetNewValue()) Then

                                    If Not Directory.Exists(Common.Parse_Quotes(OptionObj.GetNewValue())) Then
                                        If (OptionObj.GetNewValue.Contains("\")) Then
                                            Directory.CreateDirectory(Common.Parse_Quotes(OptionObj.GetNewValue()))
                                        Else
                                            Throw New Exception("term.invalidpath")
                                        End If
                                    Else
                                        Try
                                            Using fs As FileStream = File.Create(Path.Combine(Common.Parse_Quotes(OptionObj.GetNewValue()), "Access.txt"), 1, FileOptions.DeleteOnClose)
                                            End Using

                                        Catch
                                            InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ":" & Copient.PhraseLib.Lookup("term.writeaccess", LanguageID)
                                            Send_Page(InfoMessage)
                                            Return
                                        End Try
                                    End If
                                End If
                            Catch dnf As DirectoryNotFoundException
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ":" & Copient.PhraseLib.Lookup("term.invalidpath", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            Catch ioe As IOException
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ":" & Copient.PhraseLib.Lookup("term.invalidpath", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            Catch uae As UnauthorizedAccessException
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ":" & Copient.PhraseLib.Lookup("term.invalidpath", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            Catch ae As ArgumentException
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ":" & Copient.PhraseLib.Lookup("term.invalidpath", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            Catch nse As NotSupportedException
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ":" & Copient.PhraseLib.Lookup("term.invalidpath", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            Catch ex As Exception
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ":" & Copient.PhraseLib.Lookup("term.invalidpath", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End Try
                        End If

                        If row.Item("OptionID") = 184 Then
                            Option184Value = Common.Parse_Quotes(OptionObj.GetNewValue())
                            If Option184Value.Trim <> "" Then
                                If IsNumeric(Option184Value) Then
                                    intOption184Value = Convert.ToInt32(Option184Value)
                                Else
                                    InfoMessage = "Error: Invalid " & Copient.PhraseLib.Lookup("settings.184", LanguageID)
                                    Send_Page(InfoMessage)
                                    Return
                                End If
                            Else
                                InfoMessage = "Error: Invalid " & Copient.PhraseLib.Lookup("settings.184", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            If intOption184Value <= 0 Or intOption184Value > 99 Then
                                InfoMessage = "Error: Invalid " & Copient.PhraseLib.Lookup("settings.184", LanguageID) & " ; Max value is 99."
                                Send_Page(InfoMessage)
                                Return
                            End If
                        End If
                        If row.Item("OptionID") = 308 Then
                            Dim dvz2 As Integer = 0
                            Option308Value = Common.Parse_Quotes(OptionObj.GetNewValue())
                            If Option308Value.Trim <> "" Then
                                If IsNumeric(Option308Value) AndAlso Integer.TryParse(Option308Value, dvz2) = True Then
                                    intOption308Value = Convert.ToInt32(Option308Value)
                                Else
                                    InfoMessage = "Error: Invalid " & Copient.PhraseLib.Lookup("settings.308", LanguageID)
                                    Send_Page(InfoMessage)
                                    Return
                                End If
                            Else
                                InfoMessage = "Error: Invalid " & Copient.PhraseLib.Lookup("settings.308", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            If intOption308Value <= 0 Then
                                InfoMessage = "Error: Invalid " & Copient.PhraseLib.Lookup("settings.308", LanguageID) & " ; " & Copient.PhraseLib.Lookup("error.setting308", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                        End If
                        If row.Item("OptionID") = 296 Then
                            Option296Value = Common.Parse_Quotes(OptionObj.GetNewValue())
                            If Option296Value.Trim <> "" Then
                                If IsNumeric(Option296Value) And Integer.TryParse(Option296Value, intOption296Value) Then
                                    intOption296Value = Convert.ToInt32(Option296Value)
                                Else
                                    InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.passwordexpirysetting", LanguageID)
                                    Send_Page(InfoMessage)
                                    Return
                                End If
                            Else
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.passwordexpirysetting", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            If intOption296Value < 0 Then
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.passwordexpirysetting", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            OptionObj.SetNewValue(intOption296Value)
                        End If
                        If (row.Item("OptionID") = 299) Then
                            OTPValue = Common.Parse_Quotes(OptionObj.GetNewValue())
                            If OTPValue.Trim <> "" Then
                                If IsNumeric(OTPValue) And Integer.TryParse(OTPValue, intOTPValue) Then
                                    intOTPValue = Convert.ToInt32(OTPValue)
                                Else
                                    InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.otpdurationnumeric", LanguageID)
                                    Send_Page(InfoMessage)
                                    Return
                                End If
                            Else
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.otpdurationnumeric", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            If intOTPValue <= 0 Then
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.otpdurationnumeric", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            OptionObj.SetNewValue(intOTPValue)
                        End If
                        If (row.Item("OptionID") = 300) Then
                            OTPAttemptsValue = Common.Parse_Quotes(OptionObj.GetNewValue())
                            If OTPAttemptsValue.Trim <> "" Then
                                If IsNumeric(OTPAttemptsValue) And Integer.TryParse(OTPAttemptsValue, intOTPAttemptsValue) Then
                                    intOTPAttemptsValue = Convert.ToInt32(OTPAttemptsValue)
                                Else
                                    InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.otpretrynumeric", LanguageID)
                                    Send_Page(InfoMessage)
                                    Return
                                End If
                            Else
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.otpretrynumeric", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            If intOTPAttemptsValue < 0 Then
                                InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.otpretrynumeric", LanguageID)
                                Send_Page(InfoMessage)
                                Return
                            End If
                            OptionObj.SetNewValue(intOTPAttemptsValue)
                        End If
                        If row.Item("OptionID") = 232 Then
                            Option232Value = OptionObj.GetNewValue()
                            If (Option232Value.Trim().Length > 0) Then
                                Option232Value = Crypt.SQL_StringEncrypt(Option232Value)
                                OptionObj.SetNewValue(Option232Value)
                            Else
                                OptionObj.SetNewValue(OptionObj.GetOldValue)
                            End If
                        End If
                        'Checking when the system option 191 is disabled'
                        If row.Item("OptionID") = 191 Then
                            Option192Value = Common.Parse_Quotes(OptionObj.GetNewValue())
                            If Option192Value.Trim = "" Then
                                OptionObj.SetNewValue("1")
                            End If
                        End If
                        If row.Item("OptionID") = 335 Or row.Item("OptionID") = 336 Then
                            OptionValue = Common.Parse_Quotes(OptionObj.GetNewValue())
                            Dim isValid As Boolean = False
                            If IsNumeric(OptionValue) = True Then
                                If Integer.TryParse(OptionValue, intOptionValue) = True Then
                                    If Integer.Parse(OptionValue) > 0 Then
                                        isValid = True
                                    End If
                                End If
                            End If
                            If isValid = False Then
                                If row.Item("OptionID") = 335 Then
                                    InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.consecutivefailedlogin", LanguageID)
                                ElseIf row.Item("OptionID") = 336 Then
                                    InfoMessage = Copient.PhraseLib.Lookup("term.error", LanguageID) & ": " & Copient.PhraseLib.Lookup("error.autounlockperiod", LanguageID)
                                End If
                                Send_Page(InfoMessage)
                                Return
                            End If
                        End If

                        Common.QueryStr = "Update SystemOptions with (RowLock) set OptionValue=@OptionValue, LastUpdate=getdate() where Visible=1 and OptionID=@OptionID;"
                        Common.DBParameters.Add("@OptionValue", SqlDbType.NVarChar, 255).Value = OptionObj.GetNewValue()
                        Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID()
                        Common.ExecuteNonQuery(Copient.DataBases.LogixRT)

                        If (Common.RowsAffected > 0) Then
                            If InfoMessage = "" Then InfoMessage = Copient.PhraseLib.Lookup("term.changessaved", LanguageID)

                            OldOptionValueStr = OptionObj.GetOldValue
                            Common.QueryStr = "select isnull(PhraseID, 0) as PhraseID from SystemOptionValues where OptionID=@OptionID and OptionValue=@OptionValue;"
                            Common.DBParameters.Add("@OptionValue", SqlDbType.NVarChar, 255).Value = OldOptionValueStr
                            Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID
                            OVdst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                            If OVdst.Rows.Count > 0 Then
                                OldOptionValueStr = "{" & OVdst.Rows(0).Item("PhraseID") & "}"
                            End If

                            NewOptionValueStr = OptionObj.GetNewValue
                            Common.QueryStr = "select isnull(PhraseID, 0) as PhraseID from SystemOptionValues where OptionID=@OptionID and OptionValue=@OptionValue;"
                            Common.DBParameters.Add("@OptionValue", SqlDbType.NVarChar, 255).Value = NewOptionValueStr
                            Common.DBParameters.Add("@OptionID", SqlDbType.Int).Value = OptionObj.GetOptionID
                            OVdst = Common.ExecuteQuery(Copient.DataBases.LogixRT)
                            If OVdst.Rows.Count > 0 Then
                                NewOptionValueStr = "{" & OVdst.Rows(0).Item("PhraseID") & "}"
                            End If
                            If OptionObj.GetOptionID() = 232 Then
                                HistoryStr = "Edited system settings '" & Common.NZ(row.Item("OptionName"), "") & "' from: ****** to: ****** "
                                HistoryPhraseStr = "[history.editsetting] '{" & row.Item("PhraseID") & "}' [term.from]: ****** [term.to]: ******"
                            Else
                                HistoryStr = "Edited system settings '" & Common.NZ(row.Item("OptionName"), "") & "' from: " & OptionObj.GetOldValue() & " to: " & OptionObj.GetNewValue()
                                HistoryPhraseStr = "[history.editsetting] '{" & row.Item("PhraseID") & "}' [term.from]: " & OldOptionValueStr & " [term.to]: " & NewOptionValueStr
                            End If

                            Common.Activity_Log(24, 0, 0, AdminUserID, HistoryStr, HistoryPhraseStr)
                        End If
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
  
    Common.AppName = "settings.aspx"
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