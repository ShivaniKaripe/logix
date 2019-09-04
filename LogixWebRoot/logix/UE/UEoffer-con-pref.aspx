<%@ Page Language="vb" Debug="true" CodeFile="ue-cb.vb" Inherits="UECB" %>

<%@ Import Namespace="System.Collections.Generic" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: UEoffer-con-pref.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2011.  All rights reserved by:
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

  Response.Expires = 0
  Send_Page()

  %>
<script runat="server">
    Structure OfferInfo
        Dim ID As Long
        Dim ROID As Long
        Dim Name As String
        Dim IsTemplate As Boolean
        Dim FromTemplate As Boolean
        Dim DisallowEdit As Boolean
        Dim EngineID As Integer
        Dim TierCount As Integer
    End Structure

    Structure FlagData
        Dim BannersEnabled As Boolean
        Dim CloseAfterSave As Boolean
    End Structure

    Structure Preference
        Dim ID As Long
        Dim Name As String
        Dim DataTypeID As Integer
        Dim DataTypeName As String
        Dim MultiValue As Boolean
    End Structure

    Structure TierData
        Dim PKID As Integer
        Dim Level As Integer
        Dim ValueComboID As Integer
        Dim Values As List(Of TierValue)
    End Structure

    Structure TierValue
        Dim OperatorTypeID As Integer
        Dim Value As String
    End Structure

    Structure PreferenceCondition
        Dim IncentivePrefsID As Integer
        Dim Preferences As List(Of Preference)
        Dim SelectedPref As Preference
        Dim Offer As OfferInfo
        Dim Flags As FlagData
        Dim Tiers As List(Of TierData)
    End Structure

    Dim Common As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim PrefCon As New PreferenceCondition

    '------------------------------------------------------------------------------------------------------------------  
    Dim bEnableAdditionalLockoutRestrictionsOnOffers As Boolean = False
    Dim bOfferEditable As Boolean = False

    Sub Populate_Offer_Data()
        Dim dt As DataTable

        Common.QueryStr = "select INC.IncentiveName, INC.IsTemplate, INC.FromTemplate, RO.RewardOptionID, INC.EngineID, RO.TierLevels " & _
                          "from CPE_Incentives as INC with (NoLock) " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID = INC.IncentiveID " & _
                          "where INC.Deleted=0 and RO.Deleted=0 and RO.TouchResponse=0 and INC.IncentiveID=" & PrefCon.Offer.ID
        dt = Common.LRT_Select
        If dt.Rows.Count > 0 Then
            With PrefCon.Offer
                .Name = Common.NZ(dt.Rows(0).Item("IncentiveName"), "")
                .IsTemplate = Common.NZ(dt.Rows(0).Item("IsTemplate"), False)
                .FromTemplate = Common.NZ(dt.Rows(0).Item("FromTemplate"), False)
                .ROID = Common.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
                .EngineID = Common.NZ(dt.Rows(0).Item("EngineID"), 2)
                .TierCount = Common.NZ(dt.Rows(0).Item("TierLevels"), 1)
            End With
        End If
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Populate_Preferences()
        Dim dt As DataTable
        Dim Pref As Preference

        Common.QueryStr = "select PREF.PreferenceID, PREF.DataTypeID, PREF.MultiValue, PREF.Name as PrefName " & _
                          "from Preferences as PREF with (NoLock) " & _
                          "inner join PreferenceChannels as PC with (NoLock) on PC.PreferenceID=PREF.PreferenceID " & _
                          "where PREF.Deleted=0 and PREF.DataTypeID in (1,2,4,5,7,8) and PC.ChannelID=1 order by PrefName;"
        dt = Common.PMRT_Select
        If dt.Rows.Count > 0 Then
            PrefCon.Preferences = New List(Of Preference)

            For Each row As DataRow In dt.Rows
                Pref = New Preference
                With Pref
                    .ID = Common.NZ(row.Item("PreferenceID"), 0)
                    .Name = Common.NZ(row.Item("PrefName"), "")
                    .DataTypeID = Common.NZ(row.Item("DataTypeID"), 0)
                    .MultiValue = Common.NZ(row.Item("MultiValue"), False)
                End With
                PrefCon.Preferences.Add(Pref)
            Next
        End If

    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Populate_Flags()
        PrefCon.Flags.BannersEnabled = (Common.Fetch_SystemOption(66) = "1")
        PrefCon.Flags.CloseAfterSave = (Common.Fetch_SystemOption(48) = "1")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Load_Saved_Prefence()
        Dim dt As DataTable

        If PrefCon.IncentivePrefsID > 0 Then
            ' find the selected preference for this condition
            Common.QueryStr = "select PreferenceID, DisallowEdit from CPE_IncentivePrefs with (NoLock) " & _
                              "where IncentivePrefsID=" & PrefCon.IncentivePrefsID
            dt = Common.LRT_Select
            If dt.Rows.Count > 0 Then
                PrefCon.SelectedPref.ID = Common.NZ(dt.Rows(0).Item("PreferenceID"), 0)
                Dim m_EditOfferRegardlessOfBuyer = (Logix.UserRoles.EditOffersRegardlessBuyer Or Common.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, PrefCon.Offer.ID))
                If Not PrefCon.Offer.IsTemplate Then
                    PrefCon.Offer.DisallowEdit = ((Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer) = False) _
                                        OrElse (PrefCon.Offer.FromTemplate AndAlso Common.NZ(dt.Rows(0).Item("DisallowEdit"), False))
                Else
                    PrefCon.Offer.DisallowEdit = (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer = False)
                End If

            End If
        End If
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Page()

        Common.AppName = "UEoffer-con-pref.aspx"
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()

        If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
            If Common.PMRTadoConn.State = ConnectionState.Closed Then Common.Open_PrefManRT()
            AdminUserID = Verify_AdminUser(Common, Logix)

            Long.TryParse(GetCgiValue("OfferID"), PrefCon.Offer.ID)
            Integer.TryParse(GetCgiValue("IncentivePrefsID"), PrefCon.IncentivePrefsID)

            Populate_Flags()
            Populate_Offer_Data()
            Populate_Preferences()
            Load_Saved_Prefence()
            bEnableAdditionalLockoutRestrictionsOnOffers = IIf(Common.Fetch_SystemOption(260) = "1", True, False)
            bOfferEditable = Common.IsOfferEditablePastLockOutPeriod(Logix.UserRoles.EditOfferPastLockoutPeriod, Common, PrefCon.Offer.ID)
            Send_Head()
            Send_Body()

            If Common.PMRTadoConn.State <> ConnectionState.Closed Then Common.Close_PrefManRT()
        Else
            Send_Not_Installed()
        End If

        If Common.LRTadoConn.State <> ConnectionState.Closed Then Common.Close_LogixRT()

    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Head()
        Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
        Dim CopientFileVersion As String = "7.3.1.138972"
        Dim CopientProject As String = "Copient Logix"
        Dim CopientNotes As String = ""
        Dim Handheld As Boolean = False

        If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
            Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
        End If

        'This would redirect to summary page if offer is awaiting deployment or if offer is awaiting recommendations.
        CheckIfValidOffer(Common, PrefCon.Offer.ID)

        Send_HeadBegin("term.offer", "term.preference", PrefCon.Offer.ID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts(New String() {"ajaxSubmit.js", "prefentry.js"})
        Send("<script type=""text/javascript"">")
        Send("function ChangeParentDocument() { ")
        Send("  if (opener != null) {")
        Send("    var newlocation = 'UEoffer-con.aspx?OfferID=" & PrefCon.Offer.ID & "'; ")
        Send("  if (opener.location.href.indexOf(newlocation) >-1) {")
        Send("  opener.location = 'UEoffer-con.aspx?OfferID=" & PrefCon.Offer.ID & "'; ")
        Send("  }")
        Send("  }")
        Send("} ")
        Send("</" & "script>")
        Send_Selector_JS()
        Send_Value_Load_JS()
        Send_HeadEnd()
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Body()
        Dim ShowPrefPage As Boolean = True

        If (PrefCon.Offer.IsTemplate) Then
            Send_BodyBegin(12)
        Else
            Send_BodyBegin(2)
        End If

        If (Logix.UserRoles.AccessOffers = False AndAlso Not PrefCon.Offer.IsTemplate) Then
            Send_Denied(2, "perm.offers-access")
            ShowPrefPage = False
        End If

        If (Logix.UserRoles.AccessTemplates = False AndAlso PrefCon.Offer.IsTemplate) Then
            Send_Denied(2, "perm.offers-access-templates")
            ShowPrefPage = False
        End If

        If (PrefCon.Flags.BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, PrefCon.Offer.ID)) Then
            Send("<script type=""text/javascript"" language=""javascript"">")
            Send("  function ChangeParentDocument() { return true; } ")
            Send("</" & "script>")
            Send_Denied(1, "banners.access-denied-offer")
            Send_BodyEnd()
        Else
            Send_Form()
            Send_Closing_JS()
            Send_BodyEnd()
        End If
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Form()
        Send("<form action=""#"" id=""mainform"" name=""mainform"" method=""post"">")
        Send_Intro()
        Send_Main()
        Send("</form>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Intro()
        Send("<div id=""intro"">")
        Send("  <input type=""hidden"" id=""OfferID"" name=""OfferID"" value=""" & PrefCon.Offer.ID & """ />")
        Send("  <input type=""hidden"" id=""Name"" name=""Name"" value=""" & PrefCon.Offer.Name & """ />")
        Send("  <input type=""hidden"" id=""roid"" name=""roid"" value=""" & PrefCon.Offer.ROID & """ />")
        Send("  <input type=""hidden"" id=""IncentivePrefsID"" name=""IncentivePrefsID"" value=""" & PrefCon.IncentivePrefsID & """ />")
        Send("  <input type=""hidden"" id=""IsTemplate"" name=""IsTemplate"" value=""" & IIf(PrefCon.Offer.IsTemplate, "IsTemplate", "Not") & """ />")
        Send("  <input type=""hidden"" id=""TierCount"" name=""TierCount"" value=""" & PrefCon.Offer.TierCount & """ />")

        If (PrefCon.Offer.IsTemplate) Then
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & PrefCon.Offer.ID & " " & StrConv(Copient.PhraseLib.Lookup("term.preferencecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
        Else
            Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & PrefCon.Offer.ID & " " & StrConv(Copient.PhraseLib.Lookup("term.preferencecondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
        End If

        Send_Controls()
        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Controls()
        Send("<div id=""controls"">")

        If (PrefCon.Offer.IsTemplate) Then
            Send("<span class=""temp"">")
            Send("<input type=""checkbox"" class=""tempcheck"" id=""Disallow_Edit"" name=""Disallow_Edit""" & IIf(PrefCon.Offer.DisallowEdit, " checked=""checked""", "") & " value=""1"" />")
            Send("<label for=""Disallow_Edit"">" & Copient.PhraseLib.Lookup("term.locked", LanguageID) & "</label>")
            Send("</span>")
        End If
        Dim m_EditOfferRegardlessOfBuyer = Logix.UserRoles.EditOffersRegardlessBuyer Or Common.IsOfferCreatedWithUserAssociatedBuyer(AdminUserID, PrefCon.Offer.ID)
        If Not (PrefCon.Offer.IsTemplate) Then
            If (Logix.UserRoles.EditOffer And m_EditOfferRegardlessOfBuyer And Not (PrefCon.Offer.FromTemplate And PrefCon.Offer.DisallowEdit) AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable) And Not IsOfferWaitingForApproval(PrefCon.Offer.ID)) Then
                'Send_Save(" onclick=""saveForm();""")
                Send("<input type=""button"" accesskey=""s"" class=""regular"" id=""save"" name=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ onclick=""saveForm();"" />")
            End If
        Else
            If (Logix.UserRoles.EditTemplates And m_EditOfferRegardlessOfBuyer AndAlso (Not bEnableAdditionalLockoutRestrictionsOnOffers OrElse bOfferEditable)) Then
                'Send_Save(" onclick=""saveForm();""")
                Send("<input type=""button"" accesskey=""s"" class=""regular"" id=""save"" name=""save"" value=""" & Copient.PhraseLib.Lookup("term.save", LanguageID) & """ onclick=""saveForm();"" />")
            End If
        End If

        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Main()
        Send("<div id=""main"">")
        Send("  <div id=""infobar""></div>")
        Send_Column1()
        Send_Gutter()
        Send_Column2()
        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Column1()
        Send("<div id=""column1"">")
        Send_Preference_Box()
        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Gutter()
        Send("<div id=""gutter"">")
        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Column2()
        Send("<div id=""column2"">")
        Send_Value_Box()
        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Preference_Box()
        Dim DisabledAttribute As String = IIf(PrefCon.Offer.DisallowEdit, " disabled=""disabled""", "")

        Send("<div class=""box"" id=""preference"">")
        Send("  <h2>")
        Send("    <span>" & Copient.PhraseLib.Lookup("term.preference", LanguageID) & "</span>")
        Send("  </h2>")
        Send("  <br class=""half"" />")
        Send("  <input type=""radio"" id=""functionradio1"" name=""functionradio"" " & IIf(Common.Fetch_SystemOption(175)= "1", "checked=""checked""", "") & " " & DisabledAttribute & " /><label for=""functionradio1"">" & Copient.PhraseLib.Lookup("term.startingwith", LanguageID) & "</label>")
        Send("  <input type=""radio"" id=""functionradio2"" name=""functionradio"" " & IIf(Common.Fetch_SystemOption(175)= "2", "checked=""checked""", "") & " " & DisabledAttribute & " /><label for=""functionradio2"">" & Copient.PhraseLib.Lookup("term.containing", LanguageID) & "</label><br />")
        Send("  <input type=""text"" class=""medium"" id=""functioninput"" name=""functioninput"" maxlength=""100"" onkeyup=""handleKeyUp(200);"" value=""""" & DisabledAttribute & " /><br />")
        Send("  <div id=""pgList"">")
        Send("    <select class=""longer"" id=""functionselect"" name=""functionselect"" size=""12""" & DisabledAttribute & " ondblclick=""handleSelectClick('select1');"">")
        For Each p As Preference In PrefCon.Preferences
            Send("     <option value=""" & p.ID & "|" & p.DataTypeID & "|" & IIf(p.MultiValue, 1, 0) & """>" & p.Name & "</option>")
            If (p.ID = PrefCon.SelectedPref.ID) Then
                PrefCon.SelectedPref.DataTypeID = p.DataTypeID
                PrefCon.SelectedPref.DataTypeName = Get_Data_Type_Name(p.DataTypeID)
                PrefCon.SelectedPref.Name = p.Name
                PrefCon.SelectedPref.MultiValue = p.MultiValue
            End If
        Next
        Send("    </select>")
        Send("  </div>")
        Send("  <br />")
        Send("  <br class=""half"" />")
        Send("  <b>" & Copient.PhraseLib.Lookup("term.selectedpreference", LanguageID) & ":</b><br />")
        Send("  <input type=""button"" class=""regular select"" id=""select1"" name=""select1"" value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ onclick=""handleSelectClick('select1');""" & DisabledAttribute & " />&nbsp;")
        Send("  <br class=""half"" />")
        Send("  <select class=""longer"" id=""selected"" name=""selected"" size=""2""" & DisabledAttribute & ">")
        If PrefCon.SelectedPref.ID > 0 Then
            Send("     <option value=""" & PrefCon.SelectedPref.ID & "|" & PrefCon.SelectedPref.DataTypeID & "|" & IIf(PrefCon.SelectedPref.MultiValue, 1, 0) & """>" & PrefCon.SelectedPref.Name & "</option>")
        Else
            PrefCon.SelectedPref.DataTypeName = Copient.PhraseLib.Lookup("term.none", LanguageID)
        End If
        Send("  </select>")
        Send("  <input type=""hidden"" id=""preferenceid"" name=""preferenceid"" value=""" & PrefCon.SelectedPref.ID & """ />")
        Send("  <br />")
        Send("  <br class=""half"" />")
        Send("  <div>" & Copient.PhraseLib.Lookup("term.datatype", LanguageID) & ":<span id=""dataType"" style=""padding-left:5px;"">" & PrefCon.SelectedPref.DataTypeName & "</span></div>")
        Send("  <div>" & Copient.PhraseLib.Lookup("term.multiple-values", LanguageID) & ":<span id=""multipleValues"" style=""padding-left:5px;"">" & Copient.PhraseLib.Lookup(IIf(PrefCon.SelectedPref.MultiValue, "term.yes", "term.no"), LanguageID) & "</span></div>")
        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Value_Box()
        Send("<div class=""box"" id=""prefvalues""" & IIf(PrefCon.IncentivePrefsID > 0, "", " style=""display:none;""") & ">")
        Send("  <h2>")
        Send("    <span>" & Copient.PhraseLib.Lookup("term.values", LanguageID) & "</span>")
        Send("  </h2>")
        Send("  <br class=""half"" />")
        Send("  <div id=""valueLoading"" style=""display:none;"">")
        Send("    <center>")
        Send("      <img src=""/images/loadingAnimation.gif"" />")
        Send("      <br />" & Copient.PhraseLib.Lookup("term.loading", LanguageID))
        Send("    </center>")
        Send("  </div>")
        Send("  <div id=""valueEntry""></div>")
        Send("  <br />")
        Send("</div>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Selector_JS()
        Send("<script type=""text/javascript"">")
        Send("  var fullselect = null;")
        Send("  var currentDataTypeID = 0;")

        ' load up all the preference into arrays for later searching.
        If PrefCon.Preferences Is Nothing OrElse PrefCon.Preferences.Count = 0 Then
            Send("  var functionlist = Array();")
            Send("  var vallist = Array();")
        Else
            Sendb("  var functionlist = Array(")
            For Each p As Preference In PrefCon.Preferences
                Sendb("""" & p.Name.Replace("""", "\""") & """,")
            Next
            Send(""""");")

            Sendb("  var vallist = Array(")
            For Each p As Preference In PrefCon.Preferences
                Sendb("""" & p.ID & "|" & p.DataTypeID & """,")
            Next
            Send(""""");")
        End If

        Send("")
        Send("  // This is the function that refreshes the list after a keypress.")
        Send("  // The maximum number to show can be limited to improve performance with")
        Send("  // huge lists (1000s of entries).")
        Send("  // The function clears the list, and then does a linear search through the")
        Send("  // globally defined array and adds the matches back to the list.")
        Send("  function handleKeyUp(maxNumToShow) {")
        Send("    var selectObj, textObj, functionListLength;")
        Send("    var i,  numShown;")
        Send("    var searchPattern;")
        Send("    ")
        Send("    document.getElementById(""functionselect"").size = ""12"";")
        Send("    ")
        Send("    // Set references to the form elements")
        Send("    selectObj = document.forms[0].functionselect;")
        Send("    textObj = document.forms[0].functioninput;")
        Send("    ")
        Send("    // Remember the function list length for loop speedup")
        Send("    functionListLength = functionlist.length;")
        Send("    ")
        Send("    // Set the search pattern depending")
        Send("    if(document.forms[0].functionradio[0].checked == true) {")
        Send("      searchPattern = ""^""+textObj.value;")
        Send("    } else {")
        Send("      searchPattern = textObj.value;")
        Send("    }")
        Send("    searchPattern = cleanRegExpString(searchPattern);")
        Send("    ")
        Send("    // Create a regular expression")
        Send("    re = new RegExp(searchPattern,""gi"");")
        Send("    ")
        Send("    // Loop through the array and re-add matching options")
        Send("    numShown = 0;")
        Send("    if (textObj.value == '' && fullselect != null) {")
        Send("      var newSelectBox = fullselect.cloneNode(true);")
        Send("      document.getElementById('pgList').replaceChild(newSelectBox, selectObj);")
        Send("    } else {")
        Send("      var newSelectBox = selectObj.cloneNode(false);")
        Send("      document.getElementById('pgList').replaceChild(newSelectBox, selectObj);")
        Send("      selectObj = document.getElementById(""functionselect"");")
        Send("      for(i = 0; i < functionListLength; i++) {")
        Send("        if(functionlist[i].search(re) != -1) {")
        Send("          if (vallist[i] != """") {")
        Send("            selectObj[numShown] = new Option(functionlist[i], vallist[i]);")
        Send("            if (vallist[i] == 1) {")
        Send("              selectObj[numShown].style.fontWeight = 'bold';")
        Send("              selectObj[numShown].style.color = 'brown';")
        Send("            }")
        Send("            numShown++;")
        Send("          }")
        Send("        }")
        Send("        // Stop when the number to show is reached")
        Send("        if(numShown == maxNumToShow) {")
        Send("          break;")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("    removeUsed(true);")
        Send("    // When options list whittled to one, select that entry")
        Send("    if(selectObj.length == 1) {")
        Send("      selectObj.options[0].selected = true;")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function removeUsed(bSkipKeyUp) {")
        Send("    if (!bSkipKeyUp) handleKeyUp(99999);")
        Send("    // this function will remove items from the functionselect box that are used in ")
        Send("    // selected box")
        Send("    var funcSel = document.getElementById('functionselect');")
        Send("    var elSel = document.getElementById('selected');")
        Send("    var i,j;")
        Send("    for (i = elSel.length - 1; i>=0; i--) {")
        Send("      for(j=funcSel.length-1;j>=0; j--) {")
        Send("        if(funcSel.options[j].value == elSel.options[i].value){")
        Send("          funcSel.options[j] = null;")
        Send("        }")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("  ")
        Send("  // this function gets the selected value and loads the appropriate")
        Send("  // php reference page in the display frame")
        Send("  // it can be modified to perform whatever action is needed, or nothing")
        Send("  function handleSelectClick(itemSelected) {")
        Send("    var isAnyProductSelected = false;")
        Send("    var item, parent;")
        Send("    var newDataTypeID;")
        Send("    ")
        Send("    textObj = document.forms[0].functioninput;")
        Send("    ")
        Send("    selectObj = document.forms[0].functionselect;")
        Send("    selectedValue = document.getElementById(""functionselect"").value;")
        Send("    if(selectedValue != """"){ selectedText = selectObj[document.getElementById(""functionselect"").selectedIndex].text; }")
        Send("    ")
        Send("    selectboxObj = document.forms[0].selected;")
        Send("    selectedboxValue = document.getElementById(""selected"").value;")
        Send("    if(selectedboxValue != """"){ selectedboxText = selectboxObj[document.getElementById(""selected"").selectedIndex].text; }")
        Send("    ")
        Send("    if(itemSelected == ""select1"") {")
        Send("      if(selectedValue != """") {")
        Send("        // remove any existing items")
        Send("        while (selectboxObj.options.length > 0) {")
        Send("          selectboxObj.options[0] = null;")
        Send("        }        ")
        Send("        // add items to selected box")
        Send("        selectedText = selectObj.options[selectObj.selectedIndex].text;")
        Send("        selectedValue = selectObj.options[selectObj.selectedIndex].value;")
        Send("        selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);")
        Send("        selectObj[selectObj.selectedIndex].selected = false;")
        Send("      }")
        Send("    }")
        Send("    ")
        Send("    // remove items from large list that are in the other lists")
        Send("    removeUsed(false);")
        Send("    updateButtons();")
        Send("    ")
        Send("    newDataTypeID = getSelectedValue(1);")
        Send("    if ((currentDataTypeID != newDataTypeID) || newDataTypeID == 1) {")
        Send("      loadValues();")
        Send("      currentDataTypeID = newDataTypeID;")
        Send("      writeDataTypeName(getDataTypeName(currentDataTypeID));")
        Send("      writeMultiValue(parseInt(getSelectedValue(2)));")
        Send("    }")
        Send("")
        Send("    return true;")
        Send("  }")
        Send("  ")
        Send("  function updateButtons() {")
        Send("    var elemSelect1 = document.getElementById('select1');")
        Send("    var elemSave = document.getElementById('save');")
        Send("    ")
        Send("    var selectboxObj = document.forms[0].selected;")
        Send("    var selectedboxValue = selectboxObj.value;")
        Send("    ")
        Send("    if (selectboxObj != null) {")
        Send("      if (elemSave != null)  { elemSave.disabled = (selectboxObj.length == 0) ? true : false;}")
        Send("    } else {")
        Send("      elemSelect1.disabled = false;")
        Send("    }")
        Send("  }")
        Send("  ")
        Send("  function writeDataTypeName(name) {")
        Send("    var elem = document.getElementById('dataType');")
        Send("    if (elem != null) {")
        Send("      elem.innerHTML=name;")
        Send("    }")
        Send("  }")
        Send("  function writeMultiValue(bitFlag) {")
        Send("    var elem = document.getElementById('multipleValues');")
        Send("    if (elem != null) {")
        Send("      if (bitFlag==1) {")
        Send("        elem.innerHTML='" & Copient.PhraseLib.Lookup("term.yes", LanguageID) & "';")
        Send("      } else {")
        Send("        elem.innerHTML='" & Copient.PhraseLib.Lookup("term.no", LanguageID) & "';")
        Send("      }")
        Send("    }")
        Send("  }")
        Send("</" & "script>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Value_Load_JS()
        Send("<script type=""text/javascript"">")
        Send("  function loadValues() {")
        Send("    toggleLoading(true);")
        Send("    sendValueRequest();")
        Send("  }")
        Send("")
        Send("  function toggleLoading(setOn) {")
        Send("    var elemBox = document.getElementById(""prefvalues"");")
        Send("    var elemImg = document.getElementById(""valueLoading"");")
        Send("    var elemEntry = document.getElementById(""valueEntry"");")
        Send("")
        Send("    if (elemBox != null) {")
        Send("      elemBox.style.display = 'block';")
        Send("    }")
        Send("    if (elemImg != null) {")
        Send("      elemImg.style.display = (setOn) ? 'block' : 'none';")
        Send("    }")
        Send("")
        Send("    if (setOn && elemEntry != null) {")
        Send("      elemEntry.innerHTML='';")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function getSelectedValue(index) {")
        Send("    var value = 0;")
        Send("    var elem = document.getElementById(""selected"");")
        Send("")
        Send("    if (elem != null) {")
        Send("      if (elem.options.length >= 1) {")
        Send("        value = elem.options[0].value.split('|')[index];")
        Send("      }")
        Send("    }")
        Send("    return value;")
        Send("  }")
        Send("")
        Send("  function sendValueRequest() {")
        Send("    var formdata = 'mode=loadPrefConValues&offerid=" & PrefCon.Offer.ID & "&incentiveprefsid=" & PrefCon.IncentivePrefsID & "&prefid=' + getSelectedValue(0) + '&edit=" & IIf(PrefCon.Offer.DisallowEdit, "-1", "1") & "';")
        Send("")
        Send("    xmlhttpPostData('../preference-feeds.aspx', formdata, 'valueEntry', 'handleValueCallback()');")
        Send("  }")
        Send("")
        Send("  function handleValueCallback() {")
        Send("    toggleLoading(false);")
        Send("  }")
        Send("")
        Send("  function getDataTypeName(id) {")
        Send("    var name = '';")
        Send("    switch (parseInt(id)) {")
        Send("      case 1:  ")
        Send("        name = '" & Copient.PhraseLib.Lookup("term.list", LanguageID) & "';")
        Send("        break;")
        Send("     case 2:")
        Send("        name = '" & Copient.PhraseLib.Lookup("term.numeric-range", LanguageID) & "';")
        Send("        break;")
        Send("      case 4:")
        Send("        name = '" & Copient.PhraseLib.Lookup("term.numeric", LanguageID) & "';")
        Send("        break;")
        Send("      case 5:")
        Send("        name = '" & Copient.PhraseLib.Lookup("term.boolean", LanguageID) & "';")
        Send("        break;")
        Send("      case 7:")
        Send("        name = '" & Copient.PhraseLib.Lookup("term.date", LanguageID) & "';")
        Send("        break;")
        Send("      case 8:")
        Send("        name = '" & Copient.PhraseLib.Lookup("term.likert", LanguageID) & "';")
        Send("        break;")
        Send("      default:")
        Send("        name = '" & Copient.PhraseLib.Lookup("term.unsupported", LanguageID) & "';")
        Send("     }")
        Send("     return name;")
        Send("  }")
        Send("")
        Send("  function saveForm() {")
        Send("    prepareForSave();")
        Send("    if (validateForm()) { ")
        Send("      xmlhttpPost('../preference-feeds.aspx?mode=savePrefCon', 'mainform', 'infobar', '" & Copient.PhraseLib.Lookup("message.Saving", LanguageID) & "');")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function validateForm() {")
        Send("    var tierlevels = parseInt(document.getElementById('TierCount').value);")
        Send("    var elemPref = document.getElementById('preferenceid');")
        Send("    var elemVals = null;")
        Send("    var msg = '';")
        Send("    var tokenValues = [];")
        Send("")
        Send("    if (elemPref == null || elemPref.value == '') { ")
        Send("      alert('" & Copient.PhraseLib.Lookup("ueoffer-con-pref.SelectPreference", LanguageID) & "');")
        Send("      return false;")
        Send("    }")
        Send("")
        Send("    for (var i=1; i<=tierlevels; i ++) {")
        Send("      elemVals = document.getElementById('allvalues_tier' + i)")
        Send("      if (elemVals == null || elemVals.value == '') { ")
        Send("        msg = '" & Copient.PhraseLib.Lookup("ueoffer-con-pref.SupplyTierValue", LanguageID) & "';")
        Send("        tokenValues[0] = i;")
        Send("        msg = detokenizeString(msg, tokenValues);")
        Send("        alert(msg);")
        Send("        return false;")
        Send("      }")
        Send("    }")
        Send("")
        Send("    return true;")
        Send("  }")
        Send("")
        Send("  function prepareForSave() {")
        Send("    var tierlevels = parseInt(document.getElementById('TierCount').value);")
        Send("    var elemSlct = document.getElementById('selected');")
        Send("    var elemPref = document.getElementById('preferenceid');")
        Send(" ")
        Send("    for (i=1; i<=tierlevels; i++) { ")
        Send("      writeTierValues(i)")
        Send("    }")
        Send("")
        Send("    if (elemSlct != null && elemPref!= null && elemSlct.options.length > 0) { ")
        Send("      elemPref.value = elemSlct.options[0].value")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function writeTierValues(tierlevel) {")
        Send("    var elemSlct = document.getElementById('values_tier' + tierlevel);")
        Send("    var elemVals = document.getElementById('allvalues_tier' + tierlevel)")
        Send("    var vals = '';")
        Send("    ")
        Send("    if (elemSlct != null && elemVals != null) { ")
        Send("      for (var i=0; i < elemSlct.options.length; i++) { ")
        Send("        if (i > 0) { vals += ',';}")
        Send("        vals += elemSlct.options[i].value;")
        Send("      }")
        Send("      elemVals.value = vals;")
        Send("    }")
        Send("  }")
        Send("")
        Send("  function processAjaxResponseData(responseData) {")
        Send("     if (responseData.indexOf('[DATA]:Saved=1') > -1)  {")
        Send("       opener.location = 'UEoffer-con.aspx?OfferID=" & PrefCon.Offer.ID & "'; ")
        If PrefCon.Flags.CloseAfterSave Then
            Send("       if (responseData.indexOf('[DATA]:Saved=1') > -1)  {")
            Send("         window.close();")
            Send("       }")
        End If
        Send("     }")
        Send("  }")
        Send("")
        Send("</" & "script>")

    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Closing_JS()
        Send("<script type=""text/javascript"">")
        Send("  if (document.getElementById(""functionselect"") != null) {")
        Send("    fullselect = document.getElementById(""functionselect"").cloneNode(true);")
        Send("  }")
        Send("  removeUsed(true);")
        Send("  updateButtons();")
        If PrefCon.IncentivePrefsID > 0 Then
            Send("  loadValues();")
        End If
        Send("</" & "script>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Sub Send_Not_Installed()
        Send("<html><head></head><body><p>" & Copient.PhraseLib.Lookup("ueoffer-con-pref.NotInstalled", LanguageID) & "</p></body></html>")
    End Sub

    '------------------------------------------------------------------------------------------------------------------

    Function Get_Data_Type_Name(ByVal DataTypeID As Integer) As String
        Dim Name As String = ""
        Dim dt As DataTable

        Common.QueryStr = "select " & _
                          "  case when PT.Phrase is null then PDT.Name " & _
                          "  else CONVERT(nvarchar(200), PT.Phrase) end as PhraseText " & _
                          "from PreferenceDataTypes as PDT with (NoLock) " & _
                          "left join PM_PhraseText as PT with (NoLock)on PT.PhraseID = PDT.NamePhraseID " & _
                          "  and PT.LanguageID=" & LanguageID & " " & _
                          "where PDT.DataTypeID=" & DataTypeID
        dt = Common.PMRT_Select
        If dt.Rows.Count > 0 Then
            Name = Common.NZ(dt.Rows(0).Item("PhraseText"), Copient.PhraseLib.Lookup("term.none", LanguageID))
        End If

        Return Name
    End Function

    '------------------------------------------------------------------------------------------------------------------

    Sub Save_Pref_Con()
        Dim IncentivePrefsID As Integer

        Integer.TryParse(GetCgiValue("IncentivePrefsID"), IncentivePrefsID)

        ' remove any existing incentive preferences
        If IncentivePrefsID > 0 Then
            Common.QueryStr = "dbo.pt_CPE_IncentivePrefs_Delete"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@IncentivePrefsID", SqlDbType.Int).Value = 0
            Common.LRTsp.ExecuteNonQuery()
        End If

    End Sub

</script>
