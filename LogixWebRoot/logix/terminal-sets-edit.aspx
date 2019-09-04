<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: terminal-sets-edit.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2011.  All rights reserved by:
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
  ' * Version : 5.12b1.0 
  ' *
  ' *****************************************************************************
%>
<script runat="server">

  
  Dim Common As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim Handheld As Boolean = False
  Dim TerminalSetID As Integer = 0
  Dim TerminalSet As New TerminalSetStruct
  Dim InStoreLocations As New List(Of SelectionItem)
  Dim PrinterTypes As New List(Of SelectionItem)
  Dim OpDisplayTypes As New List(Of SelectionItem)
  Dim ErrorMessage As String = ""
  Dim InfoMessage As String = ""
  Dim BrowserType As String = ""
  
  Structure TerminalSetStruct
    Public TerminalSetID As Integer
    Public Name As String
    Public TerminalSetTypeID As Integer
    Public PromoEngineID As Integer
    Public LastUpdate As Date
    Public TerminalSetItems As List(Of TerminalSetItem)
  End Structure
  
  Structure TerminalSetItem
    Public PKID As Integer
    Public TerminalSetID As Integer
    Public TerminalID As Integer
    Public TerminalTypeID As Integer
    Public PrinterTypeID As Integer
    Public OpDisplayTypeID As Integer
    Public Deleted As Boolean
  End Structure
    
  Structure SelectionItem
    Public Name As String
    Public Value As String
    Public Selected As Boolean
    Public IsDefault As Boolean
  End Structure

  Enum SetTypes As Integer
    STANDARD_SET = 1
    DEFAULT_SET = 2
  End Enum

  
  '-------------------------------------------------------------------------------------------------------------  
  
  Sub Send_Controls()
    Send("    <div id=""controls"">")
    Sendb("      ")
    Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
    Send("<div class=""actionsmenu"" id=""actionsmenu"">")
    Send_Save()
    Send_Delete()
    Send_New()
    If Request.Browser.Type = "IE6" Then
      Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:145px;""></iframe>")
    End If
    Send("</div>")
        If (Common.Fetch_SystemOption(75)) Then
            If (Logix.UserRoles.AccessNotes) Then
                Send_NotesButton(26, TerminalSetID, AdminUserID)
            End If
        End If
    Send("    </div>")
  End Sub
  

  '-------------------------------------------------------------------------------------------------------------  
  
  Sub Send_Intro()
    Dim Title As String = ""
    Dim FullName As String = ""
    Dim dt As DataTable
    
    Send("  <div id=""intro"">")
    Send("    <input type=""hidden"" id=""TerminalSetID"" name=""TerminalSetID"" value=""" & TerminalSetID & """ />")

    If TerminalSetID = 0 Then
      Title = Copient.PhraseLib.Lookup("term.newterminalset", LanguageID)
    Else
      Common.QueryStr = "SELECT Name FROM TerminalSets with (NoLock) WHERE TerminalSetID = " & TerminalSetID & ";"
      dt = Common.LRT_Select
      If (dt.Rows.Count > 0) Then
        FullName = Common.NZ(dt.Rows(0).Item("Name"), "")
        Title = FullName
        If (Len(FullName) > 30) Then
          Title = Left(FullName, 27) & "..."
        End If
      End If
      Title = Copient.PhraseLib.Lookup("term.terminal-set", LanguageID) & " #" & TerminalSetID & ": " & Title
    End If

    Send("<h1 id=""title"" title=""" & FullName.Replace("""", "	&quot;") & """>")
    Send(Title)
    Send("</h1>")

    Send_Controls()
    Send("  </div>")
  End Sub
  

    
  
  '-------------------------------------------------------------------------------------------------------------  

  
  Sub Send_Identification()
    Dim Name As String = ""
    Dim EditedDisplay As String = ""
    Dim DefaultChecked As Boolean = False
    Dim CurrentDefaultSetID As Integer = 0
    Dim BoxHeight As Integer = 200
    
    Name = TerminalSet.Name
    If TerminalSet.Name Is Nothing Then Name = ""
    
    EditedDisplay = Logix.ToLongDateTimeString(TerminalSet.LastUpdate, Common)
    DefaultChecked = (TerminalSet.TerminalSetTypeID = 2)
    If BrowserType = "IE" Then BoxHeight = 204
    
    Send("      <div class=""box"" id=""identification"" style=""height:" & BoxHeight & "px;"">")
    Send("        <h2>")
    Send("          <span>" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & "</span>")
    Send("        </h2>")
    Send("        <label for=""SetName"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
    Send("        <input type=""text"" class=""longest"" id=""SetName"" name=""SetName"" maxlength=""255"" value=""" & Name.Replace("""", "&quot;") & """ /><br />")
    Send("        <input type=""hidden"" id=""PromoEngineID"" name=""PromoEngineID"" value=""9"" />")
    Send("        <br class=""half"" />")
    Send("        <input type=""checkbox"" id=""defaultSet"" name=""defaultSet"" value=""1""" & IIf(DefaultChecked, " checked=""checked""", "") & " />")
    Send("        <label for=""defaultSet"">" & Copient.PhraseLib.Lookup("terminal-sets.use-as-default", LanguageID) & "</label>")

    If Not DefaultChecked Then
      CurrentDefaultSetID = Find_Default_TerminalSetID()
      If CurrentDefaultSetID > 0 Then
        Send("<font style=""color:gray;""><i>&nbsp;&nbsp;(" & Copient.PhraseLib.Lookup("terminal-sets.current-default", LanguageID) & " " & Find_Default_TerminalSetID() & ")</i></font>")
      Else
        Send("<font style=""color:gray;""><i>&nbsp;&nbsp;(" & Copient.PhraseLib.Lookup("terminal-sets.no-default", LanguageID) & ")</i></font>")
      End If
    End If
    Send("<br />")
    
    If TerminalSet.LastUpdate <> Nothing Then
      Send("        <br class=""half"" />")
      Send(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & EditedDisplay)
    End If
    Send("        <br />")
    Send("      </div>")
    
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------  

  
  Sub Send_AssociatedStores()
    Dim dtStores As DataTable
    Dim row As DataRow
    
    dtStores = Retrieve_Associated_Stores(TerminalSet.TerminalSetID)
    
    Send("      <div class=""box"" id=""stores"">")
    Send("        <h2>")
    Send("          <span>" & Copient.PhraseLib.Lookup("term.associatedstores", LanguageID) & "</span>")
    Send("        </h2>")

    If Not dtStores Is Nothing AndAlso dtStores.Rows.Count > 0 Then
      Send("      <div class=""boxscroll"">")
      For Each row In dtStores.Rows
        Send("     <a href=""store-detail.aspx?LocationID=" & Common.NZ(row.Item("LocationID"), 0) & """>" & Common.NZ(row.Item("LocationName"), "") & "</a><br />")
      Next
      Send("      </div>")
    Else
      Send("      <div class=""boxscroll"">")
      Sendb(Copient.PhraseLib.Lookup("terminal-sets.nostores", LanguageID) & "<br />")
      Send("      </div>")
    End If

    Send("        <hr class=""hidden"" />")
    Send("     </div>")

  End Sub

  
  '-------------------------------------------------------------------------------------------------------------  


  Sub Send_Lanes_Table()

    Send("      <div class=""box"" id=""lanes"">")
    Send("        <h2>")
    Send("          <span>" & Copient.PhraseLib.Lookup("term.lanes", LanguageID) & "</span>")
    Send("        </h2>")
    Send("        <input type=""button"" name=""AddLane"" id=""AddLane"" class=""mediumshort"" style=""width:auto"" value=""" & Copient.PhraseLib.Lookup("term.addnewlane", LanguageID) & """ onclick=""javascript:addLane();"" />")
    Send("        <br /><br class=""half"" />")

    Send("        <table id=""tblLanes"" summary=""" & Copient.PhraseLib.Lookup("term.lanetypes", LanguageID) & """>")
    Send("          <thead>")
    Send("            <tr>")
    Send("            <th>" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & "</th>")
    Send("            <th>" & Copient.PhraseLib.Lookup("term.terminal", LanguageID) & "#</th>")
    Send("            <th>" & Copient.PhraseLib.Lookup("term.lanetype", LanguageID) & "</th>")
    Send("            <th>" & Copient.PhraseLib.Lookup("term.printertype", LanguageID) & "</th>")
    Send("            <th>" & Copient.PhraseLib.Lookup("term.displaytype", LanguageID) & "</th>")
    Send("            </tr>")
    Send("          </thead>")
    Send("          <tbody>")

    ' create a sample row used only to create a new row in the table when the "Add a New Lane" button is clicked.
    Send_Lane(New TerminalSetItem)
    
    ' existing lane types for the store
    For Each tsi As TerminalSetItem In TerminalSet.TerminalSetItems
      Send_Lane(tsi)
    Next
    
    Send("          </tbody>")
    Send("        </table>")
    Send("      </div>")
    
  End Sub
    
    
  '-------------------------------------------------------------------------------------------------------------  

  
  Sub Send_Lane(ByVal tsi As TerminalSetItem)
    Dim Selected As Boolean = False
    

    Send("            <tr " & IIf(tsi.PKID = 0, "style=""display:none;""", "") & ">")
    Send("              <td>")
    Send("                <input type=""hidden"" name=""PKID"" value=""" & tsi.PKID & """ />")
    Send("                <input type=""hidden"" name=""markedAsDeleted"" value=""" & IIf(tsi.PKID = 0, "1", "0") & """ />")
    Send("                <input type=""button"" class=""ex"" name=""laneDelete"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ onClick=""if(confirm('" & Copient.PhraseLib.Lookup("terminals-set.confirm-lane-delete", LanguageID) & "')){deleteLane(this)}"" value=""X"" />")
    Send("              </td>")
    Send("              <td><input type=""text"" class=""shorter"" name=""terminalID"" value=""" & IIf(tsi.PKID = 0, "", tsi.TerminalID) & """ /></td>")
    Send("              <td>")
    Send("                <select class=""medium"" name=""terminalType"">")
    For Each si As SelectionItem In InStoreLocations
      Selected = (tsi.TerminalTypeID = si.Value)
      Send("                  <option value=""" & si.Value & """" & IIf(Selected, " selected=""selected""", "") & ">" & si.Name & "</option>")
    Next
    Send("                </select>")
    Send("              </td>")
    Send("              <td>")
    Send("                <select class=""medium"" name=""printerType"">")
    For Each si As SelectionItem In PrinterTypes
      Selected = (tsi.PrinterTypeID = si.Value) OrElse (tsi.PKID = 0 AndAlso si.IsDefault)
      Send("                  <option value=""" & si.Value & """" & IIf(Selected, " selected=""selected""", "") & ">" & si.Name & "</option>")
    Next
    Send("                </select>")
    Send("              </td>")
    Send("              <td>")
    Send("                <select class=""medium"" name=""displayType"">")
    For Each si As SelectionItem In OpDisplayTypes
      Selected = (tsi.OpDisplayTypeID = si.Value)
      Send("                  <option value=""" & si.Value & """" & IIf(Selected, " selected=""selected""", "") & ">" & si.Name & "</option>")
    Next
    Send("                </select>")
    Send("              </td>")
    Send("            </tr>")

  End Sub
    
  
  '-------------------------------------------------------------------------------------------------------------  

  
  Sub Send_Main()
    Dim IsNew As Boolean = False
    
    Send("  <div id=""main"">")
    
    If (ErrorMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & ErrorMessage & "</div>")
    ElseIf (InfoMessage <> "") Then
      Send("<div id=""infobar"" class=""green-background"">" & InfoMessage & "</div>")
    End If
    
    IsNew = TerminalSetID = 0
    
    If TerminalSet.TerminalSetID > 0 OrElse IsNew Then
      'Send("    <div id=""column"">")

      Send("    <div id=""column1"">")
      Send_Identification()
      Send("    </div>")
      
      Send("    <div id=""gutter""></div>")
      
      If TerminalSet.TerminalSetID > 0 Then
        Send("    <div id=""column2"">")
        Send_AssociatedStores()
        Send("    </div>")
      End If

      Send("    <div id=""column"">")
      Send_Lanes_Table()
      Send("    </div>")
      Send("  </div>")
    End If
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------  


  Sub Send_Form()
    Send("<form action=""terminal-sets-edit.aspx"" id=""mainform"" name=""mainform"" method=""POST"">")
    If Not Logix.UserRoles.EditTerminalSets Then
      Send_Denied(1, "perm.189")
    Else
      Send_Intro()
      Send_Main()
    End If
    Send("</form>")
  End Sub

  
  
  '-------------------------------------------------------------------------------------------------------------  
  
  Sub Send_Page_JS()
    Dim NextTerminalID As Integer = 1
    
    If TerminalSet.TerminalSetItems.Count > 0 Then
      NextTerminalID = TerminalSet.TerminalSetItems.Item(TerminalSet.TerminalSetItems.Count - 1).TerminalID + 1
    End If
    
    Send("<script type=""text/javascript"">")
    Send("  var nextTerminalID = " & NextTerminalID & ";")
    Send("")
    Send("  function addLane() {")
    Send("    var elemTable = document.getElementById('tblLanes');")
    Send("    var elemNewRow = null;")
    Send("    var elemBody = null;")
    Send("")
    Send("    if (elemTable != null) {")
    Send("      elemNewRow = elemTable.getElementsByTagName('tr')[1].cloneNode(true);")
    Send("      elemNewRow.style.display = '';")
    Send("      elemTable.tBodies[0].appendChild(elemNewRow);")
    Send("      markRowDeletedValue(elemNewRow, '0');")
    Send("    }")
    Send("    assignNextBoxID();")
    Send("  }")
    Send("")
    Send("  function deleteLane(elem) {")
    Send("    var elems = null;     ")
    Send("    var elemTr, elemTd;")
    Send("")
    Send("    if (elem!=null && elem.parentNode!=null) {")
    Send("      markRowDeletedValue(elem.parentNode, '1');")
    Send("      elemTr = elem.parentNode.parentNode;")
    Send("      if (elemTr!=null) {")
    Send("        elemTr.style.display='none';")
    Send("      }")
    Send("    }")
    Send("  }")
    Send("")
    Send("  function markRowDeletedValue(elem, newValue) {")
    Send("    var elems = null;")
    Send("")
    Send("    if (elem!=null) { ")
    Send("      elems = elem.getElementsByTagName('input');")
    Send("      if (elems!=null) {")
    Send("        for (var i=0; i <elems.length; i++) {")
    Send("          if (elems[i].name == 'markedAsDeleted') { elems[i].value = newValue; }")
    Send("        }")
    Send("      }")
    Send("    }")
    Send("  }")
    Send("")
    Send("  function assignNextBoxID() {")
    Send("    var elems = document.getElementsByName('terminalID');")
    Send("")
    Send("    if (elems!=null && elems.length > 0) {")
    Send("      elems[elems.length-1].value = nextTerminalID;")
    Send("      nextTerminalID++;")
    Send("    }")
    Send("  }")
    Send("")
    Send("  function toggleDropdown() {")
    Send("    if (document.getElementById(""actionsmenu"") != null) {")
    Send("      var bOpen = (document.getElementById(""actionsmenu"").style.visibility != 'visible')")
    Send("      if (bOpen) {")
    Send("        document.getElementById(""actionsmenu"").style.visibility = 'visible';")
    Send("        document.mainform.actions.value = '" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & "▲';")
    Send("      } else {")
    Send("        document.getElementById(""actionsmenu"").style.visibility = 'hidden';")
    Send("        document.mainform.actions.value = '" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & "▼';")
    Send("      }")
    Send("    }")
    Send("  }")
    Send("")
    Send("<" & "/script>")
    
  End Sub

  '-------------------------------------------------------------------------------------------------------------  
  
  
  Sub Send_Page()
    
    Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Logix"
    Dim CopientNotes As String = ""
    
    Load_Page_Data()
    
    Send_HeadBegin("term.terminalsets", "term.edit", TerminalSetID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_Page_JS()
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 7)
    Send_Subtabs(Logix, 74, 4, , TerminalSetID)

    If Not Logix.UserRoles.EditTerminalSets Then
      Send_Denied(1, "perm.189")
    Else
      Send_Form()
    End If
        If (Common.Fetch_SystemOption(75)) Then
            If (TerminalSetID > 0 And Logix.UserRoles.AccessNotes) Then
                Send_Notes(26, TerminalSetID, AdminUserID)
            End If
        End If
    Send_BodyEnd()

  End Sub
  
  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Load_TerminalSet()
    Dim dt As DataTable
    Dim tsi As TerminalSetItem
    
    TerminalSet = New TerminalSetStruct
    TerminalSet.TerminalSetItems = New List(Of TerminalSetItem)
    
    Common.QueryStr = "select TS.Name, TS.TerminalSetTypeID, TS.PromoEngineID, TS.LastUpdate " & _
                      "from TerminalSets as TS with (NoLock) " & _
                      "where TerminalSetID=" & TerminalSetID & ";"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      With TerminalSet
        .TerminalSetID = TerminalSetID
        .Name = Common.NZ(dt.Rows(0).Item("Name"), "")
        .TerminalSetTypeID = Common.NZ(dt.Rows(0).Item("TerminalSetTypeID"), 1)
        .PromoEngineID = Common.NZ(dt.Rows(0).Item("PromoEngineID"), 1)
        .LastUpdate = Common.NZ(dt.Rows(0).Item("LastUpdate"), "1900-01-01")
      End With
      
      Common.QueryStr = "select TSI.PKID, TSI.TerminalSetID, TSI.TerminalID, TSI.TerminalTypeID, TSI.PrinterTypeID, TSI.OpDisplayTypeID " & _
                        "from TerminalSetItems as TSI with (NoLock) " & _
                        "where TSI.TerminalSetID=" & TerminalSetID & " " & _
                        "order by TerminalID;"
    
      dt = Common.LRT_Select
      For Each row As DataRow In dt.Rows
        tsi = New TerminalSetItem
        With tsi
          .PKID = Common.NZ(row.Item("PKID"), 0)
          .TerminalSetID = Common.NZ(row.Item("TerminalSetID"), 0)
          .TerminalID = Common.NZ(row.Item("TerminalID"), 0)
          .TerminalTypeID = Common.NZ(row.Item("TerminalTypeID"), 0)
          .PrinterTypeID = Common.NZ(row.Item("PrinterTypeID"), 0)
          .OpDisplayTypeID = Common.NZ(row.Item("OpDisplayTypeID"), 0)
        End With
        TerminalSet.TerminalSetItems.Add(tsi)
      Next
      
    ElseIf TerminalSetID > 0 Then
      ErrorMessage = Copient.PhraseLib.Lookup("terminals-sets.not-found", LanguageID) & _
                     "(" & Copient.PhraseLib.Lookup("term.id", LanguageID) & " = " & TerminalSetID & ")"
    End If

  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Load_InStoreLocations()
    Dim dt As DataTable
    Dim SelOpt As New SelectionItem
    
    Common.QueryStr = "select TerminalTypeID, Name, PhraseID from TerminalTypes with (NoLock) " & _
                      "where EngineID=9 and AnyTerminal=0 and Deleted=0;"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
    For Each row As DataRow In dt.Rows
      SelOpt = New SelectionItem
      With SelOpt
        .Value = Common.NZ(row.Item("TerminalTypeID"), 0)
        .Name = Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseID"), 0), LanguageID, Common.NZ(row.Item("Name"), "")) & " / " & .Value
      End With
      InStoreLocations.Add(SelOpt)
    Next
    Else
      SelOpt = New SelectionItem
      With SelOpt
        .Value = 0
        .Name = Copient.PhraseLib.Lookup("term.none", LanguageID)
      End With
      InStoreLocations.Add(SelOpt)
    End If

  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Load_PrinterTypes()
    Dim dt As DataTable
    Dim SelOpt As New SelectionItem
    
    Common.QueryStr = "select PT.PrinterTypeID, Name, PhraseID, DefaultPrinter from PrinterTypes as PT with (NoLock)  " & _
                      "inner join PromoEnginePrinterTypes as PEPT with (NoLock) on PEPT.PrinterTypeID = PT.PrinterTypeID " & _
                      "where PEPT.EngineID=9 and PT.Installed=1;"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
    For Each row As DataRow In dt.Rows
      SelOpt = New SelectionItem
      With SelOpt
        .Value = Common.NZ(row.Item("PrinterTypeID"), 0)
        .Name = Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseID"), 0), LanguageID, Common.NZ(row.Item("Name"), "")) & " / " & .Value
        .IsDefault = Common.NZ(row.Item("DefaultPrinter"), False)
      End With
      PrinterTypes.Add(SelOpt)
    Next
    Else
      SelOpt = New SelectionItem
      With SelOpt
        .Value = 0
        .Name = Copient.PhraseLib.Lookup("term.none", LanguageID)
        .IsDefault = False
      End With
      PrinterTypes.Add(SelOpt)
    End If
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Load_DisplayTypes()
    Dim dt As DataTable
    Dim SelOpt As New SelectionItem
    
    Common.QueryStr = "select OpDisplayTypeID, Name, PhraseID from CPE_OpDisplayTypes with (NoLock)"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
    For Each row As DataRow In dt.Rows
      SelOpt = New SelectionItem
      With SelOpt
        .Value = Common.NZ(row.Item("OpDisplayTypeID"), 0)
        .Name = Copient.PhraseLib.Lookup(Common.NZ(row.Item("PhraseID"), 0), LanguageID, Common.NZ(row.Item("Name"), "")) & " / " & .Value
      End With
      OpDisplayTypes.Add(SelOpt)
    Next
    Else
      SelOpt = New SelectionItem
      With SelOpt
        .Value = 0
        .Name = Copient.PhraseLib.Lookup("term.none", LanguageID)
      End With
      OpDisplayTypes.Add(SelOpt)
    End If

  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Load_Page_Data()
    Load_InStoreLocations()
    Load_PrinterTypes()
    Load_DisplayTypes()

    ' only load these if the page hasn't already created them via a form submission
    If ErrorMessage = "" Then
      Load_TerminalSet()
    End If
    
  End Sub
    

  '-------------------------------------------------------------------------------------------------------------    
  
  
  Sub Build_Terminal_Set()
    TerminalSet = New TerminalSetStruct
    
    With TerminalSet
      .Name = GetCgiValue("SetName")
      .PromoEngineID = Common.Extract_Val(GetCgiValue("PromoEngineID"))
      .TerminalSetID = Common.Extract_Val(GetCgiValue("TerminalSetID"))
      .TerminalSetTypeID = IIf(GetCgiValue("DefaultSet") = "1", SetTypes.DEFAULT_SET, SetTypes.STANDARD_SET)
    End With
    Build_Lanes()
  End Sub
      
  
  '-------------------------------------------------------------------------------------------------------------    
  
    
  Sub Build_Lanes()
    Dim PKIDs(), MarkedAsDeleted(), TerminalIDs() As String
    Dim TerminalTypes(), PrtTypes(), DisplayTypes() As String
    Dim i As Integer
    Dim yb As TerminalSetItem
    
    
    TerminalSet.TerminalSetItems = New List(Of TerminalSetItem)
    
    ' load up all the form values
    PKIDs = Request.Form.GetValues("PKID")
    MarkedAsDeleted = Request.Form.GetValues("markedAsDeleted")
    TerminalIDs = Request.Form.GetValues("terminalID")
    TerminalTypes = Request.Form.GetValues("terminalType")
    PrtTypes = Request.Form.GetValues("printerType")
    DisplayTypes = Request.Form.GetValues("displayType")
    
    If PKIDs IsNot Nothing AndAlso MarkedAsDeleted IsNot Nothing AndAlso TerminalIDs IsNot Nothing AndAlso TerminalTypes IsNot Nothing _
    AndAlso PrtTypes IsNot Nothing AndAlso DisplayTypes IsNot Nothing AndAlso (PKIDs.Length = MarkedAsDeleted.Length) AndAlso (PKIDs.Length = TerminalIDs.Length) _
    AndAlso (PKIDs.Length = TerminalTypes.Length) AndAlso (PKIDs.Length = PrtTypes.Length) AndAlso (PKIDs.Length = DisplayTypes.Length) Then
      For i = 0 To PKIDs.GetUpperBound(0)
        yb = New TerminalSetItem
        With yb
          .TerminalSetID = TerminalSetID
          .PKID = Common.Extract_Val(PKIDs(i))
          .TerminalID = Common.Extract_Decimal(TerminalIDs(i), Common.GetAdminUser.Culture)
          .TerminalTypeID = Common.Extract_Val(TerminalTypes(i))
          .PrinterTypeID = Common.Extract_Val(PrtTypes(i))
          .OpDisplayTypeID = Common.Extract_Val(DisplayTypes(i))
          .Deleted = (Common.Extract_Val(MarkedAsDeleted(i)) = 1)
        End With
        TerminalSet.TerminalSetItems.Add(yb)
      Next
    Else
      ErrorMessage = Copient.PhraseLib.Lookup("terminal-set.invalid-form", LanguageID)
    End If
    
  End Sub

  
  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Handle_New_Click()
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "terminal-sets-edit.aspx")
    Response.End()
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------    
  
  
  Function Is_Valid_Entry() As Boolean
    Dim Valid As Boolean = True
    Dim i, j As Integer

    ' validate the terminal set name
    If TerminalSet.Name Is Nothing OrElse TerminalSet.Name.Trim = "" Then
      Valid = False
      ErrorMessage = Copient.PhraseLib.Lookup("terminal-set.noname", LanguageID)
    End If
    
    ' validate each terminal set item
    If Valid Then
    For i = 0 To TerminalSet.TerminalSetItems.Count - 1
      If Not TerminalSet.TerminalSetItems(i).Deleted Then
        If TerminalSet.TerminalSetItems(i).TerminalID = 0 Then
          Valid = False
          ErrorMessage = Copient.PhraseLib.Lookup("terminal-set.invalid-terminalid", LanguageID) & i
        End If
      
        For j = (i + 1) To TerminalSet.TerminalSetItems.Count - 1
          If Not TerminalSet.TerminalSetItems(j).Deleted Then
            If TerminalSet.TerminalSetItems(i).TerminalID = TerminalSet.TerminalSetItems(j).TerminalID Then
              Valid = False
              ErrorMessage = Copient.PhraseLib.Lookup("terminal-set.duplicate-terminalid", LanguageID) & j
            ElseIf TerminalSet.TerminalSetItems(j).TerminalID = 0 Then
              Valid = False
              ErrorMessage = Copient.PhraseLib.Lookup("terminal-set.invalid-terminalid", LanguageID) & j
            End If
          End If
        Next
      
        If Not Valid Then Exit For
      End If
    Next
    End If
    
    Return Valid
  End Function
  
  '-------------------------------------------------------------------------------------------------------------    
  
    
  Sub Add_Terminal_Set(ByVal TermSet As TerminalSetStruct)
    Dim RetCode As Integer
    
    Try
      Common.QueryStr = "BEGIN TRAN"
      Common.LRT_Execute()
    
      ' first create the terminal set record
      Common.QueryStr = "dbo.pt_TerminalSets_Insert"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 255).Value = TermSet.Name
      Common.LRTsp.Parameters.Add("@TerminalSetTypeID", SqlDbType.Int).Value = TermSet.TerminalSetTypeID
      Common.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = TermSet.PromoEngineID
      Common.LRTsp.Parameters.Add("@RetCode", SqlDbType.Int).Direction = ParameterDirection.Output
      Common.LRTsp.Parameters.Add("@TerminalSetID", SqlDbType.Int).Direction = ParameterDirection.Output
      Common.LRTsp.ExecuteNonQuery()
      RetCode = Common.NZ(Common.LRTsp.Parameters("@RetCode").Value, -99)
      TerminalSetID = Common.NZ(Common.LRTsp.Parameters("@TerminalSetID").Value, 0)
      TermSet.TerminalSetID = TerminalSetID
      Common.Close_LRTsp()

      Select Case RetCode
        Case 2
          InfoMessage = Copient.PhraseLib.Lookup("terminal-set.saved-default-changed", LanguageID)
        Case 1
          InfoMessage = Copient.PhraseLib.Lookup("terminal-set.saved", LanguageID)
        Case -1
          ErrorMessage = Copient.PhraseLib.Lookup("terminal-set.duplicate-name", LanguageID)
        Case Else
          ErrorMessage = Copient.PhraseLib.Lookup("terminal-sets.save-failed", LanguageID)
      End Select
    
      ' then create the terminal set items records
      For Each tsi As TerminalSetItem In TermSet.TerminalSetItems
        If Not tsi.Deleted Then
          Common.QueryStr = "insert into TerminalSetItems (TerminalID, TerminalSetID, TerminalTypeID, PrinterTypeID, OpDisplayTypeID) " & _
                            "values (" & tsi.TerminalID & ", " & TerminalSetID & ", " & tsi.TerminalTypeID & "," & tsi.PrinterTypeID & ", " & _
                            "        " & tsi.OpDisplayTypeID & ")"
          Common.LRT_Execute()
        End If
      Next
      
      If TermSet.TerminalSetTypeID = SetTypes.DEFAULT_SET Then
        Update_Unassigned_Locations(TermSet)
      End If
      
      Common.Activity_Log(48, TerminalSetID, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-set-create", LanguageID))

      Common.QueryStr = "COMMIT TRAN"
      Common.LRT_Execute()
    Catch ex As Exception
      Common.QueryStr = "ROLLBACK TRAN"
      Common.LRT_Execute()
      ErrorMessage = Copient.PhraseLib.Lookup("terminal-sets.save-failed", LanguageID) & " " & _
                     Copient.PhraseLib.Lookup("term.reason", LanguageID) & ": " & ex.ToString
    End Try
    
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
    
  Sub Modify_Terminal_Set(ByVal TermSet As TerminalSetStruct)
    Dim SetID As Integer = 0
    
    Try
      Common.QueryStr = "BEGIN TRAN"
      Common.LRT_Execute()
    
      SetID = Find_TerminalSetID_By_Name(TermSet.Name)
      
      If SetID > 0 AndAlso SetID <> TermSet.TerminalSetID Then
        ErrorMessage = Copient.PhraseLib.Lookup("terminal-set.duplicate-name", LanguageID) & "(" & _
                       Copient.PhraseLib.Lookup("term.id", LanguageID) & SetID & ")"
      Else
        ' if this terminal set is now the default, then switch any existing default back to a standard terminal set.
        If TermSet.TerminalSetTypeID = SetTypes.DEFAULT_SET Then
          Common.QueryStr = "update TerminalSets with (RowLock) set TerminalSetTypeID=" & SetTypes.STANDARD_SET & ", LastUpdate=getdate(), UpdateLevel=UpdateLevel+1 " & _
                            "where TerminalSetTypeID=" & SetTypes.DEFAULT_SET & " and TerminalSetID<>" & TermSet.TerminalSetID & ";"
          Common.LRT_Execute()
          
          Update_Unassigned_Locations(TermSet)
        End If
        
        ' save changes to the terminal set record
        Common.QueryStr = "update TerminalSets with (RowLock) set Name=N'" & Common.Parse_Quotes(TermSet.Name) & "', PromoEngineID=" & TermSet.PromoEngineID & ", " & _
                          "  TerminalSetTypeID=" & TermSet.TerminalSetTypeID & ", LastUpdate=getdate(), UpdateLevel=UpdateLevel+1 " & _
                          "where TerminalSetID=" & TermSet.TerminalSetID & ";"
        Common.LRT_Execute()
              
        ' save changes to each of the terminal set items records
        For Each tsi As TerminalSetItem In TermSet.TerminalSetItems
          If tsi.PKID = 0 AndAlso Not tsi.Deleted Then
            Common.QueryStr = "insert into TerminalSetItems with (RowLock) (TerminalID, TerminalSetID, TerminalTypeID, PrinterTypeID, OpDisplayTypeID) " & _
                              "values (" & tsi.TerminalID & ", " & tsi.TerminalSetID & ", " & tsi.TerminalTypeID & "," & tsi.PrinterTypeID & ", " & _
                              "        " & tsi.OpDisplayTypeID & ")"
            Common.LRT_Execute()
          ElseIf tsi.PKID > 0 AndAlso Not tsi.Deleted Then
            Common.QueryStr = "update TerminalSetItems with (RowLock) set TerminalID=" & tsi.TerminalID & ", TerminalSetID=" & tsi.TerminalSetID & ", " & _
                              "  TerminalTypeID=" & tsi.TerminalTypeID & ", PrinterTypeID=" & tsi.PrinterTypeID & ", OpDisplayTypeID=" & tsi.OpDisplayTypeID & " " & _
                              "where PKID=" & tsi.PKID
            Common.LRT_Execute()
          ElseIf tsi.PKID > 0 AndAlso tsi.Deleted Then
            Common.QueryStr = "delete from TerminalSetItems with (RowLock) where PKID = " & tsi.PKID
            Common.LRT_Execute()
          End If
        Next
      End If
      
      Common.Activity_Log(48, TerminalSetID, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-set-edit", LanguageID))

      Common.QueryStr = "COMMIT TRAN"
      Common.LRT_Execute()
    Catch ex As Exception
      Common.QueryStr = "ROLLBACK TRAN"
      Common.LRT_Execute()
      ErrorMessage = Copient.PhraseLib.Lookup("terminal-sets.save-failed", LanguageID) & " " & _
                     Copient.PhraseLib.Lookup("term.reason", LanguageID) & ": " & ex.ToString
    End Try

  End Sub

      
  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Handle_Delete_Terminal_Set(ByVal TerminalSetID As Integer)
    Dim HasStores As Boolean = False
    
    Retrieve_Associated_Stores(TerminalSetID, HasStores)
    If HasStores Then
      ErrorMessage = Copient.PhraseLib.Lookup("terminal-sets.inuse", LanguageID)
      Build_Terminal_Set()
      Send_Page()
    Else
      Remove_Terminal_Set(TerminalSetID)
    End If
    
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------    
  
    
  Sub Remove_Terminal_Set(ByVal TerminalSetID As Integer)
    
    Try
      Common.QueryStr = "BEGIN TRAN"
      Common.LRT_Execute()
    
      Common.QueryStr = "delete from TerminalSetItems with (RowLock) " & _
                        "where TerminalSetID=" & TerminalSetID & ";"
      Common.LRT_Execute()

      Common.QueryStr = "delete from TerminalSets with (RowLock) " & _
                        "where TerminalSetID=" & TerminalSetID & ";"
      Common.LRT_Execute()
      

      Common.Activity_Log(48, TerminalSetID, AdminUserID, Copient.PhraseLib.Lookup("history.terminal-set-delete", LanguageID))

      Common.QueryStr = "COMMIT TRAN"
      Common.LRT_Execute()
    
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "terminal-sets-list.aspx")
      Response.End()
    Catch ex As Exception
      Common.QueryStr = "ROLLBACK TRAN"
      Common.LRT_Execute()
      ErrorMessage = Copient.PhraseLib.Lookup("terminal-sets.delete-failed", LanguageID) & " " & _
                     Copient.PhraseLib.Lookup("term.reason", LanguageID) & ": " & ex.ToString
    End Try

  End Sub
    
  
  '-------------------------------------------------------------------------------------------------------------    

  Sub Update_Unassigned_Locations(ByVal TermSet As TerminalSetStruct)
    
    ' find any locations that don't have a terminal set associated and update them to this default.
    Common.QueryStr = "insert into LocationTerminals with (RowLock) (LocationID, TerminalSetID, UpdateLevel, LastUpdate) " & _
                      "  select LOC.LocationID, " & TermSet.TerminalSetID & " as TerminalSetID, 1 as UpdateLevel, GETDATE() as LastUpdate" & _
                      "  from Locations as LOC with (NoLock) " & _
                      "  left join LocationTerminals as LT with (NoLock) on LT.LocationID = LOC.LocationID " & _
                      "  where LT.LocationID is null and LOC.Deleted=0 and LOC.EngineID = " & TermSet.PromoEngineID
    Common.LRT_Execute()
  End Sub
  
  
  '-------------------------------------------------------------------------------------------------------------    
  
    

  Sub Save_TerminalSet()
    
    Build_Terminal_Set()

    If Is_Valid_Entry() Then
      
      If TerminalSet.TerminalSetID = 0 Then
        Add_Terminal_Set(TerminalSet)
      ElseIf TerminalSet.TerminalSetID > 0 Then
        Modify_Terminal_Set(TerminalSet)
      End If
      
    End If

    If ErrorMessage.Trim = "" Then
      ' save was success, so reload page without form data to prevent form postback if user clicks the refresh/reload button.
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "terminal-sets-edit.aspx?TerminalSetID=" & TerminalSetID & "&infomsg=" & Server.UrlEncode(Copient.PhraseLib.Lookup("terminal-set.saved", LanguageID)))
      Response.End()
    Else
      Send_Page()
    End If
    
  End Sub
    

  '-------------------------------------------------------------------------------------------------------------    
    
  
  Function Find_TerminalSetID_By_Name(ByVal Name As String) As Integer
    Dim dt As DataTable
    Dim SetID As Integer = 0
    
    Common.QueryStr = "select TerminalSetID from TerminalSets with (NoLock) where Name = N'" & Common.Parse_Quotes(Name) & "';"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      SetID = Common.NZ(dt.Rows(0).Item("TerminalSetID"), 0)
    End If
    
    Return SetID
  End Function
    
  
  '-------------------------------------------------------------------------------------------------------------    
    
  
  Function Find_Default_TerminalSetID() As Integer
    Dim dt As DataTable
    Dim SetID As Integer = 0
    
    Common.QueryStr = "select TerminalSetID from TerminalSets with (NoLock) where TerminalSetTypeID=" & SetTypes.DEFAULT_SET & ";"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      SetID = Common.NZ(dt.Rows(0).Item("TerminalSetID"), 0)
    End If
    
    Return SetID
  End Function
    
  
  '-------------------------------------------------------------------------------------------------------------    
    
  
  Function Retrieve_Associated_Stores(ByVal TermSetID As Integer, Optional ByRef HasStores As Boolean = False) As DataTable
    Dim dtStores As DataTable
    
    Common.QueryStr = "select LOC.LocationID, LOC.LocationName from LocationTerminals as LT with (NoLock) " & _
                      "inner join Locations as LOC with (NoLock) on LOC.LocationID = LT.LocationID " & _
                      "where(LT.TerminalSetID = " & TermSetID & " And LOC.Deleted = 0)"
    dtStores = Common.LRT_Select
    
    HasStores = (dtStores.Rows.Count > 0)
    
    Return dtStores
  End Function
    
  
  '-------------------------------------------------------------------------------------------------------------    
    
  
  Sub Set_Browser_Type()
    Dim FullBrowserText As String = ""
    
    FullBrowserText = Request.Browser.Browser

    If FullBrowserText.IndexOf("IE") > -1 Then
      BrowserType = "IE"
    ElseIf FullBrowserText.IndexOf("Firefox") > -1 Then
      BrowserType = "Firefox"
    ElseIf FullBrowserText.IndexOf("Chrome") > -1 Then
      BrowserType = "Chrome"
    ElseIf FullBrowserText.IndexOf("Safari") > -1 Then
      BrowserType = "Safari"
    ElseIf FullBrowserText.IndexOf("Opera") > -1 Then
      BrowserType = "Opera"
    End If
      
  End Sub
</script>
<%
  '-------------------------------------------------------------------------------------------------------------    
  ' Execution starts here ... 
  
 
  Common.AppName = "terminal-sets-edit.aspx"
  
  Response.Expires = 0
  On Error GoTo ErrorTrap
  If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
  If Common.LXSadoConn.State = ConnectionState.Closed Then Common.Open_LogixXS()
  
  AdminUserID = Verify_AdminUser(Common, Logix)
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  Set_Browser_Type()
  
  TerminalSetID = Common.Extract_Val(GetCgiValue("TerminalSetID"))
  InfoMessage = GetCgiValue("infomsg")
  If InfoMessage Is Nothing Then InfoMessage = ""
  
  If GetCgiValue("save") <> "" Then
    Save_TerminalSet()
  ElseIf GetCgiValue("new") <> "" Then
    Handle_New_Click()
  ElseIf GetCgiValue("delete") <> "" Then
    Handle_Delete_Terminal_Set(TerminalSetID)
  Else
    Send_Page()
  End If
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  Common = Nothing
  Logix = Nothing

  Response.End()


ErrorTrap:
  Response.Write("<pre>" & Common.Error_Processor() & "</pre>")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  Common = Nothing
  Logix = Nothing
  
%>
