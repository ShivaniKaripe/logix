<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: uom-sets-edit.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2012.  All rights reserved by:
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
  Dim UOMSet As New UOMSetStruct
  Dim ErrorMessage As String = ""
  Dim InfoMessage As String = ""
  Dim BrowserType As String = ""
  Dim UnitsOfMeasure As New List(Of UOM)
  
  Structure UOMSetStruct
    Public UOMSetID As Integer
    Public Name As String
    Public SetItems As List(Of UOMSetItem)
  End Structure
  
  Structure UOMSetItem
    Public PKID As Integer
    Public UOMSetID As Integer
    Public UOMTypeID As Integer
    Public UOMSubTypeID As Integer
    Public SubTypeName As String
  End Structure
    
  Structure UOM
    Dim ID As Integer
    Dim Name As String
    Dim PhraseTerm As String
    Dim NameDisplayText As String
    Dim DefaultUOMSubTypeID As Integer
    Dim SubTypes As List(Of UOMSubType)
    
    Sub New(ByVal UOMTypeID As Integer)
      Me.ID = UOMTypeID
      SubTypes = New List(Of UOMSubType)
    End Sub
  End Structure

  Structure UOMSubType
    Dim ID As Integer
    Dim Name As String
    Dim PhraseTerm As String
    Dim NameDisplayText As String
    Dim Abbreviation As String
    Dim AbbrPhraseTerm As String
    Dim AbbrDisplayText As String
    Dim Precision As Integer
    
    Sub New(ByVal UOMSubTypeID As Integer)
      Me.ID = UOMSubTypeID
    End Sub
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
    Send("    </div>")
  End Sub
  

  '-------------------------------------------------------------------------------------------------------------  
  
  Sub Send_Intro()
    Dim Title As String = ""
    Dim FullName As String = ""
    Dim dt As DataTable
    
    Send("  <div id=""intro"">")
    Send("    <input type=""hidden"" id=""UOMSetID"" name=""UOMSetID"" value=""" & UOMSet.UOMSetID & """ />")

    If UOMSet.UOMSetID = 0 Then
      Title = Copient.PhraseLib.Lookup("term.newuomset", LanguageID)
    Else
      Common.QueryStr = "SELECT Name FROM UOMSets with (NoLock) WHERE UOMSetID = " & UOMSet.UOMSetID & ";"
      dt = Common.LRT_Select
      If (dt.Rows.Count > 0) Then
        FullName = Common.NZ(dt.Rows(0).Item("Name"), "")
        Title = FullName
        If (Len(FullName) > 30) Then
          Title = Left(FullName, 27) & "..."
        End If
      End If
      Title = Copient.PhraseLib.Lookup("term.unitsofmeasureset", LanguageID) & " #" & UOMSet.UOMSetID & ": " & Title
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
    
    Name = UOMSet.Name
    If UOMSet.Name Is Nothing Then Name = ""
        
    Send("      <div class=""box"" id=""identification"">")
    Send("        <h2>")
    Send("          <span>" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & "</span>")
    Send("        </h2>")
    Send("        <label for=""SetName"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
    Send("        <input type=""text"" class=""longest"" id=""SetName"" name=""SetName"" maxlength=""200"" value=""" & Name.Replace("""", "&quot;") & """ /><br />")
    Send("        <br class=""half"" />")
    Send("        <br />")
    Send("      </div>")
    
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------  

  
  Sub Send_AssociatedStores()
    Dim dtStores As DataTable
    Dim row As DataRow
    
    dtStores = Retrieve_Associated_Stores(UOMSet.UOMSetID)
    
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
      Sendb(Copient.PhraseLib.Lookup("uom-sets.nostores", LanguageID) & "<br />")
      Send("      </div>")
    End If

    Send("        <hr class=""hidden"" />")
    Send("     </div>")

  End Sub

  
  '-------------------------------------------------------------------------------------------------------------  


  Sub Send_Items()
    Dim SubTypeIDs As String = ""

    Send("      <div class=""box"" id=""items"">")
    Send("        <h2>")
    Send("          <span>" & Copient.PhraseLib.Lookup("term.items", LanguageID) & "</span>")
    Send("        </h2>")

    For Each u As UOM In UnitsOfMeasure
      SubTypeIDs = ""
      
      Send("<input type=""hidden"" name=""uomtype"" value=""" & u.ID & """ />")
      Send("<b><u>" & Copient.PhraseLib.Lookup(u.PhraseTerm, LanguageID) & "</u></b><br />")
      Send("        <table id=""items" & u.ID & """ style=""width:650px;"" summary=""" & Copient.PhraseLib.Lookup(u.PhraseTerm, LanguageID) & """>")
      Send("          <tr>")
      Send("            <td style=""width: 280px;"">" & Copient.PhraseLib.Lookup("term.available", LanguageID) & ":</td>")
      Send("            <td></td>")
      Send("            <td style=""width: 280px;"">" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</td>")
      Send("          </tr>")
      Send("          <tr>")
      Send("            <td>")
      Send("              <select id=""availuomtype" & u.ID & """ name=""availuomtype" & u.ID & """ class=""long"" size=""5"" multiple=""multiple"" ondblclick=""moveItems(" & u.ID & ",1);"">")
      For Each s As UOMSubType In u.SubTypes
        If Not Is_SubType_Selected(u.ID, s.ID) Then
        Send("                <option value=""" & s.ID & """>" & s.NameDisplayText & "(" & s.AbbrDisplayText & ")</option>")
        End If
      Next
      Send("              </select>")
      Send("            </td>")
      Send("            <td style=""vertical-align:center;"">")
      Send("              <input type=""button"" class=""arrowadd"" value=""&gt;"" onclick=""moveItems(" & u.ID & ",1);"" /><br /><br />")
      Send("              <input type=""button"" class=""arrowrem"" value=""&lt;"" onclick=""moveItems(" & u.ID & ",0);"" />")
      Send("            </td>")
      Send("            <td>")
      Send("              <select id=""seluomtype" & u.ID & """ name=""seluomtype" & u.ID & """ class=""long"" size=""5"" multiple=""multiple"" ondblclick=""moveItems(" & u.ID & ",0);"">")
      If UOMSet.SetItems IsNot Nothing Then
        For Each Item As UOMSetItem In UOMSet.SetItems
          If Item.UOMTypeID = u.ID Then
            Send("                <option value=""" & Item.UOMSubTypeID & """>" & Item.SubTypeName & "</option>")
            If SubTypeIDs.Length > 0 Then SubTypeIDs &= ","
            SubTypeIDs &= Item.UOMSubTypeID
          End If
        Next
      End If
      Send("              </select>")
      Send("            </td>")
      Send("          </tr>")
      Send("        </table>")
      Send("        <input type=""hidden"" id=""seltype" & u.ID & """ name=""seltype" & u.ID & """ value=""" & SubTypeIDs & """ />")
      Send("        <br />")
      Send("        <br class=""half"" />")
    Next
    Send("      </div>")
    
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
    
    IsNew = UOMSet.UOMSetID = 0
    
    If UOMSet.UOMSetID > 0 OrElse IsNew Then
      'Send("    <div id=""column"">")

      Send("    <div id=""column1"">")
      Send_Identification()
      Send("    </div>")
      
      Send("    <div id=""gutter""></div>")
      
      If UOMSet.UOMSetID > 0 Then
        Send("    <div id=""column2"">")
        Send_AssociatedStores()
        Send("    </div>")
      End If

      Send("    <div id=""column"">")
      Send_Items()
      Send("    </div>")
      Send("  </div>")
    End If
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------  


  Sub Send_Form()
    Send("<form action=""uom-sets-edit.aspx"" id=""mainform"" name=""mainform"" method=""POST"">")
    If Not Logix.UserRoles.EditUOMSets Then
      Send_Denied(1, "perm.190")
    Else
      Send_Intro()
      Send_Main()
    End If
    Send("</form>")
  End Sub

  
  
  '-------------------------------------------------------------------------------------------------------------  
  
  Sub Send_Page_JS()
    Send("<script type=""text/javascript"">")
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
    Send("  function moveItems(uomtypeid, direction) { ")
    Send("    var elemSource, elemDest = null;")
    Send("")
    Send("    if (direction == 1) { ")
    Send("      elemSource = document.getElementById('availuomtype' + uomtypeid);")
    Send("      elemDest = document.getElementById('seluomtype' + uomtypeid);")
    Send("    } else {")
    Send("      elemSource = document.getElementById('seluomtype' + uomtypeid);")
    Send("      elemDest = document.getElementById('availuomtype' + uomtypeid);")
    Send("    }")
    Send("    ")
    Send("    for (var i=0; i < elemSource.options.length; i++) {")
    Send("      if (elemSource.options[i].selected) { ")
    Send("        elemDest.options[elemDest.options.length] = new Option(elemSource.options[i].text, elemSource.options[i].value);")
    Send("        elemSource.options[i] = null;")
    Send("        i--;")
    Send("      }")
    Send("    }")
    Send("    sortSelect(elemDest);")
    Send("    updateSelected(uomtypeid);")
    Send("  }")
    Send("")
    Send("  function sortSelect(selElem) {")
    Send("    var tmpAry = new Array();")
    Send("    for (var i=0;i<selElem.options.length;i++) {")
    Send("      tmpAry[i] = new Array();")
    Send("      tmpAry[i][0] = selElem.options[i].text;")
    Send("      tmpAry[i][1] = selElem.options[i].value;")
    Send("    }")
    Send("    tmpAry.sort();")
    Send("    while (selElem.options.length > 0) {")
    Send("      selElem.options[0] = null;")
    Send("    }")
    Send("    for (var i=0;i<tmpAry.length;i++) {")
    Send("      var op = new Option(tmpAry[i][0], tmpAry[i][1]);")
    Send("      selElem.options[i] = op;")
    Send("    }")
    Send("    return;")
    Send("  }")
    Send("")
    Send("  function updateSelected(uomtypeid) {")
    Send("    var elemHidden = document.getElementById('seltype' + uomtypeid);")
    Send("    var elemDisplay = document.getElementById('seluomtype' + uomtypeid);")
    Send("    var newValue = '';")
    Send("")
    Send("    if (elemHidden != null && elemDisplay != null) {")
    Send("      for (var i=0; i < elemDisplay.options.length; i++) { ")
    Send("        if (i >0) { newValue += ','};")
    Send("        newValue += elemDisplay.options[i].value;")
    Send("      }")
    Send("      elemHidden.value = newValue;")
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
    
    Send_HeadBegin("term.unitsofmeasuresets", "term.edit", UOMSet.UOMSetID)
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
    Send_Subtabs(Logix, 75, 4, , UOMSet.UOMSetID)

    If Not Logix.UserRoles.EditUOMSets Then
      Send_Denied(1, "perm.190")
    Else
      Send_Form()
    End If

    Send_BodyEnd()

  End Sub
  
  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Load_UOMSet()
    Dim dt As DataTable
    Dim Item As UOMSetItem
    
    UOMSet.SetItems = New List(Of UOMSetItem)
    
    Common.QueryStr = "select UOM.Name, UOM.UOMSetID " & _
                      "from UOMSets as UOM with (NoLock) " & _
                      "where UOMSetID=" & UOMSet.UOMSetID & ";"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      With UOMSet
        .Name = Common.NZ(dt.Rows(0).Item("Name"), "")
      End With
      
      Common.QueryStr = "select UOMI.PKID, UOMI.UOMSetID, UOMI.UOMTypeID, UOMI.UOMSubTypeID " & _
                        "from UOMSetItems as UOMI with (NoLock) " & _
                        "where UOMI.UOMSetID=" & UOMSet.UOMSetID & " " & _
                        "order by UOMSubTypeID;"
    
      dt = Common.LRT_Select
      For Each row As DataRow In dt.Rows
        Item = New UOMSetItem
        With Item
          .PKID = Common.NZ(row.Item("PKID"), 0)
          .UOMSetID = Common.NZ(row.Item("UOMSetID"), 0)
          .UOMTypeID = Common.NZ(row.Item("UOMTypeID"), 0)
          .UOMSubTypeID = Common.NZ(row.Item("UOMSubTypeID"), 0)
          .SubTypeName = Get_SubType_Text(.UOMTypeID, .UOMSubTypeID)
        End With
        UOMSet.SetItems.Add(Item)
      Next
      
      UOMSet.SetItems.Sort(AddressOf SetItemCompare)
      
    ElseIf UOMSet.UOMSetID > 0 Then
      ErrorMessage = Copient.PhraseLib.Lookup("uom-sets.not-found", LanguageID) & _
                     "(" & Copient.PhraseLib.Lookup("term.id", LanguageID) & " = " & UOMSet.UOMSetID & ")"
    End If

  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
  Private Function SetItemCompare(ByVal u1 As UOMSetItem, ByVal u2 As UOMSetItem) As Integer
    Return u1.SubTypeName.CompareTo(u2.SubTypeName)
  End Function
  
  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Load_UOM_Types()
    Dim Unit As New UOM
    Dim dt As DataTable
    
    UnitsOfMeasure = New List(Of UOM)
    
    Common.QueryStr = "select distinct UOMT.UOMTypeID, UOMT.Name, UOMT.PhraseTerm, " & _
                        "  UOMT.DefaultUOMSubTypeID " & _
                        "from UOMTypes as UOMT with (NoLock) " & _
                        "inner join UOMSubTypes as UOMST with (NoLock) " & _
                        "  on UOMST.UOMTypeID = UOMT.UOMTypeID;"
    dt = Common.LRT_Select
    For Each row As DataRow In dt.Rows
      Unit = New UOM(Common.NZ(row.Item("UOMTypeID"), 0))
      With Unit
        .Name = Common.NZ(row.Item("Name"), "")
        .PhraseTerm = Common.NZ(row.Item("PhraseTerm"), "")
        .NameDisplayText = Copient.PhraseLib.Lookup(.PhraseTerm, LanguageID)
        .DefaultUOMSubTypeID = Common.NZ(row.Item("DefaultUOMSubTypeID"), 0)
        .SubTypes = Get_UOM_SubTypes(.ID)
      End With
      UnitsOfMeasure.Add(Unit)
    Next
    
    ' sort the display text as it could be out of order due to language differences.
    UnitsOfMeasure.Sort(AddressOf UOMCompare)
    
  End Sub

  '-------------------------------------------------------------------------------------------------------------    
  
  Private Function UOMCompare(ByVal u1 As UOM, ByVal u2 As UOM) As Integer
    Return u1.NameDisplayText.CompareTo(u2.NameDisplayText)
  End Function

  
  '-------------------------------------------------------------------------------------------------------------    
  
  Function Get_UOM_SubTypes(ByVal UOMTypeID As Integer) As List(Of UOMSubType)
    Dim UnitSubTypes As New List(Of UOMSubType)
    Dim SubType As New UOMSubType
    Dim dt As DataTable
    
    Common.QueryStr = "select UOMST.UOMSubTypeID, UOMST.Name, " & _
                        "  UOMST.NamePhraseTerm, UOMST.Abbreviation, " & _
                        "  UOMST.AbbreviationPhraseTerm, UOMST.Precision " & _
                        "from UOMTypes as UOMT with (NoLock)  " & _
                        "inner join UOMSubTypes as UOMST with (NoLock) " & _
                        "  on UOMST.UOMTypeID = UOMT.UOMTypeID " & _
                        "where UOMT.UOMTypeID = " & UOMTypeID & ";"
         
    dt = Common.LRT_Select
    For Each row As DataRow In dt.Rows
      SubType = New UOMSubType(Common.NZ(row.Item("UOMSubTypeID"), 0))
      With SubType
        .Name = Common.NZ(row.Item("Name"), "")
        .PhraseTerm = Common.NZ(row.Item("NamePhraseTerm"), "")
        .NameDisplayText = Copient.PhraseLib.Lookup(.PhraseTerm, LanguageID)
        .Abbreviation = Common.NZ(row.Item("Abbreviation"), "")
        .AbbrPhraseTerm = Common.NZ(row.Item("AbbreviationPhraseTerm"), "")
        .AbbrDisplayText = Copient.PhraseLib.Lookup(.AbbrPhraseTerm, LanguageID)
        .Precision = Common.NZ(row.Item("Precision"), 0)
      End With
      UnitSubTypes.Add(SubType)
    Next
   
    ' sort the display text as it could be out of order due to language differences.
    UnitSubTypes.Sort(AddressOf UOMSubTypeCompare)
    
    Return UnitSubTypes
  End Function

  '-------------------------------------------------------------------------------------------------------------    

  Private Function UOMSubTypeCompare(ByVal u1 As UOMSubType, ByVal u2 As UOMSubType) As Integer
    Return u1.NameDisplayText.CompareTo(u2.NameDisplayText)
  End Function
  
  '-------------------------------------------------------------------------------------------------------------    

  Private Function Is_SubType_Selected(ByVal UOMTypeID As Integer, ByVal SubTypeID As Integer) As Boolean
    Dim Selected As Boolean
    
    For Each Item As UOMSetItem In UOMSet.SetItems
      Selected = (Item.UOMTypeID = UOMTypeID AndAlso Item.UOMSubTypeID = SubTypeID)
      If Selected Then Exit For
    Next
      
    Return Selected
  End Function
  
  '-------------------------------------------------------------------------------------------------------------    
  
  Function Get_SubType_Text(ByVal UOMTypeID As Integer, ByVal UOMSubTypeID As Integer) As String
    Dim Text As String = ""
    
    For Each u As UOM In UnitsOfMeasure
      If u.ID = UOMTypeID Then
        For Each st As UOMSubType In u.SubTypes
          If st.ID = UOMSubTypeID Then
            Text = st.NameDisplayText & "(" & st.AbbrDisplayText & ")"
          End If
        Next
      End If
    Next
    
    Return Text
  End Function
  
  '-------------------------------------------------------------------------------------------------------------    

  
  Sub Load_Page_Data()
    Load_UOM_Types()
    Load_UOMSet()
  End Sub
    

  '-------------------------------------------------------------------------------------------------------------    
  
  
  Sub Build_UOM_Set()
    UOMSet = New UOMSetStruct
    
    With UOMSet
      .Name = GetCgiValue("SetName")
      .UOMSetID = Common.Extract_Val(GetCgiValue("UOMSetID"))
    End With
    Build_Items()
  End Sub
      
  
  '-------------------------------------------------------------------------------------------------------------    
  
    
  Sub Build_Items()
    Dim UOMTypes() As String
    Dim UOMSubTypes() As String
    Dim UOMTypeID As Integer
    Dim UOMSubTypeID As Integer
    Dim Item As UOMSetItem
    Dim Temp As String
    
    UOMSet.SetItems = New List(Of UOMSetItem)
    
    ' load up all the UOMTypes
    UOMTypes = Request.Form.GetValues("uomtype")
    If UOMTypes IsNot Nothing Then
      For Each s As String In UOMTypes
        If Integer.TryParse(s, UOMTypeID) Then
          ' get their associated selected sub types
          Temp = GetCgiValue("seltype" & UOMTypeID)
          If Temp IsNot Nothing Then
            UOMSubTypes = Temp.Split(",")
            For Each st As String In UOMSubTypes
              If Integer.TryParse(st, UOMSubTypeID) Then
                Item = New UOMSetItem
                Item.UOMSetID = UOMSet.UOMSetID
                Item.UOMTypeID = UOMTypeID
                Item.UOMSubTypeID = UOMSubTypeID
                UOMSet.SetItems.Add(Item)
              End If
            Next
          End If
        End If
      Next
    End If
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Handle_New_Click()
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "uom-sets-edit.aspx")
    Response.End()
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Add_UOM_Set()
    Dim RetCode As Integer
  
    Try
      Common.QueryStr = "BEGIN TRAN"
      Common.LRT_Execute()

      ' first create the uom set record
      Common.QueryStr = "dbo.pt_UOMSets_Insert"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 255).Value = UOMSet.Name
      Common.LRTsp.Parameters.Add("@RetCode", SqlDbType.Int).Direction = ParameterDirection.Output
      Common.LRTsp.Parameters.Add("@UOMSetID", SqlDbType.Int).Direction = ParameterDirection.Output
      Common.LRTsp.ExecuteNonQuery()
      RetCode = Common.NZ(Common.LRTsp.Parameters("@RetCode").Value, -99)
      UOMSet.UOMSetID = Common.NZ(Common.LRTsp.Parameters("@UOMSetID").Value, 0)
      Common.Close_LRTsp()
    
      Select Case RetCode
        Case 1
          InfoMessage = Copient.PhraseLib.Lookup("uom-set.saved", LanguageID)
        Case -1
          ErrorMessage = Copient.PhraseLib.Lookup("uom-set.duplicate-name", LanguageID)
        Case Else
          ErrorMessage = Copient.PhraseLib.Lookup("uom-sets.save-failed", LanguageID)
      End Select
      
      InfoMessage &= " Item Count = " & UOMSet.SetItems.Count
      ' then create the uom set items records
      For Each usi As UOMSetItem In UOMSet.SetItems
        Common.QueryStr = "insert into UOMSetItems (UOMSetID, UOMTypeID, UOMSubTypeID) " & _
                          "values (" & UOMSet.UOMSetID & ", " & usi.UOMTypeID & ", " & usi.UOMSubTypeID & ")"
        Common.LRT_Execute()
      Next
      
      Common.Activity_Log(51, 1, UOMSet.UOMSetID, AdminUserID, Copient.PhraseLib.Lookup("history.uom-set-create", LanguageID), "")
    
      Common.QueryStr = "COMMIT TRAN"
      Common.LRT_Execute()
    Catch ex As Exception
      Common.QueryStr = "ROLLBACK TRAN"
      Common.LRT_Execute()
      ErrorMessage = Copient.PhraseLib.Lookup("uom-sets.save-failed", LanguageID) & " " & _
                     Copient.PhraseLib.Lookup("term.reason", LanguageID) & ": " & ex.ToString
    End Try
  
  End Sub

  
  '-------------------------------------------------------------------------------------------------------------    
  
    
  Sub Modify_UOM_Set()
    Dim SetID As Integer = 0
    
    Try
      Common.QueryStr = "BEGIN TRAN"
      Common.LRT_Execute()
    
      SetID = Find_UOMSetID_By_Name(UOMSet.Name)
      
      If SetID > 0 AndAlso SetID <> UOMSet.UOMSetID Then
        ErrorMessage = Copient.PhraseLib.Lookup("uom-set.duplicate-name", LanguageID) & "(" & _
                       Copient.PhraseLib.Lookup("term.id", LanguageID) & SetID & ")"
      Else
        ' save changes to the uom set record
        Common.QueryStr = "update UOMSets with (RowLock) set Name=N'" & Common.Parse_Quotes(UOMSet.Name) & "' " & _
                          "where UOMSetID=" & UOMSet.UOMSetID & ";"
        Common.LRT_Execute()
          
        ' remove all the pre-existing list items for this uom set
        Common.QueryStr = "delete from UOMSetItems with (RowLock) where UOMSetID=" & UOMSet.UOMSetID & ";"
        Common.LRT_Execute()
        
        ' save changes to each of the uom set items records
        For Each usi As UOMSetItem In UOMSet.SetItems
          Common.QueryStr = "insert into UOMSetItems with (RowLock) (UOMSetID, UOMTypeID, UOMSubTypeID) " & _
                            "values (" & UOMSet.UOMSetID & ", " & usi.UOMTypeID & ", " & usi.UOMSubTypeID & ");"
          Common.LRT_Execute()
        Next
      End If
              
      Common.Activity_Log(51, 2, UOMSet.UOMSetID, AdminUserID, Copient.PhraseLib.Lookup("history.uom-set-edit", LanguageID), "")
      
      Common.QueryStr = "COMMIT TRAN"
      Common.LRT_Execute()
    Catch ex As Exception
      Common.QueryStr = "ROLLBACK TRAN"
      Common.LRT_Execute()
      ErrorMessage = Copient.PhraseLib.Lookup("uom-sets.save-failed", LanguageID) & " " & _
                     Copient.PhraseLib.Lookup("term.reason", LanguageID) & ": " & ex.ToString
    End Try

  End Sub

      
  '-------------------------------------------------------------------------------------------------------------    
  
  Sub Handle_Delete_UOM_Set()
    Dim HasStores As Boolean = False
    
    Retrieve_Associated_Stores(HasStores)
    If HasStores Then
      ErrorMessage = Copient.PhraseLib.Lookup("uom-sets.inuse", LanguageID)
      Build_UOM_Set()
      Send_Page()
    Else
      Remove_UOM_Set()
    End If
    
  End Sub
  
  '-------------------------------------------------------------------------------------------------------------    
  
    
  Sub Remove_UOM_Set()
    
    Try
      Common.QueryStr = "BEGIN TRAN"
      Common.LRT_Execute()
    
      Common.QueryStr = "delete from UOMSetItems with (RowLock) " & _
                        "where UOMSetID=" & UOMSet.UOMSetID & ";"
      Common.LRT_Execute()

      Common.QueryStr = "delete from UOMSets with (RowLock) " & _
                        "where UOMSetID=" & UOMSet.UOMSetID & ";"
      Common.LRT_Execute()
      

      Common.Activity_Log(51, 7, UOMSet.UOMSetID, AdminUserID, Copient.PhraseLib.Lookup("history.uom-set-delete", LanguageID), "")

      Common.QueryStr = "COMMIT TRAN"
      Common.LRT_Execute()
    
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "uom-sets-list.aspx")
      Response.End()
    Catch ex As Exception
      Common.QueryStr = "ROLLBACK TRAN"
      Common.LRT_Execute()
      ErrorMessage = Copient.PhraseLib.Lookup("uom-sets.delete-failed", LanguageID) & " " & _
                     Copient.PhraseLib.Lookup("term.reason", LanguageID) & ": " & ex.ToString
    End Try

  End Sub
    
  
    '-------------------------------------------------------------------------------------------------------------    
    Sub ValidateUOMSet_Name()
        UOMSet.Name = Logix.TrimAll(UOMSet.Name)
        If ((UOMSet.Name).Equals("")) Then
            If ErrorMessage = "" Then ErrorMessage = Copient.PhraseLib.Lookup("uom-set.no-name", LanguageID)
        End If
    End Sub
    '-------------------------------------------------------------------------------------------------------------
    Sub Save_UOMSet()
  
        Build_UOM_Set()
        ValidateUOMSet_Name()
        If ErrorMessage = "" Then
            'if name is not empty only then add/modify UOM Set   
            If UOMSet.UOMSetID = 0 Then
                Add_UOM_Set()
            Else
                Modify_UOM_Set()
            End If
        End If
       
      
    If ErrorMessage.Trim = "" Then
      ' save was success, so reload page without form data to prevent form postback if user clicks the refresh/reload button.
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "uom-sets-edit.aspx?UOMSetID=" & UOMSet.UOMSetID & "&infomsg=" & Server.UrlEncode(Copient.PhraseLib.Lookup("uom-set.saved", LanguageID)))
      Response.End()
    Else
      Send_Page()
    End If
      
  End Sub
    

  '-------------------------------------------------------------------------------------------------------------    
    
  
  Function Find_UOMSetID_By_Name(ByVal Name As String) As Integer
    Dim dt As DataTable
    Dim SetID As Integer = 0
    
    Common.QueryStr = "select UOMSetID from UOMSets with (NoLock) where Name = N'" & Common.Parse_Quotes(Name) & "';"
    dt = Common.LRT_Select
    If dt.Rows.Count > 0 Then
      SetID = Common.NZ(dt.Rows(0).Item("UOMSetID"), 0)
    End If
    
    Return SetID
  End Function
    
  
  '-------------------------------------------------------------------------------------------------------------    
    
  
  Function Retrieve_Associated_Stores(Optional ByRef HasStores As Boolean = False) As DataTable
    Dim dtStores As DataTable
    
    Common.QueryStr = "select LOC.LocationID, LOC.LocationName from Locations as LOC with (NoLock) " & _
                      "where LOC.UOMSetID = " & UOMSet.UOMSetID & " And LOC.Deleted = 0"
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
  
 
  Common.AppName = "uom-sets-edit.aspx"
  
  Response.Expires = 0
  On Error GoTo ErrorTrap
  If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
  If Common.LXSadoConn.State = ConnectionState.Closed Then Common.Open_LogixXS()
  
  AdminUserID = Verify_AdminUser(Common, Logix)
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  Set_Browser_Type()
  
  UOMSet.UOMSetID = Common.Extract_Val(GetCgiValue("UOMSetID"))
  InfoMessage = GetCgiValue("infomsg")
  If InfoMessage Is Nothing Then InfoMessage = ""
  
  If GetCgiValue("save") <> "" Then
    Save_UOMSet()
  ElseIf GetCgiValue("new") <> "" Then
    Handle_New_Click()
  ElseIf GetCgiValue("delete") <> "" Then
    Handle_Delete_UOM_Set()
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
