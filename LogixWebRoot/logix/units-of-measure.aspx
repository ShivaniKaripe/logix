<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="System.Collections.Generic" %>

<%@ Import Namespace="System.Xml" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: units-of-measure.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2012.  All rights reserved by:
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
  
<script runat="server">

  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim UnitsOfMeasure As New List(Of UOM)
  Dim LogText As String = ""
  Dim UOMDict As New Dictionary(Of Integer, String)
  Dim UOMSubTypeDict As New Dictionary(Of Integer, String)
  
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
  
  Enum UOMTypeIDs As Integer
    WEIGHT = 1
    VOLUME = 2
    LENGTH = 3
    SURFACE_AREA = 4
  End Enum
  
  
  ' -----------------------------------------------------------------------------------------------
  
#Region "Load and Display"
  Sub Load_UOM_Table()
    Dim dt As DataTable
    
    If UOMDict Is Nothing OrElse UOMDict.Count = 0 Then
      MyCommon.QueryStr = "select UOMTypeID, isnull(PT.Phrase,'') as Phrase " & _
                          "from UOMTypes as UOMT with (NoLock) " & _
                          "inner join UIPhrases as UIP with (NoLock) " & _
                          "  on UIP.Name = UOMT.PhraseTerm " & _
                          "inner join PhraseText as PT with (NoLock) " & _
                          "  on PT.PhraseID = UIP.PhraseID " & _
                          "where PT.LanguageID = " & LanguageID
      dt = MyCommon.LRT_Select
      For Each row As DataRow In dt.Rows
        UOMDict.Add(row.Item("UOMTypeID"), row.Item("Phrase"))
      Next
    End If
  End Sub

  Sub Load_UOM_SubTypesTable()
    Dim dt As DataTable
    
    If UOMSubTypeDict Is Nothing OrElse UOMSubTypeDict.Count = 0 Then
      MyCommon.QueryStr = "select UOMSubTypeID, isnull(PT.Phrase,'') as Phrase " & _
                          "from UOMSubTypes as UOMST with (NoLock) " & _
                          "inner join UIPhrases as UIP with (NoLock) " & _
                          "  on UIP.Name = UOMST.NamePhraseTerm " & _
                          "inner join PhraseText as PT with (NoLock) " & _
                          "  on PT.PhraseID = UIP.PhraseID " & _
                          "where PT.LanguageID = " & LanguageID
      dt = MyCommon.LRT_Select
      For Each row As DataRow In dt.Rows
        UOMSubTypeDict.Add(row.Item("UOMSubTypeID"), row.Item("Phrase"))
      Next
    End If
  End Sub

  Function Lookup_Type(ByVal UOMTypeID As Integer) As String
    Dim Result As String = UOMTypeID.ToString
      
    If UOMDict IsNot Nothing AndAlso UOMDict.ContainsKey(UOMTypeID) Then
      Result = UOMDict.Item(UOMTypeID)
    End If
    
    Return Result
  End Function

  Function Lookup_SubType(ByVal UOMSubTypeID As Integer) As String
    Dim Result As String = UOMSubTypeID.ToString
      
    If UOMSubTypeDict IsNot Nothing AndAlso UOMSubTypeDict.ContainsKey(UOMSubTypeID) Then
      Result = UOMSubTypeDict.Item(UOMSubTypeID)
    End If
    
    Return Result
  End Function

  Sub Load_UOM_Types()
    Dim Unit As New UOM
    Dim dt As DataTable
    
    UnitsOfMeasure = New List(Of UOM)
    
    MyCommon.QueryStr = "select distinct UOMT.UOMTypeID, UOMT.Name, UOMT.PhraseTerm, " & _
                        "  UOMT.DefaultUOMSubTypeID " & _
                        "from UOMTypes as UOMT with (NoLock) " & _
                        "inner join UOMSubTypes as UOMST with (NoLock) " & _
                        "  on UOMST.UOMTypeID = UOMT.UOMTypeID;"
    dt = MyCommon.LRT_Select
    For Each row As DataRow In dt.Rows
      Unit = New UOM(MyCommon.NZ(row.Item("UOMTypeID"), 0))
      With Unit
        .Name = MyCommon.NZ(row.Item("Name"), "")
        .PhraseTerm = MyCommon.NZ(row.Item("PhraseTerm"), "")
        .NameDisplayText = Copient.PhraseLib.Lookup(.PhraseTerm, LanguageID)
        .DefaultUOMSubTypeID = MyCommon.NZ(row.Item("DefaultUOMSubTypeID"), 0)
        .SubTypes = Get_UOM_SubTypes(.ID)
      End With
      UnitsOfMeasure.Add(Unit)
    Next
  End Sub
  
  Function Get_UOM_SubTypes(ByVal UOMTypeID As Integer) As List(Of UOMSubType)
    Dim UnitSubTypes As New List(Of UOMSubType)
    Dim SubType As New UOMSubType
    Dim dt As DataTable
    
    MyCommon.QueryStr = "select UOMST.UOMSubTypeID, UOMST.Name, " & _
                        "  UOMST.NamePhraseTerm, UOMST.Abbreviation, " & _
                        "  UOMST.AbbreviationPhraseTerm, UOMST.Precision " & _
                        "from UOMTypes as UOMT with (NoLock)  " & _
                        "inner join UOMSubTypes as UOMST with (NoLock) " & _
                        "  on UOMST.UOMTypeID = UOMT.UOMTypeID " & _
                        "where UOMT.UOMTypeID = " & UOMTypeID & ";"
         
    dt = MyCommon.LRT_Select
    For Each row As DataRow In dt.Rows
      SubType = New UOMSubType(MyCommon.NZ(row.Item("UOMSubTypeID"), 0))
      With SubType
        .Name = MyCommon.NZ(row.Item("Name"), "")
        .PhraseTerm = MyCommon.NZ(row.Item("NamePhraseTerm"), "")
        .NameDisplayText = Copient.PhraseLib.Lookup(.PhraseTerm, LanguageID)
        .Abbreviation = MyCommon.NZ(row.Item("Abbreviation"), "")
        .AbbrPhraseTerm = MyCommon.NZ(row.Item("AbbreviationPhraseTerm"), "")
        .AbbrDisplayText = Copient.PhraseLib.Lookup(.AbbrPhraseTerm, LanguageID)
        .Precision = MyCommon.NZ(row.Item("Precision"), 0)
      End With
      UnitSubTypes.Add(SubType)
    Next
   
    Return UnitSubTypes
  End Function

  Sub Send_Hidden_Form_Fields()
    If UnitsOfMeasure IsNot Nothing Then
      For Each u As UOM In UnitsOfMeasure
        Send("  <input type=""hidden"" name=""default" & u.ID & "_save"" value=""" & u.DefaultUOMSubTypeID & """ />")
      Next
    End If
  End Sub
  
  Sub Send_UOM_Box(ByVal UOMTypeID As UOMTypeIDs)
    For Each u As UOM In UnitsOfMeasure
      If u.ID = UOMTypeID Then
        Send("    <div class=""box"" id=""uom" & u.ID & """>")
        Send("      <h2 style=""float:left;""><span>" & u.NameDisplayText & "</span></h2>")
        Send("      <table summary=" & u.NameDisplayText & ">")
        Send("        <tr>")
        Send("          <th>" & Copient.PhraseLib.Lookup("term.default", LanguageID) & "</th>")
        Send("          <th>" & Copient.PhraseLib.Lookup("term.type", LanguageID) & "</th>")
        Send("          <th>" & Copient.PhraseLib.Lookup("term.precision", LanguageID) & "</th>")
        Send("        </tr>")
        For Each st As UOMSubType In u.SubTypes
          Send("        <tr>")
          Send("          <td><input type=""radio"" name=""default" & u.ID & """ value=""" & st.ID & """" & IIf(u.DefaultUOMSubTypeID = st.ID, " checked=""checked""", "") & " /></td>")
          Send("          <td>" & st.NameDisplayText & " (" & st.AbbrDisplayText & ")" & "</td>")
          Send("          <td>")
          Send("             <select class=""short"" name=""precision" & st.ID & """>")
          For i As Integer = 0 To 6
            Send("               <option value=""" & i & """" & IIf(st.Precision = i, " selected=""selected""", "") & ">" & i & "</option>")
          Next
          Send("             <input type=""hidden"" name=""precision" & st.ID & "_save"" value=""" & st.Precision & """ />")
          Send("          </td>")
          Send("        </tr>")
        Next
        Send("      </table>")
        Send("    </div>")
        Exit For
      End If
    Next
  End Sub

  Sub Send_Column1()
    Send("<div id=""column1"">")
    Send_UOM_Box(UOMTypeIDs.WEIGHT)
    Send_UOM_Box(UOMTypeIDs.LENGTH)
    Send_UOM_Box(UOMTypeIDs.SURFACE_AREA)
    Send("</div>")
  End Sub
  
  Sub Send_Column2()
    Send("<div id=""column1"">")
    Send_UOM_Box(UOMTypeIDs.VOLUME)
    Send("</div>")
  End Sub
  
  Sub Send_Gutter()
    Send("<div id=""gutter"">")
    Send("</div>")
  End Sub
#End Region

#Region "Save"
  Sub Save()
    ' load tables in memory for quick cross-reference while logging changes to the activity log.    
    Load_UOM_Table()
    Load_UOM_SubTypesTable()

    Save_Selected_Defaults()
    Save_SubType_Precisions()

    If LogText <> "" Then
      Send("<!-- " & LogText & " -->")
      MyCommon.Activity_Log(50, 0, AdminUserID, Left(LogText, 1000))
    End If
  End Sub
  
  Sub Save_Selected_Defaults()
    Dim SubTypeID As Integer
    Dim SubTypeIDSaved As Integer
    Dim LogBuf As New StringBuilder()
        Dim Types As UOMTypeIDs() = {UOMTypeIDs.WEIGHT, UOMTypeIDs.VOLUME, UOMTypeIDs.LENGTH, UOMTypeIDs.SURFACE_AREA}
    
    For Each UnitTypeID As UOMTypeIDs In Types
    Integer.TryParse(GetCgiValue("default" & UnitTypeID & "_save"), SubTypeIDSaved)
    Integer.TryParse(GetCgiValue("default" & UnitTypeID), SubTypeID)
    
    ' update the default selected only if it has changed.
    If SubTypeID <> SubTypeIDSaved Then
      MyCommon.QueryStr = "update UOMTypes with (RowLock) set DefaultUOMSubTypeID = " & SubTypeID & " " & _
                          "where UOMTypeID=" & UnitTypeID
      MyCommon.LRT_Execute()
      
        If LogBuf.Length > 0 Then LogBuf.Append(";")
        LogBuf.Append(" " & Lookup_Type(UnitTypeID) & " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & _
                      " " & Lookup_SubType(SubTypeIDSaved) & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & _
                      " " & Lookup_SubType(SubTypeID))
    End If
    Next
  
    If LogBuf.Length > 0 Then
      If LogText.Length > 0 Then LogText &= " "
      LogText &= Copient.PhraseLib.Lookup("unitsofmeasure.changedefault", LanguageID) & ":" & LogBuf.ToString & "."
    End If
    
  End Sub
  
  Sub Save_SubType_Precisions()
    Dim dt As DataTable
    Dim Precision As Integer
    Dim PrecisionSaved As Integer
    Dim LogBuf As New StringBuilder()
        
    MyCommon.QueryStr = "select UOMSubTypeID, NamePhraseTerm from UOMSubTypes with (NoLock);"
    dt = MyCommon.LRT_Select
    For Each row As DataRow In dt.Rows
      Integer.TryParse(GetCgiValue("precision" & row.Item("UOMSubTypeID")), Precision)
      Integer.TryParse(GetCgiValue("precision" & row.Item("UOMSubTypeID") & "_save"), PrecisionSaved)

      ' update the precision only if it has changed for the sub type.
      If Precision <> PrecisionSaved Then
        MyCommon.QueryStr = "update UOMSubTypes set Precision = " & Precision & " " & _
                            "where UOMSubTypeID=" & row.Item("UOMSubTypeID") & ";"
        MyCommon.LRT_Execute()
      

        If LogBuf.Length > 0 Then LogBuf.Append(";")
        LogBuf.Append(" " & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("NamePhraseTerm"), ""), LanguageID) & _
                      " " & Copient.PhraseLib.Lookup("term.from", LanguageID) & " " & PrecisionSaved.ToString & _
                      " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & " " & Precision.ToString)
      End If
    Next

    If LogBuf.Length > 0 Then
      If LogText.Length > 0 Then LogText &= " "
      LogText &= Copient.PhraseLib.Lookup("unitsofmeasure.changeprecision", LanguageID) & ":" & LogBuf.ToString & "."
    End If
  End Sub
  
#End Region
  </script>
<%  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False  
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "units-of-measure.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If GetCgiValue("Save") <> "" Then
    Save()
  End If
  Load_UOM_Types()
  
  Send_HeadBegin("term.unitsofmeasure")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 4)
  
  If (Logix.UserRoles.AccessSystemSettings = False) Then
    Send_Denied(1, "perm.admin-settings")
    GoTo done
  End If
%>
<form action="#" id="mainform" name="mainform" method="post">
<% Send_Hidden_Form_Fields()%>
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.unitsofmeasure", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemOptions = True) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(28, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <%
      Send_Column1()
      Send_Gutter()
      Send_Column2()
    %>
  </div> <!-- main -->
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(28, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
