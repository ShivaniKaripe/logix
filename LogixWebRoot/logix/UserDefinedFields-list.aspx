<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:5.99.1.64187.Unstable Build - WUSNS185105-0JV %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.Commoninc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: UserDefinedFields-list.aspx 
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
  ' *
  ' * PURPOSE : 
  ' *
  ' * NOTES   : 
  ' *
  ' * Version : 7.3.1.138972
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As System.Data.DataTable
  Dim dst2 As System.Data.DataTable
  Dim idNumber As Integer
  Dim idSearch As String
  Dim idSearchText As String
  Dim sSearchQuery As String
  Dim PageNum As Integer
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim sizeOfData As Integer
  Dim i As Integer
  Dim PrctSignPos As Integer
  Dim Shaded As String = "shaded"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "UserDefinedFields-list.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then
    PageNum = 0
  End If
  MorePages = False
  
  Try
    Send_HeadBegin("term.userdefinedfields")
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
    
    If (Logix.UserRoles.EditSystemConfiguration = False) Then
      Send_Denied(1, "perm.admin-configuration")
      GoTo done
    End If
    
    
    BannersEnabled = (MyCommon.Fetch_SystemOption("66") = "1")
    
    Dim SortText As String = "UDFPK"
    Dim SortDirection As String 
    
    If (Request.QueryString("SortText") <> "") Then
      SortText = Request.QueryString("SortText")
    End If
    
    If (Request.QueryString("pagenum") = "") Then
      If (Request.QueryString("SortDirection") = "ASC") Then
        SortDirection = "DESC"
      ElseIf (Request.QueryString("SortDirection") = "DESC") Then
        SortDirection = "ASC"
      Else
        SortDirection = "DESC"
      End If
    Else
      SortDirection = Request.QueryString("SortDirection")
    End If
    
    sSearchQuery = "select  udf.UDFPK, udf.ExternalID, udf.Description, (select COUNT(offerid) from userdefinedfieldsvalues as o where o.UDFPK = udf.UDFPK) 'NumOfOffers', " & _
					"udf.LastUpdate, t.Type 'DataType' from UserDefinedFields as udf inner join UserDefinedFieldsTypes as t on udf.DataType = t.UDFTypeID"
    idSearchText = Request.QueryString("searchterms")
    If (idSearchText <> "") Then
      If (Integer.TryParse(Request.QueryString("searchterms"), idNumber)) Then
        idSearch = idNumber.ToString
      Else
        idSearch = "-1"
      End If
      PrctSignPos = idSearchText.IndexOf("%")
      If (PrctSignPos > -1) Then
        idSearch = "-1"
        idSearchText = idSearchText.Replace("%", "[%]")
      End If
      If (idSearchText.IndexOf("_") > -1) Then
        idSearchText = idSearchText.Replace("_", "[_]")
      End If
      idSearchText = MyCommon.Parse_Quotes(idSearchText)
      'sSearchQuery = sSearchQuery & " and (AttributeTypeID=" & MyCommon.Extract_Val(idSearchText)
      sSearchQuery = sSearchQuery & " where ( Description like N'%" & idSearchText & "%'"
      sSearchQuery = sSearchQuery & " or ExternalID like N'%" & idSearchText & "%') "
    End If
    MyCommon.QueryStr = sSearchQuery & " order by " & SortText & " " & SortDirection
    dst = MyCommon.LRT_Select
    sizeOfData = dst.Rows.Count
    i = linesPerPage * PageNum
    
    If (sizeOfData = 1 AndAlso Request.QueryString("searchterms") <> "") Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "userdefinedfields-edit.aspx?UDFPK=" & dst.Rows(i).Item("UDFPK"))
    End If
%>
<div id="intro">
  <h1 id="title">
   <%Copient.PhraseLib.Lookup("term.userdefinedfields", LanguageID)%>
  </h1>
  <div id="controls">
    <form action="userdefinedfields-edit.aspx" id="controlsform" name="controlsform">
    <%
      If (Logix.UserRoles.AddUserDefinedFields = True) Then
        Send_New()
      End If
    %>
    </form>
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;SortText=" & SortText & "&amp;SortDirection=" & SortDirection)
  %>
  <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.attributes", LanguageID)) %>">
    <thead>
      <tr>
	  <th align="left" class="th-id" scope="col">
          <a href="UserDefinedFields-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=UDFPK&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
          </a>
          <%
            If SortText = "UDFPK" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="UserDefinedFields-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=ExternalID&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.externalid", LanguageID))%>
          </a>
          <%
            If SortText = "ExternalID" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-name" scope="col">
          <a href="UserDefinedFields-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=Description&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          </a>
          <%
            If SortText = "Description" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left"  class="th-date" scope="col">
		  <a href="UserDefinedFields-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=NumOfOffers&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.numofoffers", LanguageID))%>
          </a>
          <%
            If SortText = "NumOfOffers" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left"  class="th-id" scope="col">
		  <a href="UserDefinedFields-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=datatype&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.datatype", LanguageID))%>
          </a>
          <%
            If SortText = "datatype" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
        <th align="left" class="th-datetime" scope="col">
          <a href="UserDefinedFields-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&amp;SortText=LastUpdate&amp;SortDirection=<% Sendb(SortDirection) %>">
            <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
          </a>
          <%
            If SortText = "LastUpdate" Then
              If SortDirection = "ASC" Then
                Sendb("<span class=""sortarrow"">&#9660;</span>")
              Else
                Sendb("<span class=""sortarrow"">&#9650;</span>")
              End If
            Else
            End If
          %>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
        Shaded = "shaded"
        While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
          Send("      <tr class=""" & Shaded & """>")
          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("UDFPK"), 0) & "</td>")
          Send("        <td>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("ExternalID"), "&nbsp;"), 25) & "</td>")
          Send("        <td><a href=""UserDefinedFields-edit.aspx?UDFPK=" & MyCommon.NZ(dst.Rows(i).Item("UDFPK"), 0) & """>" & MyCommon.SplitNonSpacedString(MyCommon.NZ(dst.Rows(i).Item("Description"), "&nbsp;"), 30) & "</a></td>")

          Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("NumOfOffers"), 0) & "</td>")
		  Send("        <td>" & MyCommon.NZ(dst.Rows(i).Item("DataType"), "") & "</td>")
          If (Not IsDBNull(dst.Rows(i).Item("LastUpdate"))) Then
            Send("        <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("LastUpdate"), MyCommon) & "</td>")
          Else
            Send("        <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
          End If
          Send("      </tr>")
          If Shaded = "shaded" Then
            Shaded = ""
          Else
            Shaded = "shaded"
          End If
          i = i + 1
        End While
      %>
    </tbody>
  </table>
</div>
<%

done:
Finally
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
End Try
Send_BodyEnd("searchform", "searchterms")
%>
