<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-cashier-report.aspx 
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
  ' * Version : 5.10b1.0 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dt, rstItems As DataTable
  Dim rst As DataTable
  Dim sizeOfData As Integer = -1
  Dim i As Integer = 0
  Dim x As Integer = 0
  Dim idProgramID As String = ""
  Dim idSearchText As String = ""
  Dim idTimeStart As String = ""
  Dim idTimeStop As String = ""
  Dim idGroupID As String = ""
  Dim GroupID As String = ""
  Dim GroupText As String = ""
  Dim PageNum As Integer = 0
  Dim MorePages As Boolean
  Dim linesPerPage As Integer = 20
  Dim Shaded As String = "shaded"
  Dim restrictLinks As Boolean = False
  Dim extraLink As String = ""
  Dim InfoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim sSearchQuery As String = ""
  Dim ShowResults As Boolean = False
  
  Dim WhereClause As String = ""
  Dim WhereBuf As New StringBuilder()
  Dim AdvSearchSQL As String = ""
  Dim CriteriaMsg As String = ""
  Dim CriteriaTokens As String = ""
  Dim CriteriaError As Boolean = False
    
  Dim form_ProdStartdate As String
  Dim form_ProdEnddate As String
  Dim ProdStartDate As String
  Dim ProdEnddate As String

  Dim sDateOnlyFormat As String = "MM/dd/yyyy"
  Dim sHourOnlyFormat As String = "HH"
  Dim sMinutesOnlyFormat As String = "mm"
  Dim tempDateTime As Date
  Dim bUseTemplateLocks As Boolean
  Dim Disallow_ProductionDates As Boolean = True

  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
 
  
  Response.Expires = 0
    MyCommon.AppName = "CM-fuel-partner-report.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  PageNum = Request.QueryString("pagenum")
  If PageNum < 0 Then PageNum = 0
  MorePages = False
  
    ' lets check the logged in user and see if they are to be restricted to this page
    MyCommon.QueryStr = "select AU.StartPageID,AUSP.PageName ,AUSP.prestrict from AdminUsers as AU with (NoLock) " & _
                        "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                        "where AU.AdminUserID=" & AdminUserID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        If (MyCommon.NZ(rst.Rows(0).Item("prestrict"), False) = True) Then
            ' ok we got in here then we need to restrict the user from seeing any other pages
            restrictLinks = True
        End If
    End If
  
 
    Send_HeadBegin("term.customer", "term.history")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts(New String() {"datePicker.js"})
    If restrictLinks Then
        Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
    End If
    Send_Scripts()
%>
<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
    If Not restrictLinks Then
        Send_Tabs(Logix, 5)
    End If

    Send_Subtabs(Logix, 50, 7)
  
  If (Logix.UserRoles.ViewHistory = False) Then
    Send_Denied(1, "perm.admin-history")
    GoTo done
    End If
    
    Dim SortText As String = ""
    Dim SortDirection As String = ""
    Dim reportAdjustments as Boolean = (MyCommon.Fetch_CM_SystemOption(69) = "1")
   
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
	SortText = Request.QueryString("SortText")
    dim dt2 as DataTable
    
    Dim tempDate As Date
    
    If (Request.QueryString("SearchTimeStart") <> "") Then
        'remove any time string currently on here
        idTimeStart = Request.QueryString("SearchTimeStart")
        Logix.ToShortDateString(idTimeStart, MyCommon)
        tempDate = Date.Parse(idTimeStart)
        If reportAdjustments = False Then
            ' idTimeStart = idTimeStart & " 23:00"
             ' tempDate = tempDate.AddDays(-1)
            tempDate = tempDate.AddHours(-1)
          
        End If
        
        'tempDate = tempDate.AddDays(-1)
        idTimeStart = tempDate.ToString() 
    End If
    
    If (Request.QueryString("SearchTimeStop") <> "") Then
        idTimeStop = Request.QueryString("SearchTimeStop")
        tempDate = Date.Parse(idTimeStop)
        tempDate = tempDate.AddDays(1)
        idTimeStop = tempDate.ToString()
    End If
    
    If (Request.QueryString("SortText") <> "") Then
        idGroupID = Request.QueryString("SortText")
    End If
    
	If reportAdjustments = False Then
      Select Case idGroupID
        Case "1"
            GroupText = "MONTH(earneddate) as groupval"
            GroupID = "MONTH(earneddate) order by month(earneddate)"
        Case "2"
            GroupText = "DATEPART(week,earneddate) as groupval"
            GroupID = "DATEPART(week,earneddate) order by datepart(week,earneddate)"
        Case "3"
            GroupText = " Convert(DateTime, Convert(VarChar, earneddate, 101)) as groupval"
            GroupID = " Convert(DateTime, Convert(VarChar, earneddate, 101)) order by  Convert(DateTime, Convert(VarChar, earneddate, 101))"
      End Select
	End IF
	
    If (Request.QueryString("ProgramID") <> "") Then
        idProgramID = Request.QueryString("ProgramID")
        ShowResults = True
    End If
  
    If (ShowResults) Then
      If reportAdjustments = False Then
	    sSearchQuery = "select  SVS.Description,count(SVProgramID) as numstatus, SUM(qtyearned * value) as maxval,  " & _
                      GroupText & " from SVHistory as SV with (NoLock), StoredValueStatus as SVS with (NoLock) " & _
                     "where SV.StatusFlag = SVS.StatusID and SV.SVProgramID = " & idProgramID & " and Convert(DateTime, Convert(VarChar, SV.earneddate, 101)) > '" & idTimeStart & "' and SV.EarnedDate < '" & idTimeStop & "' group by SVS.Description," & GroupID & _
                      " " & SortDirection
	  Else
	    sSearchQuery = "Select isnull(SVTP.LastUpdate,getdate()) 'DateTime',SVTP.TransactionID 'TransactionNumber',SVTP.ExtLocationCode 'StoreID',SVTP.PartnerID," & _
		               "C.InitialCardID 'CardID' , C.InitialCardTypeID 'CardType', SVTP.SVProgramID 'SVProgramID', SVTP.RewardAmtRedeemed 'RewardAmt', SVTP.UnitsPumped 'UnitsPumped' " & _
		               "from StoredValueThirdPartyTransactions SVTP inner join Customers C on SVTP.CustomerPK = C.CustomerPK "					   
		
		If idProgramID > 0 Then
		  sSearchQuery = sSearchQuery & " where SVTP.SVProgramID = " & idProgramID.ToString & _
		                 " and datediff(dd,SVTP.LastUpdate,'" & Request.QueryString("SearchTimeStart") & "')<=0 and  " & _
						 " datediff(dd,SVTP.LastUpdate,'" & Request.QueryString("SearchTimeStop") & "')>=0 " 
		Else
		  sSearchQuery = sSearchQuery & " where datediff(dd,SVTP.LastUpdate,'" & Request.QueryString("SearchTimeStart") & "')<=0 and  " & _
						 " datediff(dd,SVTP.LastUpdate,'" & Request.QueryString("SearchTimeStop") & "')>=0 "
		End If
		If idGroupID = "1" then 
		  idGroupID = "CardID"
		End If
        sSearchQuery = sSearchQuery & " order by '" & idGroupID & "' " & SortDirection 	    
	  End If
        
        MyCommon.QueryStr = sSearchQuery
        dt = MyCommon.LXS_Select

        sizeOfData = dt.Rows.Count()
        
        If (sizeOfData = 0) Then
            ShowResults = False
        End If
        
    End If
    
    tempDateTime = Now()
    If (idTimeStart <> "") Then
        tempDate = Date.Parse(idTimeStart)
        If reportAdjustments = False Then
            ' idTimeStart = idTimeStart & " 23:00"
            tempDate = tempDate.AddHours(1)
            'tempDate = tempDate.AddDays(1)
        End If
        
        idTimeStart = tempDate.ToString()
        tempDate = Date.Parse(idTimeStop)
        tempDate = tempDate.AddDays(-1)
        idTimeStop = tempDate.ToString()
        form_ProdStartdate = MyCommon.NZ(idTimeStart, Date.Now.ToShortDateString)
        form_ProdEnddate = MyCommon.NZ(idTimeStop, Date.Now.ToShortDateString)
        ProdEnddate = idTimeStop
        ProdStartDate = idTimeStart
    Else
        idTimeStart = tempDateTime.ToString()
        idTimeStop = tempDateTime.ToString()
        form_ProdStartdate = MyCommon.NZ(Request.QueryString("form_TestStartDate"), Date.Now.ToShortDateString)
        form_ProdEnddate = MyCommon.NZ(Request.QueryString("form_TestEnddate"), Date.Now.ToShortDateString)
        ProdEnddate = tempDateTime.ToString(sDateOnlyFormat)
        ProdStartDate = tempDateTime.ToString(sDateOnlyFormat)
    End If
 
  
    If (Request.QueryString("excel") <> "") Then
	  If reportAdjustments = True Then
	     InfoMessage = ExportListToExcel_Adjustments(dt, MyCommon, Logix, idGroupID)
      Else
	    InfoMessage = ExportListToExcel(dt, MyCommon, Logix, idGroupID)
	  End If		
        
        If InfoMessage = "" Then
            GoTo done
        End If
    End If
    Dim dtFuelSV as DataTable
	Dim FuelSVstring as String = "-1"
	If reportAdjustments = False Then
	  sSearchQuery = "SELECT SVProgramID ,Name from StoredValuePrograms"
	Else
	 'Having an UNION on the StoredValueConversion and StoredValueThirdPartyTransaction to get the list of fuel partner programs
	 'to have all fuel partner programs irrespective of their transaction history
	 MyCommon.QueryStr = "Select SVProgramID from StoredValuePointsConversion With (NoLock) "  & _
	                     "Union Select SVProgramID from StoredValueThirdPartyTransactions With (NoLock)"
	 dtFuelSV = MyCommon.LXS_Select
	 
	 If dtFuelSV.Rows.Count > 0 Then
	   For indexCnt as Integer = 0 to dtFuelSV.Rows.Count - 1
	     FuelSVstring = FuelSVstring & "," & dtFuelSV.Rows(indexCnt)(0).ToString()
	   Next
	 End If
	 sSearchQuery = "SELECT 0 SVProgramID, '0 - All' 'Name' Union " & _
	                 "SELECT SVProgramID ,convert(varchar(10),SVProgramID) + ' - ' + Name 'Name' from StoredValuePrograms " & _
                     "Where SVProgramID in (" & FuelSVstring &  ") order by SVProgramID"
	End If
    MyCommon.QueryStr = sSearchQuery
    rstItems = MyCommon.LRT_Select
        
	If reportAdjustments = False Then
      sSearchQuery = "select SVProgramID  " & _
                   "from StoredValuePrograms as SV with (NoLock) " & _
                   "where SV.FuelPartner=1 and Deleted = 0 ;" 
    Else
      sSearchQuery = "Select 0 'SVProgramID' Union Select SVProgramID  " & _
                     "from StoredValuePrograms with (NoLock) " & _
                     " Where SVProgramID in (" & FuelSVstring &  ") order by SVProgramID"
	End IF
    MyCommon.QueryStr = sSearchQuery
    dt2 = MyCommon.LRT_Select
    
    Dim filterExp As String = "SVProgramID in ("
    Dim drarray() As DataRow = nothing
    Dim NumRows = 0
    Dim filterRow As DataRow
    Dim NumPrograms as Integer = 0
    
    i = 0
    NumPrograms = dt2.Rows.Count()
    If (NumPrograms > 0 ) Then
        For Each filterRow In dt2.Rows
            filterExp = filterExp & dt2.rows(i).item("SVProgramID")
            if ( i <> NumPrograms) then
                filterExp = filterExp & ","
            End If
            i = i + 1
        Next
        
        filterExp = filterExp & ")"
        drarray = rstItems.Select(filterExp, "", DataViewRowState.CurrentRows)
        
        If (drarray.Length = 1) then
            NumPrograms = 1
        Else
            NumPrograms = drarray.Length
        End If
    End If
    i = linesPerPage * PageNum
%>

<div id="intro">
  <h1 id="title">
    <% If MyCommon.Fetch_CM_SystemOption(69) = "1" Then %> 
      <% Sendb(Copient.PhraseLib.Lookup("term.fuelpartneradjreport", LanguageID))%>
	<% Else %>
	  <% Sendb(Copient.PhraseLib.Lookup("term.FuelPartnerRprt", LanguageID))%>
    <% End IF %>
  </h1>
  <div id="controls">
    <%
       If (ShowResults) Then
            If dt.Rows.Count > 0 Then
                Send_ExportToExcel()
            End If
        End If	    
    %>
  </div>
</div>

<div id="main">
<% If MyCommon.Fetch_CM_SystemOption(69) <> "1" Then  %>  
 <% if (NumPrograms < 1) then Sendb(Copient.PhraseLib.Lookup("term.nofpprograms", LanguageID)) %>
 <span style="width: 75%" >
  <div id="column" style=" <% if (NumPrograms < 1) then Sendb("visibility:hidden;") %>">
    <div class="box" id="period" >
       <h2 style="overflow: hidden; <% if (NumPrograms < 1) then Sendb("visibility:hidden;") %>">
         <% Sendb(Copient.PhraseLib.Lookup("term.searchby", LanguageID))%>
       </h2>
      <table style="overflow: hidden; <% if (NumPrograms < 1) then Sendb("visibility:hidden;") %>">
        <tr id="labels">
           <td>
                <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID) & ": ")%>
            </td> 
            <td>
                <% Sendb(Copient.PhraseLib.Lookup("term.enter-date", LanguageID))        %>
            </td>
            <td>Select report grouping: </td>
            <td></td>
         </tr>
         <tr id = "searchControls">
             <td>
             <div id="svprogram" style="overflow: hidden; <% if (NumPrograms < 1) then Sendb("visibility:hidden;") %>">
                <select id="programs" name="form_ProgramID">
                       <% 
                           Dim descriptionItem As String
                         
                           For i = 0 to NumPrograms - 1
                               descriptionItem = MyCommon.NZ(drarray(i).Item("Name"), " ") & "     "
                               If (idProgramID <> "" And drarray(i).Item("SVProgramID").ToString() = idProgramID) Then
                                   Send("     <option value=""" & drarray(i).Item("SVProgramID") & """selected>" & descriptionItem & "</option>")
                               Else
                                   Send("     <option value=""" & drarray(i).Item("SVProgramID") & """>" & descriptionItem & "</option>")
                               End If
 
                           Next
                           i = 0
                       %>
                  </select>
                  </div>
              </td>
              <td><span>
                 <div id="datepicker"></div>
                    <input class="short" id="prod-start-date" name="form_ProdStartDate" maxlength="10" type="text" value="<% sendb(Logix.ToShortDateString(ProdStartDate,MyCommon)) %>"<% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
                    <img src="../images/calendar.png" class="calendar" id="prod-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_ProdStartDate', event);" />
                    <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
                    <input class="short" id="prod-end-date" name="form_ProdEndDate" maxlength="10" type="text" value="<% sendb(Logix.ToShortDateString(ProdEndDate,MyCommon)) %>" />
                    <img src="../images/calendar.png" class="calendar" id="prod-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_ProdEndDate', event);" />
                  <%
                      If Request.Browser.Type = "IE6" Then
                          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
                      End If
                  %>
                  </span>
              </td>
               <td>
               <div id="grouping">
                <select id="groupBy" name="form_formatID" style="overflow: hidden;">
                       <% 
                           If (idGroupID <> "" And idGroupID = "1") Then
                               Send("     <option value=""" & 1 & """ selected>" & Copient.PhraseLib.Lookup("term.groupbymonth", LanguageID) & "</option>")
                           Else
                               Send("     <option value=""" & 1 & """>" & Copient.PhraseLib.Lookup("term.groupbymonth", LanguageID) & "</option>")
                           End If
                           
                           If (idGroupID <> "" And idGroupID = "2") Then
                               Send("     <option value=""" & 2 & """ selected>" & Copient.PhraseLib.Lookup("term.groupbyweek", LanguageID) & "</option>")
                           Else
                               Send("     <option value=""" & 2 & """>" & Copient.PhraseLib.Lookup("term.groupbyweek", LanguageID) & "</option>")
                           End If
                           
                           If (idGroupID <> "" And idGroupID = "3") Then
                               Send("     <option value=""" & 3 & """ selected>" & Copient.PhraseLib.Lookup("term.groupbydate", LanguageID) & "</option>")
                           Else
                               Send("     <option value=""" & 3 & """>" & Copient.PhraseLib.Lookup("term.groupbydate", LanguageID) & "</option>")
                           End If
                           
                           
                       %>
                  </select>
                  </div>
              </td>
              <td>
                  <input type="submit" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.search", LanguageID)) %>" onclick="searchNew(0);" />
              </td>
         </tr>
           
        </table>
      </div>
     </span>
</div>
 <table class="list" summary="" style="overflow: hidden; <% if (NumPrograms < 1) then Sendb("visibility:hidden;") %>">
    <thead>
     <tr>
             <% If NumPrograms > 0 Then
                    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;ProgramID=" & idProgramID & "&amp;SortText=" & idGroupID & "&amp;SortDirection=" & SortDirection & "&amp;SearchTimeStart=" & idTimeStart & "&amp;SearchTimeStop=" & idTimeStop, "", "", True)
                End If

               %>
            </tr>
      <tr>
        <th align="left" scope="col" class="th-cashiertimedate">
          <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=<% Sendb(idGroupID)%>&amp;GroupID=<%Sendb(idGroupID)%>&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.datetime", LanguageID))%>
          </a>
          <%
              If SortDirection = "ASC" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              Else
              End If
          %>
        </th>
         <th align="left" scope="col" class="th-cashier">
         
        </th>
        <th align="left" scope="col" class="th-cashier">
             <% Sendb(Copient.PhraseLib.Lookup("cm-fuelrepo.count", LanguageID))%>
      
        </th>
        <th align="left" scope="col" class="th-store">
          <% Sendb(Copient.PhraseLib.Lookup("term.maxrewardvalue", LanguageID))%>
        </th>
      </tr>
    </thead>
    <tbody>
      <%
          Dim groupvalLookup As Integer = 0
          Dim lastmonth As String = ""
          Dim thismonth As String = ""
          Dim lastweek As string = ""
          Dim thisweek As String = ""
          Dim groupFormated As String = ""
          Dim myDATE As string
          dim specificDate as Date
         
          
          While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
             If (idGroupID = "1") Then
                  thismonth = GetGroupFormat(dt.Rows(i).Item("groupval"), Convert.ToInt32(idGroupID))
                  If (lastmonth <> thismonth ) Then
                      if (lastmonth <> "") then
                          If Shaded = "shaded" Then
                              Shaded = ""
                          Else
                              Shaded = "shaded"
                          End If
                      end if
                      Send("<tr class=""" & Shaded & """>")
                      Send("  <td>" & thismonth & "</td>")
                  Else
                      Send("<tr class=""" & Shaded & """>")
                      Send("  <td>     </td>")
                  End If
                  lastmonth = thismonth
              ElseIf (idGroupID = "2") Then
                  
                  groupvalLookup = Convert.ToInt32(dt.Rows(i).Item("groupval"))
                  
                  thisweek = GetGroupFormat(dt.Rows(i).Item("groupval"), Convert.ToInt32(idGroupID))
                  
                  myDATE = thisweek
                  
                  If (myDATE <> lastweek) Then
                      if (lastweek <> "") then
                          If Shaded = "shaded" Then
                              Shaded = ""
                          Else
                              Shaded = "shaded"
                          End If
                      end if
                      Send("<tr class=""" & Shaded & """>")
                      Send("  <td>" & myDate & "</td>")
                  Else
                      Send("<tr class=""" & Shaded & """>")
                      Send("  <td>     </td>")
                  End If
                  
                  lastweek = myDATE
                  
              Else
                  Send("<tr class=""" & Shaded & """>")
                  specificDate = MyCommon.NZ(dt.Rows(i).Item("groupval"), "")
                  Send("  <td>" & specificDate.ToShortDateString() & "</td>")
                  If Shaded = "shaded" Then
                      Shaded = ""
                  Else
                      Shaded = "shaded"
                  End If
              End If
              
              Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("Description"), "") & "</td>")
              Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("numstatus"), "") & "</td>")
              Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("maxval"), "") & "</td>")
              Send("</tr>")
              
              i = i + 1
          End While
          If (sizeOfData = 0) Then
              Send("<tr>")
              Send("  <td colspan=""3""></td>")
              Send("</tr>")
          End If
      %>
    </tbody>
  </table>
<% Else %>
 <% if (NumPrograms < 2) then Sendb(Copient.PhraseLib.Lookup("term.nofpprograms", LanguageID)) %>
 <span style="width: 75%" >
  <div id="column" style=" <% if (NumPrograms < 2) then Sendb("visibility:hidden;") %>">
    <div class="box" id="period" >
       <h2 style="overflow: hidden; <% if (NumPrograms < 2) then Sendb("visibility:hidden;") %>">
         <% Sendb(Copient.PhraseLib.Lookup("term.searchby", LanguageID))%>
       </h2>
      <table style="overflow: hidden; <% if (NumPrograms < 2) then Sendb("visibility:hidden;") %>">
        <tr id="labels">
           <td>
                <% Sendb(Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID) & ": ")%>
            </td> 
            <td>
                <% Sendb(Copient.PhraseLib.Lookup("term.enter-date", LanguageID))        %>
            </td>
			<td></td>
			<td></td>
         </tr>
         <tr id = "searchControls">
            <td>   <div id="svprogram" style="overflow: hidden; <% if (NumPrograms < 2) then Sendb("visibility:hidden;") %>">
                <select id="programs" name="form_ProgramID">
                       <% 
                           Dim descriptionItem As String
                         
                           For i = 0 to NumPrograms - 1
                               descriptionItem = MyCommon.NZ(drarray(i).Item("Name"), " ") & "     "
                               If (idProgramID <> "" And drarray(i).Item("SVProgramID").ToString() = idProgramID) Then
                                   Send("     <option value=""" & drarray(i).Item("SVProgramID") & """selected>" & descriptionItem & "</option>")
                               Else
                                   Send("     <option value=""" & drarray(i).Item("SVProgramID") & """>" & descriptionItem & "</option>")
                               End If
 
                           Next
                           i = 0
                       %>
                  </select>
                  </div>
              </td>
              <td><span>
                 <div id="datepicker"></div>
                    <input class="short" id="prod-start-date" name="form_ProdStartDate" maxlength="10" type="text" value="<% sendb(Logix.ToShortDateString(ProdStartDate,MyCommon)) %>"<% if(bUseTemplateLocks and disallow_productiondates)then sendb(" disabled=""disabled""") %> />
                    <img src="../images/calendar.png" class="calendar" id="prod-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_ProdStartDate', event);" />
                    <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
                    <input class="short" id="prod-end-date" name="form_ProdEndDate" maxlength="10" type="text" value="<% sendb(Logix.ToShortDateString(ProdEndDate,MyCommon)) %>" />
                    <img src="../images/calendar.png" class="calendar" id="prod-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('form_ProdEndDate', event);" />
                  <%
                      If Request.Browser.Type = "IE6" Then
                          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
                      End If
                  %>
                  </span>
              </td>
              <td>
                  <input type="submit" class="regular" id="search" name="search" value="<% Sendb(Copient.PhraseLib.Lookup("term.generate", LanguageID)) %>" onclick="searchNew(1);" />
              </td>
			  <td></td>
         </tr>
        </table>
      </div>
     </span>
</div>
 <table class="list" summary="" style="overflow: hidden; <% if (sizeOfData = -1) then Sendb("visibility:hidden;") %>">
    <thead>
     <tr>
             <% If NumPrograms > 0  and sizeOfData > -1 Then
                    Send_Listbar(linesPerPage, sizeOfData, PageNum, Request.QueryString("searchterms"), "&amp;ProgramID=" & idProgramID & "&amp;SortText=" & idGroupID & "&amp;SortDirection=" & SortDirection & "&amp;SearchTimeStart=" & idTimeStart & "&amp;SearchTimeStop=" & idTimeStop, "", "", True)
                End If
               %>
     </tr>
     <tr>
	 <th align="left" scope="col" class="th-cashier">
		  <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=DateTime&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.datetime", LanguageID))%>
          </a>
          <%
              If SortText = "DateTime" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>		
	 <th align="left" scope="col" class="th-cashier">
		  <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=TransactionNumber&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
           <% Sendb(Copient.PhraseLib.Lookup("term.transactionnumber", LanguageID))%>
          </a>
          <%
              If SortText = "TransactionNumber" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>
	 <th align="left" scope="col" class="th-cashier">
		  <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=StoreId&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.storeid", LanguageID))%>
          </a>
          <%
              If SortText = "StoreId" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>	
	 <th align="left" scope="col" class="th-cashier">
		  <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=PartnerID&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.partnerID", LanguageID))%>
          </a>
          <%
              If SortText = "PartnerID" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>			
     <th align="left" scope="col" class="th-cashier">
          <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=CardID&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.cardid", LanguageID))%>
          </a>
          <%
              If SortText = "CardID" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>
		<th align="left" scope="col" class="th-cashier">
		<a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=CardType&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.cardtype", LanguageID))%>
          </a>
          <%
              If SortText = "CardType" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>

		<th align="left" scope="col" class="th-cashier">
		<a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=SVProgramID&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.svprogramid", LanguageID))%>
          </a>
          <%
              If SortText = "SVProgramID" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>

		<th align="left" scope="col" class="th-cashier">
		  <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=RewardAmt&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.rewardamount", LanguageID))%>
          </a>
          <%
              If SortText = "RewardAmt" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>

		<th align="left" scope="col" class="th-cashier">
		  <a href="CM-fuel-partner-report.aspx?ProgramID=<%sendb(idProgramID) %>&amp;SearchTimeStart=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStart),MyCommon))%>&amp;SearchTimeStop=<%Sendb(Logix.ToShortDateString(date.parse(idTimeStop),MyCommon))%>&amp;SortText=UnitsPumped&amp;GroupID=1&amp;SortDirection=<%Sendb(SortDirection)%>">  
            <% Sendb(Copient.PhraseLib.Lookup("term.unitspumped", LanguageID))%>
          </a>
          <%
              If SortText = "UnitsPumped" Then
                  If SortDirection = "ASC" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                  Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                  End If
              End If
          %>
        </th>
      </tr>
    </thead>
    <tbody>
	      <%
          Dim groupvalLookup As Integer = 0
          dim rst_SV as DataTable
          dim svProgramId as Integer = 1
          dim prevSVProgramId as Integer = -1
          dim cardTypeId as Integer = 0
          dim prevcardTypeId as Integer = -1
		  dim cardType as String = ""
          dim SvName as String = ""
		  i = linesPerPage * PageNum
		  
	      While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum) 
		    Send("<tr class=""" & Shaded & """>")
			Send("  <td>" & MyCommon.NZ(dt.Rows(i)(0), "") & "</td>")
			Send("  <td>" & MyCommon.NZ(dt.Rows(i)(1), "") & "</td>")
			Send("  <td>" & MyCommon.NZ(dt.Rows(i)(2), "") & "</td>")
			Send("  <td>" & MyCommon.NZ(dt.Rows(i)(3), "") & "</td>")
			Send("  <td>" & MyCommon.NZ(dt.Rows(i)(4), "") & "</td>")
			cardTypeId = dt.Rows(i)(5)
			If cardTypeId <> prevcardTypeId Then
			  MyCommon.QueryStr = "select Description from cardtypes where cardtypeID= " & dt.Rows(i)(5)
			  rst_SV = MyCommon.LXS_Select
			  If rst_SV.Rows.Count > 0 Then
			    cardType = rst_SV.Rows(0)(0)
			    Send("  <td>" & cardType & "</td>")
			  Else
			     Send("  <td>" & MyCommon.NZ(dt.Rows(i)(5), "CardType") & "</td>")
			  End IF
			Else
			  Send("  <td>" & cardType & "</td>")
			End If
			prevcardTypeId = cardTypeId
			'Send("  <td>" & MyCommon.NZ(dt.Rows(i)(5), "CardType") & "</td>")
			svProgramId = dt.Rows(i)(6)
			If svProgramId <> prevSVProgramId Then
			  MyCommon.QueryStr = "SELECT  Name from StoredValuePrograms where Deleted =0 and SVProgramID = " & svProgramId
			  rst_SV = MyCommon.LRT_Select
			  If rst_SV.Rows.Count > 0 Then
			    SvName = rst_SV.Rows(0)(0)
			    Send("  <td>" &  svProgramId & " - " & SvName  & "</td>")
			  Else
			    Send("  <td>" &  svProgramId & " - "  & "</td>")
			  End If
			Else
               Send("  <td>" &  svProgramId & " - " & SvName  & "</td>")
			End If
			prevSVProgramId = svProgramId
			'Send("  <td>" & MyCommon.NZ(dt.Rows(i).Item("SVProgramID"), "") & "</td>")
		    Send("  <td>" & MyCommon.NZ(dt.Rows(i)(7), "") & "</td>")
            Send("  <td>" & MyCommon.NZ(dt.Rows(i)(8), "") & "</td>")
            Send("</tr>")
            i = i + 1
		  End While
          If (sizeOfData = 0) Then
              Send("<tr>")
              Send("  <td colspan=""3""></td>")
              Send("</tr>")
          End If
      %>
    </tbody>
  </table>
<% End If %>
</div>

<script type="text/javascript">  
 
  function handleExcel() {
    var sUrl = document.getElementById("ExcelUrl");
    var form = document.forms['excelform'];
    
    form.action = sUrl.value;
    form.method = "Post";
    form.submit();
  }

    window.name="fuel-repo"
    var datePickerDivID = "datepicker";
    
    <% Send_Calendar_Overrides(MyCommon) %>

    <% If (Logix.UserRoles.EditOffer) Then %>
    window.onunload= function(){
        handleNavAway(document.mainform)
    };
    <% End If %>
    
    function disableUnload() {
        window.onunload = null;
    }
    
    // callback function for save changes on unload during navigate away
    function handleAutoFormSubmit() {
        elmName();
        if (ValidateOfferForm('<% Sendb(MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern) %>') && ValidateTimes()) {
            document.mainform.submit();
        }
    }
    
    function elmName(){
        window.onunload = null;
        for(i=0; i<document.mainform.elements.length; i++)
        {
            document.mainform.elements[i].disabled=false;
            //alert(document.mainform.elements[i].name)
        }
        return true;
    }
    
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
    
    function handlePageClick(e) {
      var calFrame = document.getElementById('calendariframe');
      var el=(typeof event!=='undefined')? event.srcElement : e.target        
      
      if (el != null) {
        var pickerDiv = document.getElementById(datePickerDivID);
        if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
          if (el.id!="prod-start-picker" && el.id!="prod-end-picker" && el.id!="test-start-picker" && el.id!="test-end-picker") {
            if (!isDatePickerControl(el.className)) {
              pickerDiv.style.visibility = "hidden";
              pickerDiv.style.display = "none";  
              if (calFrame != null) {
                calFrame.style.visibility = 'hidden';
                calFrame.style.display = 'none';
              }
            }
          } else  {
              pickerDiv.style.visibility = "visible";            
              pickerDiv.style.display = "block";            
              if (calFrame != null) {
                calFrame.style.visibility = 'visible';
                calFrame.style.display = 'block';
              }
          }
        }
      }
    }

    function isDatePickerControl(ctrlClass) {
      var retVal = false;
      
      if (ctrlClass != null && ctrlClass.length >= 2) {
        if (ctrlClass.substring(0,2) == "dp") {
          retVal = true;
        }
      }

      return retVal;
    }
    
   function searchNew(pvalue) {
     var currentDate = new Date()
     var day = currentDate.getDate()
     var month =  currentDate.getMonth() + 1 
     var year = currentDate.getFullYear()
     if (month < 10) {
       month = '0'+ month
     }  
     if (day < 10) {
       day = '0'+ day 
     } 
     var startTime = document.getElementById('prod-start-date');
	 if (!validDateFormat(startTime.value))
	 {
	   startTime.value = month + "/" + day + "/" + year;
	 }
     var stopTime = document.getElementById('prod-end-date');
	 if (!validDateFormat(stopTime.value))
	 {
	   stopTime.value = month + "/" + day + "/" + year;
	 }     
  
     var programID = document.getElementById('programs');
     if (pvalue == 0)
	 {
       var formatID = document.getElementById('groupBy');
       window.location = 'CM-fuel-partner-report.aspx?ProgramID=' + programID.value + '&SearchTimeStart=' + startTime.value +  '&SearchTimeStop=' + stopTime.value + '&SortText=' + formatID.value + '&SortDirection=ASC';
     }
	else
	{
       var formatID = 'CardID';
       window.location = 'CM-fuel-partner-report.aspx?ProgramID=' + programID.value + '&SearchTimeStart=' + startTime.value +  '&SearchTimeStop=' + stopTime.value + '&SortText=' + formatID + '&SortDirection=ASC';
	}
  }
function validDateFormat(pvalue)
{
var retVal = false;
var date_month = 0;
var date_day = 1;
var dateArray = new Array();
dateArray = pvalue.split('/');
if (dateArray.length == 3) {
 if (IsNumeric(dateArray[0])) {
    date_month = parseFloat(dateArray[0]);
   if (date_month > 0 && date_month <= 12) {
     if (IsNumeric(dateArray[1]) == true) {
       date_day = parseFloat(dateArray[1]);
       if ((date_day > 0 && date_day <= 31 && (date_month==1 || date_month==3 || date_month==5 || date_month==7 || date_month==8 || date_month==10 || date_month==12)) || (date_day > 0 && date_day <= 30 && (date_month==4 || date_month==6 || date_month==9 || date_month==11))  || (date_day > 0 && date_day <= 28 && date_month==2)) {
         if (IsNumeric(dateArray[2]) == true) {
           retVal = true;
         }  
       }
     }
   }
 }
}
return retVal;
}

function IsNumeric(val) {
    if (isNaN(parseFloat(val))) {
          return false;
     }
     return true
} 
</script>

<script runat="server">
    
    Private Function GetGroupFormat(ByVal groupVal As String, ByVal groupType As Integer)
        
        Dim newVal As String = ""
        Dim groupLookup
        Dim myDATE,myEOW As Date
        Dim myYear, myWeek As Integer
        
        groupLookup = Convert.ToInt32(groupVal)
        
        If (groupType = 1) Then
            
            Select Case groupLookup
                Case 1
                    newVal = Copient.PhraseLib.Lookup("term.january", LanguageID)
                Case 2
                    newVal = Copient.PhraseLib.Lookup("term.february", LanguageID)
                Case 3
                    newVal = Copient.PhraseLib.Lookup("term.march", LanguageID)
                Case 4
                    newVal = Copient.PhraseLib.Lookup("term.april", LanguageID)
                Case 5
                    newVal = Copient.PhraseLib.Lookup("term.may", LanguageID)
                Case 6
                    newVal = Copient.PhraseLib.Lookup("term.june", LanguageID)
                Case 7
                    newVal = Copient.PhraseLib.Lookup("term.july", LanguageID)
                Case 8
                    newVal = Copient.PhraseLib.Lookup("term.august", LanguageID)
                Case 9
                    newVal = Copient.PhraseLib.Lookup("term.september", LanguageID)
                Case 10
                    newVal = Copient.PhraseLib.Lookup("term.october", LanguageID)
                Case 11
                    newVal = Copient.PhraseLib.Lookup("term.november", LanguageID)
                Case 12
                    newVal = Copient.PhraseLib.Lookup("term.december", LanguageID)
            End Select
        Else
            
            myYear = Now().Year
            myWeek = groupLookup
                  
            myDATE = Date.Parse("1.1." & myYear.ToString)
            myDATE = myDATE.Date.AddDays(7 * (myWeek - 1) - myDATE.DayOfWeek)
            myEOW = mydate.AddDays(6)

            newVal = myDATE.ToShortDateString.ToString() & "-" & myEOW.ToShortDateString.ToString()
            
        End If
        
        Return newVal
    End Function
  

    Private Function ExportListToExcel(ByRef dst As DataTable, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc, ByVal groupID As String) As String
        Dim bStatus As Boolean
        Dim sMsg As String = ""
        Dim CmExport As New Copient.ExportXml
        Dim sFileFullPath As String
        Dim sFullPathFileName As String
        Dim sFileName As String = "Fuel_Partner_Report.xls"
        Dim dtExport As DataTable
        Dim dr As DataRow
        Dim drExport As DataRow
        
        Dim groupvalLookup As Integer = 0
        Dim lastmonth As String = ""
        Dim thismonth As String = ""
        Dim lastweek As string = ""
        Dim thisweek As String = ""
        Dim myDATE as String = ""
        Dim groupFormated As String = ""
       
        Dim i As Integer = 0
        Dim newval As String = ""


        If dst.Rows.Count > 0 Then
      
            'select  SVS.Description,count(SVProgramID) as numstatus, SUM(qtyearned * value) as maxval,groupval
            dtExport = New DataTable()
            dtExport.Columns.Add("Time Interval", Type.GetType("System.String"))
            dtExport.Columns.Add("Status", Type.GetType("System.String"))
            dtExport.Columns.Add("Count", Type.GetType("System.String"))
            dtExport.Columns.Add("Reward Total Value", Type.GetType("System.Decimal"))
      
            For Each dr In dst.Rows
                drExport = dtExport.NewRow()
                newval = ""
                If (groupID = "1") Then
                    thismonth = GetGroupFormat(dst.Rows(i).Item("groupval"), Convert.ToInt32(groupID))
                    If (lastmonth <> thismonth) Then
                        newval = thismonth
                    Else
                        newval = "     "
                    End If
                    lastmonth = thismonth
                ElseIf (groupID = 2) Then
                  
                  groupvalLookup = Convert.ToInt32(dst.Rows(i).Item("groupval"))
                  
                  thisweek = GetGroupFormat(dst.Rows(i).Item("groupval"), Convert.ToInt32(groupID))
                  
                  myDATE = thisweek
                  
                  If (myDATE <> lastweek) Then
                      newval =  myDate
                  Else
                      newval = ""
                  End If
                  
                  lastweek = myDATE
                  
                Else
                    newval = MyCommon.NZ(dst.Rows(i).Item("groupval"), "")
                End If
                drExport.Item("Time Interval") = newval
                drExport.Item("Status") = dr.Item("description")
                drExport.Item("Count") = dr.Item("numstatus")
                drExport.Item("Reward Total Value") = Convert.ToDecimal(dr.Item("maxval"))
                dtExport.Rows.Add(drExport)
                i = i + 1
            Next

            sFileFullPath = MyCommon.Fetch_SystemOption(29)
            sFullPathFileName = sFileFullPath & "\" & sFileName

            bStatus = CmExport.ExportToExcel(sFullPathFileName, dtExport)
            If bStatus Then
                Dim oRead As System.IO.StreamReader
                Dim LineIn As String
                Dim Bom As String = ChrW(65279)
                oRead = System.IO.File.OpenText(sFullPathFileName)
                Response.Clear()
                Response.ContentEncoding = Encoding.Unicode
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment; filename=" & sFileName)
        
                'force little endian fffe bytes at front, why?  i dont know but is required.
                Sendb(Bom)
                While oRead.Peek <> -1
                    LineIn = oRead.ReadLine()
                    Send(LineIn)
                End While
                oRead.Close()
                Response.End()
                System.IO.File.Delete(sFullPathFileName)
            Else
                sMsg = CmExport.GetStatusMsg
            End If
        Else
            sMsg = Copient.PhraseLib.Lookup("sv-list.empty", LanguageID)
        End If
  
        Return sMsg
    End Function

    Private Function ExportListToExcel_Adjustments(ByRef dst As DataTable, ByRef MyCommon As Copient.CommonInc, ByRef Logix As Copient.LogixInc, ByVal groupID As String) As String
        Dim bStatus As Boolean
        Dim sMsg As String = ""
        Dim CmExport As New Copient.ExportXml
        Dim sFileFullPath As String
        Dim sFullPathFileName As String
        Dim sFileName As String = Now.ToString("yyyyMMdd_HHmmss") & "_FuelPartnerAdjustmentsReport.csv" '"Fuel_Partner_Adjustments_Report_" & Now.ToString("MMddyyyy_HHmm") & ".csv"
        Dim dtExport As DataTable
        Dim dr As DataRow
        Dim drExport As DataRow
		Dim rst_SV as DataTable
        Dim svProgramId as Integer = 1
        Dim prevSVProgramId as Integer = -1
        Dim cardTypeId as Integer = 0
        Dim prevcardTypeId as Integer = -1			
        Dim cardTypeText as String = ""
		Dim SVnameText as String = ""
       
        If dst.Rows.Count > 0 Then
      
            'select  SVS.Description,count(SVProgramID) as numstatus, SUM(qtyearned * value) as maxval,groupval
            dtExport = New DataTable()
			dtExport.Columns.Add("DateTime", Type.GetType("System.String"))
			dtExport.Columns.Add("TransactionNumber", Type.GetType("System.String"))
			dtExport.Columns.Add("StoreId", Type.GetType("System.String"))
			dtExport.Columns.Add("PartnerID", Type.GetType("System.String"))
            dtExport.Columns.Add("CardID", Type.GetType("System.String"))
            dtExport.Columns.Add("CardType", Type.GetType("System.String"))
            dtExport.Columns.Add("SVProgramID", Type.GetType("System.String"))
            dtExport.Columns.Add("RewardAmt", Type.GetType("System.Decimal"))
			dtExport.Columns.Add("UnitsPumped", Type.GetType("System.Decimal"))

            For Each dr In dst.Rows
              drExport = dtExport.NewRow()
              
                drExport.Item("DateTime") = dr.Item("DateTime")
                drExport.Item("TransactionNumber") = dr.Item("TransactionNumber")
                drExport.Item("StoreId") = dr.Item("StoreId")
                drExport.Item("PartnerID") = dr.Item("PartnerID")

                drExport.Item("CardID") = dr.Item("CardID")
				cardTypeId = dr.Item("CardType")
				If cardTypeId <> prevcardTypeId Then
				  MyCommon.QueryStr = "select Description from cardtypes where cardtypeID= " & dr.Item("CardType")
				  rst_SV = MyCommon.LXS_Select
				  If rst_SV.Rows.Count > 0 Then
				    cardTypeText = rst_SV.Rows(0)(0)
					drExport.Item("CardType") = cardTypeText
				  Else
                    drExport.Item("CardType") = dr.Item("CardType")
                  End IF
				Else
				  drExport.Item("CardType") = cardTypeText
				End If
				prevcardTypeId = cardTypeId
                'drExport.Item("CardType") = dr.Item("CardType")
                svProgramId =dr.Item("SVProgramID")
				If svProgramId <> prevSVProgramId Then
				  MyCommon.QueryStr = "SELECT  Name from StoredValuePrograms where Deleted =0 and SVProgramID = " & svProgramId
				  rst_SV = MyCommon.LRT_Select
				  If rst_SV.Rows.Count > 0 Then
				    SVnameText = rst_SV.Rows(0)(0)
					drExport.Item("SVProgramID") = svProgramId & " - " & SVnameText
				  Else
				    drExport.Item("SVProgramID") = svProgramId
				  End IF
				Else
				  drExport.Item("SVProgramID") = svProgramId & " - " & SVnameText
				End If
                drExport.Item("RewardAmt") = Convert.ToDecimal(dr.Item("RewardAmt"))
				drExport.Item("UnitsPumped") = Convert.ToDecimal(dr.Item("UnitsPumped"))
                dtExport.Rows.Add(drExport)
            Next

            sFileFullPath = MyCommon.Fetch_SystemOption(29)
            sFullPathFileName = sFileFullPath & "\" & sFileName

            'bStatus = CmExport.ExportToExcel(sFullPathFileName, dtExport)
			sFullPathFileName = CreateCSVFile(sFullPathFileName)
			WriteInExportedFile(dtExport,sFullPathFileName)
            If sFullPathFileName <> "" Then
                Dim oRead As System.IO.StreamReader
                Dim LineIn As String
                Dim Bom As String = ChrW(65279)
                oRead = System.IO.File.OpenText(sFullPathFileName)
                Response.Clear()
                Response.ContentEncoding = Encoding.Unicode
                Response.ContentType = "application/octet-stream"
                Response.AddHeader("Content-Disposition", "attachment; filename=" & sFileName)
                'force little endian fffe bytes at front, why?  i dont know but is required.
                Sendb(Bom)
                While oRead.Peek <> -1
                    LineIn = oRead.ReadLine()
                    Send(LineIn)
                End While
                oRead.Close()
                Response.End()
                System.IO.File.Delete(sFullPathFileName)
            Else
                sMsg = CmExport.GetStatusMsg
            End If
		  Else
            sMsg = Copient.PhraseLib.Lookup("sv-list.empty", LanguageID)
        End If
  
        Return sMsg
    End Function
	
  Public Function CreateCSVFile(ByVal sFileName As String) As String
   Try
      Dim fsNewFile As System.IO.FileStream
      fsNewFile = New System.IO.FileStream(sFileName, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
      fsNewFile.Close()
      Return sFileName
    Catch ex As Exception
      sFileName = ""
      Return sFileName
    End Try
  End Function

  'Populating the .csv file in the specified path
  Public Function WriteInExportedFile(ByVal dtExportData As DataTable, ByVal strpath As String) As String
    Dim strMessage As String = ""
    Dim File As System.IO.StreamWriter = New System.IO.StreamWriter(strpath)
    Dim rowscreated As Integer = 0
    Dim ContentString As String = ""
    Dim HeaderString As String = ""
    Try
        For ColmIndex As Integer = 0 To dtExportData.Columns.Count - 1
          If HeaderString = "" Then
            HeaderString = dtExportData.Columns.Item(ColmIndex).ColumnName
          Else
            HeaderString = HeaderString & "," & dtExportData.Columns.Item(ColmIndex).ColumnName
          End If
        Next
        File.WriteLine(HeaderString)
        For Index As Integer = 1 To HeaderString.Length
          If Mid(HeaderString, Index, 1) = "," Then
            File.Write(" ")
          Else
            File.Write("-")
          End If
        Next
        File.WriteLine()
      If dtExportData.Rows.Count > 0 Then
        For rowIndex As Integer = 0 To dtExportData.Rows.Count - 1
          ContentString = ""
          For ColmIndex As Integer = 0 To dtExportData.Columns.Count - 1
            If ContentString = "" Then
              ContentString = dtExportData.Rows(rowIndex)(ColmIndex).ToString
            Else
              ContentString = ContentString & "," & dtExportData.Rows(rowIndex)(ColmIndex).ToString
            End If
          Next
          File.WriteLine(ContentString)
          rowscreated += 1
        Next
      End If
      'File.WriteLine("")
      'File.WriteLine("(" & dtExportData.Rows.Count.ToString & " rows affected)")
      strMessage = ""
    Catch ex As Exception
      strMessage = "Error at Record Number: " & rowscreated & vbNewLine & "Message: " & ex.Message
    Finally
      File.Close()
    End Try
    Return strMessage
  End Function
	
  Function TryParseLocalizedDate(ByVal DateStr As String, ByRef LocalizedDate As Date, ByRef MyCommon As Copient.CommonInc) As Boolean
    Return Date.TryParseExact(DateStr, MyCommon.GetAdminUser.Culture.DateTimeFormat.ShortDatePattern, _
                           MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, LocalizedDate)
  End Function
  
   

</script>

<form id="frmIter" name="frmIter" method="post" action="">
  
</form>


<%
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  MyCommon = Nothing
  Logix = Nothing
%>
