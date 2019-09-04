<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-reports.aspx 
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
  
  Dim MyCommon As New Copient.CommonInc
    Dim MyLookup As New Copient.CustomerLookup
    Dim MyCryptLib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim dtAdj As System.Data.DataTable
  Dim dtOff As System.Data.DataTable
  Dim dt As System.Data.DataTable
  Dim dt2 As System.Data.DataTable
  Dim dt3 As System.Data.DataTable
  Dim row As DataRow
  Dim row2 As DataRow
  Dim row3 As DataRow
  Dim AdminUserID As Long
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim restrictLinks As Boolean = False
  Dim Item As DictionaryEntry
  Dim i As Integer = 0
  Dim Wheres As String = ""
  
  ' Form fields
  Dim Download As Boolean = False
  Dim ReportStart As String = ""
  Dim ReportEnd As String = ""
  Dim R As Boolean = True
  Dim RoleID As Integer = 0
  Dim AdminID As Integer = 0
  Dim OfferOrProgram As Integer = 0
  Dim PP As Boolean = True
  Dim SV As Boolean = True
  Dim A As Boolean = True
  Dim O As Boolean = True
  Dim PPID As Integer = 0
  Dim SVID As Integer = 0
  Dim OID As Integer = 0
  Dim Number As Integer = 0
  Dim NumberOp As Integer = 0
  Dim Value As Decimal = 0
  Dim ValueOp As Integer = 0
  Dim AddO As Boolean = False
  Dim RemO As Boolean = False
  Dim AddHH As Boolean = False
  Dim RemHH As Boolean = False
  
  ' Data that gets returned
  Dim FirstName As String = ""
  Dim LastName As String = ""
  Dim CustomerPK As Integer = 0
  Dim ExtCardID As String = ""
  Dim CustomerFirstName As String = ""
  Dim CustomerMiddleName As String = ""
  Dim CustomerLastName As String = ""
  Dim CustomerType As Integer = 0
  Dim HHPK As Integer = 0
  Dim HHExtCardID As String = ""
  Dim ProgramName As String = ""
  Dim CustomerGroupName As String = ""
  Dim OfferName As String = ""
  Dim SubjectType As Integer = 0
  Dim SubjectID As Integer = 0
  Dim ProgramIDs As New Hashtable()
  Dim CustomerGroupIDs As New Hashtable()
  Dim AdminIDs As New Hashtable()
  Dim AdminIDsString As String = ""
  Dim Adjustment As Decimal = 0
  Dim TotalAdjustments As Integer = 0
  Dim Reason As String = ""
  Dim ReasonID As Integer = 0
  Dim ReasonText As String = ""
  Dim Comments As String = ""
  
  Dim TempString As String = ""
  Dim TempLen As Integer = 0
  Dim TempPos As Integer = 0
  Dim StartDate, EndDate As Date
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-reports.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  
  ' See if the logged-in user should be restricted to this page
  MyCommon.QueryStr = "select AU.StartPageID, AUSP.PageName, AUSP.Prestrict from AdminUsers as AU with (NoLock) " & _
                      "inner join AdminUserStartPages as AUSP with (NoLock) on AU.StartPageID=AUSP.StartPageID " & _
                      "where AU.AdminUserID=" & AdminUserID & ";"
  dt = MyCommon.LRT_Select
  If dt.Rows.Count > 0 Then
    If (MyCommon.NZ(dt.Rows(0).Item("prestrict"), False) = True) Then
      restrictLinks = True
    End If
  End If
  
  If (Request.QueryString("download") <> "") Then
    If (Request.QueryString("reportstart") = "" OrElse Request.QueryString("reportend") = "") Then
      infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.date", LanguageID)
    Else
      If Date.TryParse(GetCgiValue("reportstart"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, StartDate) AndAlso Date.TryParse(GetCgiValue("reportend"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, EndDate) Then
        If StartDate > EndDate Then
          infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.date", LanguageID)
        Else
          Download = True
        End If
      Else
        infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.date", LanguageID)
      End If
    End If
    
    If ((Request.QueryString("value") <> "") AndAlso Not (IsNumeric(Request.QueryString("value"))) OrElse (Request.QueryString("number") <> "") AndAlso Not (IsNumeric(Request.QueryString("number")))) Then
      infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.notanumber", LanguageID)
      Download = False
    End If
    If (MyCommon.Extract_Decimal(Request.QueryString("number"), MyCommon.GetAdminUser.Culture) < 0) Then
      infoMessage = Copient.PhraseLib.Lookup("customer-inquiry.nonegatives", LanguageID)
      Download = False
    End If
  End If
  
  If Download = True Then
    ' For ease of use, assign all the form data to variables
    ReportStart = Request.QueryString("ReportStart")
    ReportEnd = Request.QueryString("ReportEnd")
    R = IIf(Request.QueryString("userorrole") = "1", True, False)
    RoleID = MyCommon.Extract_Val(Request.QueryString("RoleID"))
    AdminID = MyCommon.Extract_Val(Request.QueryString("AdminID"))
    OfferOrProgram = Request.QueryString("OfferOrProgram")
    PP = IIf(Request.QueryString("PP") = "on", True, False)
    SV = IIf(Request.QueryString("SV") = "on", True, False)
    A = IIf(Request.QueryString("A") = "on", True, False)
    AddO = IIf(Request.QueryString("AddO") = "on", True, False)
    RemO = IIf(Request.QueryString("RemO") = "on", True, False)
    AddHH = IIf(Request.QueryString("AddHH") = "on", True, False)
    RemHH = IIf(Request.QueryString("RemHH") = "on", True, False)
    O = IIf(Request.QueryString("offerorprogram") = "1", True, False)
    PPID = MyCommon.Extract_Val(Request.QueryString("PPID"))
    SVID = MyCommon.Extract_Val(Request.QueryString("SVID"))
    OID = MyCommon.Extract_Val(Request.QueryString("OID"))
    Number = MyCommon.Extract_Decimal(Request.QueryString("Number"), MyCommon.GetAdminUser.Culture)
    NumberOp = MyCommon.Extract_Val(Request.QueryString("NumberOp"))
    Value = MyCommon.Extract_Decimal(Request.QueryString("Value"), MyCommon.GetAdminUser.Culture)
    ValueOp = MyCommon.Extract_Val(Request.QueryString("ValueOp"))
    ReasonID = MyCommon.Extract_Val(Request.QueryString("ReasonID"))
    
    ' If a role ID is specified, get all the associated user IDs and put them into a hash table:
    If RoleID > 0 Then
      MyCommon.QueryStr = "select AU.AdminUserID, AUR.RoleID from AdminUsers as AU with (NoLock) " & _
                          "inner join AdminUserRoles as AUR with (NoLock) on AUR.AdminUserID=AU.AdminUserID " & _
                          "where AUR.RoleID=" & RoleID & ";"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        i = 0
        For Each row In dt.Rows
          AdminIDsString &= IIf(i > 0, ",", "") & MyCommon.NZ(row.Item("AdminUserID"), 0)
          i += 1
        Next
      Else
        AdminIDsString = "0"
      End If
    End If
    
    ' If an OfferID is specified, get all the programs associated with it and put them into a hash table
    If OID > 0 Then
      MyCommon.QueryStr = "select OfC.LinkID as OfferProgramID from OfferConditions as OfC with (NoLock) " & _
                          "where OfC.ConditionTypeID = 3 And OfC.OfferID = " & OID & " And Deleted = 0 " & _
                          " UNION " & _
                          "select RP.ProgramID as OfferProgramID from RewardPoints as RP with (NoLock) " & _
                          "inner join OfferRewards as OfR with (NoLock) on OfR.LinkID=RP.RewardPointsID " & _
                          "where OfR.RewardTypeID = 2 And OfR.OfferID = " & OID & " And Deleted = 0 " & _
                          " UNION " & _
                          "select IPG.ProgramID as OfferProgramID from CPE_IncentivePointsGroups as IPG with (NoLock) " & _
                          "left join PointsPrograms as PP with (NoLock) on PP.ProgramID=IPG.ProgramID " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                          "where IPG.Deleted = 0 And RO.IncentiveID = " & OID & " " & _
                          " UNION " & _
                          "select PP.ProgramID as OfferProgramID from PointsPrograms as PP with (NoLock) " & _
                          "inner join CPE_DeliverablePoints as DP with (NoLock) on PP.ProgramID=DP.ProgramID and DP.Deleted=0 and PP.Deleted=0 " & _
                          "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DP.DeliverableID and D.Deleted=0 and D.RewardOptionPhase=3 " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                          "where RO.IncentiveID=" & OID & " " & _
                          " " & _
                          " UNION " & _
                          " " & _
                          "select OfC.LinkID as OfferProgramID from OfferConditions as OfC with (NoLock) " & _
                          "where OfC.ConditionTypeID = 6 And OfC.OfferID = " & OID & " And Deleted = 0 " & _
                          " UNION " & _
                          "select RSV.ProgramID as OfferProgramID from CM_RewardStoredValues as RSV with (NoLock) " & _
                          "inner join OfferRewards as OfR with (NoLock) on OfR.LinkID=RSV.RewardStoredValuesID " & _
                          "where OfR.RewardTypeID = 10 And OfR.OfferID = " & OID & " And Deleted = 0 " & _
                          " UNION " & _
                          "select ISVP.SVProgramID as OfferProgramID from CPE_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                          "left join StoredValuePrograms as SVP with (NoLock) on SVP.SVProgramID=ISVP.SVProgramID " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ISVP.RewardOptionID " & _
                          "where ISVP.Deleted = 0 And RO.IncentiveID = " & OID & " " & _
                          " UNION " & _
                          "select SVP.SVProgramID as OfferProgramID from StoredValuePrograms as SVP with (NoLock) " & _
                          "inner join CPE_DeliverableStoredValue as DSV with (NoLock) on SVP.SVProgramID=DSV.SVProgramID and DSV.Deleted=0 and SVP.Deleted=0 " & _
                          "inner join CPE_Deliverables as D with (NoLock) on D.DeliverableID=DSV.DeliverableID and D.Deleted=0 and D.RewardOptionPhase=3 " & _
                          "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                          "where RO.IncentiveID=" & OID & ";"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        ProgramIDs = New Hashtable(dt.Rows.Count)
        For Each row In dt.Rows
          ProgramIDs.Add(MyCommon.NZ(row.Item("OfferProgramID"), 0), MyCommon.NZ(row.Item("OfferProgramID"), 0))
        Next
      End If
    End If
    
    ' Also if an OfferID is specified, get all the customer groups associated with it and put them into a hash table
    ' (This is in case the user wants to find "add to offer" or "remove from offer" actions.)
    If OID > 0 Then
      MyCommon.QueryStr = "select ICG.CustomerGroupID from CPE_IncentiveCustomerGroups as ICG with (NoLock) " & _
                          "inner join CPE_RewardOptions as RO on RO.RewardOptionID=ICG.RewardOptionID " & _
                          "inner join CPE_Incentives as I on I.IncentiveID=RO.IncentiveID " & _
                          "where I.IncentiveID=" & OID & " and ICG.ExcludedUsers=0 and ICG.Deleted=0;"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        CustomerGroupIDs = New Hashtable(dt.Rows.Count)
        For Each row In dt.Rows
          CustomerGroupIDs.Add(MyCommon.NZ(row.Item("CustomerGroupID"), 0), MyCommon.NZ(row.Item("CustomerGroupID"), 0))
        Next
      End If
    End If
    
    
    ' Now we build up the main queries to dig adjustments and related info out of the activity log.
    
    
    ' ~~~ dtAdj query (the first of two), which will get all the adjustment-related activity ~~~~~~~~~~~~~
    ' First, build a list of "where" clauses, since we'll be sticking them twice into the query:
    Wheres = ""
    ' ...only customer inquiry activity
    Wheres &= " where ActivityTypeID=25 "
    ' ...only within the specified date range
    Wheres &= " and ActivityDate between '" & StartDate.ToString("yyyy-MM-dd") & "T00:00:00' and '" & EndDate.ToString("yyyy-MM-dd") & "T23:59:59' "
    ' ...only for a specific reason?
    If ReasonID > 0 Then
      Wheres &= " and ReasonID=" & ReasonID & " "
    End If
    
    ' ...including points / stored value / accumulation?
    If (PP Or SV Or A) Then
      Wheres &= " and ("
      If PP Then
        Wheres &= "(ActivitySubTypeID=12"
        If (PPID <> 0) Then
          Wheres &= " and LinkID2=" & PPID
        End If
        Wheres &= ")"
      End If
      If SV Then
        If PP Then
          Wheres &= " or "
        End If
        Wheres &= "(ActivitySubTypeID=13"
        If (SVID <> 0) Then
          Wheres &= " and LinkID2=" & SVID
        End If
        Wheres &= ")"
      End If
      If A Then
        If PP Or SV Then
          Wheres &= " or "
        End If
        Wheres &= "(ActivitySubTypeID=14)"
      End If
      Wheres &= ")"
    Else
      If OfferOrProgram = 1 Then
        Wheres &= " and ActivitySubTypeID in (12,13,14) "
      Else
        Wheres &= " and ActivitySubTypeID in (0) "
      End If
    End If
    
    ' ...limited to a particular offer?
    If OID > 0 Then
      i = 0
      If ProgramIDs.Count > 0 Then
        Wheres &= " and (ActivitySubTypeID in (12, 13) and LinkID2 in ("
        For Each Item In ProgramIDs
          Wheres &= IIf(i > 0, ",", "") & Item.Value
          i += 1
        Next
        i = 0
        Wheres &= ") or ActivitySubTypeID=14 and LinkID2 in (" & OID & ")) "
      Else
        Wheres &= " and LinkID=0 "
      End If
    End If
    
    ' ...by a particular value?
    If Request.QueryString("Value") <> "" Then
      If ValueOp = 1 Then
        Wheres &= " and cast(ActivityValue as numeric) >= " & Value & " "
      ElseIf ValueOp = 2 Then
        Wheres &= " and cast(ActivityValue as numeric) = " & Value & " "
      ElseIf ValueOp = 3 Then
        Wheres &= " and cast(ActivityValue as numeric) <= " & Value & " "
      End If
    End If
    
    MyCommon.QueryStr = "select AL.ActivityID, AL.LinkID, AL.ActivityTypeID, AL.AdminID, AL.Description, " & _
                        "AL.ActivityDate, AL.ActivitySubTypeID, AL.LinkID2, AL.ActivityValue, " & _
                        "AL.SessionID, AL.PreAdjustBalance, AL.Adjustment, AL.PostAdjustBalance, " & _
                        "AL.LinkID3, AL.LinkID4, AL.LinkID5, AL.LinkID6, AL.ReasonID, AL.ReasonText, AL.Comments, " & _
                        "AU.FirstName, AU.LastName, AU.UserName, ADJ.Number " & _
                        "from ActivityLog as AL with (NoLock) " & _
                        "inner join AdminUsers as AU with (NoLock) on AU.AdminUserID=AL.AdminID " & _
                        "inner join (" & _
                        "    select AdminID, count(ActivityID) as Number from ActivityLog with (NoLock) " & _
                        "    " & Wheres & " " & _
                        "    group by AdminID" & _
                        ") as ADJ on ADJ.AdminID=AL.AdminID " & _
                        " " & Wheres & " "
    ' ...by user or by role?
    If AdminID <> 0 Then
      MyCommon.QueryStr &= " and AL.AdminID=" & AdminID & " "
    ElseIf RoleID <> 0 Then
      MyCommon.QueryStr &= " and AL.AdminID in (" & AdminIDsString & ") "
    End If
    ' ...by a particular number of adjustments made?
    If Number > 0 Then
      If NumberOp = 1 Then
        MyCommon.QueryStr &= " and Number >= " & Number
      ElseIf NumberOp = 2 Then
        MyCommon.QueryStr &= " and Number = " & Number
      ElseIf NumberOp = 3 Then
        MyCommon.QueryStr &= " and Number <= " & Number
      End If
    End If
    
    MyCommon.QueryStr &= " order by ActivityDate;"
    dtAdj = MyCommon.LRT_Select
    
    ' ~~~ dtOff query (the second of two), which will get all the offer- and household-related activity ~~~
    ' First, build a list of "where" clauses, since we'll be sticking them twice into the query:
    Wheres = ""
    ' ...only customer inquiry activity
    Wheres &= " where ActivityTypeID=25 "
    ' ...only within the specified date range
    Wheres &= " and ActivityDate between '" & StartDate.ToString("yyyy-MM-dd") & "T00:00:00' and '" & EndDate.ToString("yyyy-MM-dd") & "T23:59:59' "
    
    ' ...including add/rem offers or add/rem household?
    If (AddO Or RemO Or AddHH Or RemHH) Then
      Wheres &= " and ActivitySubTypeID in ("
      If AddO Then
        Wheres &= "15"
        If (RemO Or AddHH Or RemHH) Then
          Wheres &= ","
        End If
      End If
      If RemO Then
        Wheres &= "16"
        If (AddHH Or RemHH) Then
          Wheres &= ","
        End If
      End If
      If AddHH Then
        Wheres &= "17"
        If (RemHH) Then
          Wheres &= ","
        End If
      End If
      If RemHH Then
        Wheres &= "18"
      End If
      Wheres &= ") "
    Else
      Wheres &= " and ActivitySubTypeID in (0) "
    End If
    
    ' ...limited to a particular offer?
    If OID > 0 Then
      i = 0
      If CustomerGroupIDs.Count > 0 Then
        Wheres &= " and (ActivitySubTypeID in (15, 16) and LinkID2 in ("
        For Each Item In CustomerGroupIDs
          Wheres &= IIf(i > 0, ",", "") & Item.Value
          i += 1
        Next
        Wheres &= "))"
        i = 0
      Else
        Wheres &= " and ActivitySubTypeID in (15, 16)"
      End If
    End If
    
    MyCommon.QueryStr = "select AL.ActivityID, AL.LinkID, AL.ActivityTypeID, AL.AdminID, AL.Description, " & _
                        "AL.ActivityDate, AL.ActivitySubTypeID, AL.LinkID2, AL.ActivityValue, " & _
                        "AL.SessionID, AL.PreAdjustBalance, AL.Adjustment, AL.PostAdjustBalance, " & _
                        "AL.LinkID3, AL.LinkID4, AL.LinkID5, AL.LinkID6, AL.ReasonID, AL.ReasonText, AL.Comments, " & _
                        "AU.FirstName, AU.LastName, AU.UserName, ADJ.Number " & _
                        "from ActivityLog as AL with (NoLock) " & _
                        "inner join AdminUsers as AU with (NoLock) on AU.AdminUserID=AL.AdminID " & _
                        "inner join (" & _
                        "    select AdminID, count(ActivityID) as Number from ActivityLog with (NoLock) " & _
                        "    " & Wheres & " " & _
                        "    group by AdminID" & _
                        ") as ADJ on ADJ.AdminID = AL.AdminID " & _
                        " " & Wheres & " "
    ' ...by user or by role?
    If AdminID <> 0 Then
      MyCommon.QueryStr &= " and AL.AdminID=" & AdminID & " "
    ElseIf RoleID <> 0 Then
      MyCommon.QueryStr &= " and AL.AdminID in (" & AdminIDsString & ") "
    End If
    
    MyCommon.QueryStr &= " order by ActivityDate;"
    dtOff = MyCommon.LRT_Select
    
    
    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    If (dtAdj.Rows.Count > 0) OrElse (dtOff.Rows.Count > 0) Then
      
      Response.AddHeader("Content-Disposition", "attachment; filename=Adjustments" & ".csv")
      Response.ContentType = "text/csv"
      Response.ContentEncoding = Encoding.GetEncoding("iso-8859-1")
      
      ' Write the column headers
      Sendb(Copient.PhraseLib.Lookup("term.userid", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.firstname", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.lastname", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.customerpk", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.customerid", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.householdid", LanguageID).Replace("&#39;", "'") & ",")
      If Logix.UserRoles.AccessCustomerIdData_FirstName Then
        Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID).Replace("&#39;", "'") & " " & StrConv(Copient.PhraseLib.Lookup("term.firstname", LanguageID).Replace("&#39;", "'"), VbStrConv.Lowercase) & ",")
      End If
      If Logix.UserRoles.AccessCustomerIdData_LastName Then
        Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID).Replace("&#39;", "'") & " " & StrConv(Copient.PhraseLib.Lookup("term.lastname", LanguageID).Replace("&#39;", "'"), VbStrConv.Lowercase) & ",")
      End If
      Sendb(Copient.PhraseLib.Lookup("term.datetime", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.programid", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.programname", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.groupID", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.groupname", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.offername", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.adjustment", LanguageID).Replace("&#39;", "'") & ",")
      Sendb(Copient.PhraseLib.Lookup("term.totaladjustments", LanguageID).Replace("&#39;", "'") & ",")
	  Sendb(Copient.PhraseLib.Lookup("term.comments", LanguageID).Replace("&#39;", "'"))
      If MyCommon.Fetch_SystemOption(108) = "1" Then
        Sendb(",")
        Sendb(Copient.PhraseLib.Lookup("term.reason", LanguageID).Replace("&#39;", "'") & ",")
        Send(Copient.PhraseLib.Lookup("term.reasontext", LanguageID).Replace("&#39;", "'"))
      Else
        Send("")
      End If
      
      ' Write the adjustments activity from dtAdj ~~~~~~~~~~
      If dtAdj.Rows.Count > 0 Then
        For Each row In dtAdj.Rows
          ' This row should be displayed, so we'll continue by setting the basic data
          AdminID = MyCommon.NZ(row.Item("AdminID"), "")
          FirstName = MyCommon.NZ(row.Item("FirstName"), "")
          LastName = MyCommon.NZ(row.Item("LastName"), "")
          CustomerPK = MyCommon.NZ(row.Item("LinkID"), 0)
          SubjectType = MyCommon.NZ(row.Item("ActivitySubTypeID"), 0)
          SubjectID = MyCommon.NZ(row.Item("LinkID2"), 0)
          Adjustment = MyCommon.NZ(row.Item("ActivityValue"), 0)
          TotalAdjustments = MyCommon.NZ(row.Item("Number"), 0)
 		  Comments = MyCommon.NZ(row.Item("Comments"), "") 
          If MyCommon.Fetch_SystemOption(108) = "1" Then
            ReasonID = MyCommon.NZ(row.Item("ReasonID"), 0)
            ReasonText = MyCommon.NZ(row.Item("ReasonText"), "")
            MyCommon.QueryStr = "select Description from AdjustmentReasons with (NoLock) where ReasonID=" & ReasonID & ";"
            dt2 = MyCommon.LXS_Select
            If dt2.Rows.Count > 0 Then
              Reason = MyCommon.NZ(dt2.Rows(0).Item("Description"), "")
            End If
          End If
          
          ' Next, query for customer-related details
          MyCommon.QueryStr = "select CustomerPK, FirstName, MiddleName, LastName, Employee, CustomerTypeID, HHPK " & _
                              "from Customers with (NoLock) where CustomerPK=" & MyCommon.NZ(row.Item("LinkID"), "") & ";"
          dt2 = MyCommon.LXS_Select
          If dt2.Rows.Count > 0 Then
            CustomerFirstName = MyCommon.NZ(dt2.Rows(0).Item("FirstName"), "")
            CustomerMiddleName = MyCommon.NZ(dt2.Rows(0).Item("MiddleName"), "")
            CustomerLastName = MyCommon.NZ(dt2.Rows(0).Item("LastName"), "")
            CustomerType = MyCommon.NZ(dt2.Rows(0).Item("CustomerTypeID"), 0)
            HHPK = MyCommon.NZ(dt2.Rows(0).Item("HHPK"), 0)
            HHExtCardID = ""
            If HHPK <> 0 Then
                            MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK=" & HHPK & " order by ExtCardID;"
              dt3 = MyCommon.LXS_Select
              If dt3.Rows.Count > 0 Then
                i = 0
                For Each row3 In dt3.Rows
                  HHExtCardID &= MyCryptLib.SQL_StringDecrypt(dt3.Rows(i).Item("ExtCardID").ToString())
                  If i < (dt3.Rows.Count - 1) Then
                    HHExtCardID &= "|"
                  End If
                  i += 1
                Next
              Else
                HHExtCardID = 0
              End If
            Else
              HHExtCardID = 0
            End If
          End If
          
          ' Next, query for the program names
          If (SubjectID > 0) Then
            If SubjectType = 12 Then
              MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID=" & SubjectID & ";"
              dt2 = MyCommon.LRT_Select
              If dt2.Rows.Count > 0 Then
                ProgramName = MyCommon.NZ(dt2.Rows(0).Item("ProgramName"), "")
              End If
            ElseIf SubjectType = 13 Then
              MyCommon.QueryStr = "select Name from StoredValuePrograms with (NoLock) where SVProgramID=" & SubjectID & ";"
              dt2 = MyCommon.LRT_Select
              If dt2.Rows.Count > 0 Then
                ProgramName = MyCommon.NZ(dt2.Rows(0).Item("Name"), "")
              End If
            ElseIf SubjectType = 14 Then
              MyCommon.QueryStr = "select Name from Offers with (NoLock) where OfferID=" & SubjectID & " " & _
                                  " union " & _
                                  "select IncentiveName as Name from CPE_Incentives where IncentiveID=" & SubjectID & ";"
              dt2 = MyCommon.LRT_Select
              If dt2.Rows.Count > 0 Then
                ProgramName = MyCommon.NZ(dt2.Rows(0).Item("Name"), "")
              End If
            End If
          End If
          
          ' Finally, write the whole thing to a CSV
          If (SubjectID > 0) Then
            Sendb(AdminID & ",")
            Sendb(MyCommon.Strip_Commas(FirstName) & ",")
            Sendb(MyCommon.Strip_Commas(LastName) & ",")
            Sendb(CustomerPK & ",")
            If CustomerType = 0 Then
              Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID).Replace("&#39;", "'") & ",")
            ElseIf CustomerType = 1 Then
              Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID).Replace("&#39;", "'") & ",")
            ElseIf CustomerType = 2 Then
              Sendb(Copient.PhraseLib.Lookup("term.cam", LanguageID).Replace("&#39;", "'") & ",")
            End If
            ExtCardID = ""
                        MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by ExtCardID;"
            dt2 = MyCommon.LXS_Select
            If dt2.Rows.Count > 0 Then
              i = 0
              For Each row2 In dt2.Rows
                ExtCardID &= MyCryptLib.SQL_StringDecrypt(dt2.Rows(i).Item("ExtCardID").ToString())
                If i < (dt2.Rows.Count - 1) Then
                  ExtCardID &= "|"
                End If
                i += 1
              Next
            End If
            If CustomerType = 1 Then
              Sendb("0,")
              Sendb(ExtCardID & ",")
            Else
              Sendb(ExtCardID & ",")
              Sendb(HHExtCardID & ",")
            End If
            If Logix.UserRoles.AccessCustomerIdData_FirstName Then
              Sendb(MyCommon.Strip_Commas(CustomerFirstName) & ",")
            End If
            If Logix.UserRoles.AccessCustomerIdData_LastName Then
              Sendb(MyCommon.Strip_Commas(CustomerLastName) & ",")
            End If
            Sendb(MyCommon.NZ(row.Item("ActivityDate"), "") & ",")
            If SubjectType = 12 Then
              Sendb(Copient.PhraseLib.Lookup("term.points", LanguageID).Replace("&#39;", "'") & ",")
              Sendb(SubjectID & ",")
              Sendb(MyCommon.Strip_Commas(ProgramName) & ",")
              Sendb(",,,,")
            ElseIf SubjectType = 13 Then
              Sendb(Copient.PhraseLib.Lookup("term.storedvalue", LanguageID).Replace("&#39;", "'") & ",")
              Sendb(SubjectID & ",")
              Sendb(MyCommon.Strip_Commas(ProgramName) & ",")
              Sendb(",,,,")
            ElseIf SubjectType = 14 Then
              Sendb(Copient.PhraseLib.Lookup("term.accumulation", LanguageID).Replace("&#39;", "'") & ",")
              Sendb(",,,,")
              Sendb(SubjectID & ",")
              Sendb(MyCommon.Strip_Commas(ProgramName) & ",")
            Else
              Sendb(",")
            End If
			'RT #5624 
			Sendb(Adjustment & ",")
            Sendb(TotalAdjustments & ",") 
			Sendb(Comments)
            If MyCommon.Fetch_SystemOption(108) = "1" Then
              Sendb(",")
              Sendb(MyCommon.Strip_Commas(Reason) & ",")
              Send(MyCommon.Strip_Commas(ReasonText))
            Else
              Send("")
            End If
          End If
        Next
        ' End of report-writing for dtAdj
      End If
      
      ' Write the offer and household activity from dtOff ~~~~~~~~~~
      If dtOff.Rows.Count > 0 Then
        For Each row In dtOff.Rows
          ' This row should be displayed, so we'll continue by setting the basic data
          AdminID = MyCommon.NZ(row.Item("AdminID"), "")
          FirstName = MyCommon.NZ(row.Item("FirstName"), "")
          LastName = MyCommon.NZ(row.Item("LastName"), "")
          CustomerPK = MyCommon.NZ(row.Item("LinkID"), 0)
          SubjectType = MyCommon.NZ(row.Item("ActivitySubTypeID"), 0)
          SubjectID = MyCommon.NZ(row.Item("LinkID2"), 0)
          Adjustment = MyCommon.NZ(row.Item("ActivityValue"), 0)
          TotalAdjustments = MyCommon.NZ(row.Item("Number"), 0)
		  Comments = MyCommon.NZ(row.Item("Comments"), "")		  
          If MyCommon.Fetch_SystemOption(108) = "1" Then
            ReasonID = MyCommon.NZ(row.Item("ReasonID"), 0)
            ReasonText = MyCommon.NZ(row.Item("ReasonText"), "")
            MyCommon.QueryStr = "select Description from AdjustmentReasons with (NoLock) where ReasonID=" & ReasonID & ";"
            dt2 = MyCommon.LXS_Select
            If dt2.Rows.Count > 0 Then
              Reason = MyCommon.NZ(dt2.Rows(0).Item("Description"), "")
            End If
          End If
          
          ' Next, query for customer-related details
          MyCommon.QueryStr = "select CustomerPK, FirstName, MiddleName, LastName, Employee, CustomerTypeID, HHPK " & _
                              "from Customers with (NoLock) where CustomerPK=" & MyCommon.NZ(row.Item("LinkID"), "") & ";"
          dt2 = MyCommon.LXS_Select
          If dt2.Rows.Count > 0 Then
            CustomerFirstName = MyCommon.NZ(dt2.Rows(0).Item("FirstName"), "")
            CustomerMiddleName = MyCommon.NZ(dt2.Rows(0).Item("MiddleName"), "")
            CustomerLastName = MyCommon.NZ(dt2.Rows(0).Item("LastName"), "")
            CustomerType = MyCommon.NZ(dt2.Rows(0).Item("CustomerTypeID"), 0)
            HHPK = MyCommon.NZ(dt2.Rows(0).Item("HHPK"), 0)
            HHExtCardID = ""
            If HHPK <> 0 Then
                            MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK=" & HHPK & " order by ExtCardID;"
              dt3 = MyCommon.LXS_Select
              If dt3.Rows.Count > 0 Then
                i = 0
                For Each row3 In dt3.Rows
                  HHExtCardID &=MyCryptLib.SQL_StringDecrypt(dt3.Rows(i).Item("ExtCardID").ToString())
                  If i < (dt3.Rows.Count - 1) Then
                    HHExtCardID &= "|"
                  End If
                  i += 1
                Next
              Else
                HHExtCardID = 0
              End If
            Else
              HHExtCardID = 0
            End If
          End If
          
          ' Next, query for the customer group name
          If (SubjectID > 0) Then
            If SubjectType = 15 OrElse SubjectType = 16 Then
              MyCommon.QueryStr = "select Name from CustomerGroups with (NoLock) where CustomerGroupID=" & SubjectID & ";"
              dt2 = MyCommon.LRT_Select
              If dt2.Rows.Count > 0 Then
                CustomerGroupName = MyCommon.NZ(dt2.Rows(0).Item("Name"), "")
              End If
              ' Also, find the associated offer IDs
              TempString = MyCommon.NZ(row.Item("Description"), "")
              If TempString <> "" Then
                TempLen = Len(TempString)
                TempPos = InStr(1, TempString, "(")
                TempString = Mid(TempString, TempPos + 1, (TempLen - TempPos) - 1)
                TempString = Replace(TempString, ", ", "|")
              End If
              ' If there's something in the offer ID string, look up the name of the first offer
              If (TempString <> "") Then
                If (InStr(1, TempString, "|") = 0) Then
                  MyCommon.QueryStr = "select IncentiveName as OfferName from CPE_Incentives with (NoLock) where IncentiveID=" & MyCommon.Extract_Val(TempString) & _
                                      " union " & _
                                      "select Name as OfferName from Offers with (NoLock) where OfferID=" & MyCommon.Extract_Val(TempString) & ";"
                Else
                  MyCommon.QueryStr = "select IncentiveName as OfferName from CPE_Incentives with (NoLock) where IncentiveID=" & MyCommon.Extract_Val(TempString) & _
                                      " union " & _
                                      "select Name as OfferName from Offers with (NoLock) where OfferID=" & MyCommon.Extract_Val(Left(TempString, InStr(1, TempString, "|"))) & ";"
                End If
                dt2 = MyCommon.LRT_Select
                If dt2.Rows.Count > 0 Then
                  OfferName = MyCommon.NZ(dt2.Rows(0).Item("OfferName"), "")
                Else
                  OfferName = "(" & Copient.PhraseLib.Lookup("term.unknown", LanguageID).Replace("&#39;", "'") & ")"
                End If
              Else
                OfferName = ""
              End If
            ElseIf SubjectType = 17 OrElse SubjectType = 18 Then
              CustomerGroupName = ""
            End If
          End If
          
          ' Finally, write the whole thing to a CSV
          Sendb(AdminID & ",")
          Sendb(MyCommon.Strip_Commas(FirstName) & ",")
          Sendb(MyCommon.Strip_Commas(LastName) & ",")
          Sendb(CustomerPK & ",")
          If CustomerType = 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.customer", LanguageID).Replace("&#39;", "'") & ",")
          ElseIf CustomerType = 1 Then
            Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID).Replace("&#39;", "'") & ",")
          ElseIf CustomerType = 2 Then
            Sendb(Copient.PhraseLib.Lookup("term.cam", LanguageID).Replace("&#39;", "'") & ",")
          End If
          ExtCardID = ""
                    MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " order by ExtCardID;"
          dt2 = MyCommon.LXS_Select
          If dt2.Rows.Count > 0 Then
            i = 0
            For Each row2 In dt2.Rows
              ExtCardID &= MyCryptLib.SQL_StringDecrypt(dt2.Rows(i).Item("ExtCardID").ToString())
              If i < (dt2.Rows.Count - 1) Then
                ExtCardID &= "|"
              End If
              i += 1
            Next
          End If
          If CustomerType = 1 Then
            Sendb("0,")
            Sendb(HHExtCardID & ",")
          Else
            Sendb(ExtCardID & ",")
            Sendb(HHExtCardID & ",")
          End If
          Sendb(MyCommon.Strip_Commas(CustomerFirstName) & ",")
          Sendb(MyCommon.Strip_Commas(CustomerLastName) & ",")
          Sendb(MyCommon.NZ(row.Item("ActivityDate"), "") & ",")
          If SubjectType = 15 Then
            Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID).Replace("&#39;", "'") & " " & StrConv(Copient.PhraseLib.Lookup("term.added", LanguageID).Replace("&#39;", "'"), VbStrConv.Lowercase) & ",")
            Sendb(",,")
            Sendb(SubjectID & ",")
            Sendb(MyCommon.Strip_Commas(CustomerGroupName) & ",")
            Sendb(TempString & ",")
            Sendb(OfferName & ",")
          ElseIf SubjectType = 16 Then
            Sendb(Copient.PhraseLib.Lookup("term.offer", LanguageID).Replace("&#39;", "'") & " " & StrConv(Copient.PhraseLib.Lookup("term.removed", LanguageID).Replace("&#39;", "'"), VbStrConv.Lowercase) & ",")
            Sendb(",,")
            Sendb(SubjectID & ",")
            Sendb(MyCommon.Strip_Commas(CustomerGroupName) & ",")
            Sendb(TempString & ",")
            Sendb(OfferName & ",")
          ElseIf SubjectType = 17 Then
            Sendb(Copient.PhraseLib.Lookup("term.householded", LanguageID).Replace("&#39;", "'") & ",")
            Sendb(",,,,,,")
          ElseIf SubjectType = 18 Then
            Sendb(Copient.PhraseLib.Lookup("term.unhouseholded", LanguageID).Replace("&#39;", "'") & ",")
            Sendb(",,,,,,")
          Else
            Sendb(",")
          End If
		  	Sendb(Adjustment & ",")
            Sendb(TotalAdjustments & ",") 
		    Sendb(Comments)
          If MyCommon.Fetch_SystemOption(108) = "1" Then
            Sendb(",")
            Sendb(MyCommon.Strip_Commas(Reason) & ",")
            Send(MyCommon.Strip_Commas(ReasonText))
          Else
            Send("")
          End If
        Next
        ' End of report-writing for dtOff
      End If
      ' End of all report-writing
      GoTo done
    End If
  Else
    ReportStart = Logix.ToShortDateString(DateAdd("d", -7, Date.Now.Date), MyCommon)
    ReportEnd = Logix.ToShortDateString(Date.Now.Date, MyCommon)
  End If
  
  Send_HeadBegin("term.customer", "term.reports")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  If restrictLinks Then
    Send("<link rel=""stylesheet"" href=""../css/logix-restricted.css"" type=""text/css"" media=""screen"" />")
  End If
  Send_Scripts(New String() {"datePicker.js"})
%>

<script type="text/javascript">
  var datePickerDivID = "datepicker";
  
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
  
<% Send_Calendar_Overrides(MyCommon) %>

  function handlePageClick(e) {
    var calFrame = document.getElementById('calendariframe');
    var el=(typeof event!=='undefined')? event.srcElement : e.target
    
    if (el != null) {
      var pickerDiv = document.getElementById(datePickerDivID);
      if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
        if (el.id!="prod-start-picker" && el.id!="prod-end-picker") {
          if (!isDatePickerControl(el.className)) {
            pickerDiv.style.visibility = "hidden";
            pickerDiv.style.display = "none";
            if (calFrame != null) {
              calFrame.style.visibility = 'hidden';
              calFrame.style.display = "none";
            }
          }
        } else  {
          pickerDiv.style.visibility = "visible";            
          pickerDiv.style.display = "block";            
          if (calFrame != null) {
            calFrame.style.visibility = 'visible';
            calFrame.style.display = "block";
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
  
  function handleSearch() {
    var rptStart = document.getElementById("reportstart");
    var rptEnd = document.getElementById("reportend");
    
    if (!isDate(rptStart.value)) {
      rptStart.focus();
      rptStart.select();
      return;
    }
    if (!isDate(rptEnd.value)) {
      rptEnd.focus();
      rptEnd.select();
      return;
    }
    if (!isValidDateRange(new Date(rptStart.value), new Date(rptEnd.value), rptStart)) {
      alert('<% Sendb(Copient.PhraseLib.Lookup("reports.startdate", LanguageID)) %>');
      rptStart.focus();
      rptStart.select();
      return;
    }
    
    generateReport();
  }
  
  function isValidDateRange(dtStart, dtEnd, rptStart) {
    var retVal = true;
    
    if (dtStart!=null && dtEnd!=null) {
      if (dtStart > dtEnd) {
        retVal = false;                    
      }
    }
    
    return retVal;
  }
  
  function toggleR() {
    var elemR = document.getElementById("R");             // "Role" radio button
    var elemU = document.getElementById("U");             // "User" radio button
    var elemRID = document.getElementById("roleID");      // The role <select>
    var elemUID = document.getElementById("adminID");     // The user <select>
    var elemRdiv = document.getElementById("rolediv");    // The div containing the role <select>
    var elemUdiv = document.getElementById("userdiv");    // The div containing the user <select>
    var elemOdiv = document.getElementById("offerdiv");   // The div containing the offer <select>
    var elemPdiv = document.getElementById("programdiv"); // The div containing the programs <select>s
    
    if (elemR.checked) {
      elemRID.disabled = false;
      elemUID.disabled = true;
      elemUID.selectedIndex = 0;
      elemRdiv.style.display = "";
      elemUdiv.style.display = "none";
    }
    if (elemU.checked) {
      elemRID.disabled = true;
      elemUID.disabled = false;
      elemRID.selectedIndex = 0;
      elemRdiv.style.display = "none";
      elemUdiv.style.display = "";
    }
    elemOdiv.style.display = "none";
    elemPdiv.style.display = "none";
    updateSummary();
  }
  
  function toggleO() {
    var elemO = document.getElementById("O");
    var elemP = document.getElementById("P");
    var elemOID = document.getElementById("OID");
    
    var elemRdiv = document.getElementById("rolediv");
    var elemUdiv = document.getElementById("userdiv");
    
    var elemOdiv = document.getElementById("offerdiv");
    var elemPdiv = document.getElementById("programdiv");
    
    var elemPPdiv = document.getElementById("ppdiv");
    var elemPPID = document.getElementById("PPID");
    var elemPP = document.getElementById("PP");
    var elemSVdiv = document.getElementById("svdiv");
    var elemSVID = document.getElementById("SVID");
    var elemSV = document.getElementById("SV");
    var elemA = document.getElementById("A");
    
    if (elemO.checked) {
      elemOID.disabled = false;
      elemOdiv.style.display = "";
      elemPdiv.style.display = "none";
      elemPP.checked = true;
      elemPP.disabled = true;
      elemSV.checked = true;
      elemSV.disabled = true;
      elemA.checked = true;
      elemA.disabled = true;
      elemPPID.selectedIndex = 0;
      elemSVID.selectedIndex = 0;
    }
    if (elemP.checked) {
      elemOID.disabled = true;
      elemOdiv.style.display = "none";
      elemPdiv.style.display = "";
      elemPP.disabled = false;
      elemSV.disabled = false;
      elemA.disabled = false;
      elemOID.selectedIndex = 0;
    }
    elemRdiv.style.display = "none";
    elemUdiv.style.display = "none";
    updateSummary();
  }
  
  function togglePP() {
    var elemPP = document.getElementById("PP");
    var elemPPID = document.getElementById("PPID");
    
    if (elemPP.checked) {
      elemPPID.disabled = false;
      elemPPID.selectedIndex = 0;
    } else {
      elemPPID.disabled = true;
      elemPPID.selectedIndex = -1;
    }
    updateSummary();
  }
  
  function toggleSV() {
    var elemSV = document.getElementById("SV");
    var elemSVID = document.getElementById("SVID");
    
    if (elemSV.checked) {
      elemSVID.disabled = false;
      elemSVID.selectedIndex = 0;
    } else {
      elemSVID.disabled = true;
      elemSVID.selectedIndex = -1;
    }
    updateSummary();
  }
  
  function toggleA() {
    updateSummary();
  }
  
  function toggleAddO() {
    updateSummary();
  }
  function toggleRemO() {
    updateSummary();
  }
  function toggleAddHH() {
    updateSummary();
  }
  function toggleRemHH() {
    updateSummary();
  }
  
  function daysBetween(d1, d2) {
    var date1 = d1.split("/");
    var date2 = d2.split("/");
    var start = new Date(date1[0]+"/"+date1[1]+"/"+date1[2]);
    var end = new Date(date2[0]+"/"+date2[1]+"/"+date2[2]);
    var daysApart = (Math.abs(Math.round((start-end)/86400000)) + 1);
    return daysApart;
  }
  
  function updateSummary() {
    var elemstart = document.getElementById("reportstart");
    var elemend = document.getElementById("reportend");
    var elemuser = document.getElementById("adminID");
    var elemrole = document.getElementById("roleID");
    var elempp = document.getElementById("PP");
    var elemppid = document.getElementById("PPID");
    var elemsv = document.getElementById("SV");
    var elemsvid = document.getElementById("SVID");
    var elema = document.getElementById("A");
    var elemaddo = document.getElementById("AddO");
    var elemremo = document.getElementById("RemO");
    var elemaddhh = document.getElementById("AddHH");
    var elemremhh = document.getElementById("RemHH");
    var elemo = document.getElementById("O");
    var elemoid = document.getElementById("OID");
    var elemvalue = document.getElementById("value");
    var elemvalueop = document.getElementById("valueOp");
    var elemnumber = document.getElementById("number");
    var elemnumberop = document.getElementById("numberOp");
    var elemreasonid = document.getElementById("reasonID");
    
    var txtintro = document.getElementById("txtintro");
    var txtuser = document.getElementById("txtuser");
    var txtperiod = document.getElementById("txtperiod");
    var txtconsisting = document.getElementById("txtconsisting");
    var txtpp = document.getElementById("txtpp");
    var txtand = document.getElementById("txtand");
    var txtsv = document.getElementById("txtsv");
    var txtaccum = document.getElementById("txtaccum");
    var txtaddremo = document.getElementById("txtaddremo");
    var txtaddremhh = document.getElementById("txtaddremhh");
    var txtoffer = document.getElementById("txtoffer");
    var txtwhere = document.getElementById("txtwhere");
    var txtvalue = document.getElementById("txtvalue");
    var txtnumber = document.getElementById("txtnumber");
    var txtreason = document.getElementById("txtreason");
    
    // Intro
    txtintro.innerHTML = "Reporting activity";
    
    // User
    if (elemuser.value == 0) {
      if (elemrole.value == 0) {
        txtuser.innerHTML = "&nbsp;by all users";
      } else {
        txtuser.innerHTML = "&nbsp;by " + elemrole.options[elemrole.selectedIndex].text + " users";
      }
    } else {
      txtuser.innerHTML = "&nbsp;by " + elemuser.options[elemuser.selectedIndex].text;
    }
    
    // Dates
    if (elemstart.value != "" && elemend.value != "") {
      if (elemstart.value == elemend.value) {
        txtperiod.innerHTML = "&nbsp;on " + elemstart.value;
      } else {
        txtperiod.innerHTML = "&nbsp;for the ";
        txtperiod.innerHTML += daysBetween(elemstart.value, elemend.value);
        txtperiod.innerHTML += "-day period from " + elemstart.value + " to " + elemend.value;
      }
    }

    // Reason
    if (elemreasonid.value > 0) {
      txtreason.innerHTML = "&nbsp;due to " + elemreasonid.options[elemreasonid.selectedIndex].text.toLowerCase() + " and";
    } else {
      txtreason.innerHTML = "";
    }
    
    // Consisting
    txtconsisting.innerHTML ="&nbsp;consisting of the following:"
    
    // Programs
    if (elempp.checked == 0 && elemsv.checked == 0) {
      txtpp.innerHTML = "";
      txtand.innerHTML = "";
      txtsv.innerHTML = "";
    } else if ((elempp.checked == 1 && elemsv.checked == 1) && (elemppid.value == 0 && elemsvid.value == 0)) {
      txtpp.innerHTML = "&nbsp;adjustments to all programs";
      txtand.innerHTML = "";
      txtsv.innerHTML = "";
    } else {
      if (elempp.checked == 1) {
        if (elemppid.value == 0) {
          txtpp.innerHTML = "&nbsp;adjustments to all points programs";
        } else {
          txtpp.innerHTML = "&nbsp;adjustments to points program " + elemppid.value + " (&quot;" + elemppid.options[elemppid.selectedIndex].text + "&quot;)";
        }
      } else {
        txtpp.innerHTML = "";
      }
      if (elempp.checked == 1 && elemsv.checked == 1) {
        txtand.innerHTML = "&nbsp;and";
      } else {
        txtand.innerHTML = "";
      }
      if (elemsv.checked == 1) {
        if (elemsvid.value == 0) {
          txtsv.innerHTML = "&nbsp;adjustments to all stored value programs";
        } else {
          txtsv.innerHTML = "&nbsp;adjustments to stored value program " + elemsvid.value + " (&quot;" + elemsvid.options[elemsvid.selectedIndex].text + "&quot;)";
        }
      } else {
        txtsv.innerHTML = "";
      }
    }
    
    // Accumulations
    if (elema.checked == 1) {
      if ((txtpp.innerHTML != "") || (txtand.innerHTML != "") || (txtsv.innerHTML != "")) {
        txtaccum.innerHTML = "&nbsp;and accumulations";
      } else {
        txtaccum.innerHTML = "&nbsp;to accumulation balances";
      }      
    } else {
      txtaccum.innerHTML = "";
    }
    
    // Offer
    if (elemo.checked == 1) {
      if (elemoid.value == 0) {
        txtoffer.innerHTML = "";
      } else {
        txtoffer.innerHTML = "&nbsp;associated with offer " + elemoid.value + " (&quot;" + elemoid.options[elemoid.selectedIndex].text + "&quot;)";
      }
    } else {
      txtoffer.innerHTML = "";
    }
    
    // Value and Number
    if (elemvalue.value != "") {
      txtvalue.innerHTML = "&nbsp;where the adjustment value is";
      if (elemvalueop.value == "1") {
        txtvalue.innerHTML += "&nbsp;at least";
      } else if (elemvalueop.value == "2") {
        txtvalue.innerHTML += "&nbsp;exactly";
      } else if (elemvalueop.value == "3") {
        txtvalue.innerHTML += "&nbsp;at most";
      }
      txtvalue.innerHTML += " " + elemvalue.value;
    } else {
      txtvalue.innerHTML = "";
    }
    if (elemnumber.value != "") {
      if ((elemvalue.value == "") || (elemvalue.value == null)) {
        txtnumber.innerHTML = "&nbsp;where";
      } else {
        txtnumber.innerHTML = "&nbsp;and";
      }
      txtnumber.innerHTML += "&nbsp;the total number of such adjustments made by each user is";
      if (elemnumberop.value == "1") {
        txtnumber.innerHTML += "&nbsp;at least";
      } else if (elemnumberop.value == "2") {
        txtnumber.innerHTML += "&nbsp;exactly";
      } else if (elemnumberop.value == "3") {
        txtnumber.innerHTML += "&nbsp;at most";
      }
      txtnumber.innerHTML += " " + elemnumber.value;
    } else {
      txtnumber.innerHTML = "";
    }
    
    // Add and Remove Offers
    if ((elemaddo.checked == 1) || (elemremo.checked == 1)) {
      if (elempp.checked == 1 || elemsv.checked == 1 || elema.checked == 1) {
        txtaddremo.innerHTML = ";";
      } else {
        txtaddremo.innerHTML = "";
      }
      if (elemaddo.checked == 1) {
        txtaddremo.innerHTML += " additions to";
      }
      if ((elemaddo.checked == 1) && (elemremo.checked == 1)) {
        txtaddremo.innerHTML += " or ";
      }
      if (elemremo.checked == 1) {
        txtaddremo.innerHTML += " removals from";
      }
      if (elemo.checked == 1) {
        if (elemoid.value == 0) {
          txtaddremo.innerHTML += " offers";
        } else {
          txtaddremo.innerHTML += " offer " + elemoid.value + " (&quot;" + elemoid.options[elemoid.selectedIndex].text + "&quot;)";
        }
      } else {
        txtaddremo.innerHTML += " offers";
      }
    } else {
      txtaddremo.innerHTML = "";
    }
    
    // Add and Remove Households
    if ((elemaddhh.checked == 1) || (elemremhh.checked == 1)) {
      if (elempp.checked == 1 || elemsv.checked == 1 || elema.checked == 1 || elemaddo.checked == 1 || elemremo.checked == 1) {
        txtaddremhh.innerHTML = ";";
      } else {
        txtaddremhh.innerHTML = "";
      }
      if (elemaddhh.checked == 1) {
        txtaddremhh.innerHTML += " additions to";
      }
      if ((elemaddhh.checked == 1) && (elemremhh.checked == 1)) {
        txtaddremhh.innerHTML += " or ";
      }
      if (elemremhh.checked == 1) {
        txtaddremhh.innerHTML += " removals from";
      }
      txtaddremhh.innerHTML += " households";
    } else {
      txtaddremhh.innerHTML = "";
    }
  }
</script>

<%
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  If (Not restrictLinks) Then
    Send_Tabs(Logix, 3)
    Send_Subtabs(Logix, 32, 2, , 0)
  Else
    Send_Subtabs(Logix, 91, 1, , 0)
  End If
  
  If (Logix.UserRoles.AccessCustomerInquiryReporting = False) Then
    Send_Denied(1, "perm.customers-reports")
    GoTo done
  End If
%>
<form action="customer-reports.aspx" id="mainform" name="mainform">
<div id="intro">
  <h1 id="title">
    <% Sendb(Copient.PhraseLib.Lookup("term.reports", LanguageID))%>
  </h1>
  <div id="controls">
    <input type="submit" class="regular" id="download" name="download" value="<% Sendb(Copient.PhraseLib.Lookup("term.download", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("reports.downloadnote", LanguageID)) %>" />
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div style="background-color:#ddddff; border:1px solid aaaaff; padding:8px; width:730px;<%Sendb(IIf(LanguageID = 1, "", " display:none;"))%>">
    <span id="txtintro"></span><span id="txtuser"></span><span id="txtperiod"></span><span id="txtreason"></span><span id="txtconsisting"></span><span id="txtpp"></span><span id="txtand"></span><span id="txtsv"></span><span id="txtaccum"></span><span id="txtoffer"></span><span id="txtvalue"></span><span id="txtnumber"></span><span id="txtaddremo"></span><span id="txtaddremhh"></span><span>.</span>
  </div>
  <br />
  <div id="column1x">
    <div id="period">
      <label for="reportstart">
        <b><% Sendb(Copient.PhraseLib.Lookup("term.period", LanguageID))%>:</b>
      </label><br />
      <br class="half" />
      <input type="text" class="short" id="reportstart" name="reportstart" maxlength="10" value="<% Sendb(ReportStart) %>" onchange="updateSummary()" />
      <img src="../images/calendar.png" class="calendar" id="prod-start-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('reportstart', event);" />
      <% Sendb(Copient.PhraseLib.Lookup("term.to", LanguageID))%>
      <input type="text" class="short" id="reportend" name="reportend" maxlength="10" value="<% Sendb(ReportEnd) %>" onchange="updateSummary()" />
      <img src="../images/calendar.png" class="calendar" id="prod-end-picker" alt="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" title="<% Sendb(Copient.PhraseLib.Lookup("term.datepicker", LanguageID)) %>" onclick="displayDatePicker('reportend', event);" />
      <br />
      <br />
      <div id="datepicker" class="dpDiv">
      </div>
    </div>
    <div id="adjustments" style="line-height: 22px;">
      <%
        Send("<label for=""number""><b>" & Copient.PhraseLib.Lookup("customer-report.number", LanguageID) & "</b></label><br />")
        Send("<select id=""numberOp"" name=""numberOp"" onchange=""updateSummary()"">")
        Send("  <option value=""1""" & IIf(NumberOp = 1, " selected=""selected""", "") & ">&gt;=</option>")
        Send("  <option value=""2""" & IIf(NumberOp = 2, " selected=""selected""", "") & ">=</option>")
        Send("  <option value=""3""" & IIf(NumberOp = 3, " selected=""selected""", "") & ">&lt;=</option>")
        Send("</select>")
        Send("<input type=""text"" class=""short"" id=""number"" name=""number"" maxlength=""9"" value=""" & IIf(Number <> 0, Number, "") & """ onchange=""updateSummary()"" /><br />")
        Send("<br class=""half"" />")
        Send("<label for=""value""><b>" & Copient.PhraseLib.Lookup("customer-report.value", LanguageID) & "</b></label><br />")
        Send("<select id=""valueOp"" name=""valueOp"" onchange=""updateSummary()"">")
        Send("  <option value=""1""" & IIf(ValueOp = 1, " selected=""selected""", "") & ">&gt;=</option>")
        Send("  <option value=""2""" & IIf(ValueOp = 2, " selected=""selected""", "") & ">=</option>")
        Send("  <option value=""3""" & IIf(ValueOp = 3, " selected=""selected""", "") & ">&lt;=</option>")
        Send("</select>")
        Send("<input type=""text"" class=""short"" id=""value"" name=""value"" maxlength=""9"" value=""" & IIf(Value.CompareTo(0) <> 0, Value, "") & """ onchange=""updateSummary()"" /><br />")
        Send("<br />")
      %>
    </div>
    <div id="searchby">
      <%
        Send("<b>" & Copient.PhraseLib.Lookup("customer-report.selectusers", LanguageID) & "</b><br />")
        Send("<input type=""radio"" id=""R"" name=""userorrole"" value=""1""" & IIf(R, " checked=""checked""", "") & " onclick=""toggleR()"" /> ")
        Send("<label for=""R"" style=""color:#0000ff;"">" & Copient.PhraseLib.Lookup("term.role", LanguageID) & "</label><br />")
        Send("<input type=""radio"" id=""U"" name=""userorrole"" value=""2""" & IIf(R, "", " checked=""checked""") & " onclick=""toggleR()"" /> ")
        Send("<label for=""U"" style=""color:#0000ff;"">" & Copient.PhraseLib.Lookup("term.user", LanguageID) & "</label><br />")
        Send("<br />")
        Send("<b>" & Copient.PhraseLib.Lookup("customer-report.selectadj", LanguageID) & "</b><br />")
        Send("<input type=""radio"" id=""O"" name=""offerorprogram"" value=""1""" & IIf(O, " checked=""checked""", "") & " onclick=""toggleO()"" /> ")
        Send("<label for=""O"" style=""color:#0000ff;"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & "</label><br />")
        Send("<input type=""radio"" id=""P"" name=""offerorprogram"" value=""2""" & IIf(O, "", " checked=""checked""") & " onclick=""toggleO()"" /> ")
        Send("<label for=""P"" style=""color:#0000ff;"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & "</label><br />")
        Send("  <input type=""checkbox"" id=""PP"" name=""PP"" style=""margin-left:20px;""" & IIf(PP, " checked=""checked""", "") & " onclick=""togglePP()"" /> ")
        Send("  <label for=""PP"">" & Copient.PhraseLib.Lookup("term.pointsprograms", LanguageID) & "</label><br />")
        Send("  <input type=""checkbox"" id=""SV"" name=""SV"" style=""margin-left:20px;""" & IIf(SV, " checked=""checked""", "") & " onclick=""toggleSV()"" /> ")
        Send("  <label for=""SV"">" & Copient.PhraseLib.Lookup("term.storedvalueprograms", LanguageID) & "</label><br />")
        Send("  <input type=""checkbox"" id=""A"" name=""A"" style=""margin-left:20px;""" & IIf(A, " checked=""checked""", "") & " onclick=""toggleA()"" /> ")
        Send("  <label for=""A"">" & Copient.PhraseLib.Lookup("term.accumulation", LanguageID) & "</label><br />")
        Send("<br />")
        If MyCommon.Fetch_SystemOption(108) = "1" Then
          Send("<label for=""reasonID""><b>" & Copient.PhraseLib.Lookup("customer-report.selectreason", LanguageID) & "</b></label><br />")
          Send("<select id=""reasonID"" name=""reasonID"" style=""width:210px;"" onchange=""updateSummary()"">")
          MyCommon.QueryStr = "select ReasonID, Description from AdjustmentReasons with (NoLock) order by ReasonID;"
          dt2 = MyCommon.LXS_Select
          Send("  <option value=""0""" & IIf(ReasonID = 0, " selected=""selected""", "") & ">Any reason</option>")
          If dt2.Rows.Count > 0 Then
            For Each row In dt2.Rows
              Send("  <option value=""" & MyCommon.NZ(row.Item("ReasonID"), 0) & """" & IIf(ReasonID = MyCommon.NZ(row.Item("ReasonID"), 0), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Description"), "") & "</option>")
            Next
          End If
          Send("</select><br />")
          Send("<br />")
        Else
          Send("<select id=""reasonID"" name=""reasonID"" style=""display:none;"">")
          Send("  <option value=""0"">" & Copient.PhraseLib.Lookup("customer-reports.AnyReason", LanguageID) & "</option>")
          Send("</select>")
        End If
        Send("<b>" & Copient.PhraseLib.Lookup("customer-report.selectother", LanguageID) & "</b><br />")
        Send("<input type=""checkbox"" id=""AddO"" name=""AddO""" & IIf(AddO, " checked=""checked""", "") & " onclick=""toggleAddO()"" /> ")
        Send("<label for=""AddO"">" & Copient.PhraseLib.Lookup("customer-report.addoffer", LanguageID) & "</label><br />")
        Send("<input type=""checkbox"" id=""RemO"" name=""RemO""" & IIf(RemO, " checked=""checked""", "") & " onclick=""toggleRemO()"" /> ")
        Send("<label for=""RemO"">" & Copient.PhraseLib.Lookup("customer-report.remoffer", LanguageID) & "</label><br />")
        Send("<input type=""checkbox"" id=""AddHH"" name=""AddHH""" & IIf(AddHH, " checked=""checked""", "") & " onclick=""toggleAddHH()"" /> ")
        Send("<label for=""AddHH"">" & Copient.PhraseLib.Lookup("customer-report.addhh", LanguageID) & "</label><br />")
        Send("<input type=""checkbox"" id=""RemHH"" name=""RemHH""" & IIf(RemHH, " checked=""checked""", "") & " onclick=""toggleRemHH()"" /> ")
        Send("<label for=""RemHH"">" & Copient.PhraseLib.Lookup("customer-report.remhh", LanguageID) & "</label><br />")
      %>
    </div>
  </div>
  
  <div id="gutter">
  </div>
  
  <div id="column2x">
    <div id="rolediv" style="position: absolute; display: none;">
      <%
        Send("<label for=""roleID""><b>" & Copient.PhraseLib.Lookup("term.role", LanguageID) & ":</b></label><br />")
        Send("<br class=""half"" />")
        Send("<select id=""roleID"" name=""roleID"" size=""18"" style=""width:350px;"" onchange=""updateSummary()"">")
        MyCommon.QueryStr = "select RoleID, RoleName, PhraseID from AdminRoles with (NoLock) order by RoleName;"
        dt2 = MyCommon.LRT_Select
        Send("  <option value=""0""" & IIf(RoleID = 0, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.allroles", LanguageID) & "</option>")
        If dt2.Rows.Count > 0 Then
          For Each row In dt2.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("RoleID"), 0) & """" & IIf(RoleID = MyCommon.NZ(row.Item("RoleID"), 0), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("RoleName"), "") & "</option>")
          Next
        End If
        Send("</select>")
      %>
    </div>
    <div id="userdiv" style="position: absolute; display: none;">
      <%
        Send("<label for=""adminID""><b>" & Copient.PhraseLib.Lookup("term.user", LanguageID) & ":</b></label><br />")
        Send("<br class=""half"" />")
        Send("<select id=""adminID"" name=""adminID"" size=""18"" style=""width:350px;"" onchange=""updateSummary()"">")
        MyCommon.QueryStr = "select AdminUserID, FirstName, LastName from AdminUsers order by LastName;"
        dt2 = MyCommon.LRT_Select
        Send("  <option value=""0""" & IIf(AdminID = 0, " selected=""selected""", "") & ">" & Copient.PhraseLib.Lookup("term.allusers", LanguageID) & "</option>")
        If dt2.Rows.Count > 0 Then
          For Each row In dt2.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("AdminUserID"), 0) & """" & IIf(AdminID = MyCommon.NZ(row.Item("AdminUserID"), 0), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("FirstName"), "") & " " & MyCommon.NZ(row.Item("LastName"), "") & "</option>")
          Next
        End If
        Send("</select>")
      %>
    </div>
    <div id="offerdiv" style="position: absolute; display: none;">
      <%
        Send("<label for=""OID""><b>" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & ":</b></label><br />")
        Send("<br class=""half"" />")
        Send("<select id=""OID"" name=""OID"" size=""18"" style=""width:460px;""" & IIf(O, "", " disabled=""disabled""") & " onchange=""updateSummary()"">")
        MyCommon.QueryStr = "select AOLV.* from AllOffersListview AOLV with (NoLock) order by Name;"
        dt2 = MyCommon.LRT_Select
        Sendb("  <option value=""0""")
        If (OID = 0 AndAlso O) Then
          Sendb(" selected=""selected""")
        End If
        Send(">" & Copient.PhraseLib.Lookup("term.alloffers", LanguageID) & "</option>")
        If dt2.Rows.Count > 0 Then
          For Each row In dt2.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("OfferID"), 0) & """" & IIf(OID = MyCommon.NZ(row.Item("OfferID"), 0), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Name"), "&nbsp;") & "</option>")
          Next
        End If
        Send("</select>")
      %>
    </div>
    <div id="programdiv" style="position: absolute; display: none;">
      <%
        Send("<div id=""ppdiv"" style=""float:left;margin-right:10px;"">")
        Send("<label for=""PPID""><b>" & Copient.PhraseLib.Lookup("term.pointsprogram", LanguageID) & ":</b></label><br />")
        Send("<br class=""half"" />")
        Send("<select id=""PPID"" name=""PPID"" size=""18"" style=""width:225px;""" & IIf(PP, "", " disabled=""disabled""") & " onchange=""updateSummary()"">")
        MyCommon.QueryStr = "select ProgramID, ProgramName from PointsPrograms where Deleted=0 order by ProgramName;"
        dt2 = MyCommon.LRT_Select
        Sendb("  <option value=""0""")
        If (PPID = 0 AndAlso PP) Then
          Sendb(" selected=""selected""")
        End If
        Send(">" & Copient.PhraseLib.Lookup("term.allprograms", LanguageID) & "</option>")
        If dt2.Rows.Count > 0 Then
          For Each row In dt2.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("ProgramID"), 0) & """" & IIf(PPID = MyCommon.NZ(row.Item("ProgramID"), 0), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("ProgramName"), "") & "</option>")
          Next
        End If
        Send("</select>")
        Send("</div>")
        Send("<div id=""svdiv"" style=""float:left;"">")
        Send("<label for=""SVID""><b>" & Copient.PhraseLib.Lookup("term.storedvalueprogram", LanguageID) & ":</b></label><br />")
        Send("<br class=""half"" />")
        Send("<select id=""SVID"" name=""SVID"" size=""18"" style=""width:225px;""" & IIf(SV, "", " disabled=""disabled""") & " onchange=""updateSummary()"">")
        MyCommon.QueryStr = "select SVProgramID, Name from StoredValuePrograms where Deleted=0 order by Name;"
        dt2 = MyCommon.LRT_Select
        Sendb("  <option value=""0""")
        If (SVID = 0 AndAlso SV) Then
          Sendb(" selected=""selected""")
        End If
        Send(">" & Copient.PhraseLib.Lookup("term.allprograms", LanguageID) & "</option>")
        If dt2.Rows.Count > 0 Then
          For Each row In dt2.Rows
            Send("  <option value=""" & MyCommon.NZ(row.Item("SVProgramID"), 0) & """" & IIf(SVID = MyCommon.NZ(row.Item("SVProgramID"), 0), " selected=""selected""", "") & ">" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
          Next
        End If
        Send("</select>")
        Send("</div>")
      %>
    </div>
    <%
      If Request.Browser.Type = "IE6" Then
        Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
      End If
      If Download = True Then
        If (dtAdj.Rows.Count = 0 AndAlso dtOff.Rows.Count = 0) OrElse (OID > 0 AndAlso ProgramIDs.Count = 0) Then
          Send("<script type=""text/javascript"">")
          Send("  alert('" & Copient.PhraseLib.Lookup("customer-inquiry.noactivity", LanguageID) & "');")
          Send("</script>")
        End If
      End If
    %>
  </div>
</div>
</form>

<script type="text/javascript">
  updateSummary();
  toggleO();
  toggleR();
</script>

<%
  Send_BodyEnd("mainform", "reportstart")
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
