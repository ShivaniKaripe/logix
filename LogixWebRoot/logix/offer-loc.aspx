<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: offer-loc.aspx 
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
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim rstSelected As DataTable
  Dim rstExcluded As DataTable
  Dim row As DataRow
  Dim row2 As DataRow
  Dim OfferID As Long
  Dim PromoID As String
  Dim Name As String = ""
  Dim IsTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  'Dim Disallow_Stores As Boolean = False
  'Dim Disallow_Terminals As Boolean = False
  'Dim PaddingValue As Integer = 0
  'Dim BtnPaddingTop As Integer = 30
  Dim EngineID As Integer = 0
  Dim EngineSubTypeID As Integer = 0
  Dim SelectSize As Integer = 6
  Dim infoMessage As String = ""
  Dim modMessage As String = ""
  Dim Handheld As Boolean = False
  Dim BannersEnabled As Boolean = False
  Dim NameValue As String = ""
  Dim PrevNameValue As String = ""
  Dim BannerList As New ArrayList(10)
  Dim BannerListStr As String = ""
  Dim LoopCtr As Integer = 0
  Dim AllBannersPermission As Boolean = False
  Dim StatusCode As Copient.LogixInc.STATUS_FLAGS = Copient.LogixInc.STATUS_FLAGS.STATUS_UNKNOWN
  Dim StatusText As String = ""
  Dim Popup As Boolean = False
  Dim FocusType As Integer = 0
  Dim LocationGroupID As Integer = 0
  Dim ExcludedStores As String
  Dim IsCam As Boolean = False
  Dim IsUE As Boolean = False
  Dim TierLevels As Integer = 1
  Dim CloseAfterSave As Boolean = False
  Dim bStoreUser As Boolean = False
  Dim sValidLocIDs As String = ""
  Dim wherestr As String = "" 
  Dim sUnion As String = ""
  Dim sAllLoc As String = ""  
  Dim iLen As Integer = 0
  Dim int As Integer
  Dim bCreateGroupOrProgramFromOffer As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(134) ="1",True,False)
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "offer-loc.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  'Store User
  If(MyCommon.Fetch_CM_SystemOption(131) = "1") Then
    MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
    rst = MyCommon.LRT_Select
    iLen = rst.Rows.Count
    If iLen > 0 Then
      bStoreUser = True
      For int=0 to (iLen-1)
        If int=0 Then 
          sValidLocIDs = rst.Rows(0).Item("LocationID")
        Else 
          sValidLocIDs &= "," & rst.Rows(int).Item("LocationID")
        End If
      Next
    End If
  End If
  
  OfferID = Request.QueryString("OfferID")
  
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  Popup = IIf((Request.QueryString("Popup") <> "") AndAlso (Request.QueryString("Popup") <> "0"), True, False)
  FocusType = IIf(Request.QueryString("Focus") = "1", 1, 0)
  
  ' Determine the engine type
  MyCommon.QueryStr = "Select EngineID from OfferIDs with (NoLock) where OfferID=" & OfferID
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    EngineID = MyCommon.NZ(row.Item("EngineID"), 0)
  Next
  'If EngineID = 7 Then
  '  Popup = True
  'End If
  ' CAM works like CPE, so to simplify, consider CAM CPE
  If EngineID = 6 Then
    EngineID = 2
    IsCAM = True
  End If
  
  ' UE works like CPE, so to simplify, consider UE CPE
  'If EngineID = 9 Then
  '  EngineID = 2
  '  IsUE = True
  'End If
  
  'Determine the engine subtype
  If EngineID = 2 Then
    MyCommon.QueryStr = "select EngineSubTypeID from CPE_Incentives with (NoLock) where IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("EngineSubTypeID"), 0)
    End If
  End If
  
  If (OfferID = 0) Then
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "offer-gen.aspx?new=New")
    GoTo done
  End If
  
  'Get the tier level
  If EngineID = 2 OrElse EngineID = 6 Then
    MyCommon.QueryStr = "select TierLevels from CPE_RewardOptions with (NoLock) where TouchResponse=0 and IncentiveID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      TierLevels = rst.Rows(0).Item("TierLevels")
    End If
  ElseIf EngineID = 0 Then
    MyCommon.QueryStr = "select NumTiers from Offers with (NoLock) where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      TierLevels = rst.Rows(0).Item("NumTiers")
      If TierLevels = 0 Then
        TierLevels = 1
      End If
    End If
  End If
  
  If EngineID = 2 OrElse EngineID = 6 Then
    MyCommon.QueryStr = "select IncentiveName as Name, IsTemplate, FromTemplate from CPE_Incentives with (NoLock) " & _
                        "where IncentiveID=" & OfferID & " and Deleted=0;"
  Else
    MyCommon.QueryStr = "select OfferID, Name, IsTemplate, FromTemplate, O.Description, O.OfferCategoryID, CG.Description as cgDescription, " & _
                        "OfferTypeID, ProdStartDate, ProdEndDate, TestStartDate, TestEndDate, TierTypeID, NumTiers, DistPeriod, DistPeriodLimit, DistPeriodVarID, " & _
                        "EmployeeFiltering, NonEmployeesOnly, CRMRestricted, O.LastUpdate, StatusFlag, PriorityLevel, O.EngineID, PE.Description as eDescription " & _
                        "from Offers as O with (NoLock) " & _
                        "left join OfferCategories as CG with (NoLock) on CG.OfferCategoryID=O.OfferCategoryID " & _
                        "left join PromoEngines as PE with (NoLock) on PE.EngineID=O.EngineID " & _
                        "where offerID=" & OfferID & " and O.Deleted=0 and Visible=1;"
  End If
  
  rst = MyCommon.LRT_Select()
  For Each row In rst.Rows
    Name = MyCommon.NZ(row.Item("Name"), "")
    IsTemplate = row.Item("IsTemplate")
    FromTemplate = row.Item("FromTemplate")
  Next
  
  If (Request.QueryString("stores-add1") <> "" And Request.QueryString("sgroups-available") <> "") Then
    Dim i As String = Request.QueryString("sgroups-available")
    Dim a() As String
    Dim j As Integer
    a = i.Split(",")
    MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where LocationGroupID=1 and OfferID=" & OfferID
    rst = MyCommon.LRT_Select
    
    If (rst.Rows.Count > 0) Then
      ' All cardholders is already in the selected box, so lose it
      MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, StatusFlag=2, TCRMAStatusFlag=3 where LocationGroupID=1 and OfferID=" & OfferID
      MyCommon.LRT_Execute()
      If EngineID = 2 Or EngineID = 6 Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate() where IncentiveID=" & OfferID
      Else
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
      End If
      MyCommon.LRT_Execute()
    End If
    
    If (a(0) = 1) Then
      MyCommon.QueryStr = "delete from OfferLocations with (RowLock) where Deleted=1 and OfferID=" & OfferID
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, TCRMAStatusFlag=3 where OfferID=" & OfferID
      MyCommon.LRT_Execute()
      
      MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = a(0)
      MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
      MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
      MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      
      MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & a(0)
      MyCommon.LRT_Execute()
      If EngineID = 2 OrElse EngineID = 6 Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
      Else
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
      End If
      
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addstore", LanguageID))
    Else
      For j = 0 To a.GetUpperBound(0)
        MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
        MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = a(j)
        MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
        MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
        MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
        
        MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & a(j)
        MyCommon.LRT_Execute()
        
        If EngineID = 2 Or EngineID = 6 Then
          MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate() where IncentiveID=" & OfferID
        Else
          MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
        End If
        
        MyCommon.LRT_Execute()
        'MyCommon.QueryStr = "insert into OfferLocations (OfferID,LocationGroupID,LastUpdate) values(" & OfferID & "," & a(j) & "," & "getdate())"
        'MyCommon.LRT_Execute()            
      Next
      
      ' clear out the excluded during an add 
      MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where OfferID=" & OfferID & " and Excluded=1"
      rst = MyCommon.LRT_Select
      For Each row In rst.Rows
        MyCommon.QueryStr = "dbo.pt_OfferLocations_Delete"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = row.Item("PKID")
        MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()
      Next
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addstore", LanguageID))
    End If
    
  ElseIf (Request.QueryString("stores-rem1") <> "" And Request.QueryString("sgroups-select") <> "") Then
    MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where OfferID=" & OfferID & " and LocationGroupID in (" & Request.QueryString("sgroups-select") & ")"
    rst = MyCommon.LRT_Select
    
    For Each row In rst.Rows
      MyCommon.QueryStr = "dbo.pt_OfferLocations_Delete"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = row.Item("PKID")
      MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
    Next
    
    MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2,TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID in (" & Request.QueryString("sgroups-select") & ")"
    MyCommon.LRT_Execute()
    If EngineID = 2 OrElse EngineID = 6 Then
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, lastupdate=getdate() where IncentiveID=" & OfferID
    Else
      MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
    End If
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-removestore", LanguageID))
    
  ElseIf (Request.QueryString("stores-add2") <> "" And Request.QueryString("sgroups-available") <> "" And Request.QueryString("sgroups-available") <> "1") Then
    ' Someone clicked excluded, so force all locations into the selected box
    MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
    MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = 1
    MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 0
    MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
    MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    MyCommon.Close_LRTsp()
    
    MyCommon.QueryStr = "select PKID from OfferLocations with (NoLock) where Excluded<>1 and Deleted=0 and OfferID=" & OfferID
    rst = MyCommon.LRT_Select
    
    If (rst.Rows.Count > 0) Then
      ' All cardholders is already in the selected box, so lose it
      MyCommon.QueryStr = "update OfferLocations with (RowLock) set Deleted=1, StatusFlag=2, TCRMAStatusFlag=3 where LocationGroupID<>1 and OfferID=" & OfferID
      MyCommon.LRT_Execute()
      If EngineID = 2 Or EngineID = 6 Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate() where IncentiveID=" & OfferID
      Else
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
      End If
      MyCommon.LRT_Execute()
    End If
    
    'Response.Write("Adding " & Request.QueryString("sgroups-avail") & " to excluded")
    Dim i As String = Request.QueryString("sgroups-available")
    Dim a() As String
    Dim j As Integer = 0
    Dim locGroupID As String = ""
    a = i.Split(",")
    
    'find the first group that's not All Locations(i.e. LocationGroupdID= "1") and use it
    For j = 0 To a.GetUpperBound(0)
      If (a(j) <> "1") Then
        locGroupID = a(j)
        Exit For
      End If
    Next
    
    If (locGroupID <> "") Then
      MyCommon.QueryStr = "dbo.pt_OfferLocations_Insert"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@LocationGroupID", SqlDbType.BigInt).Value = locGroupID
      MyCommon.LRTsp.Parameters.Add("@Excluded", SqlDbType.Bit).Value = 1
      MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
      MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2, TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & a(j)
      MyCommon.LRT_Execute()
      If EngineID = 2 Or EngineID = 6 Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
      Else
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
      End If
      MyCommon.LRT_Execute()
      'MyCommon.QueryStr = "insert into OfferLocations (OfferID,LocationGroupID,LastUpdate) values(" & OfferID & "," & a(j) & "," & "getdate())"
      'MyCommon.LRT_Execute()
    End If
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-excludestore", LanguageID))
  ElseIf (Request.QueryString("stores-rem2") <> "" And Request.QueryString("sgroups-exclude") <> "") Then
    MyCommon.QueryStr = "select PKID from OfferLocations where OfferID=" & OfferID & " and LocationGroupID=" & Request.QueryString("sgroups-exclude")
    rst = MyCommon.LRT_Select
    
    For Each row In rst.Rows
      MyCommon.QueryStr = "dbo.pt_OfferLocations_Delete"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = row.Item("PKID")
      MyCommon.LRTsp.Parameters.Add("@TCRMAStatusFlag", SqlDbType.Int).Value = 3
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
    Next
    
    MyCommon.QueryStr = "update OfferLocations with (RowLock) set StatusFlag=2, TCRMAStatusFlag=3 where OfferID=" & OfferID & " and LocationGroupID=" & Request.QueryString("sgroups-exclude")
    MyCommon.LRT_Execute()
    If EngineID = 2 OrElse EngineID = 6 Then
      MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
    Else
      MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
    End If
    MyCommon.LRT_Execute()
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-unexcludestore", LanguageID))
  ElseIf (Request.QueryString("terminals-add1") <> "" And Request.QueryString("terminals-available") <> "") Then
    ' Response.Write("terminals-available:" & Request.QueryString("terminals-available"))
    Dim i As String = Request.QueryString("terminals-available")
    Dim a() As String
    Dim j As Integer
    
    ' First delete the any terminals, if present
    MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=0 and TerminalTypeID in " & _
                        "(select TerminalTypeID from TerminalTypes with (NoLock) where AnyTerminal=1);"
    MyCommon.LRT_Execute()
    a = i.Split(",")
        For j = 0 To a.GetUpperBound(0)
            MyCommon.QueryStr = "select 1 from OfferTerminals where OfferID = @OfferID And TerminalTypeID = @TerminalTypeID"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            MyCommon.DBParameters.Add("@TerminalTypeID", SqlDbType.BigInt).Value = a(j)
            Dim dt As New DataTable
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If Not dt.Rows.Count > 0 Then
                MyCommon.QueryStr = "insert into OfferTerminals with (RowLock) (OfferID,TerminalTypeID,LastUpdate) values(" & OfferID & "," & a(j) & "," & "getdate())"
                MyCommon.LRT_Execute()
                If EngineID = 2 OrElse EngineID = 6 Then
                    MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
                Else
                    MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
                End If
                MyCommon.LRT_Execute()
            End If
        Next
    
    ' clear out all other terminals if anyterminals terminal type is selected
    MyCommon.QueryStr = "select TT.TerminalTypeId from OfferTerminals OT with (NoLock) " & _
                        "inner join TerminalTypes TT with (NoLock) on TT.TerminalTypeId=OT.TerminalTypeID " & _
                        "where OT.Excluded=0 and OfferID=" & OfferID & " and TT.AnyTerminal=1;"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=0 and TerminalTypeID not in " & _
                          "(select TerminalTypeID from TerminalTypes where AnyTerminal=1);"
      MyCommon.LRT_Execute()
    End If
    
    ' clear out all excluded terminals if a regular terminal type is selected
    MyCommon.QueryStr = "select TT.TerminalTypeId from OfferTerminals OT with (NoLock) " & _
                        "inner join TerminalTypes TT with (NoLock) on TT.TerminalTypeId=OT.TerminalTypeID " & _
                        "where OT.Excluded=0 and OfferID=" & OfferID & " and TT.AnyTerminal=0;"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=1;"
      MyCommon.LRT_Execute()
    End If
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-addterminal", LanguageID))
  ElseIf (Request.QueryString("terminals-rem1") <> "" And Request.QueryString("terminals-select") <> "") Then
    Dim i As String = Request.QueryString("terminals-select")
    Dim a() As String
    Dim j As Integer
    
    ' first delete all excluded terminals
    MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=1;"
    MyCommon.LRT_Execute()
    
    a = i.Split(",")
    For j = 0 To a.GetUpperBound(0)
      MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and TerminalTypeID=" & a(j)
      MyCommon.LRT_Execute()
      If EngineID = 2 Or EngineID = 6 Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
      Else
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
      End If
      MyCommon.LRT_Execute()
    Next
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-removeterminal", LanguageID))
  ElseIf (Request.QueryString("terminals-add2") <> "" And Request.QueryString("terminals-available") <> "") Then
    Dim i As String = Request.QueryString("terminals-available")
    Dim a() As String
    Dim j As Integer
    
    ' first check if any terminals is selected before allowing an excluded terminal
    MyCommon.QueryStr = "select OT.* from OfferTerminals OT with (NoLock) " & _
                        "inner join TerminalTypes TT with (NoLock) on OT.TerminalTypeID = TT.TerminalTypeID " & _
                        "where OfferID=" & OfferID & " and Excluded=0 and AnyTerminal=1;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      a = i.Split(",")
            For j = 0 To a.GetUpperBound(0)
                MyCommon.QueryStr = "select 1 from OfferTerminals where OfferID = @OfferID And TerminalTypeID = @TerminalTypeID"
                MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                MyCommon.DBParameters.Add("@TerminalTypeID", SqlDbType.BigInt).Value = a(j)
                Dim dt As New DataTable
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If Not dt.Rows.Count > 0 Then
                    MyCommon.QueryStr = "insert into OfferTerminals with (RowLock) (OfferID, TerminalTypeID, LastUpdate, Excluded) values(" & OfferID & ", " & a(j) & ", " & "getdate(), 1)"
                    MyCommon.LRT_Execute()
                    If EngineID = 2 OrElse EngineID = 6 Then
                        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
                    Else
                        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
                    End If
                    MyCommon.LRT_Execute()
                End If
            Next
      ' delete just in case the any terminal was selected
      MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where OfferID=" & OfferID & " and Excluded=1 and TerminalTypeID in " & _
                          "(select TerminalTypeID from TerminalTypes with (NoLock) where AnyTerminal=1);"
      MyCommon.LRT_Execute()
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-excludeterminal", LanguageID))
    Else
      infoMessage = Copient.PhraseLib.Lookup("offer-loc.excludewarning", LanguageID)
    End If
  ElseIf (Request.QueryString("terminals-rem2") <> "" And Request.QueryString("terminals-exclude") <> "") Then
    Dim i As String = Request.QueryString("terminals-exclude")
    Dim a() As String
    Dim j As Integer
    a = i.Split(",")
    For j = 0 To a.GetUpperBound(0)
      MyCommon.QueryStr = "delete from OfferTerminals with (RowLock) where PKID=" & a(j)
      MyCommon.LRT_Execute()
      If EngineID = 2 Or EngineID = 6 Then
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where IncentiveID=" & OfferID
      Else
        MyCommon.QueryStr = "update Offers with (RowLock) set StatusFlag=1, LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & " where OfferID=" & OfferID
      End If
      MyCommon.LRT_Execute()
    Next
    MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-unexcludeterminal", LanguageID))
  End If
  
  Send_HeadBegin("term.offer", "term.locations", OfferID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_HeadEnd()
  If (IsTemplate) Then
    Send_BodyBegin(IIf(Popup, 13, 11))
  Else
    Send_BodyBegin(IIf(Popup, 3, 1))
  End If
  
  If (Not Popup) Then
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 2)
    If IsCam Then
      EngineID = 6
    ElseIf IsUE Then
      EngineID = 9
    End If
    
    If (IsTemplate) Then
      Select Case EngineID
        Case 0, 1
          Send_Subtabs(Logix, 22, 7, , OfferID)
        Case 2
          Send_Subtabs(Logix, 25, 8, , OfferID)
          'Case 4
          '  Send_Subtabs(Logix, 201, 7, , OfferID)
        Case 6
          Send_Subtabs(Logix, 205, 8, , OfferID)
        Case 9
          Send_Subtabs(Logix, 209, 8, , OfferID)
        Case Else
          Send_Subtabs(Logix, 22, 7, , OfferID)
      End Select
    Else
      Select Case EngineID
        Case 0, 1
          Send_Subtabs(Logix, 21, 7, , OfferID)
        Case 2
          Send_Subtabs(Logix, 24, 8, , OfferID)
        'Case 4
        '  Send_Subtabs(Logix, 200, 7, , OfferID)
        Case 6
          Send_Subtabs(Logix, 205, 8, , OfferID)
        Case 9
          Send_Subtabs(Logix, 208, 8, , OfferID)
        Case Else
          Send_Subtabs(Logix, 21, 7, , OfferID)
      End Select
    End If
    If IsCam Then
      EngineID = 2
    ElseIf IsUE Then
      EngineID = 2
    End If
  End If
  
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send_Denied(1, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send_Denied(1, "perm.offers-access-templates")
    GoTo done
  End If
  If (Logix.UserRoles.AccessInstantWinOffers = False AndAlso EngineID = 2 AndAlso EngineSubTypeID = 1) Then
    Send_Denied(1, "perm.offers-accessinstantwin")
    GoTo done
  End If
  
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function updateCookie() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
  
  If (infoMessage = "") Then
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
  Else
    CloseAfterSave = False
  End If
  If (Request.QueryString("SaveThenClose") = "true") Then
    CloseAfterSave = True
  End If
%>
<script type="text/javascript">

var searchTextVal=0;
function handleCreateClick(createbtn)
    {
    
        var  alertMessage='<% Sendb(Copient.PhraseLib.Lookup("term.enter", LanguageID))%>' +' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.valid", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID).ToLower())%>'+' ' + '<% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID).ToLower())%>';
        if(document.getElementById(createbtn)!= undefined && document.getElementById(createbtn) != null)
        {
            var searchText= document.getElementById('functioninput').value;
            if(searchText != null && searchText!="")
            {
                var isExistinginSelectedList= isLocationExisting('sgroups-select');
                var isExistinginExcludedList= isLocationExisting('sgroups-exclude');
                var isExistinginAvailableList = isLocationExisting('sgroups-available');
                
                if(isExistinginSelectedList == true || isExistinginExcludedList == true)
                {
                     alert('<% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID))%>' + ": '"+ searchText + "' " + '<% Sendb(Copient.PhraseLib.Lookup("term.is", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.already", LanguageID).ToLower())%>'+'  '+'<% Sendb(Copient.PhraseLib.Lookup("term.selected", LanguageID).ToLower())%>');
                     return false;
                }
                else if(isExistinginAvailableList == true)
                {
                    if(parseInt(searchTextVal) != 1)
                        {
                            alert ('<% Sendb(Copient.PhraseLib.Lookup("term.existing", LanguageID) )%>'  +'  '+ '<% Sendb( Copient.PhraseLib.Lookup("term.group", LanguageID).ToLower())%>'+ ": '"+ searchText+"'  "+'<% Sendb(Copient.PhraseLib.Lookup("offer.message", LanguageID))%>');
                            xmlhttpPost_CreateGroupOrProgramFromOffer('OfferFeeds.aspx', 'CreateGroupOrProgramFromOffer');
                        }
                    else
                        {
                            alert(alertMessage);
                            return false;
                        }
                }
                else
                {
                    xmlhttpPost_CreateGroupOrProgramFromOffer('OfferFeeds.aspx', 'CreateGroupOrProgramFromOffer');
                }
                
            }
            else
            {
                alert(alertMessage);
                document.getElementById('functioninput').focus();
                return false;
            }
        }
        return true;
    }

    function isLocationExisting(ctlName)
    {
        var isExisting =false;
        var searchText= document.getElementById('functioninput').value;
        var x = document.getElementById(ctlName);
            var txt = "";
            var val = "";
            for (var i = 0; i < x.length; i++) 
            {
                if(String(searchText).toLowerCase()== String(x[i].text).toLowerCase())
                {
                        isExisting =true;
                        searchTextVal = x[i].value;
                        break;
                }
            }
        return isExisting;
    }


    function xmlhttpPost_CreateGroupOrProgramFromOffer(strURL,mode) {
      var xmlHttpReq = false;
      var self = this;
      document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
      if (window.XMLHttpRequest) {
        self.xmlHttpReq = new XMLHttpRequest();
      }
      // IE
      else if (window.ActiveXObject) {
        self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
      }
      var qryStr = getcreatequery(mode);
      self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
      self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
      self.xmlHttpReq.onreadystatechange = function() {
        if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
          updatepage_creategroupOrprogramfromoffer(self.xmlHttpReq.responseText);
        }
      }

      self.xmlHttpReq.send(qryStr);
    }
    function getcreatequery(mode,parameters)
    {
        return "Mode=" + mode + "&CreateType=Location&Name=" + document.getElementById('functioninput').value+"&OfferID="+document.getElementById('OfferID').value;
    }
    function updatepage_creategroupOrprogramfromoffer(str) 
    {
       if(str.length > 0)
      {
          var status ="";
          var responseArr = str.split('~');
          if(responseArr.length >0)
          {
            status=responseArr[0];
            if(status=="Ok")
            {
               window.location.href="offer-loc.aspx?OfferID="+document.getElementById('OfferID').value ;
            }
            else if(status =="Error")
            {
                alert(responseArr[1]);
                return false;
            }
         }
      }
       document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
   }

  var isIE =(navigator.appName.indexOf("Microsoft")!=-1) ? true : false;
  var isOpera = (navigator.appName.indexOf("Opera")!=-1) ? true : false;
  var fullAvailList = null;
  var hasUnsavedChanges = <%Sendb(IIf(Request.QueryString("condChanged") = "true", "true","false")) %>;
  
  // This is the javascript array holding the function list
  // The PrintJavascriptArray ASP function can be used to print this array.
  // Locations array
  <% 
    If (BannersEnabled) Then
      If (EngineID = 0) Then
        MyCommon.QueryStr = "select BE.BannerID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join Banners BAN with (NoLock) on BAN.BannerID = AUB.BannerID " & _
                            "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                            "inner join BannerOffers BO with (NoLock) on BO.BannerID = BE.BannerID " & _
                            "where AdminUserID = " & AdminUserID & " and BE.EngineID = " & EngineID & _
                            " and BAN.AllBanners=1 and BO.OfferId=" & OfferID & ";" 
      Else
        MyCommon.QueryStr = "select BE.BannerID from AdminUserBanners AUB with (NoLock) " & _
                            "inner join Banners BAN with (NoLock) on BAN.BannerID = AUB.BannerID " & _
                            "inner join BannerEngines BE with (NoLock) on BE.BannerID = BAN.BannerID " & _
                            "where AdminUserID = " & AdminUserID & " and BE.EngineID = " & EngineID & " and BAN.AllBanners=1;" 
      End If
      rst = MyCommon.LRT_Select
      AllBannersPermission = (rst.Rows.Count > 0)
      
      ' build up a string with the Banner IDs, if all banners then get all the banners assigned to that engine
      MyCommon.QueryStr  = "select Distinct AUB.BannerID, BAN.AllBanners from Banners BAN with (NoLock)	" & _
                           " inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                           " inner join BannerOffers BO with (NoLock) on BO.BannerID = BAN.BannerID " & _
                           "where AUB.AdminUserID =" & AdminUserID & " and BO.OfferID=" & OfferID & " and BAN.Deleted=0"
      rst = MyCommon.LRT_Select
      If (rst.rows.count>0 )
        For Each row In rst.Rows
          BannerList.Add(MyCommon.NZ(row.Item("BannerID"), -1))
          If (MyCommon.NZ(row.Item("AllBanners"), False)) then
            ' get all the banners assigned to the engine
            MyCommon.QueryStr = "select BannerID from BannerEngines BE with (NoLock) where EngineID = " & EngineID & " " & _
                                " and BannerID <> " & MyCommon.NZ(row.Item("BannerID"), -1) & ";"
            rst3 = MyCommon.LRT_Select
            For Each row2 In rst3.Rows
              BannerList.Add(MyCommon.NZ(row2.Item("BannerID"), -1))
            Next
          End If
        Next
      Else
        BannerList.Add("-1")
      End If
      
      For LoopCtr = 0 to BannerList.Count-1
        If (LoopCtr >0) Then BannerListStr &= ","
        BannerListStr &= BannerList(LoopCtr)
      Next LoopCtr

      If bStoreUser Then
        wherestr = " and (LocationGroupID in (select LocationGroupID from LocGroupItems where LocationID in (" & sValidLocIDs & "))) "
      else
        sUnion =  "union " & _
                  "select LocationGroupID, Name, PhraseID, AllLocations, 0 as BannerID, '" & Copient.PhraseLib.Lookup("term.all", LanguageID) & "' as BannerName from LocationGroups as LG with (NoLock) where Deleted=0 and LocationGroupID not in " & _
                  "(select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ")" & " " & _
                  "and Deleted=0 and AllLocations=1 " & _
                  "order by AllLocations desc, BannerName, Name"
      End If
      
      If EngineID = 0 Then
        MyCommon.QueryStr = "select LocationGroupID, LG.Name, PhraseID, AllLocations, BAN.BannerID, BAN.Name as BannerName from LocationGroups as LG with (NoLock) " & _
                            "left join Banners BAN with (NoLock) on BAN.BannerID = LG.BannerID and BAN.Deleted=0 " & _
                            "where LG.Deleted=0 and LocationGroupID not in (select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ") " & wherestr & _
                            "and (EngineID=" & EngineID & " and AllLocations=0) and (" & IIf(AllBannersPermission, "LG.BannerID is NULL or", "") & " LG.BannerID in (" & BannerListStr & ")) " & sUnion & ";"
      Else
        MyCommon.QueryStr = "select LocationGroupID, LG.Name,PhraseID, BAN.BannerID, BAN.Name as BannerName from LocationGroups as LG with (NoLock) " & _
                            "left join Banners BAN with (NoLock) on BAN.BannerID = LG.BannerID and BAN.Deleted=0 " & _
                            "where LG.Deleted=0 and LocationGroupID not in (select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ")" & _
                            "   and (EngineID=" & EngineID & " and AllLocations=0) and (" & IIf(AllBannersPermission, "LG.BannerID is NULL or", "") & " LG.BannerID in (" & BannerListStr & ")) " & _
                            "order by AllLocations desc, BannerName, Name;"
      End If
    Else
    
      If bStoreUser Then
        wherestr = " and (LocationGroupID in (select LocationGroupID from LocGroupItems where LocationID in (" & sValidLocIDs & "))) "
      Else
        sAllLoc = " or AllLocations=1 "
      End If
    
      MyCommon.QueryStr = "select LocationGroupID,Name,PhraseID from LocationGroups as LG with (NoLock) where Deleted=0 and LocationGroupID not in " & _
                          "(select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ") " & wherestr & _
                          "and Deleted=0 and (EngineID=" & EngineID &  sAllLoc & " ) order by AllLocations Desc, Name"
    End If
    rst = MyCommon.LRT_Select
    
    If (rst.rows.count>0 )
      Sendb("var functionlist = Array(")
      For Each row In rst.Rows
        If IsDBNull(row.Item("PhraseID")) Then
          Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
        Else
          If (row.Item("PhraseID") = 0) Then
            Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
          Else
            Sendb("""" & Copient.PhraseLib.Lookup(row.item("PhraseID"), LanguageID) & """,")
          End If
        End If
      Next
      Send(""""");")
      Sendb("  var vallist = Array(")
      For Each row In rst.Rows
        Sendb("""" & row.item("LocationGroupID") & """,")
      Next
      Send(""""");")
      If (BannersEnabled) Then
        Sendb("  var bannerlist = Array(")
        For Each row In rst.Rows
          Sendb("""" & MyCommon.NZ(row.item("BannerName"), "") & """,")
        Next
        Send(""""");")
      End If
    Else
      Sendb("var functionlist = Array(")
      Send("""" & "" & """);")
      Sendb("  var vallist = Array(")
      Send("""" & "" & """);")
      If (BannersEnabled) Then
        Sendb("  var bannerlist = Array(")
        Send("""" & "" & """);")
      End If
    End If
  %>
  
  // Terminals Array 
  <%
    If (BannersEnabled) Then
      If EngineID = 0 Then
        MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, TT.AnyTerminal, TT.FuelProcessing, TT.PhraseID, BAN.Name as BannerName from TerminalTypes TT with (NoLock) " & _ 
                            "left join Banners BAN with (NoLock) on BAN.BannerID = TT.BannerID and BAN.Deleted=0 " & _
                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                            "and (TT.BannerID in (" & BannerListStr & ")) " & _
                            "union " & _
                            "select TT.TerminalTypeID, TT.Name, TT.AnyTerminal, TT.FuelProcessing, TT.PhraseID, '" & Copient.PhraseLib.Lookup("term.allbanners", LanguageID) & "' as BannerName from TerminalTypes TT with (NoLock) " & _ 
                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                            "and ((TT.BannerID=0 and AnyTerminal=0) or AnyTerminal=1)" & _
                            "order by BannerName, AnyTerminal desc, TT.Name;"
      Else
        MyCommon.QueryStr = "select TT.TerminalTypeID, TT.Name, TT.AnyTerminal, TT.FuelProcessing, TT.PhraseID, BAN.Name as BannerName from TerminalTypes TT with (NoLock) " & _ 
                            "left join Banners BAN with (NoLock) on BAN.BannerID = TT.BannerID and BAN.Deleted=0 " & _
                            "where TT.Deleted=0 and TT.TerminalTypeID not in " & _
                            "  (select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") and EngineID=" & EngineID & " " & _
                            "and (" & IIf(AllBannersPermission, "TT.BannerID=0 or", "") & " (TT.BannerID in ( " & BannerListStr & ") or AnyTerminal=1) ) " & _
                            "order by BannerName, AnyTerminal desc, TT.Name;"
      End If
    Else
      MyCommon.QueryStr = "select TerminalTypeID, FuelProcessing, Name, PhraseID from TerminalTypes with (NoLock) " & _
                          "where Deleted=0 and TerminalTypeID not in " & _
                          "(select TerminalTypeID from OfferTerminals with (NoLock) where OfferID=" & OfferID & ") " & _
                          "and EngineID=" & EngineID & " order by AnyTerminal desc, Name"
    End If
    
    rst = MyCommon.LRT_Select
    If (rst.rows.count>0)
      Sendb("var functionlist2 = Array(")
      For Each row In rst.Rows
        If IsDBNull(row.Item("PhraseID")) Then
          Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
        Else
          If (row.Item("PhraseID") = 0) Then
            Sendb("""" & MyCommon.NZ(row.item("Name"), "").ToString().Replace("""", "\""") & """,")
          Else
            Sendb("""" & Copient.PhraseLib.Lookup(row.item("PhraseID"), LanguageID) & """,")
          End If
        End If
      Next
      Send(""""");")
      If (BannersEnabled) Then
        Sendb("  var bannerlist2 = Array(")
        For Each row In rst.Rows
          Sendb("""" & MyCommon.NZ(row.item("BannerName"), "").ToString().Replace("""", "\""") & """,")
        Next
        Send(""""");")
      End If
      Sendb("var vallist2 = Array(")
      For Each row In rst.Rows
        Sendb("""" & row.item("TerminalTypeID") & """,")
      Next
      Send(""""");")
    Else
      Sendb("var functionlist2 = Array(")
      Send("""" & "" & """);")
      Sendb("var vallist2 = Array(")
      Send("""" & "" & """);")
      If (BannersEnabled) Then
        Sendb("  var bannerlist2 = Array(")
        Send("""" & "" & """);")
      End If
    End If
  %>
  
  // This is the function that refreshes the list after a keypress.
  // The maximum number to show can be limited to improve performance with
  // huge lists (1000s of entries).
  // The function clears the list, and then does a linear search through the
  // globally defined array and adds the matches back to the list.
  function handleKeyUp(maxNumToShow) {
    var selectObj, textObj, functionListLength;
    var i,  numShown;
    var searchPattern;
    var newOpt, optGp;
    
    document.getElementById("sgroups-available").size = "10";
    
    // Set references to the form elements
    selectObj = document.getElementById("sgroups-available");
    textObj = document.forms[0].functioninput;
    
    // Remember the function list length for loop speedup
    functionListLength = functionlist.length;
    
    // Set the search pattern depending
    if(document.forms[0].functionradio[0].checked == true) {
      searchPattern = "^"+textObj.value;
    } else {
      searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);
    
    // Create a regular expression
    re = new RegExp(searchPattern,"gi");
    
    // Clear the options list (for IE, cloning the select box without its option and then replacing the 
    // existing one, is significantly faster than removing each option.
    if (textObj.value == '' && fullAvailList != null)  {
      document.getElementById('sgrouplist').replaceChild(fullAvailList, selectObj);
    } else {
      //var newSelectBox = selectObj.cloneNode(false);
      if (isIE) {
        <%
          Send("var newSelectBox = document.createElement('select');")
          Send("newSelectBox.setAttribute('id', 'sgroups-available');")
          Send("newSelectBox.setAttribute('name', 'sgroups-available');")
          Send("newSelectBox.setAttribute('class', 'longest');")
          Send("newSelectBox.setAttribute('size', '10');")
          Send("newSelectBox.setAttribute('multiple', 'multiple');")
        %>
      } else {
        var newSelectBox = document.createElement('select');
        newSelectBox.id = 'sgroups-available';
        newSelectBox.name = 'sgroups-available';
        newSelectBox.className = 'longest';
        newSelectBox.size = '10';
        newSelectBox.multiple = true;
        <%

        %>
      }
      
      document.getElementById('sgrouplist').replaceChild(newSelectBox, selectObj);
      selectObj = document.getElementById("sgroups-available");
     
      // Loop through the array and re-add matching options
      numShown = 0;
      for(i = 0; i < functionListLength; i++) {
        if(functionlist[i].search(re) != -1) {
          if (vallist[i] != "") {
            var newOpt = document.createElement('OPTION');
            newOpt.value = vallist[i];
            if (isIE) { newOpt.innerText = functionlist[i]}; 
            newOpt.text =  functionlist[i]; 
            
            <% If (BannersEnabled) Then %>
              if (!isOpera) {
                optGp = GetOptionGroup(bannerlist[i], selectObj);
                if (optGp != null) {
                  optGp.appendChild(newOpt);
                  selectObj.appendChild(optGp);
                } else {
                  selectObj[numShown] = newOpt;
                }                
              } else {
                selectObj[numShown] = newOpt;
              }
            <% Else %>
              selectObj[numShown] = new Option(newOpt.text, newOpt.value);
            <% End If %>
            if (vallist[i] == 1) {
              selectObj[numShown].style.fontWeight = 'bold';
              selectObj[numShown].style.color = 'brown';
            }
            numShown++;
          }
        }
        // Stop when the number to show is reached
        if(numShown == maxNumToShow) {
          break;
        }
      }
      // When options list whittled to one, select that entry
      if(selectObj.length == 1) {
        try {
          selectObj.options[0].selected = true;
          enableEditButton(1);
        } catch (ex) {
          // ignore if unable to select (workaround for problem in IE 6)
        }
      }
    }
  }
  
  function GetOptionGroup(bannerName, elemSlct) {
    var elemGroup = null;
    
    if (bannerName == "") {
      elemGroup = document.getElementById("gpAllBanners");
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.unassigned", LanguageID)) %>";
        elemGroup.id = "gpAllBanners";
        elemSlct.appendChild(elemGroup);
      }
    } else {
      elemGroup = document.getElementById("gp" + bannerName);
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        elemGroup.label = bannerName;
        elemGroup.id = "gp" + bannerName;
        elemSlct.appendChild(elemGroup);
      }
    }
    return elemGroup;
  }
  
  // This is the function that refreshes the list after a keypress.
  // The maximum number to show can be limited to improve performance with
  // huge lists (1000s of entries).
  // The function clears the list, and then does a linear search through the
  // globally defined array and adds the matches back to the list.
  function handleKeyUp2(maxNumToShow) {
    var selectObj, textObj, functionList2Length;
    var i,  numShown;
    var searchPattern;
    var newOpt, optGp;
    var isAll = false;
    
    document.getElementById("terminals-available").size = "10";
    
    // Set references to the form elements
    selectObj = document.getElementById("terminals-available");
    textObj = document.forms[0].functioninput2;
    
    // Remember the function list length for loop speedup
    functionList2Length = functionlist2.length;
    
    // Set the search pattern depending
    if(document.forms[0].functionradio3[0].checked == true) {
      searchPattern = "^"+textObj.value;
    } else {
      searchPattern = textObj.value;
    }
    searchPattern = cleanRegExpString(searchPattern);
    
    // Create a regulare expression
    
    re = new RegExp(searchPattern,"gi");
    // Clear the options list
    selectObj.length = 0;
    
    // Loop through the array and re-add matching options
    numShown = 0;
    for(i = 0; i < functionList2Length; i++) {
      if(functionlist2[i].search(re) != -1) {
        if (vallist2[i] != "") {
          isAll = (functionlist2[i].toLowerCase() == "all cpe terminals" || functionlist2[i].toLowerCase() == "all cm terminals" || functionlist2[i].toLowerCase() == "all dp terminals")
          var newOpt = document.createElement('OPTION');
          newOpt.value = vallist2[i];
          if (isIE) { newOpt.innerText = functionlist2[i]}; 
          newOpt.text =  functionlist2[i]; 
          <% If (BannersEnabled) Then %>
            if (!isOpera) {
              optGp = GetTerminalOptionGroup(bannerlist2[i], selectObj, isAll);
              if (optGp != null) {
                optGp.appendChild(newOpt);
                selectObj.appendChild(optGp);
              } else {
                selectObj[numShown] = newOpt
              }                
            } else {
              selectObj[numShown] = newOpt
            }
          <% Else %>
            selectObj[numShown] = new Option(newOpt.text, newOpt.value);
          <% End If %>
          
          if (isAll) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
          }
          numShown++;
        }                        
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
    // When options list whittled to one, select that entry
    if(selectObj.length == 1) {
      selectObj.options[0].selected = true;
    }
  }
  
  function GetTerminalOptionGroup(bannerName, elemSlct, isAll) {
    var elemGroup = null;
    
    if (bannerName == "") {
      elemGroup = document.getElementById("gpTermAllBanners");
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        if (isAll) {
          elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.allbanners", LanguageID)) %>";
        } else {
          elemGroup.label = "<% Sendb(Copient.PhraseLib.Lookup("term.unassigned", LanguageID)) %>";
        }
        elemGroup.id = "gpTermAllBanners";
        elemSlct.appendChild(elemGroup);
      }
    } else {
      elemGroup = document.getElementById("gpTerm" + bannerName);
      if (elemGroup == null) {
        elemGroup = document.createElement("OPTGROUP");
        elemGroup.label = bannerName;
        elemGroup.id = "gpTerm" + bannerName;
        elemSlct.appendChild(elemGroup);
      }
    }
    return elemGroup;
  }
  
  function handleKeyDown(e, slctName) {
    var key = e.which ? e.which : e.keyCode;
    
    if (key == 40) {
      var elemSlct = document.getElementById(slctName);
      if (elemSlct != null) { elemSlct.focus(); }
    }
  }
  
//  function handleSlctKeyDown(e, slctName, btnName) {
//    var key = e.which ? e.which : e.keyCode;
//    
//    if (key == 13) {
//      var elemSlct = document.getElementById(slctName);
//      if (elemSlct != null && elemSlct.disabled == false) { 
//        var fireOnThis = document.getElementById(btnName);
//        if( document.createEvent ) {
//          var evObj = document.createEvent('MouseEvents');
//          evObj.initEvent( 'click', true, false );
//          fireOnThis.dispatchEvent(evObj);
//        } else if( document.createEventObject ) {
//          fireOnThis.fireEvent('onclick');
//        }
//      }
//      e.returnValue=false;
//      return false;
//    }
//  }
  
  function ConfirmGroupExclude() { 
    var elem = document.getElementById('sgroups-available');
    var selectedCt = 0;
    var i = 0;
    var excludedGroup = "";
    
    if (elem != null) {
      for (i=0; i<elem.length; i++) {
        if(elem.options[i].selected) {
          selectedCt++;
          if (excludedGroup=="" && elem.options[i].value != "1") {
      	    excludedGroup = elem.options[i].text;
          }
        }
      }
      if (selectedCt > 1) {
        var response = confirm("<% Sendb(Copient.PhraseLib.Lookup("message.alertexcludegroup", LanguageID)) %>" + " '" + excludedGroup + "'?");
        if (response) {
          //' do something
        } else {
          return false;
        }
      }
    }
  }
    
<%
  If EngineID <> 7 Then
    goto skipPDEjavascript
  End If
%>
  function saveForm() {
    var funcSel = document.getElementById('sgroups-available');
    var elSel = document.getElementById('sgroups-select');
    var exSel = document.getElementById('sgroups-exclude');
    var i,j;
    var selectList = "";
    var excludedList = "";
    var htmlContents = "";
    
    // assemble the list of values from the selected box
    for (i = elSel.length - 1; i>=0; i--) {
      if(elSel.options[i].value != ""){
        if(selectList != "") { selectList = selectList + ","; }
        selectList = selectList + elSel.options[i].value;
      }
    }
    for (i = exSel.length - 1; i>=0; i--) {
      if(exSel.options[i].value != ""){
        if(excludedList != "") { excludedList = excludedList + ","; }
        excludedList = excludedList + exSel.options[i].value;
      }
    }
    
    // time to build up the hidden variables to pass for saving
    htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
    htmlContents = htmlContents + "<input type=\"hidden\" name=\"exGroups\" value=" + excludedList + ">";
    document.getElementById("hiddenVals").innerHTML = htmlContents;
    
    document.getElementById("LocationGroupID").value = selectList;
    document.getElementById("ExcludedStores").value = excludedList; 
    
    // alert(htmlContents);
    return true;
  }
  
  function isExceptionGroup(groupID) {
    var bRetVal = false;
    
    for (var i=0; i < exceptlist.length && !bRetVal; i++) {
      bRetVal = (exceptlist[i] == groupID) 
    }
    
    return bRetVal;
  }
  
  function selectCreatedGroup(groupID) {
    var funcSel = document.getElementById('sgroups-available');
    var i,j;
    
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == groupID){
        funcSel.options[j].selected = true;
      }
    }
  }
  
  function createGroup(groupID) {
    document.location = "lgroup-edit.aspx?OfferID=<%Sendb(OfferID) %>&EngineID=7" + getGroupTokens();
  }
  
  function editGroup(mode) {
    var elem = null;
    var index = -1;
    var groupID = 0;
    
    if (mode == 1) {
      elem = document.getElementById("sgroups-available")
    } else if (mode == 2) {
      elem = document.getElementById("sgroups-select")
    } else if (mode == 3) {
      elem = document.getElementById("sgroups-exclude")
    }
    
    if (elem != null) {
      index = elem.selectedIndex;
      if (index > -1) {
        groupID = elem.options[index].value;
        
        if (groupID > 1) {
          document.location = "lgroup-edit.aspx?LocationGroupID=" + groupID + "&OfferID=<%Sendb(OfferID) %>&EngineID=7" + getGroupTokens();
        }
      }
    }
  }
  
  function enableEditButton(mode) {
    var elem = null;
    var elemButton = null;
    var index = -1;
    var groupID = 0;
    
    if (mode == 1) {
      elem = document.getElementById("sgroups-available");
      elemButton = document.getElementById("editAvailable");
      elemButton.disabled = false;
    } else if (mode == 2) {
      elem = document.getElementById("sgroups-select");
      elemButton = document.getElementById("editSelected");
    } else if (mode == 3) {
      elem = document.getElementById("sgroups-exclude");
      elemButton = document.getElementById("editExcluded");
    }
    
    if (elem != null) {
      index = elem.selectedIndex;
      
      if (index > -1) {
        groupID = elem.options[index].value;
        if (groupID <= 1) {
          elemButton.disabled = true;
        } else {
          elemButton.disabled = false;
        }
      } else {
        elemButton.disabled = true;
      }
    }
  }
  
  function handleCloseClick() {
    var saveChanges = false;
    var elem = null;
    
    if (hasUnsavedChanges) { 
      saveChanges = confirm('<%Sendb(Copient.PhraseLib.Lookup("PDEoffer-loc.save-on-close", LanguageID)) %>')
    }
    
    if (saveChanges) {
      // save the changes
      saveForm();
      elem = document.getElementById('SaveThenClose');
      if (elem != null) {
        elem.value = "true";
      }
      document.mainform.submit();
    } else {
      // close the customer window.
      top.closeIframePopup(true);
    }
  }
  
  function getGroupTokens() {
    var funcSel = document.getElementById('sgroups-available');
    var elSel = document.getElementById('sgroups-select');
    var exSel = document.getElementById('sgroups-exclude');
    var i,j;
    var selectList = "-1";
    var excludedList = "-1";
    var groupTokens = "";
    
    // assemble the list of values from the selected box
    for (i = elSel.length - 1; i>=0; i--) {
      if(elSel.options[i].value != ""){
        if(selectList != "") { selectList = selectList + ","; }
        selectList = selectList + elSel.options[i].value;
      }
    }
    for (i = exSel.length - 1; i>=0; i--) {
      if(exSel.options[i].value != ""){
        if(excludedList != "") { excludedList = excludedList + ","; }
        excludedList = excludedList + exSel.options[i].value;
      }
    }
    
    if (selectList != '') {
      groupTokens = '&slct=' + selectList;
    }
    if (excludedList != '') {
      groupTokens += '&ex=' + excludedList
    }
    if (hasUnsavedChanges) {
      groupTokens += '&condChanged=true';
    }
    
    return groupTokens;
  }
<%
skipPDEjavascript:
%>
</script>

<form action="#" id="mainform" name="mainform">
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
  <input type="hidden" id="Popup" name="Popup" value="<% Sendb(IIf(Popup, 1, 0)) %>" />
  <input type="hidden" id="PromoID" name="PromoID" value="<% Sendb(PromoID) %>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
  <input type="hidden" id="CustomerGroupID" name="CustomerGroupID" value="<% Sendb(LocationGroupID) %>" />
  <input type="hidden" id="ExcludedStores" name="ExcludedStores" value="<% Sendb(ExcludedStores) %>" />
  <input type="hidden" id="SaveThenClose" name="SaveThenClose" value="false" />
  <%
    Send("<div id=""intro"">")
    If (IsTemplate) Then
      Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID)
    Else
      Sendb("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID)
    End If
    'If EngineID = 7 Then
    '  Send(" " & StrConv(Copient.PhraseLib.Lookup("term.locations", LanguageID), VbStrConv.Lowercase) & "</h1>")
    'Else
      Send(": " & MyCommon.TruncateString(Name, 50) & "</h1>")
    'End If
    Send("<div id=""controls"">")
    'If (IsTemplate) AndAlso Logix.UserRoles.EditTemplates Then
      'Send_Save()
    'End If
    'If EngineID = 7 Then
    '  'If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then Send_Save()
      
    '  If Not IsTemplate Then
    '    If (Logix.UserRoles.EditOffer And Not (FromTemplate)) Then Send_Save()
    '  Else
    '    If (Logix.UserRoles.EditTemplates) Then Send_Save()
    '  End If

    '  Send(" <input type=""button"" name=""btnClose"" id=""btnClose"" class=""regular"" value=""" & Copient.PhraseLib.Lookup("term.close", LanguageID) & """onclick = ""handleCloseClick();"" />")
    'End If
    If MyCommon.Fetch_SystemOption(75) Then
      If (OfferID > 0 AndAlso Logix.UserRoles.AccessNotes AndAlso Not Popup) Then
        Send_NotesButton(3, OfferID, AdminUserID)
      End If
    End If
    Send("</div>")
    Send("</div>")
  %>
  <a name="h00" id="h00"></a>
  <div id="main">
    <%
      Select Case EngineID
        Case 2, 3, 5, 6, 9
        MyCommon.QueryStr = "select StatusFlag from CPE_Incentives where IncentiveID=" & OfferID & ";"
        Case Else
        MyCommon.QueryStr = "select StatusFlag from Offers where OfferID=" & OfferID & ";"
      End Select
      rst2 = MyCommon.LRT_Select
      StatusText = Logix.GetOfferStatus(OfferID, LanguageID, StatusCode)
      If Not IsTemplate Then
        If (rst2.Rows.Count > 0 AndAlso MyCommon.NZ(rst2.Rows(0).Item("StatusFlag"), 0) <> 2) Then
          If (StatusCode = Copient.LogixInc.STATUS_FLAGS.STATUS_ACTIVE) AndAlso (MyCommon.NZ(rst2.Rows(0).Item("StatusFlag"), 0) > 0) Then
            modMessage = Copient.PhraseLib.Lookup("alert.modpostdeploy", LanguageID)
            Send("<div id=""modbar"">" & modMessage & "</div>")
          End If
        End If
      End If
      If (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      End If
      
      ' Send the status bar, but not if it's a new offer or a template, or if there's already a modMessage being shown.
      If (Not IsTemplate) AndAlso (rst.Rows.Count > 0) AndAlso (modMessage = "") Then
        Select Case EngineID
          Case 2, 3, 6, 9
          MyCommon.QueryStr = "select incentiveId from CPE_Incentives with (NoLock) where CreatedDate = LastUpdate and IncentiveID=" & OfferID
          rst3 = MyCommon.LRT_Select
          If (rst3.Rows.Count = 0) Then
            Send_Status(OfferID, EngineID)
          End If
          Case Else
          MyCommon.QueryStr = "select OfferID from Offers with (NoLock) where CreatedDate = LastUpdate and OfferID=" & OfferID
          rst3 = MyCommon.LRT_Select
          If (rst3.Rows.Count = 0) Then
            Send_Status(OfferID)
          End If
        End Select
        End If
    %>
    <div id="<%Sendb(IIf(EngineID = 7, "column", "column1")) %>">
      <div class="box" id="storegroups">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.storegroups", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate) Then%>
        <!--
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="temp-stores" name="Disallow_Stores" />
          <label for="temp-stores"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID)) %></label>
        </span>
        -->
        <% End If%>
        <label for="sgroups-available">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID) & ":")%>
          </b>
        </label>
        <br clear="all" />
        <input type="radio" id="functionradio1" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <%
          'If (EngineID = 7 And (Logix.UserRoles.CreateStoreGroups Or Logix.UserRoles.EditStoreGroups)) Then
          '  Send("<div style=""float:right;"">")
          '  If Logix.UserRoles.CreateStoreGroups Then
          '    Send("  <input type=""button"" class=""short"" id=""create"" name=""create"" value=""" & Copient.PhraseLib.Lookup("term.create", LanguageID) & """ onclick=""createGroup(" & LocationGroupID & ")"" />")
          '  End If
          '  If Logix.UserRoles.EditStoreGroups Then
          '    Send("  <input type=""button"" class=""short"" id=""editAvailable"" name=""editAvailable"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""editGroup(1)"" disabled=""disabled"" />")
          '  End If
          '  Send("</div>")
          'End If
        %>
        <input class="medium" onkeydown="handleKeyDown(event, 'sgroups-available');" onkeyup="handleKeyUp(200);" id="functioninput" name="functioninput" maxlength="100" type="text" value="" />
          <% If (bCreateGroupOrProgramFromOffer AndAlso Logix.UserRoles.CreateStoreGroups) Then%>
        <input class="regular" name="btncreate" id="btncreate" type="button" value="<% Sendb(Copient.PhraseLib.Lookup("term.create", LanguageID))%>" onclick="javascript:handleCreateClick('btncreate');" />
         <div id="searchLoadDiv" style="display: block;">
        &nbsp;</div>
        <% End If%>
        
        <br class="half" />
        <%
          Send("<span id=""sgrouplist"">")
          Send("<select class=""longest"" multiple=""multiple"" id=""sgroups-available"" name=""sgroups-available"" size=""10""" & IIf(EngineID = 7, " onchange=""enableEditButton(1)"" style=""width:650px;""", "") & ">")
          If bStoreUser Then
            wherestr = " and (LocationGroupID in (select LocationGroupID from LocGroupItems where LocationID in (" & sValidLocIDs & "))) "
          Else
            sAllLoc = " or AllLocations=1 "
          End If
          MyCommon.QueryStr = "select LocationGroupID,Name,PhraseID from LocationGroups as LG with (NoLock) where Deleted=0 and LocationGroupID not in " & _
                              "(select LocationGroupID from OfferLocations with (NoLock) where Deleted=0 and OfferID=" & OfferID & ")" & wherestr & " and Deleted=0 and (EngineID=" & EngineID & sAllLoc & ") order by AllLocations desc, Name"
          rst = MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            For Each row In rst.Rows
              If IsDBNull(row.Item("PhraseID")) Then
                Send("<option value=""" & row.Item("LocationGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
              Else
                If (row.Item("PhraseID") = 0) Then
                  Send("<option value=""" & row.Item("LocationGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                Else
                  Send("<option value=""" & row.Item("LocationGroupID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                End If
              End If
            Next
          End If
          Send("</select>")
          Send("</span>")
          Send("<br />")
          Send("<br />")
          
          ' First off: queries to get both the selected and excluded store groups
          If (BannersEnabled) Then
            MyCommon.QueryStr = "select O.LocationGroupID, LG.Name, LG.PhraseID, BAN.Name as BannerName from OfferLocations as O with (NoLock) " & _
                                "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=O.LocationGroupID " & _
                                "left join Banners as BAN with (NoLock) on LG.BannerID = BAN.BannerID and BAN.Deleted=0 " & _
                                "where Excluded=0 and O.Deleted=0 and O.OfferID=" & OfferID & " order by BAN.Name, LG.Name"
          Else
            MyCommon.QueryStr = "select O.LocationGroupID, LG.Name, LG.PhraseID from OfferLocations as O with (NoLock) " & _
                                "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=O.LocationGroupID " & _
                                "where Excluded=0 and O.Deleted=0 and O.OfferID=" & OfferID & " order by LG.Name"
          End If
          rstSelected = MyCommon.LRT_Select
          MyCommon.QueryStr = "select O.LocationGroupID,LG.Name,LG.PhraseID from OfferLocations as O with (NoLock) left join LocationGroups as LG with (NoLock) on " & _
                              "LG.LocationGroupID=O.LocationGroupID where Excluded=1 and O.deleted=0 and O.OfferID=" & OfferID & " order by LG.Name"
          rstExcluded = MyCommon.LRT_Select
          
          ' SELECTED STORE GROUPS
          Send("<label for=""sgroups-select""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label>")
          Send("<br />")
          
          ' Buttons
          'If (EngineID = 7 And Logix.UserRoles.EditCustomerGroups) Then
          '  Send("<div style=""float:right;"">")
          '  Send("  <input type=""button"" class=""short"" id=""editSelected"" name=""editSelected"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""editGroup(2)"" disabled=""disabled"" />")
          '  Send("</div>")
          'End If
          Sendb("<input type=""submit"" class=""regular select"" id=""stores-add1"" name=""stores-add1"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
          If Not IsTemplate Then
            If (Not Logix.UserRoles.CRUDLocationsToOffers) Then
              Sendb(" disabled=""disabled""")
            End If
          Else
             If (Not Logix.UserRoles.CRUDLocationsToTemplates) Then
              Sendb(" disabled=""disabled""")
            End If
          End If
          Send(" />")
          Sendb("<input type=""submit"" class=""regular deselect"" id=""stores-rem1"" name=""stores-rem1"" title=""" & Copient.PhraseLib.Lookup("term.unselect", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
          If Not IsTemplate Then
            If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) Then
              Sendb(" disabled=""disabled""")
            End If
          Else
             If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) Then
               Sendb(" disabled=""disabled""")
             End If
          End If
          Send(" />")
          Send("<br />")
          
          ' List
          Send("<select class=""longest"" multiple=""multiple"" id=""sgroups-select"" name=""sgroups-select"" size=""" & SelectSize & """" & IIf(EngineID = 7, " onchange=""enableEditButton(2)"" style=""height:38px; margin-top:2px; width:650px;""", "") & ">")
          If (rstSelected.Rows.Count > 0) Then
            For Each row In rstSelected.Rows
              If (BannersEnabled) Then
                If row.Item("LocationGroupID") = 1 Then
                  NameValue = Copient.PhraseLib.Lookup("term.all", LanguageID)
                Else
                  NameValue = MyCommon.NZ(row.Item("BannerName"), Copient.PhraseLib.Lookup("term.unassigned", LanguageID))
                End If
                If (NameValue <> PrevNameValue) Then
                  If (PrevNameValue <> "") Then Send("</optgroup>")
                  Send("   <optgroup label=""" & NameValue & """>  ")
                End If
                PrevNameValue = NameValue
              End If
              If IsDBNull(row.Item("PhraseID")) Then
                Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
              Else
                If (row.Item("PhraseID") = 0) Then
                  Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
                Else
                  Send("<option value=""" & row.Item("LocationGroupID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
                End If
              End If
            Next
            If (BannersEnabled) Then Send("</optgroup>")
          End If
          Send("</select>")
          Send("<br />")
          Send("<br class=""half"" />")
          
          ' EXCLUDED STORE GROUPS
          Send("<label for=""sgroups-exclude""><b>" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ":</b></label>")
          Send("<br />")
          
          ' Buttons
          'If (EngineID = 7 And Logix.UserRoles.EditCustomerGroups) Then
          '  Send("<div style=""float:right;"">")
          '  Send("  <input type=""button"" class=""short"" id=""editExcluded"" name=""editExcluded"" value=""" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & """ onclick=""editGroup(3)"" disabled=""disabled"" />")
          '  Send("</div>")
          'End If
          Sendb("<input type=""submit"" class=""regular select"" id=""stores-add2"" name=""stores-add2"" onclick=""return ConfirmGroupExclude();"" title=""" & Copient.PhraseLib.Lookup("term.exclude", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
          If Not IsTemplate Then
            If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("LocationGroupID") <> 1) Then
              Sendb(" disabled=""disabled""")
            End If
          Else
            If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("LocationGroupID") <> 1) Then
              Sendb(" disabled=""disabled""")
            End If 
          End If     
          Send(" />")
          Sendb("<input type=""submit"" class=""regular deselect"" id=""stores-rem2"" name=""stores-rem2"" title=""" & Copient.PhraseLib.Lookup("term.unexclude", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
          If Not IsTemplate Then
            If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstExcluded.Rows.Count = 0) Then
              Sendb(" disabled=""disabled""")
            End If  
          Else
            If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstExcluded.Rows.Count = 0) Then
              Sendb(" disabled=""disabled""")
            End If
          End If
          Send(" />")
          Send("<br />")
          
          ' List
          Send("<select class=""longest"" id=""sgroups-exclude"" name=""sgroups-exclude"" size=""5""" & IIf(EngineID = 7, " onchange=""enableEditButton(3)"" style=""height:38px; margin-top:2px; width:650px;""", "") & ">")
          For Each row In rstExcluded.Rows
            If IsDBNull(row.Item("PhraseID")) Then
              Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
            Else
              If (row.Item("PhraseID") = 0) Then
                Send("<option value=""" & row.Item("LocationGroupID") & """>" & row.Item("Name") & "</option>")
              Else
                Send("<option value=""" & row.Item("LocationGroupID") & """>" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
              End If
            End If
          Next
          Send("</select>")
        %>
        <hr class="hidden" />
      </div>
      <a name="h01"></a>
      <%
        'If (EngineID = 4) Then
        '  PaddingValue = 50
        '  BtnPaddingTop += PaddingValue + 30
        '  SelectSize = 10
        'End If
      %>
    </div>
    
    <div id="gutter"<%Sendb(IIf(EngineID = 7, " style=""display:none;""", "")) %>>
    </div>
    
    <div id="column2"<%Sendb(IIf(EngineID = 7, " style=""display:none;""", "")) %>>
      <div class="box" id="terminals">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.terminals", LanguageID))%>
          </span>
        </h2>
        <% If (IsTemplate) Then%>
        <!--
        <span class="temp">
          <input type="checkbox" class="tempcheck" id="temp-terminals" name="Disallow_Terminals" />
          <label for="temp-terminals"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID)) %></label>
        </span>
        -->
        <% End If%>
        <label for="terminals-available">
          <b>
            <% Sendb(Copient.PhraseLib.Lookup("term.available", LanguageID) & ":")%>
          </b>
        </label>
        <br clear="all" />
        <input type="radio" id="functionradio3" name="functionradio3" <% if(MyCommon.Fetch_SystemOption(175)= "1" ) then sendb(" checked=""checked""") %> /><label for="functionradio3"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio4" name="functionradio3" <% if(MyCommon.Fetch_SystemOption(175)= "2" ) then sendb(" checked=""checked""") %> /><label for="functionradio4"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
        <input class="longer" onkeydown="handleKeyDown(event, 'terminals-available');" onkeyup="handleKeyUp2(9999);" id="functioninput2" name="functioninput2" maxlength="100" type="text" value="" /><br />
        <br class="half" />
        <select class="longest" multiple="multiple" id="terminals-available" name="terminals-available" size="10">
        </select>
        <br />
        <br />
        <%
          ' First off: queries to get both the selected and excluded terminals
          If (BannersEnabled) Then
            MyCommon.QueryStr = "select OT.TerminalTypeID as TID, T.Name, T.FuelProcessing, T.PhraseID, BAN.Name as BannerName, T.AnyTerminal,T.BannerID from OfferTerminals as OT with (NoLock) " & _
                                "left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                "left join Banners BAN with (NoLock) on BAN.BannerID = T.BannerID and BAN.Deleted=0 " & _
                                "where OfferID=" & OfferID & " and Excluded=0" & _
                                "order by AnyTerminal desc, BAN.Name, T.Name"
          Else
            MyCommon.QueryStr = "select OT.TerminalTypeID as TID, T.Name, T.FuelProcessing, T.PhraseID, T.AnyTerminal from OfferTerminals as OT with (NoLock) " & _
                                "left join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                                "where OfferID=" & OfferID & " and Excluded=0" & _
                                "order by AnyTerminal desc, T.Name"
          End If
          rstSelected = MyCommon.LRT_Select
          MyCommon.QueryStr = "select OT.PKID, T.Name, T.FuelProcessing from OfferTerminals as OT with (nolock) " & _
                              "left join TerminalTypes as T with (NoLock) on T.TerminalTypeID=OT.TerminalTypeID " & _
                              "where OT.OfferID=" & OfferID & " and OT.Excluded=1" & _
                              "order by T.Name"
          rstExcluded = MyCommon.LRT_Select
          
          ' SELECTED TERMINALS
          Send("<label for=""terminals-select""><b>" & Copient.PhraseLib.Lookup("term.selected", LanguageID) & ":</b></label>")
          Send("<br />")
          
          ' Buttons
          Sendb("<input type=""submit"" class=""regular select"" id=""terminals-add1"" name=""terminals-add1"" title=""" & Copient.PhraseLib.Lookup("term.select", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
          If Not IsTemplate Then
            If (Not Logix.UserRoles.CRUDLocationsToOffers) Then
              Sendb(" disabled=""disabled""")
            End If
          Else
            If (Not Logix.UserRoles.CRUDLocationsToTemplates) Then
              Sendb(" disabled=""disabled""")
            End If
          End If
          Send(" />")
          Sendb("<input type=""submit"" class=""regular deselect"" id=""terminals-rem1"" name=""terminals-rem1"" title=""" & Copient.PhraseLib.Lookup("term.unselect", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
          If Not IsTemplate Then
             If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) Then
               Sendb(" disabled=""disabled""")
             End If
          Else
             If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) Then
               Sendb(" disabled=""disabled""")
             End If
          End If
          Send(" />")
          Send("<br />")
          
          ' List
          Send("<select class=""longest"" multiple=""multiple"" id=""terminals-select"" name=""terminals-select"" size=""" & SelectSize & """>")
          If rstSelected.Rows.Count = 0 Then
            'Send("<option value="""">All Terminals</option>")
          Else
            PrevNameValue = ""
            For Each row In rstSelected.Rows
              If (BannersEnabled) Then
                If (MyCommon.NZ(row.Item("AnyTerminal"), False)) Then
                  NameValue = Copient.PhraseLib.Lookup("term.allbanners", LanguageID)
                Else
                  If (EngineID = 0 And (MyCommon.NZ(row.Item("BannerId"), 0)) = 0) Then
                    NameValue = Copient.PhraseLib.Lookup("term.allbanners", LanguageID)
                  Else
                    NameValue = MyCommon.NZ(row.Item("BannerName"), Copient.PhraseLib.Lookup("term.unassigned", LanguageID))
                  End If
                End If
                If (NameValue <> PrevNameValue) Then
                  If (PrevNameValue <> "") Then Send("</optgroup>")
                  Send("   <optgroup label=""" & NameValue & """>  ")
                End If
                PrevNameValue = NameValue
              End If
              If (row.Item("Name").ToString.Trim = "All CPE Terminals" OrElse row.Item("Name").ToString.Trim = "All CM Terminals" OrElse row.Item("Name").ToString.Trim = "All DP Terminals") Then
                Send("<option value=""" & row.Item("TID") & """ style=""font-weight:bold;color:brown;"">" & Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "</option>")
              Else
                Send("<option value=""" & row.Item("TID") & """>" & row.Item("Name") & "</option>")
              End If
            Next
            If (BannersEnabled) Then Send("</optgroup>")
          End If
          Send("</select>")
          Send("<br />")
          Send("<br class=""half"" />")
          
          ' EXCLUDED TERMINALS
          Send("<div>")
          Send("<label for=""terminals-exclude""><b>" & Copient.PhraseLib.Lookup("term.excluded", LanguageID) & ":</b></label>")
          Send("<br />")
          
          'Buttons
          Sendb("<input type=""submit"" class=""regular select"" id=""terminals-add2"" name=""terminals-add2"" title=""" & Copient.PhraseLib.Lookup("term.exclude", LanguageID) & """ value=""&#9660; " & Copient.PhraseLib.Lookup("term.select", LanguageID) & """")
          If Not IsTemplate Then
            If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("AnyTerminal") = False) Then
              Sendb(" disabled=""disabled""")
            End If
          Else
            If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstSelected.Rows.Count = 0) OrElse (rstSelected.Rows(0).Item("AnyTerminal") = False) Then
              Sendb(" disabled=""disabled""")
            End If
          End If
          Send(" />")
          Sendb("<input type=""submit"" class=""regular deselect"" id=""terminals-rem2"" name=""terminals-rem2"" title=""" & Copient.PhraseLib.Lookup("term.unexclude", LanguageID) & """ value=""" & Copient.PhraseLib.Lookup("term.deselect", LanguageID) & " &#9650;""")
          If Not IsTemplate Then
            If (Not Logix.UserRoles.CRUDLocationsToOffers) OrElse (rstExcluded.Rows.Count = 0) Then
              Sendb(" disabled=""disabled""")
            End If
          Else
            If (Not Logix.UserRoles.CRUDLocationsToTemplates) OrElse (rstExcluded.Rows.Count = 0) Then
              Sendb(" disabled=""disabled""")
            End If
          End If
          Send(" />")
          Send("<br />")
          
          ' List
          Send("<select class=""longest"" multiple=""multiple"" id=""terminals-exclude"" name=""terminals-exclude"" size=""5"">")
          For Each row In rstExcluded.Rows
            If (row.Item("Name").ToString.Trim = "All CPE Terminals" OrElse row.Item("Name").ToString.Trim = "All CM Terminals" OrElse row.Item("Name").ToString.Trim = "All DP Terminals") Then
              Send("<option value=""" & row.Item("PKID") & """ style=""font-weight:bold;color:brown;"">" & row.Item("Name") & "</option>")
            Else
              Send("<option value=""" & row.Item("PKID") & """>" & row.Item("Name") & "</option>")
            End If
          Next
          Send("</select>")
          Send("</div>")
        %>
      </div>
    </div>
    <br clear="all" />
  </div>
</form>

<script type="text/javascript">

<% If (Request.QueryString("save") <> "" OrElse Request.QueryString("SaveThenClose") = "true") Then %>
  top.drawStores(<%Sendb(OfferID)%>);
  <%If (CloseAfterSave) Then %>
   top.closeIframePopup(false);
  <% End If %>
<%  End If %>

<% If (BannersEnabled) Then %>
  handleKeyUp(9999);
<% End If %>
  handleKeyUp2(9999);
  
  if (document.getElementById("sgroups-available") != null) {
    fullAvailList = document.getElementById("sgroups-available").cloneNode(true);
  }
  
  <% If (infoMessage <> "") Then %>
    // DOM2
    if (typeof window.addEventListener != "undefined")
      window.addEventListener( "load", JumpToTop, false );
    // IE
    else if (typeof window.attachEvent != "undefined") {
      window.attachEvent( "onload", JumpToTop );
    } else {
      if (window.onload != null) {
        var oldOnload = window.onload;
        window.onload = function ( e ) {
        oldOnload(e);
          JumpToTop();
        };
      } else
        window.onload = InitialiseScrollableArea;
    }
    function JumpToTop() {
      try {
        window.location.hash = 'h00';
        document.getElementById('main').scrollTop = 0;
      } catch (err) {
        // ignore
      }
    }
  <% End If %>
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (OfferID > 0 AndAlso Logix.UserRoles.AccessNotes) Then
      Send_Notes(3, OfferID, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  'Send_BodyEnd("mainform", IIf(FocusType = 0, "functioninput", "functioninput2"))
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
