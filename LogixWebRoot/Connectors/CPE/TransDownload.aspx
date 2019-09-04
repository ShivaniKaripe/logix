<%
  ' version:7.3.1.138972.Official Build (SUSDAY10202)
  ' *****************************************************************************
  ' * FILENAME: TransDownload.aspx 
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
%>

<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%@ Import Namespace="System.Data" %>
<script runat="server">
  Public Common As New Copient.CommonInc
  Public Connector As New Copient.ConnectorInc
  Public GZIP As New Copient.GZIPInc
  Public OutboundBuffer As StringBuilder
  Public TextData As String
  Public LogFile As String
  Dim StartTime As Object
  Public MacAddress As String
    Public LocalServerIP As String
    Public MyCryptLib As New Copient.CryptLib
  ' -------------------------------------------------------------------------------------------------


  Sub SD(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr & vbCrLf)
  End Sub


  ' -------------------------------------------------------------------------------------------------


  Sub SDb(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr)
  End Sub


  ' -----------------------------------------------------------------------------------------------


  Function Parse_Bit(ByVal BooleanField As Boolean) As String
    If BooleanField Then
      Parse_Bit = "1"
    Else
      Parse_Bit = "0"
    End If
  End Function


  ' -----------------------------------------------------------------------------------------------


  Sub Construct_Table(ByVal TableName As String, ByVal Operation As String, ByVal DelimChar As String, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal dst As DataTable, Optional ByVal ColumnNamesOverride As String = "")
    
    Dim TempResults As String
    Dim NumRecs As Long
    Dim row As DataRow
    Dim SQLCol As DataColumn
    Dim TempOut As String
    Dim Index As Integer
    Dim FieldList As String
    
    TempOut = ""
    TempResults = ""
    NumRecs = 0
    FieldList = ""
    If dst.Rows.Count > 0 Then
      If ColumnNamesOverride = "" Then
      For Each SQLCol In dst.Columns
        If Not (FieldList = "") Then FieldList = FieldList & DelimChar
        FieldList = FieldList & SQLCol.ColumnName
      Next
      Else
        FieldList = ColumnNamesOverride
      End If
      TempOut = "1:" & TableName & vbCrLf
      TempOut = TempOut & "2:" & Operation & vbCrLf
      TempOut = TempOut & "3:" & FieldList
      SD(TempOut)
      Common.Write_Log(LogFile, TempOut)
      
      If UCase(TableName) = "USERS" And Operation = 1 Then
        For Each row In dst.Rows
                    SD(Common.NZ(row.Item("UserID"), 0) & DelimChar & MyCryptLib.SQL_StringDecrypt(Common.NZ(row.Item("ClientUserID1"), " ")) & DelimChar & Common.NZ(row.Item("HHPrimaryID"), 0) & DelimChar & Common.NZ(row.Item("HHrec"), 0) & DelimChar & Common.NZ(row.Item("CustomerTypeID"), 0) & DelimChar & row.Item("CustomerStatusID") & DelimChar & row.Item("AlternateID") & DelimChar & row.Item("Verifier") & DelimChar & Parse_Bit(row.Item("Employee")) & DelimChar & row.Item("AltIDOptOut") & DelimChar & row.Item("FirstName") & DelimChar & row.Item("LastName") & DelimChar & row.Item("EmployeeID") & DelimChar & row.Item("AirmileMemberID") & DelimChar & row.Item("Prefix") & DelimChar & row.Item("Suffix"))
          Common.Write_Log(LogFile, Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("ClientUserID1"), " ") & DelimChar & Common.NZ(row.Item("HHPrimaryID"), 0) & DelimChar & Common.NZ(row.Item("HHrec"), 0) & DelimChar & Common.NZ(row.Item("CustomerTypeID"), 0) & DelimChar & row.Item("CustomerStatusID") & DelimChar & row.Item("AlternateID") & DelimChar & row.Item("Verifier") & DelimChar & Parse_Bit(row.Item("Employee")) & DelimChar & row.Item("AltIDOptOut") & DelimChar & row.Item("FirstName") & DelimChar & row.Item("LastName") & DelimChar & row.Item("EmployeeID") & DelimChar & row.Item("AirmileMemberID") & DelimChar & row.Item("Prefix") & DelimChar & row.Item("Suffix"))
        Next
        
      ElseIf UCase(TableName) = "CARDIDS" And Operation = 2 Then
        For Each row In dst.Rows
          SD(row.Item("UserID"))
          Common.Write_Log(LogFile, row.Item("UserID"))
        Next
        
      ElseIf UCase(TableName) = "CARDIDS" And (Operation = 1 Or Operation = 6) Then
        For Each row In dst.Rows
                    SD(row.Item("CardPK") & DelimChar & row.Item("UserID") & DelimChar & MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID")) & DelimChar & row.Item("CardStatusID") & DelimChar & row.Item("CardTypeID"))
          Common.Write_Log(LogFile, row.Item("CardPK") & DelimChar & row.Item("UserID") & DelimChar & MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID")) & DelimChar & row.Item("CardStatusID") & DelimChar & row.Item("CardTypeID"))
        Next
        
      ElseIf UCase(TableName) = "CUSTOMERATTRIBUTES" And Operation = 2 Then
        For Each row In dst.Rows
          SD(row.Item("AttributeTypeID"))
          Common.Write_Log(LogFile, row.Item("AttributeTypeID"))
        Next
        
      ElseIf UCase(TableName) = "CUSTOMERATTRIBUTES" And Operation = 12 And dst.Columns(0).ColumnName = "AttributeTypeID" Then
        For Each row In dst.Rows
          SD(row.Item("AttributeTypeID") & DelimChar & row.Item("AttributeValueID"))
          Common.Write_Log(LogFile, row.Item("AttributeTypeID") & DelimChar & row.Item("AttributeValueID"))
        Next
        
      ElseIf UCase(TableName) = "CUSTOMERATTRIBUTES" And Operation = 12 And dst.Columns(0).ColumnName = "CustomerPK" Then
        For Each row In dst.Rows
          SD(row.Item("CustomerPK") & DelimChar & row.Item("AttributeTypeID"))
          Common.Write_Log(LogFile, row.Item("CustomerPK") & DelimChar & row.Item("AttributeTypeID"))
        Next
        
      ElseIf UCase(TableName) = "CUSTOMERATTRIBUTES" And Operation = 7 Then
        For Each row In dst.Rows
          SD(row.Item("CustomerPK") & DelimChar & row.Item("AttributeTypeID") & DelimChar & row.Item("AttributeValueID"))
          Common.Write_Log(LogFile, row.Item("CustomerPK") & DelimChar & row.Item("AttributeTypeID") & DelimChar & row.Item("AttributeValueID"))
        Next
        
      ElseIf UCase(TableName) = "STDADJ" And (Operation = 5) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("AdjAmount"), 0))
          Common.Write_Log(LogFile, Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("AdjAmount"), 0))
        Next
        
      ElseIf UCase(TableName) = "USERREMOVAL" And Operation = 2 Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("UserID"), 0))
          Common.Write_Log(LogFile, Common.NZ(row.Item("UserID"), 0))
        Next
        
      ElseIf UCase(TableName) = "GROUPMEMBERSHIP" And (Operation = 1 Or Operation = 7) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("UserGroupID"), 0))
          Common.Write_Log(LogFile, Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("UserGroupID"), 0))
        Next
        
      ElseIf UCase(TableName) = "REWARDDISTRIBUTION" And (Operation = 1 Or Operation = 7) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("LocalID"), 0) & DelimChar & Common.NZ(row.Item("ServerSerial"), 0) & DelimChar & Common.NZ(row.Item("IncentiveID"), 0) & DelimChar & Common.NZ(row.Item("RewardOptionID"), 0) & DelimChar & Common.NZ(row.Item("Phase"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("DistributionDate"), "") & DelimChar & Common.NZ(row.Item("WaitingACK"), 0))
          Common.Write_Log(LogFile, Common.NZ(row.Item("LocalID"), 0) & DelimChar & Common.NZ(row.Item("ServerSerial"), 0) & DelimChar & Common.NZ(row.Item("IncentiveID"), 0) & DelimChar & Common.NZ(row.Item("RewardOptionID"), 0) & DelimChar & Common.NZ(row.Item("Phase"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("DistributionDate"), "") & DelimChar & Common.NZ(row.Item("WaitingACK"), 0))
        Next
        
      ElseIf UCase(TableName) = "POINTSADJ" And (Operation = 5) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("ProgramID"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("AdjAmount"), 0))
          Common.Write_Log(LogFile, Common.NZ(row.Item("ProgramID"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("AdjAmount"), 0))
        Next
        
      ElseIf UCase(TableName) = "CUSTOMERPREFERENCES" And (Operation = 7) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("CustomerPK"), 0) & DelimChar & Common.NZ(row.Item("PreferenceID"), 0) & DelimChar & Common.NZ(row.Item("Value"), ""))
          Common.Write_Log(LogFile, Common.NZ(row.Item("CustomerPK"), 0) & DelimChar & Common.NZ(row.Item("PreferenceID"), 0) & DelimChar & Common.NZ(row.Item("Value"), ""))
        Next
        
      ElseIf (UCase(TableName) = "CUSTOMERPREFERENCES" Or UCase(TableName) = "CUSTOMERPREFERENCESMV") And (Operation = 12) Then
        For Each row In dst.Rows
          SD(row.Item("PreferenceID") & DelimChar & row.Item("Value"))
          Common.Write_Log(LogFile, row.Item("PreferenceID") & DelimChar & row.Item("Value"))
        Next

        
      ElseIf UCase(TableName) = "CUSTOMERPREFERENCESMV" And (Operation = 8) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("CustomerPK"), 0) & DelimChar & Common.NZ(row.Item("PreferenceID"), 0) & DelimChar & Common.NZ(row.Item("Value"), ""))
          Common.Write_Log(LogFile, Common.NZ(row.Item("CustomerPK"), 0) & DelimChar & Common.NZ(row.Item("PreferenceID"), 0) & DelimChar & Common.NZ(row.Item("Value"), ""))
        Next

      ElseIf UCase(TableName) = "CUSTOMERPREFERENCESMV" And (Operation = 2) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("CustomerPK"), 0) & DelimChar & Common.NZ(row.Item("PreferenceID"), 0) & DelimChar & Common.NZ(row.Item("Value"), ""))
          Common.Write_Log(LogFile, Common.NZ(row.Item("CustomerPK"), 0) & DelimChar & Common.NZ(row.Item("PreferenceID"), 0) & DelimChar & Common.NZ(row.Item("Value"), ""))
        Next

      ElseIf UCase(TableName) = "REWARDACCUMULATION" And (Operation = 1 Or Operation = 7) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("LocalID"), 0) & DelimChar & Common.NZ(row.Item("ServerSerial"), 0) & DelimChar & Common.NZ(row.Item("RewardOptionID"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("QtyPurchased"), 0) & DelimChar & Common.NZ(row.Item("TotalPrice"), 0) & DelimChar & Common.NZ(row.Item("AccumulationDate"), "") & DelimChar & Parse_Bit(row.Item("OverThreshold")) & DelimChar & Common.NZ(row.Item("WaitingACK"), 0))
          Common.Write_Log(LogFile, Common.NZ(row.Item("LocalID"), 0) & DelimChar & Common.NZ(row.Item("ServerSerial"), 0) & DelimChar & Common.NZ(row.Item("RewardOptionID"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("QtyPurchased"), 0) & DelimChar & Common.NZ(row.Item("TotalPrice"), 0) & DelimChar & Common.NZ(row.Item("AccumulationDate"), "") & DelimChar & Parse_Bit(row.Item("OverThreshold")) & DelimChar & Common.NZ(row.Item("WaitingACK"), 0))
        Next
        
      ElseIf UCase(TableName) = "STOREDVALUE" And (Operation = 7) Then
        For Each row In dst.Rows
          SD(row.Item("LocalID") & DelimChar & row.Item("ServerSerial") & DelimChar & row.Item("SVProgramID") & DelimChar & row.Item("IncentiveID") & DelimChar & row.Item("CustomerPK") & DelimChar & row.Item("QtyEarned") & DelimChar & row.Item("QtyUsed") & DelimChar & row.Item("Value") & DelimChar & row.Item("EarnedDate") & DelimChar & row.Item("EarnedLocationID") & DelimChar & row.Item("ExpireDate") & DelimChar & row.Item("ExternalID"))
          Common.Write_Log(LogFile, row.Item("LocalID") & DelimChar & row.Item("ServerSerial") & DelimChar & row.Item("SVProgramID") & DelimChar & row.Item("IncentiveID") & DelimChar & row.Item("CustomerPK") & DelimChar & row.Item("QtyEarned") & DelimChar & row.Item("QtyUsed") & DelimChar & row.Item("Value") & DelimChar & row.Item("EarnedDate") & DelimChar & row.Item("EarnedLocationID") & DelimChar & row.Item("ExpireDate") & DelimChar & row.Item("ExternalID"))
        Next
        
      ElseIf UCase(TableName) = "SVAdj" And (Operation = 5) Then
        For Each row In dst.Rows
          SD(row.Item("LocalID") & DelimChar & row.Item("ServerSerial") & DelimChar & row.Item("NewExternalID") & DelimChar & row.Item("QtyUsed") & DelimChar & row.Item("LastUpdate"))
          Common.Write_Log(LogFile, row.Item("LocalID") & DelimChar & row.Item("ServerSerial") & DelimChar & row.Item("QtyUsed") & DelimChar & row.Item("StatusFlag") & DelimChar & row.Item("LastUpdate"))
        Next
        
      ElseIf UCase(TableName) = "STOREDVALUE" And (Operation = 11) Then
        For Each row In dst.Rows
          SD(row.Item("LocalID") & DelimChar & row.Item("ServerSerial") & DelimChar & row.Item("CustomerPK"))
          Common.Write_Log(LogFile, row.Item("LocalID") & DelimChar & row.Item("ServerSerial") & DelimChar & row.Item("CustomerPK"))
        Next
        
      ElseIf UCase(TableName) = "USERRESPONSES" And (Operation = 1 Or Operation = 7) Then
        For Each row In dst.Rows
          SD(Common.NZ(row.Item("LocalID"), 0) & DelimChar & Common.NZ(row.Item("ServerSerial"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("IncentiveID"), 0) & DelimChar & Common.NZ(row.Item("OnScreenAdID"), 0) & DelimChar & Common.NZ(row.Item("RewardOptionID"), 0) & DelimChar & Common.NZ(row.Item("Response"), "") & DelimChar & Common.NZ(row.Item("WaitingACK"), 0))
          Common.Write_Log(LogFile, Common.NZ(row.Item("LocalID"), 0) & DelimChar & Common.NZ(row.Item("ServerSerial"), 0) & DelimChar & Common.NZ(row.Item("UserID"), 0) & DelimChar & Common.NZ(row.Item("IncentiveID"), 0) & DelimChar & Common.NZ(row.Item("OnScreenAdID"), 0) & DelimChar & Common.NZ(row.Item("RewardOptionID"), 0) & DelimChar & Common.NZ(row.Item("Response"), "") & DelimChar & Common.NZ(row.Item("WaitingACK"), 0))
        Next
        
      ElseIf UCase(TableName) = "STOREDFRANKING" And (Operation = 1) Then
        For Each row In dst.Rows
          SD(row.Item("userid") & DelimChar & Common.NZ(row.Item("priority"), 0) & DelimChar & row.Item("output") & DelimChar & Common.NZ(row.Item("deliverabletype"), 10) & DelimChar & row.Item("roid") & DelimChar & row.Item("status") & DelimChar & Common.NZ(row.Item("create_date"), "") & DelimChar & Common.NZ(row.Item("issue_date"), ""))
          Common.Write_Log(LogFile, row.Item("userid") & DelimChar & Common.NZ(row.Item("priority"), 0) & DelimChar & row.Item("output") & DelimChar & Common.NZ(row.Item("deliverabletype"), 10) & DelimChar & row.Item("roid") & DelimChar & row.Item("status") & DelimChar & Common.NZ(row.Item("create_date"), "") & DelimChar & Common.NZ(row.Item("issue_date"), ""))
        Next
        
      ElseIf UCase(TableName) = "STOREDFRANKING" And (Operation = 12) Then
        For Each row In dst.Rows
          SD(row.Item("userid") & DelimChar & row.Item("roid"))
          Common.Write_Log(LogFile, row.Item("userid") & DelimChar & row.Item("roid"))
        Next
        
      Else
        For Each row In dst.Rows
          Index = 0
          TempResults = ""
          For Each SQLCol In dst.Columns
            If Not (TempResults = "") Then
              TempResults = TempResults & DelimChar
            End If
            If SQLCol.DataType.Name = "Boolean" Then 'if it is a binary field 
              TempResults = TempResults & Parse_Bit(Common.NZ(row(Index), 0))
            ElseIf SQLCol.DataType.Name = "Int32" Or SQLCol.DataType.Name = "Int64" Then 'if it is an Int or BigInt field
              TempResults = TempResults & Common.NZ(row(Index), 0)
            Else 'treat it as a string
              TempResults = TempResults & Common.NZ(row(Index), "")
            End If
            Index = Index + 1
          Next
          SD(TempResults)
          Common.Write_Log(LogFile, TempResults)
        Next
      End If
      SD("###")
      Common.Write_Log(LogFile, "###")
    End If
    
  End Sub
  


  ' -----------------------------------------------------------------------------------------------
  
  Sub Construct_Output(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal LastHeard As String)
    
    Dim ColumnNamesOverride As String = ""
    
    Const MAX_RUN_TIME As Double = 240
    Const DelimChar As String = Chr(30)
    
    Common.Write_Log(LogFile, "Returned the following data:")
    
    
    Dim MustIPL As Boolean = False 

    Common.QueryStr = "dbo.pa_CPE_CheckMustIPL"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    Dim dst As DataTable = Common.LRTsp_select
    Common.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      MustIPL = Common.NZ(dst.Rows(0).Item("MustIPL"), True)
    End If
    dst = Nothing
    
    
    Dim OutStr As String = ""
    If MustIPL Then
      OutStr = "MustIPL"
    Else
      OutStr = "TransDownload"
    End If
    OutStr = OutStr & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
    OutStr = OutStr & "LocationID=" & LocationID
    SD(OutStr)
    Common.Write_Log(LogFile, OutStr)
    
    If Not (MustIPL) Then
      
      Dim TotalTime As Object = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Time Elapsed when starting fetch queries=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
       
      'send the data from the CPE_Customers_Output table
      If Not (TotalTime > MAX_RUN_TIME) Then
        Common.QueryStr = "dbo.pa_CPE_TD_CustomersOutput"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("Users", 1, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped Users_Output - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished Users_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'UserRemoval
      If Not (TotalTime > MAX_RUN_TIME) Then
        'send the data from the CPE_CustomerRemoval_Output table
        Common.QueryStr = "dbo.pa_CPE_TD_CustomerRemoval"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("UserRemoval", 2, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
          Common.Write_Log(LogFile, "Skipped UserRemoval - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished UserRemoval Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
            
      'send the data from the CPE_CardIDs_Output table
      If Not (TotalTime > MAX_RUN_TIME) Then
        'send the DELETED data from the CPE_CardIDs_Output table
        Common.QueryStr = "dbo.pa_CPE_TD_CardIDsOutput_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("CardIDs", 2, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        'query for CardIDs records that should be updated/inserted
        Common.QueryStr = "dbo.pa_CPE_TD_CardIDsOutput"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("CardIDs", 6, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped CardIDs_Output - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished CardIDs_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'send the data from the CPE_CustomerAttributes_Output table
      If Not (TotalTime > MAX_RUN_TIME) Then
        'send the DELETED data from the CPE_CustomerAttributes_Output table
        Common.QueryStr = "dbo.pa_CPE_TD_CustomerAttributes_AttributeType_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("CustomerAttributes", 2, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        'send the DELETED data from the CPE_CustomerAttributes_Output table
        Common.QueryStr = "dbo.pa_CPE_TD_CustomerAttributes_AttributeValue_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("CustomerAttributes", 12, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        'send the DELETED data from the CPE_CustomerAttributes_Output table
        Common.QueryStr = "dbo.pa_CPE_TD_CustomerAttributesOutput_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("CustomerAttributes", 12, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        'query for CPE_CustomerAttributes_Output records that should be updated/inserted
        Common.QueryStr = "dbo.pa_CPE_TD_CustomerAttributesOutput"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("CustomerAttributes", 7, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped CustomerAttributes_Output - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished CustomerAttributes_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'SavingsToDate
      'send the data from the STD_Output table
      If Not (TotalTime > MAX_RUN_TIME) Then
        Common.QueryStr = "dbo.pa_CPE_TD_STDOutput"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        ColumnNamesOverride = "UserID" & DelimChar & "AdjAmount"
        Construct_Table("STDADJ", 5, DelimChar, LocalServerID, LocationID, dst, ColumnNamesOverride)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped STD_Output - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished STD_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'GroupMembership
      If Not (TotalTime > MAX_RUN_TIME) Then
        'send active GroupMembership data
        Common.QueryStr = "dbo.pa_CPE_TD_GMOutput_Active"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("GroupMembership", 7, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        'send the DELETED data from the GroupMembership table - only sends manually deleted records
        Common.QueryStr = "dbo.pa_CPE_TD_GMOutput_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("GroupMembership", 2, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped GroupMembership - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished GroupMembership Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'RewardAccumulation
      If Not (TotalTime > MAX_RUN_TIME) Then
        'send active RewardAccumulation data
        Common.QueryStr = "dbo.pa_CPE_TD_RAOutput_Active"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("RewardAccumulation", 7, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        'send deleted RewardAccumulation data
        Common.QueryStr = "dbo.pa_CPE_TD_RAOutput_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("RewardAccumulation", 2, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped RewardAccumulation - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished RewardAccumulation Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'Points/PointsAdj
      If Not (TotalTime > MAX_RUN_TIME) Then
        Common.QueryStr = "dbo.pa_CPE_TD_PointsAdj"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("PointsAdj", 5, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped Points/PointsAdj - Total Time exceeded serial=" & LocalServerID & "  IPAddress=" & (Trim(Request.UserHostAddress)) & " server=" & Environment.MachineName)
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished PointsAdj Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      
      If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) And Common.Fetch_CPE_SystemOption(128) = "1" Then
        'send the data from the CPE_PrefValueRemoval_Output table
        If Not (TotalTime > MAX_RUN_TIME) Then
          'query for Prefs records that should be removed from the local server
          Common.QueryStr = "dbo.pa_CPE_TD_PrefValueRemovalOutput"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          dst = Common.LXSsp_select
          Common.Close_LXSsp()
          Construct_Table("CustomerPreferences", 12, DelimChar, LocalServerID, LocationID, dst)
          dst = Nothing
          'now remove the returned rows and get rid of any related records from the CPE_Prefs_Output table
          Common.QueryStr = "pa_CPE_TD_PrefValueOutputPurge"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          Common.LXSsp.ExecuteNonQuery()
          Common.Close_LEXsp()
        Else
          Common.Write_Log(LogFile, "Skipped PrefValueRemoval_Output - Total Time exceeded")
        End If
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished CPE_PrefValueRemoval_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

        'send the data from the CPE_PrefValueRemovalMV_Output table
        If Not (TotalTime > MAX_RUN_TIME) Then
          'query for Prefs records that should be removed from the local server
          Common.QueryStr = "dbo.pa_CPE_TD_PrefValueRemovalMVOutput"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          dst = Common.LXSsp_select
          Common.Close_LXSsp()
          Construct_Table("CustomerPreferencesMV", 12, DelimChar, LocalServerID, LocationID, dst)
          dst = Nothing
          'now remove the returned rows and get rid of any related records from the CPE_Prefs_Output table
          Common.QueryStr = "pa_CPE_TD_PrefValueMVOutputPurge"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          Common.LXSsp.ExecuteNonQuery()
          Common.Close_LEXsp()
        Else
          Common.Write_Log(LogFile, "Skipped PrefValueRemovalMV_Output - Total Time exceeded")
        End If
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished CPE_PrefValueRemovalMV_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

        'send the data from the CPE_Prefs_Output table
        If Not (TotalTime > MAX_RUN_TIME) Then
          'query for Prefs records that should be updated/inserted
          Common.QueryStr = "dbo.pa_CPE_TD_PrefsOutput"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          dst = Common.LXSsp_select
          Common.Close_LXSsp()
          Construct_Table("CustomerPreferences", 7, DelimChar, LocalServerID, LocationID, dst)
          dst = Nothing
        Else
          Common.Write_Log(LogFile, "Skipped Prefs_Output - Total Time exceeded")
        End If
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished CPE_Prefs_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      
        'send the data from the CPE_PrefsMV_Output table
        If Not (TotalTime > MAX_RUN_TIME) Then
          'query for Prefs Multi-Value records that should be updated/inserted
          Common.QueryStr = "dbo.pa_CPE_TD_PrefsMVOutput_Active"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          dst = Common.LXSsp_select
          Common.Close_LXSsp()
          Construct_Table("CustomerPreferencesMV", 8, DelimChar, LocalServerID, LocationID, dst)
          dst = Nothing
          'query for Prefs Multi-Value records that should be deleted
          Common.QueryStr = "dbo.pa_CPE_TD_PrefsMVOutput_Deleted"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          dst = Common.LXSsp_select
          Common.Close_LXSsp()
          Construct_Table("CustomerPreferencesMV", 2, DelimChar, LocalServerID, LocationID, dst)
          dst = Nothing
        Else
          Common.Write_Log(LogFile, "Skipped PrefsMV_Output - Total Time exceeded")
        End If
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished CPE_PrefsMV_Output Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      End If
      
      
      'RewardDistribution
      If Not (TotalTime > MAX_RUN_TIME) Then
        Common.QueryStr = "dbo.pa_CPE_TD_RDOutput_Active"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("RewardDistribution", 7, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        Common.QueryStr = "dbo.pa_CPE_TD_RDOutput_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("RewardDistribution", 2, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped RewardDistribution - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished RewardDistribution Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'StoredValue
      If Not (TotalTime > MAX_RUN_TIME) Then
        Common.QueryStr = "dbo.pa_CPE_TD_SVOutput_New"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("StoredValue", 7, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        Common.QueryStr = "dbo.pa_CPE_TD_SVOutput_Used"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("SVAdj", 5, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        ' if program balance rollup is enabled then send that data
        If (Common.Extract_Val(Common.Fetch_SystemOption(85)) = 1) Then
          Common.QueryStr = "dbo.pa_CPE_TD_SVOutput_Transfer"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
          dst = Common.LXSsp_select
          Common.Close_LXSsp()
          Construct_Table("StoredValue", 11, DelimChar, LocalServerID, LocationID, dst)
          dst = Nothing
        End If
      Else
        Common.Write_Log(LogFile, "Skipped StoredValue - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished Stored Tables=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      
      'StoredFranking
      If Not (TotalTime > MAX_RUN_TIME) Then
        'send the data from the StoredFranking table
        Common.QueryStr = "dbo.pa_CPE_TD_StoredFranking_New"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("StoredFranking", 1, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        Common.QueryStr = "dbo.pa_CPE_TD_StoredFranking_Issued"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("StoredFranking", 12, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped StoredFranking - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished StoredFranking Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      


      'UserResponses <------------------------------------------ This needs to be last <-------------------------------------------------
      If Not (TotalTime > MAX_RUN_TIME) Then
        'send the data from the UserResponses table
        Common.QueryStr = "dbo.pa_CPE_TD_CROutput_Active"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("UserResponses", 7, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
        Common.QueryStr = "dbo.pa_CPE_TD_CROutput_Deleted"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        Construct_Table("UserResponses", 2, DelimChar, LocalServerID, LocationID, dst)
        dst = Nothing
      Else
        Common.Write_Log(LogFile, "Skipped UserResponses - Total Time exceeded")
      End If
      TotalTime = DateAndTime.Timer - StartTime
      Common.Write_Log(LogFile, "Finished UserResponses Table=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

      
      SD("***") 'send the EOF marker    
    End If  'not (MustIPL)
    

  End Sub


  ' -----------------------------------------------------------------------------------------------


  Sub Process_ACK(ByVal LocalServerID As Integer, ByVal LocationID As Integer)
    
    Common.Write_Log(LogFile, "Received TransDownload ACK")
    'we got an ACK ... so we need to clear the WaitingACK bit for this LocalServer
    Common.QueryStr = "dbo.pa_CPE_TD_UpdateLastHeard"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    Common.LRTsp.ExecuteNonQuery()
    Common.Close_LRTsp()
    
    Common.QueryStr = "dbo.pa_CPE_TD_ProcessACK"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
    Common.LXSsp.ExecuteNonQuery()
    Common.Close_LXSsp()
    
    Send_Response_Header("TransDownload", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Send("ACK Received")
    
  End Sub


  ' -----------------------------------------------------------------------------------------------


  Sub Process_NAK(ByVal LocalServerID As Integer, ByVal LocationID As Integer)
        Dim ExtLocationCode As String = ""
        Dim LocationName As String = ""
        Dim rst As New DataTable
    
    MacAddress = Trim(Request.QueryString("mac"))
    
    If MacAddress = "" Then
      MacAddress = "0"
    End If
    LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
    LocalServerIP = Trim(Request.QueryString("IP"))
    If LocalServerIP = "" Or LocalServerIP = "0" Then
      LocalServerIP = Trim(Request.UserHostAddress) & " IP from requesting browser. "
    End If
    
    Dim ErrorMsg As String = Trim(Request.QueryString("errormsg"))
    Common.Write_Log(LogFile, "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Received NAK - ErrorMsg:" & ErrorMsg & " Server:" & Environment.MachineName)
    Send_Response_Header("NAK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID.ToString & ";"
        rst = Common.LRT_Select
        If rst.Rows.Count > 0 Then
            LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
            ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
        End If
        
        Dim OutBuffer As String = "Local Server TransDownload NAK ErrorMsg Received" & vbCrLf
        OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString() & vbCrLf
        OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
        OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
        OutBuffer = OutBuffer & "IP: " & LocalServerIP & vbCrLf
        OutBuffer = OutBuffer & "Server: " & Environment.MachineName & vbCrLf
        OutBuffer = OutBuffer & "ErrorMsg " & ErrorMsg & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "Subject: TransDownload NAK ErrorMsg Received"
    
        Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch NAK ErrorMsg Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
    
  End Sub


</script>
<%

  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here
  
  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim TotalTime As Object
  Dim ZipOutput As Boolean
  Dim DataFile As String
  Dim ZipFile As String
  Dim FileStamp As String
  Dim Mode As String
  Dim RawRequest As String
  Dim Index As Long
  Dim IPAddress As String = ""
  Dim CompressedArray() As Byte
  Dim BannerID As Integer
    Dim OutBuffer As String = ""
    Dim LocationName As String = ""
    Dim ExtLocationCode As String = ""
    Dim rst As New DataTable
    
  Common.AppName = "TransDownload.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer
  IPAddress = Request.UserHostAddress
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "CPE-TransUpdateLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
  LocalServerIP = Trim(Request.QueryString("IP"))
  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Then
    MacAddress = "0"
  End If
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
    LocalServerIP = Trim(Request.UserHostAddress)
  End If
  LastHeard = "1/1/1980"
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LSVersion = Common.Extract_Val(Request.QueryString("lsversion"))
  LSBuild = Common.Extract_Val(Request.QueryString("lsbuild"))
  Mode = UCase(Request.QueryString("mode"))
  
  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, LocalServerIP)
  
  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  Process running on server:" & Environment.MachineName & " Mac IPAddress=" & (Trim(Request.UserHostAddress)))
  
  If LocationID = 0 Then
    Common.Write_Log(LogFile, "Received Invalid Serial Number from ma=" & IPAddress & " IP=" & Trim(Request.UserHostAddress) & " server=" & Environment.MachineName)
    Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        
        OutBuffer = "TransDownload Received Invalid Serial from MacAddress:" & MacAddress & vbCrLf
        OutBuffer = OutBuffer & "LocationID: " & LocationID & vbCrLf
        OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
        OutBuffer = OutBuffer & "IP: " & Trim(Request.UserHostAddress) & vbCrLf
        OutBuffer = OutBuffer & "Server: " & Environment.MachineName & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "TransDownload Invalid Serial from MacAddress"
        Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Invalid Serial Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID, OutBuffer)
'  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
'    'the location calling TransDownload is not associated with the CPE promoengine
'    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
'    Send_Response_Header("This location is associated with a promotion engine other thanCPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  Else
    Select Case Mode
            
      Case "ACK"
        Process_ACK(LocalServerID, LocationID)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "ACK Total Run Time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

      Case "NAK"
        Process_NAK(LocalServerID, LocationID)

      Case "FETCH"
        OutboundBuffer = New StringBuilder
        Construct_Output(LocalServerID, LocationID, LastHeard)
        Common.Write_Log(LogFile, "Starting GZip compression ... size before zipping is " & Format(OutboundBuffer.Length, "###,###,###,###,##0") & " bytes")
        TotalTime = DateAndTime.Timer - StartTime
                
        Common.Write_Log(LogFile, "Time elapsed before starting compression=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
        CompressedArray = Encoding.Default.GetBytes(GZIP.CompressString(OutboundBuffer.ToString))
        Response.BinaryWrite(CompressedArray)
        Common.Write_Log(LogFile, "GZip compression successful ... size after zipping is " & Format(UBound(CompressedArray) + 1, "###,###,###,###,##0") & " bytes")
                
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Fetch Total Run Time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")

      Case Else
        Send_Response_Header("Invalid Request - bad mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Received invalid request!")
        Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf)
        RawRequest = Get_Raw_Form(Request.InputStream)
        Common.Write_Log(LogFile, RawRequest)

                If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
                Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
                rst = Common.LRT_Select
                If rst.Rows.Count > 0 Then
                    LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
                    ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
                End If
    
                OutBuffer = "TransDownload Received invalid request - bad mode" & vbCrLf
                OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString() & vbCrLf
                OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
                OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
                OutBuffer = OutBuffer & vbCrLf & "Subject: TransDownload Received invalid request"
                Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Invalid Request Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID.ToString(), OutBuffer)

    End Select
  End If 'locationid="0"
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
      
    If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
    Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
    rst = Common.LRT_Select
    If rst.Rows.Count > 0 Then
        LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
        ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
    End If
    
    Dim ErrorMsg As String = "TransDownload Error during Local Server Processing" & vbCrLf
    ErrorMsg = ErrorMsg & "LocationID: " & LocationID.ToString() & vbCrLf
    ErrorMsg = ErrorMsg & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
    ErrorMsg = ErrorMsg & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
    ErrorMsg = ErrorMsg & "MacAddress: " & MacAddress & vbCrLf
    ErrorMsg = ErrorMsg & "IP: " & LocalServerIP & vbCrLf
    ErrorMsg = ErrorMsg & vbCrLf & "Subject:TransDownload Error during Local Server Processing"
    Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Error in Processing Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID.ToString(), ErrorMsg)
    
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  Common = Nothing
%>