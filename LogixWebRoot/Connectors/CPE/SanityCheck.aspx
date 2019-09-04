<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: SanityCheck.aspx 
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
<script runat="server">
  Public Common As New Copient.CommonInc
  
  ' hyrdra variables
  Public Const AppCode As String = "IN"
  Public OutboundData As String
  Public OutboundBuffer As StringBuilder
  Public CompressedData As String
  Public RunTime As String
  Public RunDate As String
  
  ' don't want other modules to accidentally slaughter one of these important vars
  Public LocationID As String
  Public ServerSerialNum As String
  Public LocalServerID As String
  Public MacAddress As String
  Public LocalServerIP As String
    
  ' sanity check variables
  ' 1 to 27 are the valid table numbers being stored in CheckTables
  ' table numbers are coordinated with the local server
  Public CheckTables(0 To 27, 0 To 5) As String
  Public CheckTablesInit As Boolean
  Public Const CHECK_TABLES_LBOUND As Integer = 1
  Public Const CHECK_TABLES_UBOUND As Integer = 28
  
  ' 1 to 100 are the valid table numbers being stored in GroupTables
  ' table numbers are coordinated with the local server
  Public GroupTables(0 To 1, 0 To 4) As String
  Public Const GROUP_TABLES_LBOUND As Integer = 100
  Public Const GROUP_TABLES_UBOUND As Integer = 101
  Public GroupTablesInit As Boolean
  
  ' CheckTables index vars
  Public TableName As Integer
  Public PrimaryKey As Integer
  Public LocalServerKeys As Integer
  Public CentralServerKeys As Integer
  Public Query As Integer
  Public ExceptionQuery As Integer
  
  ' GroupTables index vars
  Public GType As Integer
  Public GThreshold As Integer
  Public GLocalServerKeys As Integer
  Public GQueryStart As Integer
  Public GQueryEnd As Integer
  Public FailoverServer As Integer
  Public VerboseLogging As Boolean
  Public LogFile As String
  Public sFormData As String
  
  Public Connector As New Copient.ConnectorInc
  Public GZIP As New Copient.GZIPInc
  Public MD5 As String
  Dim TextData As String
  Dim StartTime As Object
  Public OperateAtEnterprise As Boolean
  Public IncludeAnyCustomer As Boolean = False
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub SD(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr & vbCrLf)
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub SDb(ByVal OutStr As String)
    OutboundBuffer.Append(OutStr)
  End Sub
  
  ' -------------------------------------------------------------------------------------------------
  
  Sub CheckTables_Init()
    'Dim ActiveIncentiveIDs As String = ""
    'Dim ActiveROIDS As String = ""
    'Dim dst As DataTable
    'Dim row As DataRow
   
    Dim QueryEndDateTime as String = ""						'...CLOUDSOL-1571 determine EndDate of incentives to include in sanity check
	If Common.Fetch_CPE_SystemOption(80) = "1" Then		
		Dim QueryEndDateDecrement As Integer = -1			'...decrement by 1 day to compensate for CPE day end purge being 1 day behind, set time to midnight unless IPL was run today
		Dim QueryEndDateLastIPL as Date
		Dim dst As DataTable
		Common.QueryStr = "select LastIPL from Locations with (NoLock) where LocationID = " & LocationID
		dst = Common.LRT_Select
		If dst.Rows.Count > 0 Then
			QueryEndDateLastIPL = Common.NZ(dst.Rows(0).Item("LastIPL"), "")
			if Format(Now, "yyyy-MM-dd") = Format(QueryEndDateLastIPL, "yyyy-MM-dd") Then 
				QueryEndDateDecrement = 0					'...do not decrement if EndDate IPL was run today
			end if
		End If
		dst = Nothing
		QueryEndDateTime = Format(DateAdd("d", QueryEndDateDecrement ,Now), "yyyy-MM-dd") & " 00:00:00.000"	
		Common.Write_Log(LogFile, "LocationID: " & LocationID & " CPE_SystemOption(80)=" & Common.Fetch_CPE_SystemOption(80) & ": Last IPL was at " & QueryEndDateLastIPL & ": Setting EndDate criteria to " & QueryEndDateTime & ".")					
		QueryEndDateTime = " AND EndDate >= '" & QueryEndDateTime & "' "			
	End If													'...CLOUDSOL-1571
	
    ' important global initialization for CheckTables indices
    PrimaryKey = 0
    TableName = 1
    CentralServerKeys = 2
    LocalServerKeys = 3
    Query = 4
    ExceptionQuery = 5
    
    If Common.Fetch_CPE_SystemOption(125) = "1" Then
      IncludeAnyCustomer = True
    End If

    'Create a temp table to store the list of IncentiveIDs that need to be sent 
    Common.QueryStr = "BEGIN TRY DROP TABLE #ActiveIncentives; END TRY BEGIN CATCH END CATCH;" & _
                      "CREATE TABLE #ActiveIncentives (IncentiveID bigint);"
    Common.LRT_Execute()
    'Create a temp table to store the list of ROIDs that need to be sent 
    Common.QueryStr = "BEGIN TRY DROP TABLE #ActiveROIDs; END TRY BEGIN CATCH END CATCH;" & _
                      "CREATE TABLE #ActiveROIDs (RewardOptionID bigint);"
    Common.LRT_Execute()
    
    Common.Write_Log(LogFile, "Initializing one-to-one table definitions for location " & LocationID)
    
    'Common.QueryStr = "select IncentiveID from CPE_IncentiveLoc_Func(" & LocationID & ");"
    'If CPE System option 80 is enabled then do not find the expired offers.
    If OperateAtEnterprise Then
      ' includes offers where the location was deleted from the offer's location group but the offer is still at the store.
      Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID) " & _
                        "select distinct OLU.OfferID as IncentiveID " & _
                        "from OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                        "Inner Join CPE_IncentiveLocationsView_Enterprise as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID " & _
                        "where (OLU.LocationID=" & LocationID & ") " & QueryEndDateTime & _
                        " union " & _
                        "select distinct I.IncentiveID " & _
                        "from CPE_Incentives AS I with (NoLock) " & _
                        "  inner join OfferLocations AS OL with (NoLock) on I.IncentiveID = OL.OfferID and I.Deleted = 0 and OL.Deleted = 0 and OL.Excluded = 0 " & _
                        "    and I.IsTemplate = 0 " & _
                        "  inner join LocGroupItems AS LGI with (NoLock) on LGI.LocationGroupID = OL.LocationGroupID and LGI.Deleted = 1 " & _
                        "  inner join Locations AS L with (NoLock) on LGI.LocationID = L.LocationID and L.EngineID = 2 " & _
                        "where (L.LocationID=" & LocationID & ") " & QueryEndDateTime & _
                        "order by IncentiveID;"
    Else
      Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID) " & _
                        "select distinct OLU.OfferID as IncentiveID " & _
                        "from OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                        "Inner Join CPE_IncentiveLocationsView as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID and ILV.LocationID=" & LocationID & " " & _
                        "where (OLU.LocationID=" & LocationID & ") " & QueryEndDateTime & _
                        "order by OfferID;"
    End If
    Common.LRT_Execute()
    
    'Common.QueryStr = "select RewardOptionID from CPE_ROIDLoc_Func (" & LocationID & ");"
    Common.QueryStr = "Insert into #ActiveROIDs (RewardOptionID) " & _
                      "select distinct RO.rewardOptionID " & _
                      "from CPE_rewardOptions as RO with (NoLock) Inner Join #ActiveIncentives as AI on RO.IncentiveID=AI.IncentiveID and RO.Deleted=0 " & _
                      "order by RO.RewardOptionID;"
    Common.LRT_Execute()
    
    ' Note: the order of the following definitions is important, as the table order
    ' must be the same on the local server as it is here.  The table order is defined
    ' in SanityTables.inc on the local server.
    
    Dim tableNum As Integer
    tableNum = LBound(CheckTables)
    CheckTables(tableNum, TableName) = "CPE_ST_CashierMessageTiers"  'Table 0 in CheckTables array, Table 1 from Local Server
    CheckTables(tableNum, PrimaryKey) = "MessageID"
    CheckTables(tableNum, Query) = "select MessageID from " & _
                                   " (select distinct CMT.MessageID " & _
                                   "  from CPE_ST_CashierMessageTiers as CMT with (NoLock) " & _
                                   "  inner Join CPE_ST_Deliverables as D with (NoLock) on CMT.MessageID=D.OutputID " & _
                                   "  inner join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                                   "  where D.DeliverableTypeID=9 and D.Deleted=0 " & _
                                   " union " & _
                                   "  select distinct CMT.MessageID " & _
                                   "  from CPE_CashierMessageTiers as CMT with (NoLock) " & _
                                   "  inner join CPE_CashierMessages as CM with (NoLock) on CM.MessageID=CMT.MessageID " & _
                                   "  where CM.PLU=1) as table1 " & _
                                   "order by MessageID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_DeliverablePoints"  'Table 1 in CheckTables array, Table 2 from Local Server
    CheckTables(tableNum, PrimaryKey) = "pkid"
    CheckTables(tableNum, Query) = "select distinct DP.PKID from CPE_ST_DeliverablePoints DP with (NoLock) " & _
                                    "inner join CPE_ST_Deliverables D with (NoLock) on D.OutputID=DP.PKID and D.DeliverableTypeID=8 and D.Deleted=0 and DP.Deleted=0 " & _
                                    "inner join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                                    "order by DP.PKID;"
    CheckTables(tableNum, ExceptionQuery) = ""

    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_DeliverableRoids"   'Table 2 in CheckTables array, Table 3 from Local Server
    CheckTables(tableNum, PrimaryKey) = "pkid"
    CheckTables(tableNum, Query) = "select distinct DR.PKID " & _
                                   "from CPE_ST_DeliverableROIDS DR with (NoLock) " & _
                                   "inner join CPE_ST_Deliverables D with (NoLock) on D.DeliverableID=DR.DeliverableID and D.Deleted=0 and DR.Deleted=0 " & _
                                   "inner join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                                   "Order by DR.PKID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_Deliverables"   'Table 3 in CheckTables array, Table 4 from Local Server
    CheckTables(tableNum, PrimaryKey) = "deliverableid"
    CheckTables(tableNum, Query) = "select distinct D.DeliverableID " & _
                                   "from CPE_ST_Deliverables D with (NoLock) " & _
                                   "inner join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID and D.Deleted=0 " & _
                                   " order by D.DeliverableID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_Discounts"   'Table 4 in CheckTables array, Table 5 from Local Server
    CheckTables(tableNum, PrimaryKey) = "discountid"
    CheckTables(tableNum, Query) = "select distinct DISC.DiscountID " & _
                                   "from CPE_ST_Discounts DISC with (NoLock) " & _
                                   "inner join CPE_ST_Deliverables D with (NoLock) on D.DeliverableTypeID=2 and D.OutputID=DISC.DiscountID and D.Deleted=0 and DISC.Deleted=0 " & _
                                   "inner join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                                   "order by DISC.DiscountID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_OfferTerminals"  'Table 5 in CheckTables array, Table 6 from Local Server
    CheckTables(tableNum, PrimaryKey) = "pkid"
    CheckTables(tableNum, Query) = "select distinct OT.PKID " & _
                                   "from CPE_ST_OfferTerminals OT with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=OT.OfferID " & _
                                   "order by OT.PKID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_IncentivePointsGroups"   'Table 6 in CheckTables array, Table 7 from Local Server
    CheckTables(tableNum, PrimaryKey) = "incentivepointsid"
    CheckTables(tableNum, Query) = "select distinct IPG.IncentivePointsID " & _
                                   "from CPE_ST_IncentivePointsGroups IPG with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPG.RewardOptionID and IPG.Deleted=0 " & _
                                   "order by IPG.IncentivePointsID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_IncentiveProductGroups"  'Table 7 in CheckTables array, Table 8 from Local Server
    CheckTables(tableNum, PrimaryKey) = "incentiveproductgroupid"
    CheckTables(tableNum, Query) = "select distinct IPG.IncentiveProductGroupID " & _
                                   "from CPE_ST_IncentiveProductGroups IPG with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPG.RewardOptionID and IPG.Deleted=0 " & _
                                   "order by IPG.IncentiveProductGroupID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_IncentiveCustomerGroups"   'Table 8 in CheckTables array, Table 9 from Local Server
    CheckTables(tableNum, PrimaryKey) = "incentivecustomerid"
    CheckTables(tableNum, Query) = "select distinct ICG.IncentiveCustomerID " & _
                                   "from CPE_ST_IncentiveCustomerGroups ICG with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ICG.RewardOptionID and ICG.Deleted=0 " & _
                                   "order by ICG.IncentiveCustomerID;"
    CheckTables(tableNum, ExceptionQuery) = ""

    'tableNum = tableNum + 1
    'CheckTables(tableNum, TableName) = "OfferLocations"
    'CheckTables(tableNum, PrimaryKey) = "pkid"
    'CheckTables(tableNum, Query) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_Incentives"   'Table 9 in CheckTables array, Table 10 from Local Server
    CheckTables(tableNum, PrimaryKey) = "incentiveid"
    CheckTables(tableNum, Query) = "select distinct I.IncentiveID " & _
                                   "from CPE_ST_Incentives I with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=I.IncentiveID and I.Deleted=0 " & _
                                   "order by I.IncentiveID;"
    CheckTables(tableNum, ExceptionQuery) = ""

    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "Locations"     'Table 10 in CheckTables array, Table 11 from Local Server
    CheckTables(tableNum, PrimaryKey) = "locationid"
    If OperateAtEnterprise Then
      CheckTables(tableNum, Query) = "select LocationID from Locations with (NoLock) where Deleted=0 order by LocationID;"
    Else
      CheckTables(tableNum, Query) = "select LocationID from Locations with (NoLock) where LocationID=" & LocationID & " and Deleted=0 order by LocationID;"
    End If
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "OnScreenAds"  'Table 11 in CheckTables array, Table 12 from Local Server
    CheckTables(tableNum, PrimaryKey) = "onscreenadid"
    If FailoverServer = 1 Then
      CheckTables(tableNum, Query) = "select 1 as OnScreenAdID where 1=2;"
    Else
      CheckTables(tableNum, Query) = "select OSA.OnScreenAdID from OnScreenADs as OSA with (NoLock) Inner Join OnScreenAdLocUpdate as OSALU with (NoLock) on OSA.OnScreenAdID=OSALU.OnScreenAdID and OSA.Deleted=0 and OSALU.LocationID=" & LocationID & "  " & _
                                           "union " & _
                                           "select Distinct OSA.OnScreenAdID from OnScreenAds as OSA with (NoLock) Inner Join ScreenCells as SC with (NoLock) on OSA.OnScreenAdID=SC.BackgroundImg and OSA.Deleted=0 and SC.Deleted=0 order by OSA.OnScreenAdID;"
    End If
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "PointsPrograms"  'Table 12 in CheckTables array, Table 13 from Local Server
    CheckTables(tableNum, PrimaryKey) = "programid"
    CheckTables(tableNum, Query) = ""
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_PrintedMessages"   'Table 13 in CheckTables array, Table 14 from Local Server
    CheckTables(tableNum, PrimaryKey) = "messageid"
    CheckTables(tableNum, Query) = "select distinct PM.MessageID " & _
                                   "from CPE_ST_PrintedMessages PM with (NoLock) inner join CPE_ST_Deliverables D with (NoLock) on D.DeliverableTypeID=4 and D.OutputID=PM.MessageID and D.Deleted=0 " & _
                                   "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                                   "order by PM.MessageID;"
    CheckTables(tableNum, ExceptionQuery) = ""

    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "ProductGroups"    'Table 14 in CheckTables array, Table 15 from Local Server
    CheckTables(tableNum, PrimaryKey) = "productgroupid"
    CheckTables(tableNum, Query) = "select distinct PG.ProductGroupID from ProductGroups as PG with (NoLock) Inner Join ProductGroupLocUpdate as PGLU with (NoLock) on PG.ProductGroupID=PGLU.ProductGroupID and PG.Deleted=0 and PGLU.LocationID=" & LocationID & " and EngineID=2 " & _
                                   "union select distinct ProductGroupID from ProductGroups with (NoLock) where AnyProduct=1 or PointsNotApplyGroup=1 or NonDiscountableGroup=1 " & _
                                   "order by PG.ProductGroupID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    'tableNum = tableNum + 1
    'CheckTables(tableNum, TableName) = "LocationGroups" 
    'CheckTables(tableNum, PrimaryKey) = "locationgroupid"
    'CheckTables(tableNum, Query) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_RewardOptions"  'Table 15 in CheckTables array, Table 16 from Local Server
    CheckTables(tableNum, PrimaryKey) = "rewardoptionid"
    CheckTables(tableNum, Query) = "select distinct RO.RewardOptionID " & _
                                   "from CPE_ST_RewardOptions RO with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=RO.RewardOptionID and RO.Deleted=0 " & _
                                   "order by RO.RewardOptionID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "ScreenCells"   'Table 16 in CheckTables array, Table 17 from Local Server
    CheckTables(tableNum, PrimaryKey) = "cellid"
    CheckTables(tableNum, Query) = ""
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "ScreenLayouts"   'Table 17 in CheckTables array, Table 18 from Local Server
    CheckTables(tableNum, PrimaryKey) = "layoutid"
    CheckTables(tableNum, Query) = ""
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "TouchAreas"     'Table 18 in CheckTables array, Table 19 from Local Server
    CheckTables(tableNum, PrimaryKey) = "areaid"
    CheckTables(tableNum, Query) = "select TA.AreaID from TouchAreas TA with (NoLock) " & _
                                   "inner join OnScreenAdLocUpdate OSALU with (NoLock) on OSALU.OnScreenAdID = TA.OnScreenAdID " & _
                                   "where TA.Deleted=0 and OSALU.LocationID = " & LocationID & " order by TA.AreaID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CustomerGroups"   'Table 19 in CheckTables array, Table 20 from Local Server
    CheckTables(tableNum, PrimaryKey) = "customergroupid"
    CheckTables(tableNum, Query) = "select Distinct CG.CustomerGroupID from CustomerGroups as CG with (NoLock) Inner Join CustomerGroupLocUpdate as CGLU with (NoLock) on CG.CustomerGroupID=CGLU.CustomerGroupID and CG.Deleted=0 and CGLU.LocationID=" & LocationID & " and CG.CustomerGroupID <> 1 " & _
                                   "union " & _
                                   "select distinct CustomerGroupID from CustomerGroups with (NoLock) where AnyCardholder=1 or NewCardholders=1 or AnyCAMCardholder=1 "  & IIf(IncludeAnyCustomer, "or AnyCustomer=1 ", "") & _
                                   "order by CG.CustomerGroupID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    'tableNum = tableNum + 1
    'CheckTables(tableNum, TableName) = "LocGroupItems"
    'CheckTables(tableNum, PrimaryKey) = "pkid"
    '' CheckTables(tableNum, Query) = ""
    'CheckTables(tableNum, Query) = "select pkid from LocGroupItems where LocationID=" & LocationID & " and Deleted=0 order by pkid;"
    'CheckTablesInit = True
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_FrankingMessageTiers"   'Table 20 in CheckTables array, Table 21 from Local Server
    CheckTables(tableNum, PrimaryKey) = "FrankID"
    CheckTables(tableNum, Query) = "select distinct FMT.FrankID " & _
                                   "from CPE_ST_FrankingMessageTiers FMT with (NoLock) inner join CPE_Deliverables D with (NoLock) on FMT.FrankID=D.OutputID and D.Deleted=0 and D.DeliverableTypeID=10 " & _
                                   "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                                   "order by FMT.FrankID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "TerminalTypes"   'Table 21 in CheckTables array, Table 22 from Local Server
    CheckTables(tableNum, PrimaryKey) = "terminaltypeid"
    CheckTables(tableNum, Query) = "select distinct TerminalTypeID from TerminalTypes  with (NoLock) " & _
                                    "where Deleted = 0 and EngineID=2 order by TerminalTypeID;"
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "StoredValuePrograms"   'Table 22 in CheckTables array, Table 23 from Local Server
    CheckTables(tableNum, PrimaryKey) = "svprogramid"
    CheckTables(tableNum, Query) = ""
    CheckTables(tableNum, ExceptionQuery) = ""
    
    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_IncentiveStoredValuePrograms"    'Table 23 in CheckTables array, Table 24 from Local Server
    CheckTables(tableNum, PrimaryKey) = "incentivestoredvalueid"
    CheckTables(tableNum, Query) = "select distinct ISVP.IncentiveStoredValueID " & _
                                   "from CPE_ST_IncentiveStoredValuePrograms ISVP with (NoLock) Inner Join #ActiveROIDs as AR on ISVP.Deleted=0 and AR.RewardOptionID=ISVP.RewardOptionID " & _
                                   "order by ISVP.IncentiveStoredValueID;"
    CheckTables(tableNum, ExceptionQuery) = ""

    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_DeliverableStoredValue"   'Table 24 in CheckTables array, Table 25 from Local Server
    CheckTables(tableNum, PrimaryKey) = "pkid"
    CheckTables(tableNum, Query) = "select distinct DSV.PKID " & _
                                   "from CPE_ST_DeliverableStoredValue DSV with (NoLock) inner join CPE_ST_Deliverables D with (NoLock) on D.OutputID=DSV.PKID and D.DeliverableTypeID=11 and D.Deleted=0 and DSV.Deleted=0 " & _
                                   "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                                   "order by DSV.PKID;"
    CheckTables(tableNum, ExceptionQuery) = ""

    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "SystemOptions"   'Table 25 in CheckTables array, Table 26 from Local Server
    CheckTables(tableNum, PrimaryKey) = "optionid"
    CheckTables(tableNum, Query) = "select optionid from SystemOptions with (NoLock) order by optionid;"
    CheckTables(tableNum, ExceptionQuery) = ""

    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_SystemOptions"   'Table 26 in CheckTables array, Table 27 from Local Server
    CheckTables(tableNum, PrimaryKey) = "optionid"
    CheckTables(tableNum, Query) = "select optionid from CPE_SystemOptions with (NoLock) order by optionid;"
    CheckTables(tableNum, ExceptionQuery) = ""

    tableNum = tableNum + 1
    CheckTables(tableNum, TableName) = "CPE_ST_IncentiveTenderTypes"   'Table 27 in CheckTables array, Table 28 from Local Server
    CheckTables(tableNum, PrimaryKey) = "IncentiveTenderID"
    'CheckTables(tableNum, Query) = "select distinct ITT.IncentiveTenderID from CPE_ST_IncentiveTenderTypes ITT with (NoLock) where ITT.Deleted=0 and ITT.RewardOptionID in (" & ActiveROIDS & ") order by ITT.IncentiveTenderID;"
    CheckTables(tableNum, Query) = "select distinct ITT.IncentiveTenderID " & _
                                   "from CPE_ST_IncentiveTenderTypes ITT with (NoLock) " & _
                                   "inner join CPE_ST_IncentiveTenderTypeTiers as ITTT with (NoLock) on ITT.IncentiveTenderID=ITTT.IncentiveTenderID " & _
                                   "inner join #ActiveROIDs as AR on AR.RewardOptionID=ITT.RewardOptionID " & _
                                   "where ITTT.TierLevel = 1 And ITT.Deleted = 0 " & _
                                   "order by ITT.IncentiveTenderID;"
    CheckTables(tableNum, ExceptionQuery) = ""

    
  End Sub
  
  '-----------------------------------------------------------------------------------------------
  
  Sub GroupTables_Init()
    
    ' the GroupTables array is used to check for group membership in a given group
    ' the GType index represents what type of group
    ' the GThreshold index represents the minimum acceptable number of group members
    '     (anything less will trigger an error)
    ' the GQueryStart and GQueryEnd indices represent the sql query needed to determine
    '     the number of members in the group... the query will be a concatenation of
    '     the GQueryStart index, the actualy group id, and the GQueryEnd index
    
    ' important global initialization for CheckTables indices
    GType = 0
    GThreshold = 1
    GLocalServerKeys = 2
    GQueryStart = 3
    GQueryEnd = 4
    
    Common.Write_Log(LogFile, "Initializing group table definitions for location " & LocationID)
    
    ' Note: the order of the following definitions is important, as the table order
    ' must be the same on the local server as it is here.  The table order is defined
    ' in SanityTables.inc on the local server.
    
    Dim tableNum As Integer
    tableNum = LBound(GroupTables)
    GroupTables(tableNum, GType) = "Product Group"
    GroupTables(tableNum, GThreshold) = 1
    GroupTables(tableNum, GQueryStart) = "select count(PGI.ProductID) from ProdGroupItems as PGI with (NoLock) where PGI.Deleted=0 and PGI.ProductGroupID="
    GroupTables(tableNum, GQueryEnd) = ";"
    
    tableNum = tableNum + 1
    GroupTables(tableNum, GType) = "Customer Group"
    GroupTables(tableNum, GThreshold) = 1
    GroupTables(tableNum, GQueryStart) = "select count(GM.MembershipID) " & _
            "from GroupMembership as GM with (NoLock) Inner Join CustomerLocations as CL with (NoLock) on CL.CustomerPK=GM.CustomerPK and GM.Deleted=0 " & _
            "where GM.CustomerGroupID="
    GroupTables(tableNum, GQueryEnd) = "and CL.LocationID=" & LocationID & ";"
    
    GroupTablesInit = True
    
  End Sub
  
  '-----------------------------------------------------------------------------------------------
  
  Sub SanityTables_Load_Local(ByRef LocalServerData As String)
    
    Dim dataSize As Long
    Dim Index As Long
    Dim endPoint As Long
    Dim oneTable As String
    Dim tableData As Object
    Dim tdSize As Long
    Dim tableIndex As Long
    
    Common.Write_Log(LogFile, "Loading sanity data sent from local server for location " & LocationID)
    
    ' Making certain the CheckTables array is ready for us
    If (CheckTablesInit = False) Then
      CheckTables_Init()
    End If
    
    If (GroupTablesInit = False) Then
      GroupTables_Init()
    End If
    
    ' The following block splits the local server data into chunks of data,
    ' each representing an individual table in the local server database.
    ' The first number in each chunk of data is the table number for the table.
    ' The subsequent numbers (if any) represent a list of primary key values stored
    ' in the table.  See the Sanity Check definition document for more detail.
    
    dataSize = Len(LocalServerData)
    Index = 1
    While Index < dataSize
      
      endPoint = InStr(Index, LocalServerData, vbCrLf, vbBinaryCompare)
      
      If (VerboseLogging) Then
        Common.Write_Log(LogFile, "Parsing table data starting at " & Index & " ending with " & endPoint & " location " & LocationID & "." & " Serial:" & LocalServerID & "Mac IPAddress:" & (Trim(Request.UserHostAddress)) & "server:" & Environment.MachineName)
      End If
      
      oneTable = Mid(LocalServerData, Index, endPoint - Index)
      
      tableData = Split(oneTable, ",", 2)
      ' deallocate the oneTable string, it's not needed right now
      oneTable = ""
      
      tdSize = UBound(tableData) - LBound(tableData) + 1
      
      tableIndex = CLng(tableData(0))
      
      ' debugging...
      ' Send ("Arr size: " & TDSize & " | Table Number: " & TableData(0))
      ' Send ("Table data not array: " & TableData)
      
      If ((tableIndex >= CHECK_TABLES_LBOUND) And (tableIndex <= CHECK_TABLES_UBOUND)) Then
        'If ((tableIndex >= LBound(CheckTables)) And (tableIndex <= UBound(CheckTables))) Then
        
        If (VerboseLogging) Then
          Common.Write_Log(LogFile, "Adding table data to one-to-one tables, table number " & tableIndex & " location " & LocationID & ".")
        End If
        
        ' adjust the table number back to account for the zero-based array 
        tableIndex = tableIndex - CHECK_TABLES_LBOUND
        
        If (tdSize = 2) Then
          CheckTables(tableIndex, LocalServerKeys) = tableData(1)
        Else
          CheckTables(tableIndex, LocalServerKeys) = ""
        End If
        
      ElseIf ((tableIndex >= GROUP_TABLES_LBOUND) And (tableIndex <= GROUP_TABLES_UBOUND)) Then
        'ElseIf ((tableIndex >= LBound(GroupTables)) And (tableIndex <= UBound(GroupTables))) Then
        
        If (VerboseLogging) Then
          Common.Write_Log(LogFile, "Adding table data to group tables, table number " & tableIndex & " location " & LocationID & ".")
        End If
        
        tableIndex = tableIndex - GROUP_TABLES_LBOUND
        
        If (tdSize = 2) Then
          GroupTables(tableIndex, GLocalServerKeys) = tableData(1)
        Else
          GroupTables(tableIndex, GLocalServerKeys) = ""
        End If
        
      End If
      
      ' deallocate the tableData array
      ' tableData = Nothing
      Index = endPoint + 2 'move past the CRLF
      
    End While
    
  End Sub
  
  '-----------------------------------------------------------------------------------------------
  
  
  Function Process_Data(ByVal MD5 As String) As Boolean
    Dim CompressedData As String
    Dim InboundData As String = ""
    Dim FileData() As Byte
    Dim Checksum As String
    Dim DataRetrieved As Boolean = False
        Dim OutBuffer As String = ""
        Dim ExtLocationCode As String = ""
        Dim LocationName As String = ""
        Dim rst As New DataTable

    
    Try
      MD5 = Request.QueryString("MD5")
      LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
      LocalServerIP = Trim(Request.QueryString("IP"))
      MacAddress = Trim(Request.QueryString("mac"))
      If MacAddress = "" Or MacAddress = "0" Then
        MacAddress = Trim(Request.UserHostAddress)
      End If
      If LocalServerIP = "" Or LocalServerIP = "0" Then
        LocalServerIP = MacAddress & " IP from requesting browser. "
      End If
            
      InboundData = ""
      If Request.Files.Count > 0 Then
        ReDim FileData(Request.Files(0).ContentLength)
        Request.Files(0).InputStream.Read(FileData, 0, Request.Files(0).ContentLength)
        'uncomment to view raw data
        'Send(Encoding.Default.GetString(FileData))
        CompressedData = Encoding.Default.GetString(FileData)
        FileData = Nothing
        InboundData = GZIP.DecompressString(CompressedData)
        CompressedData = Nothing
        'uncomment to view decompressed data
        'Send(Inbounddata)
        If (MD5 <> "") Then
          'if we got an MD5 checksum, make sure it matches the commpressed data we recieved
          Checksum = Common.MD5(InboundData)
          If Checksum <> MD5 Then
            Common.Write_Log(LogFile, "Bad MD5 .. LocalServer " & LocalServerID & " sent ->" & MD5 & "     CentralServer computed ->" & Checksum)

                        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
                        Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
                        rst = Common.LRT_Select
                        If rst.Rows.Count > 0 Then
                            LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
                            ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
                        End If

                        OutBuffer = "Sanity Check Bad MD5 Message" & vbCrLf
                        OutBuffer = OutBuffer & "Bad MD5 Sent by LocalServer. " & vbCrLf
                        OutBuffer = OutBuffer & "LocationID: " & LocationID & vbCrLf
                        OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
                        OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
                        OutBuffer = OutBuffer & "CheckSum: '" & Checksum & "' computed by CentralServer.  LocalServer sent an MD5 of '" & MD5 & "'" & vbCrLf
                        OutBuffer = OutBuffer & vbCrLf & "Subject: Sanity Check Bad MD5"
                        Common.Send_Email(Common.Get_Error_Emails(6), Common.SystemEmailAddress, "Sanity Check Bad MD5 Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
            Send("NAK - Bad MD5")
            DataRetrieved = False
          Else
            Common.Write_Log(LogFile, "Valid MD5 ->" & MD5)
            Common.Write_Log(LogFile, "GZip decompression successful ... size after unzipping is " & Format(Len(InboundData), "###,###,###,###,##0") & " bytes")
            Common.Write_Log(LogFile, " Serial: " & LocalServerID & " IPAddress:" & Trim(Request.UserHostAddress) & " Received Data:" & vbCrLf & InboundData)
            InboundData = InboundData
            DataRetrieved = True
          End If
        Else
          If InboundData = "no data" Then
            Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " Server had no data to upload - no processing performed")
            Send("ACK")
            DataRetrieved = False
          End If
        End If
      Else
        Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " No files were uploaded")
        Send("NAK - No files were uploaded")
        DataRetrieved = False
      End If
    Catch ex As Exception
      Send(ex.ToString)
      Common.Error_Processor()
      Common.Write_Log(LogFile, ex.ToString() & " serial: " & LocalServerID & " MacAddress:" & MacAddress)
            
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
            Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
            rst = Common.LRT_Select
            If rst.Rows.Count > 0 Then
                LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
                ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
            End If

            OutBuffer = "Sanity Check Exception Message" & vbCrLf
            OutBuffer = OutBuffer & "LocationID: " & LocationID & vbCrLf
            OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
            OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
            OutBuffer = OutBuffer & "Exception: " & ex.ToString & vbCrLf
            OutBuffer = OutBuffer & vbCrLf & "Subject: Sanity Check Exception"
            Common.Send_Email(Common.Get_Error_Emails(6), Common.SystemEmailAddress, "Sanity Check Exception Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
    End Try
    
    Common.Write_Log(LogFile, "LocationID: " & LocationID & " MD5 check succcesful.")
    
    ' we don't need the compressed data any more, deallocate it to free up memory
    ' preferrably before we chew a bunch up loading the local server data
    CompressedData = ""
    sFormData = ""
    
    ' now that we've unzipped it and verified the md5sum,
    ' let's store the local server data in the CheckTables array
    If (DataRetrieved) Then
      SanityTables_Load_Local(InboundData)
    End If
    
    Return DataRetrieved
  End Function
  
  '-----------------------------------------------------------------------------------------------
  
  Sub Handle_Post()
    Dim CurrentTime As Date
    Dim tableNumber As Integer
    Dim centralKeyData As String
    Dim localKeyData As String
    Dim centralKeyArr As Object
    Dim localKeyArr As Object
    Dim centralKeyIndex As Long
    Dim localKeyIndex As Long
    Dim centralKeyMax As Long
    Dim localKeyMax As Long
    Dim centralKeyID As Long
    Dim localKeyID As Long
    Dim localKeyMissing As StringBuilder
    Dim centralKeyMissing As StringBuilder
    Dim maxKeysLogged As Integer
    Dim centralKeysLogged As Integer
    Dim localKeysLogged As Integer
    Dim localKeysIgnored As Boolean
    Dim centralKeysIgnored As Boolean
    Dim tmpBuffer As StringBuilder
    Dim keyValue As Long
    Dim recordCount As Long
    Dim tmpID As Object
    Dim ResultOK As Integer
    Dim rst As New DataTable
    Dim row As DataRow
        Dim OutBuffer As String = ""
        Dim ExtLocationCode As String = ""
        Dim LocationName As String = ""
    
    ' vars used strictly for processing groups (GroupTables)
    Dim zeroRecords As New DataTable
    Dim centralKeyCount As Long
    Dim groupOK As Boolean
    
    VerboseLogging = False
    Send_Response_Header("SanityCheck", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    
    Common.Open_LogixRT()
    Common.Open_LogixXS()
    
    ' Fetch the CurrentTime from the database server
    CurrentTime = Now
    
    RunTime = (Format(CurrentTime, "h:mm:ss"))
    RunDate = (Format(CurrentTime, "MM/dd/yyyy"))
    
    If LocationID = "0" Then
      Send("Invalid Serial Number")
      Send("NAK")
      Common.Write_Log(LogFile, "Serial:" & LocalServerID & " Received invalid Serial Number from MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
            
            OutBuffer = "Sanity Check Received Invalid Serial from MacAddress:" & MacAddress & vbCrLf
            OutBuffer = OutBuffer & "LocationID: " & LocationID & vbCrLf
            OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
            OutBuffer = OutBuffer & "IP: " & LocalServerIP & vbCrLf
            OutBuffer = OutBuffer & "Server: " & Environment.MachineName & vbCrLf
            OutBuffer = OutBuffer & vbCrLf & "Subject: Sanity Check Invalid Serial from MacAddress"
            Common.Send_Email(Common.Get_Error_Emails(6), Common.SystemEmailAddress, "Sanity Check Invalid Serial Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID, OutBuffer)
    Else
      Common.Write_Log(LogFile, "Serial:" & LocalServerID & " received SanityCheck from MacAdrress:" & MacAddress & " server:" & Environment.MachineName)
      
      'see if this local server is a FailoverServer
      FailoverServer = 0
      Common.QueryStr = "select FailoverServer from LocalServers with (NoLock) where LocalServerID=" & ServerSerialNum & ";"
      rst = Common.LRT_Select
      If (rst.Rows.Count > 0) Then
        If Common.NZ(rst.Rows(0).Item("FailoverServer"), False) = True Then FailoverServer = 1
      End If
      
      ' Initialize the CheckTables and GroupTables arrays, which will hold the
      ' table names and primary keys that we need to compare from the
      ' local/central servers
      CheckTables_Init()
      GroupTables_Init()
      
      ' Grab the local server data from the HTTP POST file
      ' unzips file and verifies the md5 sum
      If (Not Process_Data(MD5)) Then Exit Sub
      
      ' PART ONE: BEGIN
      ' Compare the central server data and the local server keys data from the
      ' CheckTables array, one table at a time...
      ' This will involve querying the database for each table to build the
      ' centralKeyData string
      Common.Write_Log(LogFile, "LocationID: " & LocationID & " Comparing local and central server data for one-to-one tables.")
      
      ' we're only going to show (at most) the first maxKeysLogged missing keys
      maxKeysLogged = 100
      
      ResultOK = 1
      OutboundBuffer = New StringBuilder
      
      ' Step through the array and query the "table/primary key" pairs
      For tableNumber = LBound(CheckTables) To UBound(CheckTables) Step 1
        
        If (VerboseLogging) Then
          Common.Write_Log(LogFile, "LocationID: " & LocationID & " Processing table number " & tableNumber)
        End If
        
        ' query the database for this table
        ' and store the value in centralServerData
        If CheckTables(tableNumber, Query) = "" Then
          Common.QueryStr = "select " & CheckTables(tableNumber, PrimaryKey) & " from " & _
          CheckTables(tableNumber, TableName) & " with (NoLock) where deleted=0 order by " & _
          CheckTables(tableNumber, PrimaryKey) & ";"
        Else
          ' the query is already stored in the array
          Common.QueryStr = CheckTables(tableNumber, Query)
        End If
        rst = Common.LRT_Select
        If (rst.Rows.Count > 0) Then
          ' Loop through the record set and build a string of key values
          tmpBuffer = New StringBuilder
          recordCount = 0
          For Each row In rst.Rows
            ' Fetch the zeroth column and store in keyValue
            keyValue = Common.NZ(row.Item(0), 0)
            ' DEBUG ' If tableNumber = 23 Then Write_Log LocationID, "length: " & tmpBuffer.Length & " count: " & recordCount
            If tmpBuffer.Length > 0 Then tmpBuffer.Append(",")
            tmpBuffer.Append(keyValue)
            recordCount = recordCount + 1
          Next
          
          If (VerboseLogging) Then
            Common.Write_Log(LogFile, "LocationID: " & LocationID & " Found " & recordCount & " records for " & CheckTables(tableNumber, TableName))
          End If
          
          centralKeyData = tmpBuffer.ToString
          
          ' deallocate memory
          tmpBuffer = Nothing
          
        Else
          ' the query didn't return any data, central server has no records for this table
          If (VerboseLogging) Then
            Common.Write_Log(LogFile, "LocationID: " & LocationID & " No records found for " & CheckTables(tableNumber, TableName))
          End If
          centralKeyData = ""
        End If
        
        localKeyMissing = New StringBuilder
        centralKeyMissing = New StringBuilder
        localKeyData = CheckTables(tableNumber, LocalServerKeys)
        
        ' deallocate the string from the array, since we're done with it
        CheckTables(tableNumber, LocalServerKeys) = ""
        
        ' remove modified keys from both central and local key data before comparing
        ' COMMENTED OUT TO TEST WITH CHANGE TO USE SHADOW TABLE  2/15/08
        'RemoveModifiedKeys(tableNumber, localKeyData, centralKeyData)
        
        If (centralKeyData = localKeyData) Then
          ' we found a match... log it and move on to the next table
          Common.Write_Log(LogFile, "LocationID: " & LocationID & " " & CheckTables(tableNumber, TableName) & "(" & tableNumber + 1 & "): Local/Central Match ")
          OutboundBuffer.Append(CheckTables(tableNumber, TableName) & ": Local/Central Match " & vbCrLf)
        Else
          ResultOK = 0
          ' check for a couple of simple cases
          If (localKeyData = "") Then
            Common.Write_Log(LogFile, "LocationID: " & LocationID & " " & CheckTables(tableNumber, TableName) & "(" & tableNumber + 1 & "): Local server empty, missing keys: " & centralKeyData)
            OutboundBuffer.Append(CheckTables(tableNumber, TableName) & ": Local server empty, missing keys: " & centralKeyData & vbCrLf)
          ElseIf (centralKeyData = "") Then
            Common.Write_Log(LogFile, "LocationID: " & LocationID & " " & CheckTables(tableNumber, TableName) & "(" & tableNumber + 1 & "): Central server empty, missing keys: " & localKeyData)
            OutboundBuffer.Append(CheckTables(tableNumber, TableName) & ": Central server empty, missing keys: " & localKeyData & vbCrLf)
          Else
            
            Common.Write_Log(LogFile, "LocationID: " & LocationID & " " & CheckTables(tableNumber, TableName) & "(" & tableNumber + 1 & "): Local/Central difference(s)...")
            OutboundBuffer.Append(CheckTables(tableNumber, TableName) & ": Local/Central difference(s)..." & vbCrLf)
            
            ' break the primary key strings up into individual primary keys and compare one by one
            centralKeyArr = Split(centralKeyData, ",")
            centralKeyData = ""
            localKeyArr = Split(localKeyData, ",")
            localKeyData = ""
            
            centralKeyIndex = LBound(centralKeyArr)
            localKeyIndex = LBound(localKeyArr)
            centralKeyMax = UBound(centralKeyArr)
            localKeyMax = UBound(localKeyArr)
            
            localKeysLogged = 0
            centralKeysLogged = 0
            localKeysIgnored = False
            centralKeysIgnored = False
            
            While ((centralKeyIndex <= centralKeyMax) And (localKeyIndex <= localKeyMax) And ((Not centralKeysIgnored) Or (Not localKeysIgnored)))
              
              tmpID = centralKeyArr(centralKeyIndex)
              centralKeyID = CLng(tmpID)
              tmpID = localKeyArr(localKeyIndex)
              localKeyID = CLng(tmpID)
              
              If (localKeyID = centralKeyID) Then
                ' advance both and move on
                centralKeyIndex = centralKeyIndex + 1
                localKeyIndex = localKeyIndex + 1
              ElseIf (localKeyID < centralKeyID) Then
                ' add the missing key to the message, advance localKeyIndex, and move on
                ' If centralKeyMissing.Length > 0 Then centralKeyMissing.Append (",")
                If (centralKeysLogged < maxKeysLogged) Then
                  centralKeyMissing.Append("," & localKeyID)
                  centralKeysLogged = centralKeysLogged + 1
                ElseIf Not (centralKeysIgnored) Then
                  centralKeyMissing.Append(" [more]")
                  centralKeysIgnored = True
                End If
                ' increment it even if we don't log it
                localKeyIndex = localKeyIndex + 1
              Else  'If (centralKeyID < localKeyID) Then
                ' add the missing key to the message, advance centralKeyIndex, and move on
                If (localKeysLogged < maxKeysLogged) Then
                  localKeyMissing.Append("," & centralKeyID)
                  localKeysLogged = localKeysLogged + 1
                ElseIf Not (localKeysIgnored) Then
                  localKeyMissing.Append(" [more]")
                  localKeysIgnored = True
                End If
                centralKeyIndex = centralKeyIndex + 1
              End If
            End While
            
            ' add any remaining central keys to the localKeyMissing string
            While ((centralKeyIndex <= centralKeyMax) And (Not localKeysIgnored))
              If (localKeysLogged < maxKeysLogged) Then
                localKeyMissing.Append("," & centralKeyArr(centralKeyIndex))
                centralKeyIndex = centralKeyIndex + 1
                localKeysLogged = localKeysLogged + 1
              Else
                localKeyMissing.Append(" [more]")
                localKeysIgnored = True
              End If
            End While
            
            ' add any remaining local keys to the centralKeyMissing string
            While ((localKeyIndex <= localKeyMax) And (Not centralKeysIgnored))
              If (centralKeysLogged < maxKeysLogged) Then
                centralKeyMissing.Append("," & localKeyArr(localKeyIndex))
                localKeyIndex = localKeyIndex + 1
                centralKeysLogged = centralKeysLogged + 1
              Else
                centralKeyMissing.Append(" [more]")
                centralKeysIgnored = True
              End If
            End While
            
            centralKeyArr = New Object()
            localKeyArr = New Object()
            
            If (localKeyMissing.Length > 0) Then
              ' remove the leading "," from the string and log the missing keys
              Common.Write_Log(LogFile, "LocationID: " & LocationID & " ...serial:" & LocalServerID & " with Mac IPAddress:" & (Trim(Request.UserHostAddress)) & " missing keys: " & Mid(localKeyMissing.ToString, 2))
              OutboundBuffer.Append(" ...Local server missing keys: " & Mid(localKeyMissing.ToString, 2) & vbCrLf)
              localKeyMissing = Nothing
              localKeyMissing = New StringBuilder
            End If
            
            If (centralKeyMissing.Length > 0) Then
              ' remove the leading "," from the string and log the missing keys
              Common.Write_Log(LogFile, "LocationID: " & LocationID & " ...Central server missing keys: " & Mid(centralKeyMissing.ToString, 2))
              'OutboundBuffer.AppendByVal(" ...Central server missing keys: " & Mid(centralKeyMissing.ToString, 2) & vbCrLf)
              OutboundBuffer.Append(" ...Central server missing keys: " & Mid(centralKeyMissing.ToString, 2) & vbCrLf)
              centralKeyMissing = Nothing
              centralKeyMissing = New StringBuilder
            End If
            
          End If    ' end of Else ... (localKeyData != "") And (centralKeyData != "")
          
        End If    ' end of Else ... (centralKeyData != localKeyData)
        
      Next tableNumber
      
      Erase CheckTables
      
      ' PART ONE: FINISHED
      
      ' PART TWO: BEGIN
      ' Query the central database to see if we can match the local server groups
      ' with groups from central (all local server groups that were sent over have
      ' presumably went below the acceptable minimum [sane] number of group members)
      
      Common.Write_Log(LogFile, "LocationID: " & LocationID & " Comparing local and central server data for group tables.")
      For tableNumber = LBound(GroupTables) To UBound(GroupTables) Step 1
        
        If VerboseLogging Then Common.Write_Log(LogFile, "TableNumber=" & tableNumber)
        'GroupMembership data is not sent to the local servers when operating at enterprise
        'so we may need to skip processing of any GroupMembership data
        If tableNumber = 1 And OperateAtEnterprise = True Then
          Common.Write_Log(LogFile, "LocationID: " & LocationID & " Skipping processing of GroupMembership since we are operating at enterprise")
        Else
          
          localKeyData = GroupTables(tableNumber, GLocalServerKeys)
          If VerboseLogging Then Common.Write_Log(LogFile, "LocalKeyData=" & localKeyData)
          If (localKeyData.Trim <> "") Then
            localKeyArr = Split(localKeyData, ",")
          
            groupOK = True
          
            For localKeyIndex = LBound(localKeyArr) To UBound(localKeyArr) Step 1
            
              'build the query using the key from the local server
              tmpID = localKeyArr(localKeyIndex)
              localKeyID = CLng(tmpID)
            
              Common.QueryStr = GroupTables(tableNumber, GQueryStart) & localKeyID & GroupTables(tableNumber, GQueryEnd)
              Select Case tableNumber
                Case 0
                  zeroRecords = Common.LRT_Select()
                Case 1
                  zeroRecords = Common.LXS_Select()
                Case Else
                  zeroRecords = Common.LRT_Select()
              End Select
            
              If (zeroRecords.Rows.Count > 0) Then
              
                centralKeyCount = Common.NZ(zeroRecords.Rows(0).Item(0), 0)
                If VerboseLogging And tableNumber = 1 Then Common.Write_Log(LogFile, "centralKeyCount=" & centralKeyCount & "  For CustomerGroupID " & localKeyID & "   GroupTable compare=" & GroupTables(tableNumber, GThreshold) & vbCrLf & "Query=" & Common.QueryStr)
                ' If the group membership is above the acceptable threshold, we have an error
                ' to report... since the local server was *below* the threshold
                If (centralKeyCount >= GroupTables(tableNumber, GThreshold)) Then
                
                  If (groupOK) Then
                    ' if groupOK is true, then this is the first error for this group, output header
                    Common.Write_Log(LogFile, "LocationID: " & LocationID & " " & GroupTables(tableNumber, GType) & ": Local/Central Potentially Insane...")
                    OutboundBuffer.Append(GroupTables(tableNumber, GType) & ": Local/Central Potentially Insane..." & vbCrLf)
                  End If
                
                  groupOK = False   ' indicates that this particular group type has an error
                  ResultOK = 0      ' indicates general error during sanity check
                
                  ' log the error and send an error back
                  Common.Write_Log(LogFile, "LocationID: " & LocationID & " ... [" & localKeyID & "] Central has " & centralKeyCount & _
                      " members; Local has " & (GroupTables(tableNumber, GThreshold) - 1) & " or fewer ")
                
                  OutboundBuffer.Append(" ... [" & localKeyID & "] Central has " & centralKeyCount & _
                      " members; Local has " & (GroupTables(tableNumber, GThreshold) - 1) & " or fewer " & vbCrLf)
                
                  ' DEBUGGING...
                  '
                  '          Else
                  '
                  '            OutboundBuffer.Append (" ... SANE ... [" & localKeyID & "] Central has " & centralKeyCount & _
                  '                " members; Local has < " & GroupTables(tableNumber, GThreshold) & vbCrLf)
                End If
              Else
                ' if we're in this else, that means the group was not found
                Common.Write_Log(LogFile, "LocationID: " & LocationID & " " & GroupTables(tableNumber, GType) & ": group " & localKeyID & _
                    " not found on Central")
                OutboundBuffer.Append(GroupTables(tableNumber, GType) & ": group " & localKeyID & _
                    " not found on Central" & vbCrLf)
              End If
            
            Next localKeyIndex
          
            If (groupOK) Then
            
              ' if groupOK is true, then there were no errors for this group, output successful msg
              Common.Write_Log(LogFile, "LocationID: " & LocationID & " " & GroupTables(tableNumber, GType) & ": Local/Central Sane...")
              OutboundBuffer.Append(GroupTables(tableNumber, GType) & ": Local/Central Sane..." & vbCrLf)
            
            End If
          
          End If
        End If 'operating at enterprise and table=101 (GroupMembership)
        
      Next tableNumber
      
      ' PART TWO: FINISHED
      
      ' everything has been received fine... send an ACK to signal this to the local server
      ' the ACK isn't sent until we're done so the local server will be informed of
      ' errors encountered during processing
      Send("ACK")
      Common.Write_Log(LogFile, "LocationID: " & LocationID & " ResultOK=" & ResultOK)
      
      'update SanityCheckLastHeard in the LocalServers table
      Common.QueryStr = "update LocalServers with (RowLock) set SanityCheckLastHeard=getdate() where LocalServerID=" & ServerSerialNum & ";"
      Common.LRT_Execute()
      
      Common.QueryStr = "proc_Store_SanityCheckStatus"
      Common.Open_LRTsp()
      Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
      Common.LRTsp.Parameters.Add("@LastReport", SqlDbType.NVarChar, 1073741823).Value = OutboundBuffer.ToString
      Common.LRTsp.Parameters.Add("@DBOK", SqlDbType.Bit).Value = IIf(ResultOK, 1, 0)
      Common.LRTsp.ExecuteNonQuery()
      Common.Close_LRTsp()
      
      OutboundBuffer = Nothing
            'ResultOK = 0 indicates general error during sanity check
            If ResultOK = 0 Then
                If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
                Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
                rst = Common.LRT_Select
                If rst.Rows.Count > 0 Then
                    LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
                    ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
                End If
                
                OutBuffer = "Sanity Check Report Failed" & vbCrLf
                OutBuffer = OutBuffer & "LocationID: " & LocationID & vbCrLf
                OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
                OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
                OutBuffer = OutBuffer & " IP: " & LocalServerIP & vbCrLf
                OutBuffer = OutBuffer & vbCrLf & "Subject: Failed Sanity Check Report"
                Common.Send_Email(Common.Get_Error_Emails(6), Common.SystemEmailAddress, "Sanity Check Report Failed Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
            End If
      
    End If    ' end of Else ... (valid serial number)
    Common.Write_Log(LogFile, "Serial:" & LocalServerID & " LocationID: " & LocationID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " SanityCheck Finished!")
    
  End Sub
  
  '-----------------------------------------------------------------------------------------------
  
  Sub RemoveModifiedKeys(ByVal tableNumber As Integer, ByRef localKeyData As String, ByRef centralKeyData As String)
    
    Dim dt As DataTable = Nothing
    Dim row As DataRow = Nothing
    Dim exKeys As Hashtable = Nothing
    Dim centralKeyArr() As String = Nothing
    Dim localKeyArr() As String = Nothing
    Dim i As Integer
    Dim centralBuf As New StringBuilder()
    Dim localBuf As New StringBuilder()
    Dim keyList As String = ""
    Dim key As String
    
    ExceptionQuery = 5
    
    ' Run Modified Keys query stored in CheckTable
    If CheckTables(tableNumber, ExceptionQuery).Trim <> "" Then
      If (Common.LRTadoConn.State <> ConnectionState.Open) Then Common.Open_LogixRT()
      
      Common.QueryStr = CheckTables(tableNumber, ExceptionQuery)
      
      dt = Common.LRT_Select
      If (dt.Rows.Count > 0) Then
        ' Store keys in hashtable
        exKeys = New Hashtable(dt.Rows.Count)
        For Each row In dt.Rows
          exKeys.Add(Common.NZ(row.Item(0), "").ToString.Trim, Common.NZ(row.Item(0), "").ToString.Trim)
          keyList += row.Item(0) & ","
        Next
        
        ' split localKeyData, centralKeyData into tokens
        centralKeyArr = Split(centralKeyData, ",")
        localKeyArr = Split(localKeyData, ",")
        
        ' loop through each keyData and remove if keys match keys in hashtable
        For i = centralKeyArr.GetLowerBound(0) To centralKeyArr.GetUpperBound(0)
          key = centralKeyArr(i).Trim
          If (Not exKeys.ContainsKey(key)) Then
            If (centralBuf.Length > 0) Then centralBuf.Append(",")
            centralBuf.Append(key)
          End If
        Next
        
        centralKeyData = centralBuf.ToString
        centralBuf = Nothing
        
        For i = localKeyArr.GetLowerBound(0) To localKeyArr.GetUpperBound(0)
          key = localKeyArr(i).Trim
          If (Not exKeys.ContainsKey(key)) Then
            If (localBuf.Length > 0) Then localBuf.Append(",")
            localBuf.Append(key)
          End If
        Next
        localKeyData = localBuf.ToString
        localBuf = Nothing
      End If
    End If
    
  End Sub
  
</script>
<%
  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here
  
  Dim LocalServerID As Long
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
  Dim dst As DataTable
  Dim ProcessOK As Boolean
  Dim SerialOK As Boolean
  Dim MustIPL As Boolean
  Dim BannerID As Integer

  Common.AppName = "SanityCheck.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer
  IPAddress = Request.UserHostAddress
  
  LastHeard = "1/1/1980"
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "CPE-SanityCheckLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
  LocalServerIP = Trim(Request.QueryString("IP"))
  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Or MacAddress = "0" Then
    MacAddress = "0"
  End If
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
    LocalServerIP = Trim(Request.UserHostAddress)
  End If
    
  ServerSerialNum = Common.Extract_Val(Request.QueryString("serial"))
  LSVersion = Common.Extract_Val(Request.QueryString("lsversion"))
  LSBuild = Common.Extract_Val(Request.QueryString("lsbuild"))
  LocationID = Common.Extract_Val(Request.QueryString("locationid"))
  MD5 = Trim(Request.QueryString("md5"))
  
  Mode = UCase(Request.QueryString("mode"))
  If Mode = "" Then Mode = "FETCH"  
  
  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  'Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 2, IPAddress)
  
  Common.Write_Log(LogFile, "--------------------------------------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Common.AppName & "  -  " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & " Macddress:" & MacAddress & " IP:" & LocalServerIP & " Mode: " & Mode & " server: " & Environment.MachineName)

  ProcessOK = True
  SerialOK = False
  
  If ProcessOK Then
    Common.QueryStr = "dbo.pa_CPE_Gen_CheckSerial"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
    dst = Common.LRTsp_select
    Common.Close_LRTsp()
    If dst.Rows.Count > 0 Then
      If Common.NZ(dst.Rows(0).Item("NumRecs"), 0) > 0 Then SerialOK = True
    End If
    
    Common.QueryStr = "select LocationID from LocalServers with (NoLock) where LocalServerID = " & LocalServerID
    dst = Common.LRT_Select
    If dst.Rows.Count > 0 Then
      LocationID = Common.NZ(dst.Rows(0).Item("LocationID"), 0)
    End If
    dst = Nothing
    
    If Not (SerialOK) Then
      Send_Response_Header("Invalid SerialNumber", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "Returned: Invalid Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & vbCrLf)
      ProcessOK = False
    End If
  End If
  
  OperateAtEnterprise = False
  If Common.Fetch_CPE_SystemOption(91) = "1" Then
    OperateAtEnterprise = True
  End If

  Handle_Post()
  
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()

  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Total Run Time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
   
  
%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial:" & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
    Dim LocationName As String = ""
    Dim ExtLocationCode As String = ""
    Dim rst As New DataTable

    If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
    Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
    rst = Common.LRT_Select
    If rst.Rows.Count > 0 Then
        LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
        ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
    End If
    
    Dim ErrorMsg As String = "Sanity Check Error during Local Server Processing" & vbCrLf
    ErrorMsg = ErrorMsg & "LocationID: " & LocationID.ToString() & vbCrLf
    ErrorMsg = ErrorMsg & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
    ErrorMsg = ErrorMsg & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
    ErrorMsg = ErrorMsg & "MacAddress: " & MacAddress & vbCrLf
    ErrorMsg = ErrorMsg & "IP: " & LocalServerIP & vbCrLf
    ErrorMsg = ErrorMsg & vbCrLf & "Subject: Sanity Check Error during Local Server Processing"
    Common.Send_Email(Common.Get_Error_Emails(6), Common.SystemEmailAddress, "Sanity Check Error in Processing Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID.ToString(), ErrorMsg)
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  Common = Nothing
%>
