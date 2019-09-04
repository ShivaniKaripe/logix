<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO.Compression" %>
<% 
  ' *****************************************************************************
  ' * FILENAME: IPL-Offers.aspx 
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
    Public Connector As New Copient.ConnectorInc
    Public TextData As String
    Public IPL As Boolean
    Public LogFile As String
    Public FileStamp As String
    Public FileNum As Integer
    Public StartTime As Decimal
    Public TotalTime As Decimal
    Public ApplicationName As String
    Public ApplicationExtension As String
    Public gzStream As GZipStream = Nothing
    Public UncompressedSize As Long
    Public BufferedRecs As Long
    Public FlushTime As Decimal
    Public FlushStartTime As Decimal
    Public MacAddress As String
    Public LocalServerIP As String
    Public LSVerMajor As Integer
    Public LSVerMinor As Integer
    Public LSBuildMajor As Integer
    Public LSBuildMinor As Integer


    ' -------------------------------------------------------------------------------------------------

    Sub SD(ByVal OutStr As String)
        'PrintLine(FileNum, OutStr)
        Dim Bytes As Byte()
        Bytes = Encoding.UTF8.GetBytes(OutStr & vbCrLf)
        UncompressedSize = UncompressedSize + Bytes.Length
        gzStream.Write(Bytes, 0, Bytes.Length)
        Bytes = Nothing
        BufferedRecs = BufferedRecs + 1
        If BufferedRecs >= 5000 Then
            FlushStartTime = DateAndTime.Timer
            Response.Flush()
            FlushTime = FlushTime + (DateAndTime.Timer - FlushStartTime)
            BufferedRecs = 0
        End If
    End Sub

    ' -------------------------------------------------------------------------------------------------

    Sub SDb(ByVal OutStr As String)
        'Print(FileNum, OutStr)
        Dim Bytes As Byte()
        Bytes = Encoding.UTF8.GetBytes(OutStr)
        UncompressedSize = UncompressedSize + Bytes.Length
        gzStream.Write(Bytes, 0, Bytes.Length)
        Bytes = Nothing
        BufferedRecs = BufferedRecs + 1
        If BufferedRecs >= 5000 Then
            FlushStartTime = DateAndTime.Timer
            Response.Flush()
            BufferedRecs = 0
            FlushTime = FlushTime + (DateAndTime.Timer - FlushStartTime)
        End If
    End Sub

    ' -------------------------------------------------------------------------------------------------

    Function Parse_Bit(ByVal BooleanField As Boolean) As String
        If BooleanField Then
            Parse_Bit = "1"
        Else
            Parse_Bit = "0"
        End If
    End Function

    ' -------------------------------------------------------------------------------------------------

    Function Construct_Table(ByVal TableName As String, ByVal Operation As Integer, ByVal DelimChar As Integer, ByVal LocalServerID As String, ByVal LocationID As String, ByVal DBName As String) As String

        Dim TempResults As String
        Dim NumRecs As Long
        Dim dst As DataTable
        Dim row As DataRow
        Dim SQLCol As DataColumn
        Dim TempOut As String
        Dim Index As Integer
        Dim FieldList As String
        Dim DataBack As String
        Dim OperationType As Integer
        Dim TextMsg As String
        Dim BodyText As String
        Dim QueryStartTime As Decimal
        Dim QueryTotalTime As Decimal
        Dim ConstructStartTime As Decimal
        Dim ConstructTotalTime As Decimal
        Dim DelimCH As String

        Const BUFFERED_WRITE_SIZE As Integer = 10000
        'Send ("<!-- Table=" & TableName & " " & Format$(Time, "hh:mm:ss") & Format$(Timer - Fix(Timer), ".00") & " -->")

        TempOut = ""
        TempResults = ""
        NumRecs = 0
        FieldList = ""
        DataBack = ""
        DelimCH = Chr(DelimChar)

        Common.LRTTimeout = 1200
        Common.LWHTimeout = 1200
        Common.LXSTimeout = 1200

        QueryStartTime = DateAndTime.Timer
        If UCase(DBName) = "LXS" Then
            dst = Common.LXS_Select
        ElseIf UCase(DBName) = "LWH" Then
            dst = Common.LWH_Select
        ElseIf UCase(DBName) = "PMRT" Then
            dst = Common.PMRT_Select
        Else
            dst = Common.LRT_Select
        End If
        QueryTotalTime = DateAndTime.Timer - QueryStartTime
        ConstructStartTime = DateAndTime.Timer

        If dst.Rows.Count > 0 Then
            If UCase(TableName) = "INCENTIVES" And Operation = 1 Then
                FieldList = "IncentiveID" & DelimCH & "IncentiveName" & DelimCH & "Priority" & DelimCH & "StartDate" & DelimCH & "EndDate" & DelimCH & "TestingStartDate" & DelimCH & "TestingEndDate" & DelimCH & "EveryDOW" & DelimCH & "EligibilityStartDate" & DelimCH & "EligibilityEndDate" & DelimCH & "P1DistQtyLimit" & DelimCH & "P1DistPeriod" & DelimCH & "P1DistTimeType" & DelimCH & _
                            "P2DistQtyLimit" & DelimCH & "P2DistPeriod" & DelimCH & "P2DistTimeType" & DelimCH & "P3DistQtyLimit" & DelimCH & "P3DistPeriod" & DelimCH & "P3DistTimeType" & DelimCH & "Reporting" & DelimCH & "EmployeesOnly" & DelimCH & "EmployeesExcluded" & DelimCH & "UpdateLevel" & DelimCH & "DeferCalcToEOS" & DelimCH & "EveryTOD" & DelimCH & "ChargebackVendorID" & DelimCH & _
                            "SendIssuance" & DelimCH & "ManufacturerCoupon" & DelimCH & "InboundCRMEngineID" & DelimCH & "EnableImpressRpt" & DelimCH & "ClientOfferID" & DelimCH & "VendorCouponCode" & DelimCH & "EngineID" & DelimCH & "MutuallyExclusive" & DelimCH & "EngineSubTypeID" & DelimCH & "PromoClassID" & DelimCH & "RestrictedRedemption" & DelimCH & "ScorecardID" & DelimCH & "ScorecardDesc" & DelimCH & _
                            "PromptForReward"
            ElseIf UCase(TableName) = "INCENTIVETERMINALS" And Operation = 5 Then
                FieldList = "PKID" & DelimCH & "IncentiveID" & DelimCH & "TerminalTypeID" & DelimCH & "Excluded"
            ElseIf UCase(TableName) = "REWARDOPTIONS" And Operation = 1 Then
                FieldList = "RewardOptionID" & DelimCH & "Name" & DelimCH & "IncentiveID" & DelimCH & "Priority" & DelimCH & "HHEnable" & DelimCH & "TouchResponse" & DelimCH & "ProductComboID" & DelimCH & "ExcludedTender" & DelimCH & "ExcludedTenderAmtRequired" & DelimCH & "TierLevels" & DelimCH & "AttributeComboID" & DelimCH & "PreferenceComboID"
            ElseIf UCase(TableName) = "DELIVERABLES" And Operation = 5 Then
                FieldList = "DeliverableID" & DelimCH & "RewardOptionID" & DelimCH & "RewardOptionPhase" & DelimCH & "DeliverableTypeID" & DelimCH & "OutputID" & DelimCH & "AvailabilityTypeID" & DelimCH & "Priority" & DelimCH & "ScreenCellID"
            ElseIf UCase(TableName) = "EDISCOUNTS_TIERS" And Operation = 1 Then
                FieldList = "PKID" & DelimCH & "EDiscountID" & DelimCH & "TierLevel" & DelimCH & "ReceiptDescription" & DelimCH & "DiscountAmount" & DelimCH & "ItemLimit" & DelimCH & "WeightLimit" & DelimCH & "IsWeightTotal" & DelimCH & "DollarLimit" & DelimCH & "SPRepeatLevel"
            ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS" And Operation = 5 Then
                FieldList = "IncentiveProductGroupID" & DelimCH & "RewardOptionID" & DelimCH & "ProductGroupID" & DelimCH & "QtyForIncentive" & DelimCH & "QtyUnitType" & DelimCH & "AccumMin" & DelimCH & "AccumLimit" & DelimCH & "AccumPeriod" & DelimCH & "ExcludedProducts" & DelimCH & "Disqualifier" & DelimCH & "UniqueProduct" & DelimCH & "Rounding" & DelimCH & "MinPurchAmt"
            ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS_TIERS" And Operation = 5 Then
                FieldList = "PKID" & DelimCH & "IncentiveProductGroupID" & DelimCH & "TierLevel" & DelimCH & "Quantity"
            ElseIf UCase(TableName) = "INCENTIVEUSERGROUPS" And Operation = 5 Then
                FieldList = "IncentiveUserID" & DelimCH & "RewardOptionID" & DelimCH & "UserGroupID" & DelimCH & "ExcludedUsers"
            ElseIf UCase(TableName) = "VENDORS" And Operation = 5 Then
                FieldList = "VendorID" & DelimCH & "ExtVendorID"
            ElseIf UCase(TableName) = "EIWTRIGGERS" And Operation = 5 Then
                FieldList = "TriggerID" & DelimCH & "IncentiveID" & DelimCH & "RewardOptionID" & DelimCH & "IncentiveEIWID" & DelimCH & "TriggerTime" & DelimCH & "Consumed"
            ElseIf UCase(TableName) = "INCENTIVECARDTYPES" And Operation = 5 Then
                FieldList = "IncentiveCardTypeID" & DelimCH & "RewardOptionID" & DelimCH & "CardTypeID"
            Else
                'get the field list from the query results
                For Each SQLCol In dst.Columns
                    If Not (FieldList = "") Then FieldList = FieldList & Chr(DelimChar)
                    FieldList = FieldList & SQLCol.ColumnName
                Next
            End If
            OperationType = Operation
            If OperationType = 99 Then OperationType = 2
            'send the table header
            TempOut = "1:" & TableName & vbCrLf
            TempOut = TempOut & "2:" & Trim(Str(OperationType)) & vbCrLf
            TempOut = TempOut & "3:" & FieldList
            SD(TempOut)
            Common.Write_Log(LogFile, TempOut)
            NumRecs = 0

            Dim buf As New StringBuilder()

            If UCase(TableName) = "PRINTEDMESSAGES" And Operation = 5 Then '-- need to encode PrintedMessages because it allows \n
                For Each row In dst.Rows
                    TextMsg = Common.NZ(row.Item("TextMsg"), " ")
                    buf.AppendLine(Common.NZ(row.Item("MessageID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("PrintZone"), 1) & Chr(DelimChar) & Common.URL_Encode(Common.NZ(row.Item("TextMsg"), " ")) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("SuppressZeroBalance"), 0)))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    'Common.Write_Log(LogFile, Common.NZ(row.Item("MessageID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("Name"), 0) & Chr(DelimChar) & Replace(TextMsg, vbCrLf, "|", , , vbBinaryCompare))
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "PRINTEDMESSAGES_TIERS" And Operation = 5 Then '-- need to encode PrintedMessages because it allows \n
                For Each row In dst.Rows
                    BodyText = Common.NZ(row.Item("BodyText"), " ")
                    buf.AppendLine(Common.NZ(row.Item("PKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("MessageID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("TierLevel"), 0) & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("BodyText"), " ")))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "PMTRANSLATIONS" And Operation = 5 Then '-- need to encode PMTranslations.BodyText because it allows \r\n
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("PKID") & Chr(DelimChar) & row.Item("PMTiersID") & Chr(DelimChar) & row.Item("LanguageID") & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("BodyText"), " ")))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "INCENTIVES" And Operation = 1 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "INCENTIVETERMINALS" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "REWARDOPTIONS" And Operation = 1 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "DELIVERABLES" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "EDISCOUNTS_TIERS" And Operation = 1 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS_TIERS" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "INCENTIVEUSERGROUPS" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "INCENTIVECARDTYPES" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "VENDORS" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "EIWTRIGGERS" And Operation = 5 Then
                For Each row In dst.Rows
                    buf.AppendLine(row.Item("Data"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "DELIVERABLEPASSTHRUTIERS" And Operation = 5 Then '-- need to encode PrintedMessages because it allows \n
                For Each row In dst.Rows
                    buf.AppendLine(Common.NZ(row.Item("PKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("PTPKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("TierLevel"), 0) & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("Data"), " ")) & Chr(DelimChar) & row.Item("LanguageID"))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            ElseIf UCase(TableName) = "EDISCOUNTS" And Operation = 1 Then
                For Each row In dst.Rows
                    buf.AppendLine(Common.NZ(row.Item("EdiscountID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("Name"), 0) & Chr(DelimChar) & Common.NZ(row.Item("DiscountTypeID"), 0) & Chr(DelimChar) & _
                        Common.NZ(row.Item("ReceiptDescription"), 0) & Chr(DelimChar) & Common.NZ(row.Item("DiscountedProductGroupID"), 0) & Chr(DelimChar) & _
                        Common.NZ(row.Item("ExcludedProductGroupID"), 0) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("BestDeal"), 0)) & Chr(DelimChar) & _
                        Parse_Bit(Common.NZ(row.Item("AllowNegative"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("ComputeDiscount"), 0)) & Chr(DelimChar) & _
                        Math.Round(Common.NZ(row.Item("DiscountAmount"), 0), 3) & Chr(DelimChar) & Common.NZ(row.Item("AmountTypeID"), 0) & Chr(DelimChar) & _
                        Math.Round(Common.NZ(row.Item("L1Cap"), 0), 2) & Chr(DelimChar) & Math.Round(Common.NZ(row.Item("L2DiscountAmt"), 0), 0) & Chr(DelimChar) & _
                        Common.NZ(row.Item("L2AmountTypeID"), 0) & Chr(DelimChar) & Math.Round(Common.NZ(row.Item("L2Cap"), 0), 2) & Chr(DelimChar) & _
                        Math.Round(Common.NZ(row.Item("L3DiscountAmt"), 0), 2) & Chr(DelimChar) & Common.NZ(row.Item("L3AmountTypeID"), 0) & Chr(DelimChar) & _
                        Common.NZ(row("ItemLimit"), 1) & Chr(DelimChar) & Common.NZ(row.Item("WeightLimit"), 0) & Chr(DelimChar) & _
                        IIf(row.Item("IsWeightTotal") Is Nothing, "1", IIf(row.Item("IsWeightTotal"), "1", "0")) & Chr(DelimChar) & _
                        Math.Round(Common.NZ(row.Item("DollarLimit"), 0), 2) & Chr(DelimChar) & Common.NZ(row.Item("ChargeBackDeptID"), 0) & Chr(DelimChar) & _
                        Parse_Bit(Common.NZ(row.Item("DecliningBalance"), 0)) & Chr(DelimChar) & Common.NZ(row.Item("SVProgramID"), 0) & Chr(DelimChar) & _
                        Parse_Bit(Common.NZ(row.Item("FlexNegative"), 0)) & Chr(DelimChar) & Common.NZ(row.Item("ScorecardID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("ScorecardDesc"), 0) & Chr(DelimChar) & _
                        Math.Round(Common.NZ(row.Item("PercentFixedRounding"), 0), 2))
                    If ( buf.Length > BUFFERED_WRITE_SIZE )
                        SDb( buf.ToString() )
                        buf.Clear()
                    End If
                    'Common.Write_Log(LogFile, Common.NZ(row.Item("EdiscountID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("Name"), 0) & Chr(DelimChar) & Common.NZ(row.Item("ReceiptDescription"), 0) & Chr(DelimChar) & Common.NZ(row.Item("DiscountedProductGroupID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("ExcludedProductGroupID"), 0) & Chr(DelimChar) & Math.Round(Common.NZ(row.Item("DiscountAmount"), 0), 3) & Chr(DelimChar) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("DiscountTypeID"), 0)) & Common.NZ(row.Item("AmountTypeID"), 0) & Chr(DelimChar) & Math.Round(row.Item("L1Cap"), 2) & Chr(DelimChar) & Math.Round(row.Item("L2DiscountAmt"), 0) & Chr(DelimChar) & Common.NZ(row.Item("L2AmountTypeID"), 0) & Chr(DelimChar) & Math.Round(row.Item("L2Cap"), 2) & Chr(DelimChar) & Math.Round(row.Item("L3DiscountAmt"), 2) & Chr(DelimChar) & Common.NZ(row.Item("L3AmountTypeID"), 0) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("DecliningBalance"), 0)) & Chr(DelimChar) & Common.NZ(row.Item("ChargeBackDeptID"), 0) & Chr(DelimChar) & _
                    'Common.NZ(row("ItemLimit"), 1) & Chr(DelimChar) & Common.NZ(row.Item("WeightLimit"), 0) & Chr(DelimChar) & Math.Round(row.Item("DollarLimit"), 2) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("BestDeal"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("AllowNegative"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("ComputeDiscount"), 0)))
                    NumRecs = NumRecs + 1
                Next
            Else
                For Each row In dst.Rows
                    TempOut = ""
                    Index = 0
                    For Each SQLCol In dst.Columns
                        If Not (TempOut = "") Then
                            TempOut = TempOut & Chr(DelimChar)
                            SDb(Chr(DelimChar))
                        End If
                        If SQLCol.DataType.Name = "Boolean" Then 'if it is a binary field 
                            TempOut = TempOut & Parse_Bit(Common.NZ(row(Index), 0))
                            SDb(Parse_Bit(Common.NZ(row.Item(Index), 0)))
                        ElseIf SQLCol.DataType.Name = "Int32" Or SQLCol.DataType.Name = "Int64" Then 'if it is an Int or BigInt Field
                            TempOut = TempOut & Common.NZ(row(Index), 0)
                            SDb(Common.NZ(row(Index), 0))
                        Else 'else treat it as a string
                            TempOut = TempOut & Common.NZ(row(Index), "")
                            SDb(Common.NZ(row(Index), ""))
                        End If
                        Index = Index + 1
                    Next
                    SDb(vbCrLf)
                    'Common.Write_Log(LogFile, TempOut)
                    NumRecs = NumRecs + 1
                Next
                DataBack = "Sent Data"
            End If
            buf.AppendLine("###")
            SDb( buf.ToString() )
            Common.Write_Log(LogFile, "# records: " & NumRecs)
            ConstructTotalTime = DateAndTime.Timer - ConstructStartTime
            TotalTime = DateAndTime.Timer - StartTime
            Common.Write_Log(LogFile, "Query took " & Int(QueryTotalTime) & Format$(QueryTotalTime - Fix(QueryTotalTime), ".000") & "(sec) - Constructing data took " & Int(ConstructTotalTime) & Format$(ConstructTotalTime - Fix(ConstructTotalTime), ".000") & "(sec) - Total elapsed time " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)" & vbCrLf)

        End If  'rst.recordcount=0

        Construct_Table = DataBack

    End Function

    ' -------------------------------------------------------------------------------------------------

    Sub Construct_Output(ByVal LocalServerID As Long, ByVal LocationID As Long)

        Dim FailoverServer As Boolean
        Dim dst As DataTable
        Dim OutStr As String
        Dim TempOut As String
        Dim TestingLocation As Boolean
        Dim DelimChar As Integer
        'Dim ActiveROIDs As String
        'Dim ActiveIncentiveIDs As String
        Dim row As DataRow
        Dim MissingIncentives As String
        Dim Enterprise As Boolean

        DelimChar = 30
        TempOut = ""
        Enterprise = False
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Enterprise = True
        End If

        FailoverServer = False
        Common.QueryStr = "select FailoverServer from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            FailoverServer = Common.NZ(dst.Rows(0).Item("FailoverServer"), False)
        End If

        Common.QueryStr = "select TestingLocation from Locations with (NoLock) where LocationID=" & LocationID & ";"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            TestingLocation = Common.NZ(dst.Rows(0).Item("TestingLocation"), False)
        End If

        'Create a temp table to store the list of IncentiveIDs that need to be sent 
        Common.QueryStr = "BEGIN TRY DROP TABLE #ActiveIncentives; END TRY BEGIN CATCH END CATCH;" & _
                          "CREATE TABLE #ActiveIncentives (IncentiveID bigint, STIncentiveID bigint);"
        Common.LRT_Execute()
        'Create a temp table to store the list of ROIDs that need to be sent 
        Common.QueryStr = "BEGIN TRY DROP TABLE #ActiveROIDs; END TRY BEGIN CATCH END CATCH;" & _
                          "CREATE TABLE #ActiveROIDs (RewardOptionID bigint);"
        Common.LRT_Execute()

        'get a list of the active IncentiveIDs for this location (keep Deleted column)
        'ActiveIncentiveIDs = "-77"
        MissingIncentives = ""
        'Common.QueryStr = "select IncentiveID from CPE_IncentiveLoc_Func(" & LocationID & ");"

        If Enterprise Then
            If Common.Fetch_CPE_SystemOption(80) = 0 Then  'lock offers after expiration is turned off
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "  from OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  " inner join CPE_IncentiveLocationsView_Enterprise as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID " & _
                                  "  left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                  " where(OLU.LocationID=" & LocationID & ") " & _
                                  " order by OfferID;"
            Else  'lock offers after expiration is turned on
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "from OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  "inner join CPE_IncentiveLocationsView_Enterprise as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID " & _
                                  "left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                  "where(OLU.LocationID=" & LocationID & ") and (dateadd(d, 1, STI.EndDate)>getdate() or dateadd(d, 1, STI.EligibilityEndDate)>getdate() or  dateadd(d, 1, STI.TestingEndDate)>getdate()) " & _
                                  "order by OfferID;"
            End If
        Else
            If Common.Fetch_CPE_SystemOption(80) = 0 Then   'lock offers after expiration is turned off
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "  from OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  " inner join CPE_IncentiveLocationsView as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID and ILV.LocationID=" & LocationID & " " & _
                                  "  left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                  " where(OLU.LocationID=" & LocationID & ") " & _
                                  " order by OfferID;"
            Else   'lock offers after expiration is turned on
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "from OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  "inner join CPE_IncentiveLocationsView as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID and ILV.LocationID=" & LocationID & " " & _
                                  "left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                  "where(OLU.LocationID=" & LocationID & ") and (dateadd(d, 1, STI.EndDate)>getdate() or dateadd(d, 1, STI.EligibilityEndDate)>getdate() or  dateadd(d, 1, STI.TestingEndDate)>getdate()) " & _
                                  "order by OfferID;"
            End If
        End If
        Common.LRT_Execute()

        Common.QueryStr = "select IncentiveID from #ActiveIncentives where STIncentiveID=-1;"
        dst = Common.LRT_Select
        For Each row In dst.Rows
            MissingIncentives = MissingIncentives & row.Item("IncentiveID") & ","
        Next
        If Not (MissingIncentives = "") Then
            'get rid of the comma at the end of the list of missing incentives
            MissingIncentives = Left(MissingIncentives, Len(MissingIncentives) - 1)
            Common.Write_Log(LogFile, "Serial= " & LocalServerID & " IPAddress=" & Trim(Request.UserHostAddress) & " The following offers cannot be deployed during an IPL of LocationID (" & LocationID & ") because they are missing from the shadow tables.  These offers will have to be redeployed: " & MissingIncentives & " Serial= " & LocalServerID & "; Mac IPAddress=" & (Trim(Request.UserHostAddress)) & " Server = " & Environment.MachineName, True)
            Common.LastErrorTime = "1/1/1980 00:00:00"
            Common.Error_Processor("Serial=" & LocalServerID & " IPAddress=" & Trim(Request.UserHostAddress) & " The following offers cannot be deployed during an IPL of LocationID (" & LocationID & ") because they are missing from the shadow tables.  These offers will have to be redeployed: " & MissingIncentives, "The following offers cannot be deployed during an IPL of LocationID (" & LocationID & ") because they are missing from the shadow tables.  These offers will have to be redeployed: " & MissingIncentives, , , 0)
            Common.LastErrorTime = "1/1/1980 00:00:00"
        End If

        'get rid of the rows where the IncentiveIDs are missing from the shadow tables
        Common.QueryStr = "Delete from #ActiveIncentives where STIncentiveID=-1;"
        Common.LRT_Execute()
        'Common.Write_Log(LogFile, "ActiveIncentiveIDs=" & ActiveIncentiveIDs)

        Common.QueryStr = "Insert into #ActiveROIDs (RewardOptionID) " & _
                          "  select distinct RO.rewardOptionID " & _
                          "  from CPE_ST_RewardOptions as RO with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=RO.IncentiveID and RO.Deleted=0 " & _
                          "  order by RO.RewardOptionID;"
        Common.LRT_Execute()
        'Common.Write_Log(LogFile, "ActiveROIDs=" & ActiveROIDs)

        Common.Write_Log(LogFile, "Returned the following data:")
        OutStr = ""
        OutStr = OutStr & ApplicationName & "," & Connector.CSMajorVersion & "," & Connector.CSMinorVersion & "," & Connector.CSBuild & "," & Connector.CSBuildRevision & vbCrLf
        OutStr = OutStr & "D:" & vbCrLf
        OutStr = OutStr & "T:" & vbCrLf
        OutStr = OutStr & "LocationID=" & LocationID
        SD(OutStr)
        Common.Write_Log(LogFile, OutStr)

        ' Look up TransNumSeed for LocalServer and IPLSequenceNum for Location
        '   TransNumSeed = IPLSequenceNum = MAX(TransNumSeed, IPLSequenceNum) + 1
        '   This guarantees that both numbers are always increasing
        '   and prevents us from having to make an interface change.
        Common.QueryStr = "SELECT LogixTransNumSeed FROM [dbo].[LocalServers] WHERE LocalServerID=" & LocalServerID
        Dim transNumSeedResult As DataTable = Common.LRT_Select()
        If transNumSeedResult.Rows.Count <> 1
            Throw New Exception("Found unexpected condition")
        End If
        Dim TransNumSeed As Integer = transNumSeedResult.Rows(0).Item("LogixTransNumSeed")
        Common.QueryStr = "SELECT IPLSequenceNum FROM [dbo].[LocationSeqNum] WHERE LocationID=" & LocationID
        Dim IPLSequenceNumResult As DataTable = Common.LXS_Select()
        Dim IPLSequenceNum As Integer = 1
        If IPLSequenceNumResult.Rows.Count <> 1
            ' This should _never_ happen.
            Common.QueryStr = "INSERT INTO [dbo].[LocationSeqNum] (LocationID) VALUES (" & LocationID & ")"
            Common.LXS_Execute()
        Else
            IPLSequenceNum = IPLSequenceNumResult.Rows(0).Item("IPLSequenceNum")
        End If

        Dim nextSequenceNum As Integer = Math.Max(TransNumSeed, IPLSequenceNum) + 1

        'Update the LogixTransNumSeed for this LocalServer
        Common.QueryStr = "Update LocalServers with (RowLock) set LogixTransNumSeed=" & nextSequenceNum & " WHERE LocalServerID=" & LocalServerID
        Common.LRT_Execute()

        ' Update the IPLSequenceNum(s) for this LocalServer
        Common.QueryStr = "UPDATE [dbo].[LocationSeqNum] with (RowLock) SET IPLSequenceNum=" & nextSequenceNum & " WHERE LocationID=" & LocationID
        Common.LXS_Execute()

        'LocalServers
        'send the data from the LocalServers table
        Common.QueryStr = "select LocalServerID, LocationID, IncentiveUpdateFreq, TransactionUpdateFreq, IncentiveFetchURL, ImageFetchURL, OfflineFTPUser, OfflineFTPPass, OfflineFTPPath, OfflineFTPIP, LogixTransNumSeed, PhoneHomeIPOverride from LocalServers with (NoLock) where LocalServerID=" & LocalServerID & ";"
        OutStr = OutStr & Construct_Table("LocalServers", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'Locations
        'send the data from the Locations table
        If Not (Enterprise) Then
            'send only the one row for the location performing the IPL
            Common.QueryStr = "select LocationID, LocationName, Address1, Address2, City, State, Zip, ExtLocationCode as ClientLocationCode, TestingLocation, LocationTypeID from Locations with (NoLock) where Deleted=0 and LocationID=" & LocationID & ";"
            Construct_Table("Locations", 1, DelimChar, LocalServerID, LocationID, "LRT")
            'LocationLanguages
            Common.QueryStr = "select PKID, LocationID, LanguageID, Required from LocationLanguages with (NoLock) where Deleted=0 and LocationID=" & LocationID & ";"
            Construct_Table("LocationLanguages", 5, DelimChar, LocalServerID, LocationID, "LRT")
            'YellowBoxes
            Common.QueryStr = "select BoxID, LocationID as RetailLocationID, InStoreLocationID, PrinterTypeID, OpDisplayTypeID from CPE_YellowBoxes with (NoLock) where LocationID=" & LocationID & ";"
            OutStr = OutStr & Construct_Table("YellowBoxes", 1, DelimChar, LocalServerID, LocationID, "LXS")
        Else
            'send all of the rows from the Locations table
            Common.QueryStr = "select LocationID, LocationName, Address1, Address2, City, State, Zip, ExtLocationCode as ClientLocationCode, TestingLocation, LocationTypeID, TimeZone from Locations with (NoLock) where Deleted=0;"
            Construct_Table("Locations", 1, DelimChar, LocalServerID, LocationID, "LRT")
            'LocationLanguages
            Common.QueryStr = "select PKID, LocationID, LanguageID, Required from LocationLanguages with (NoLock) where Deleted=0;"
            Construct_Table("LocationLanguages", 5, DelimChar, LocalServerID, LocationID, "LRT")
            'YellowBoxes
            Common.QueryStr = "select BoxID, LocationID as RetailLocationID, InStoreLocationID, PrinterTypeID, OpDisplayTypeID from CPE_YellowBoxes with (NoLock);"
            OutStr = OutStr & Construct_Table("YellowBoxes", 1, DelimChar, LocalServerID, LocationID, "LXS")
        End If

        'LocalServer_Seeds
        'send the data to indicate the LocalID seed numbers for RewardAccumulation, RewardDistribution, and StoredValue
        Common.QueryStr = "exec pa_CPE_LocalID_Seeds " & LocalServerID
        OutStr = OutStr & Construct_Table("LocalIDSeeds", 5, DelimChar, LocalServerID, LocationID, "LXS")

        'PromoEngines
        Common.QueryStr = "select EngineID, Description, DefaultEngine, Installed from PromoEngines with (NoLock)"
        OutStr = OutStr & Construct_Table("PromoEngines", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'PromoEngineSubTypes
        Common.QueryStr = "select PKID, PromoEngineID, SubTypeID, SubTypeName, Installed, ReplayEnabled from PromoEngineSubTypes with (NoLock)"
        OutStr = OutStr & Construct_Table("PromoEngineSubTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'ExtCRMInterfaces
        Common.QueryStr = "select ExtInterfaceID, Name, Description from ExtCRMInterfaces where Deleted=0;"
        OutStr = OutStr & Construct_Table("ExtCRMInterfaces", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'Integrations
        Common.QueryStr = "select IntegrationID, Name, Installed from Integrations;"
        OutStr = OutStr & Construct_Table("Integrations", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'OpDisplaytypes
        'send the date from the OpDisplayTypes Table
        Common.QueryStr = "select OpDisplayTypeID, Name from CPE_OpDisplayTypes ODT with (NoLock);"
        OutStr = OutStr & Construct_Table("OpDisplayTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'RemoteDataOptions
        Common.QueryStr = "select RDO.PKID, RDO.RemoteDataTypeID, RDO.StyleID, RDO.Enabled " & _
                          "from RemoteDataOptions as RDO with (NoLock) Inner Join RemoteDataTypes as RDT with (NoLock) on RDO.RemoteDataTypeID=RDT.RemoteDataTypeID and RDT.EngineID=2;"
        OutStr = OutStr & Construct_Table("RemoteDataOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_SystemOptions
        Common.QueryStr = "select OptionID, OptionName, OptionValue from CPE_SystemOptions with (NoLock);"
        TempOut = Construct_Table("CPE_SystemOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'SystemOptions
        Common.QueryStr = "select OptionID, OptionName, OptionValue from SystemOptions with (NoLock);"
        TempOut = Construct_Table("SystemOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'Languages
        Common.QueryStr = "select LanguageID, Name, MSNetCode, JavaLocaleCode, RightToLeftText, AvailableForCustFacing " & _
                          "from Languages with (NoLock);"
        Construct_Table("Languages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_TenderTypes
        Common.QueryStr = "select TenderTypeID, Name, ExtTenderType, ExtVariety, ExtBinNum from CPE_TenderTypes where Deleted=0;"
        TempOut = Construct_Table("TenderTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'ProductTypes	CLOUDSOL-1874 added table
        Common.QueryStr = "select ProductTypeID, Name, PhraseID, HasAttributes, PaddingLength, MaxLength, IsNumeric, LastUpdate from ProductTypes with (NoLock);"
        TempOut = Construct_Table("ProductTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'PrinterTypes
        Common.QueryStr = "select PrinterTypeID, PageWidth, Name, MaxLines from PrinterTypes with (NoLock);"
        TempOut = Construct_Table("PrinterTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'MarkupTags
        Common.QueryStr = "select distinct MT.MarkupID, case when MT.NumParams>0 then '|'+Tag+'[' else '|'+Tag+'|' end as Tag from Markuptags as MT with (NoLock) Inner Join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID where EngineID=2"
        TempOut = Construct_Table("MarkupTags", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'PrinterTranslation
        Common.QueryStr = "select distinct PT.TranslationID, PT.PrinterTypeID, PT.MarkupID, ControlChars from PrinterTranslation as PT with (NoLock) Inner Join MarkupTagUsage as MTU with (NoLock) on PT.MarkupID=MTU.MarkUpID where EngineID=2;"
        TempOut = Construct_Table("PrinterTranslation", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'ScreenLayouts
        Common.QueryStr = "select LayoutID, Name, Width, Height from ScreenLayouts with (NoLock) where Deleted=0;"
        TempOut = Construct_Table("ScreenLayouts", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'PassThruInterfaceTypes
        Common.QueryStr = "select LSInterfaceID, Description from PassThruInterfaceTypes with(NoLock);"
        Construct_Table("PassThruInterfaceTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")



        'CashierMessages -- leave fields for backward compatibility
        Common.QueryStr = "select CMT.MessageID PKID, CMT.Line1, CMT.Line2, CMT.Beep, CMT.BeepDuration, CM.PLU " & _
                          "  from CPE_ST_CashierMessageTiers as CMT with (NoLock) " & _
                          " Inner Join CPE_ST_Deliverables as D with (NoLock) on CMT.MessageID=D.OutputID " & _
                          " Inner Join CPE_ST_CashierMessages as CM with (NoLock) on CMT.MessageID=CM.MessageID " & _
                          "   and D.DeliverableTypeID=9 and D.Deleted=0 and CM.PLU=0 " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        TempOut = Construct_Table("CashierMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CashierMessageTiers
        Common.QueryStr = "select CMT.PKID, CMT.Line1, CMT.Line2, CMT.Beep, CMT.BeepDuration, CMT.MessageID, CMT.TierLevel, CMT.DisplayImmediate, CM.PLU " & _
                          "  from CPE_ST_CashierMessageTiers as CMT with (NoLock) " & _
                          " Inner Join CPE_ST_Deliverables as D with (NoLock) on CMT.MessageID=D.OutputID " & _
                          " Inner Join CPE_ST_CashierMessages as CM with (NoLock) on CMT.MessageID=CM.MessageID " & _
                          "   and D.DeliverableTypeID=9 and D.Deleted=0 and CM.PLU=0 " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        TempOut = Construct_Table("CashierMessages_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CashierMsgTranslations
        Common.QueryStr = "select CMTrans.PKID, CMTrans.CashierMsgTierID, CMTrans.LanguageID, isnull(CMTrans.Line1, '') as Line1, isnull(CMTrans.Line2, '') as Line2 " & _
                          "from CPE_ST_CashierMsgTranslations as CMTrans with (NoLock) Inner Join CPE_ST_CashierMessageTiers as CMT with (NoLock) on CMTrans.CashierMsgTierID=CMT.PKID " & _
                          "Inner Join CPE_ST_Deliverables as D with (NoLock) on CMT.MessageID=D.OutputID and D.DeliverableTypeID=9 and D.Deleted=0 " & _
                          "Inner Join CPE_ST_CashierMessages as CM with (NoLock) on CMT.MessageID=CM.MessageID and CM.PLU=0 " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        Construct_Table("CashierMsgTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'Send the special "PLU Not Used" (return to customer) cashier message
        Common.QueryStr = "select CMT.MessageID as PKID, CMT.Line1, CMT.Line2, CMT.Beep, CMT.BeepDuration, CM.PLU " & _
                          "  from CPE_CashierMessageTiers as CMT with (NoLock) " & _
                          " inner join CPE_CashierMessages as CM with (NoLock) on CM.MessageID=CMT.MessageID " & _
                          " where CMT.TierLevel=1 and CM.PLU=1;"
        Construct_Table("CashierMessages", 1, DelimChar, LocalServerID, LocationID, "LRT")
        Common.QueryStr = "select CMT.PKID, CMT.Line1, CMT.Line2, CMT.Beep, CMT.BeepDuration, CMT.MessageID, CMT.TierLevel, CMT.DisplayImmediate, CM.PLU " & _
                          "  from CPE_CashierMessageTiers as CMT with (NoLock) " & _
                          " inner join CPE_CashierMessages as CM with (NoLock) on CM.MessageID=CMT.MessageID " & _
                          " where CM.PLU=1;"
        Construct_Table("CashierMessages_Tiers", 1, DelimChar, LocalServerID, LocationID, "LRT")



        'PassThrus
        Common.QueryStr = "select DPT.PKID, DPT.DeliverableID, DPT.PassThruRewardID, D.RewardOptionID, DPT.LSInterfaceID, DPT.ActionTypeID " & _
                          "  from CPE_ST_PassThrus as DPT with (NoLock) Inner Join CPE_ST_Deliverables as D with (NoLock) on DPT.PKID=D.OutputID " & _
                          "    Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "  where D.DeliverableTypeID=12 and D.Deleted=0;"
        Construct_Table("DeliverablePassThrus", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PassThruTiers
        Common.QueryStr = "select DPTT.PKID, DPTT.PTPKID, DPTT.TierLevel, DPTT.Data, DPTT.LanguageID " & _
                          "  from CPE_ST_Deliverables as D with (NoLock) Inner Join CPE_ST_PassThrus as DPT with (NoLock) on DPT.PKID=D.OutputID " & _
                          "    Inner Join CPE_ST_PassThruTiers as DPTT with (NoLock) on DPTT.PTPKID=DPT.PKID " & _
                          "    Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "  where D.DeliverableTypeID=12 and D.Deleted=0;"
        Construct_Table("DeliverablePassThruTiers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'ScreenCells
        Common.QueryStr = "select CellID, LayoutID, ContentsID, X, Y, Width, Height, BackgroundImg from ScreenCells with (NoLock) where Deleted=0;"
        TempOut = Construct_Table("ScreenCells", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'ScreenCellContents
        Common.QueryStr = "select ContentsID, Name from ScreenCellContents with (NoLock) where Deleted=0;"
        TempOut = Construct_Table("ScreenCellContents", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'TouchAreas
        Common.QueryStr = "select AreaID, TA.OnScreenAdID, Name, X, Y, Width, Height " & _
                          "from TouchAreas as TA with (NoLock) Inner Join OnScreenAdLocUpdate as OSALU with (NoLock) on TA.OnScreenAdID=OSALU.OnScreenAdID and TA.Deleted=0 " & _
                          "Where OSALU.LocationID=" & LocationID & ";"
        TempOut = Construct_Table("TouchAreas", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'Scorecards
        Common.QueryStr = "select ScorecardID, ScorecardTypeID, Description, Priority, Bold, PrintTotalLine, PreferenceListItemName, PrintZeroBalance from ScoreCards where Deleted=0 and EngineID in (2,6);"
        TempOut = Construct_Table("ScoreCards", 1, DelimChar, LocalServerID, LocationID, "LRT")



        'DeliverableROIDs
        Common.QueryStr = "select DR.PKID, DR.DeliverableID, DR.AreaID, DR.RewardOptionID, DR.IncentiveID " & _
                          "  from CPE_ST_DeliverableROIDs DR with (NoLock) " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on DR.DeliverableID=D.DeliverableID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where D.DeliverableTypeID=1 and D.Deleted=0 and DR.Deleted=0;"
        TempOut = Construct_Table("DeliverableROIDs", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'PointsPrograms
        Common.QueryStr = "select ProgramID, ProgramName, ScorecardID, ScorecardDesc, ScorecardBold, AdjustmentUPC, isnull(CAMProgram, 0) as CAMProgram, ExtHostTypeID, ExtHostProgramID, ExtHostPartnerCode, ExtHostPartnerID, ExtHostFuelProgram, ExtHostCardBINMin, ExtHostCardBINMax from PointsPrograms with (NoLock) where Deleted=0;"
        TempOut = Construct_Table("PointsPrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PointsProgramTranslations
        Common.QueryStr = "select PPT.PKID, PPT.ProgramID, PPT.LanguageID, isnull(PPT.ScorecardDesc, '') as ScorecardDesc " & _
                          "from PointsProgramTranslations as PPT with (NoLock);"
        Construct_Table("PointsProgramTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'StoredValuePrograms
        Common.QueryStr = "select SVProgramID, Name, Value, OneUnitPerRec, SVExpireType, SVExpirePeriodType, ExpirePeriod, ExpireTOD, ExpireDate, " & _
                          " ScorecardID, ScorecardDesc, ScorecardBold, AdjustmentUPC " & _
                          "  from StoredValuePrograms with (NoLock) where Deleted=0;"
        TempOut = Construct_Table("StoredValuePrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'SVProgramTranslations
        Common.QueryStr = "select SVT.PKID, SVT.SVProgramID, SVT.LanguageID, isnull(SVT.ScorecardDesc, '') as ScorecardDesc " & _
                          "from SVProgramTranslations as SVT with (NoLock);"
        Construct_Table("SVProgramTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentivePointsGroups
        Common.QueryStr = "select IPtG.IncentivePointsID, IPtG.RewardOptionID, IPtG.ProgramID, IPtGT.Quantity QtyForIncentive " & _
                          "  from CPE_ST_IncentivePointsGroups as IPtG with (NoLock) " & _
                          " inner join CPE_ST_IncentivePointsGroupTiers as IPtGT with (NoLock) on IPtG.IncentivePointsID=IPtGT.IncentivePointsID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPtG.RewardOptionID " & _
                          " where IPtGT.TierLevel=1 and IPtG.Deleted=0;"
        TempOut = Construct_Table("IncentivePointsGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentivePointsGroups_Tiers
        Common.QueryStr = "select IPtGT.PKID, IPtGT.IncentivePointsID, IPtGT.TierLevel, IPtGT.Quantity " & _
                          "  from CPE_ST_IncentivePointsGroupTiers as IPtGT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPtGT.RewardOptionID;"
        TempOut = Construct_Table("IncentivePointsGroups_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT") '-- no need RewardOptionID


        'DeliverablePoints
        Common.QueryStr = "select DP.PKID, DP.DeliverableID, DP.ProgramID, DP.RewardOptionID, DPT.Quantity, DP.ChargebackDeptID, DP.ScorecardID, " & _
                          "       DP.ScorecardDesc, DP.ScorecardBold " & _
                          "from CPE_ST_DeliverablePoints as DP with (NoLock) " & _
                          "  inner join CPE_ST_DeliverablePointTiers as DPT on DP.PKID=DPT.DPPKID " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on DP.PKID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "where DPT.TierLevel=1 and D.DeliverableTypeID=8 and D.Deleted=0 and DP.Deleted=0;"
        TempOut = Construct_Table("DeliverablePoints", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverablePointsTranslations
        Common.QueryStr = "select DPT.PKID, DPT.DeliverablePointsID, DPT.LanguageID, isnull(DPT.ScorecardDesc, '') as ScorecardDesc " & _
                          "from CPE_ST_DeliverablePointsTranslations as DPT with (NoLock) inner join CPE_ST_DeliverablePoints as DP with (NoLock) on DPT.DeliverablePointsID=DP.PKID " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on DP.PKID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "where D.DeliverableTypeID=8 and D.Deleted=0 and DP.Deleted=0;"
        Construct_Table("DeliverablePointsTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverablePoints_Tiers
        Common.QueryStr = "select DPT.PKID, DPT.DPPKID, DPT.TierLevel, DPT.Quantity " & _
                          "  from CPE_ST_DeliverablePointTiers as DPT with (NoLock) " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on DPT.DPPKID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "  and D.DeliverableTypeID=8 and D.Deleted=0;"
        TempOut = Construct_Table("DeliverablePoints_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentiveStoredValuePrograms
        Common.QueryStr = "select ISVP.IncentiveStoredValueID, ISVP.RewardOptionID, ISVP.SVProgramID, ISVPT.Quantity QtyForIncentive " & _
                          "from CPE_ST_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                          "  inner join CPE_ST_IncentiveStoredValueProgramTiers as ISVPT with (NoLock) on ISVP.IncentiveStoredValueID=ISVPT.IncentiveStoredValueID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ISVP.RewardOptionID " & _
                          "where ISVPT.TierLevel=1 and Deleted=0;"
        TempOut = Construct_Table("IncentiveStoredValuePrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveStoredValuePrograms_Tiers (no need to send IPGT.RewardOptionID)
        Common.QueryStr = "select ISVPT.PKID, ISVPT.IncentiveStoredValueID, ISVPT.TierLevel, ISVPT.Quantity " & _
                          "  from CPE_ST_IncentiveStoredValueProgramTiers as ISVPT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ISVPT.RewardOptionID;"
        TempOut = Construct_Table("IncentiveStoredValuePrograms_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'DeliverableStoredValue
        Common.QueryStr = "select DSV.PKID, DSV.DeliverableID, DSV.SVProgramID, DSV.RewardOptionID, DSV.Quantity, DSV.ScorecardID, DSV.ScorecardDesc, DSV.ScorecardBold " & _
                          "from CPE_ST_DeliverableStoredValue as DSV with (NoLock) " & _
                          "  inner join CPE_ST_DeliverableStoredValueTiers as DSVT with (NoLock) on DSV.PKID=DSVT.DSVPKID " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on DSV.PKID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "where DSVT.TierLevel=1 and D.DeliverableTypeID=11 and D.Deleted=0 and DSV.Deleted=0;"
        TempOut = Construct_Table("DeliverableStoredValue", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverableSVTranslations
        Common.QueryStr = "select DSVT.PKID, DSVT.DeliverableSVID, DSVT.LanguageID, isnull(DSVT.ScorecardDesc, '') as ScorecardDesc " & _
                          "from CPE_ST_DeliverableSVTranslations as DSVT with (NoLock) inner join CPE_ST_DeliverableStoredValue as DSV with (NoLock) on DSVT.DeliverableSVID=DSV.PKID " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on DSV.PKID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "where D.DeliverableTypeID=11 and D.Deleted=0 and DSV.Deleted=0;"
        TempOut = Construct_Table("DeliverableSVTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverableStoredValueTiers
        Common.QueryStr = "select DSVT.PKID, DSVT.DSVPKID, DSVT.TierLevel, DSVT.Quantity " & _
                          "  from CPE_ST_DeliverableStoredValueTiers as DSVT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DSVT.DSVPKID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where D.DeliverableTypeID=11 and D.Deleted=0;"
        TempOut = Construct_Table("DeliverableStoredValue_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'LocationOptions
        If Not (Enterprise) Then
            Common.QueryStr = "select PKID, LocationID, OptionID, OptionValue from LocationOptions with (NoLock) where Deleted=0 and LocationID=" & LocationID & ";"
        Else
            Common.QueryStr = "select PKID, LocationID, OptionID, OptionValue from LocationOptions with (NoLock) where Deleted=0;"
        End If
        TempOut = Construct_Table("LocationOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'SiteSpecificOptions
        Common.QueryStr = "select SSO.OptionID, OptionName from SiteSpecificOptions SSO with (NoLock);"
        TempOut = Construct_Table("SiteSpecificOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'SiteSpecificOptionValues
        Common.QueryStr = "select PKID, OptionID, ValueDescription, OptionValue, DefaultVal " & _
                          "from SiteSpecificOptionValues SSOV with (NoLock);"
        TempOut = Construct_Table("SiteSpecificOptionValues", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'Incentives
        'Common.QueryStr = "select distinct I.IncentiveID, I.IncentiveName, I.Priority, I.StartDate, I.EndDate, I.TestingStartDate, I.TestingEndDate, isnull(EveryDOW, 1) as EveryDOW, isnull(I.EligibilityStartDate, I.StartDate) as EligibilityStartDate, isnull(I.EligibilityEndDate, I.EndDate) as EligibilityEndDate, P1DistQtyLimit, P1DistPeriod, P1DistTimeType, " & _
        '                  "P2DistQtyLimit, P2DistPeriod, P2DistTimeType, P3DistQtyLimit, P3DistPeriod, P3DistTimeType, isnull(EnableRedeemRpt, 1) as Reporting, EmployeesOnly, EmployeesExcluded, UpdateLevel, DeferCalcToEOS, EveryTOD, ChargebackVendorID, SendIssuance, ManufacturerCoupon, InboundCRMEngineID, EnableImpressRpt, ClientOfferID, VendorCouponCode, isnull(EngineID, 2) as EngineID, MutuallyExclusive, EngineSubTypeID, isnull(PromoClassID, 0) as PromoClassID " & _
        '                  "From CPE_ST_Incentives as I with (NoLock) where I.IncentiveID in (" & ActiveIncentiveIDs & ");"
        Common.QueryStr = "select convert(nvarchar,I.IncentiveID)+char(" & DelimChar & ")+I.IncentiveName+char(" & DelimChar & ")+convert(nvarchar,isnull(I.Priority,0))+char(" & DelimChar & ")+convert(nvarchar,I.StartDate,120)+char(" & DelimChar & ")+convert(nvarchar,I.EndDate,120)+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,I.TestingStartDate,120)+char(" & DelimChar & ")+convert(nvarchar,I.TestingEndDate,120)+char(" & DelimChar & ")+convert(nvarchar,isnull(EveryDOW, 1))+char(" & DelimChar & ")+convert(nvarchar,isnull(I.EligibilityStartDate, I.StartDate),120)+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,isnull(I.EligibilityEndDate, I.EndDate),120)+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistQtyLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistTimeType,0))+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,isnull(P2DistQtyLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P2DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P2DistTimeType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P3DistQtyLimit,0))+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,isnull(P3DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P3DistTimeType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(EnableRedeemRpt, 1))+char(" & DelimChar & ")+convert(nvarchar,EmployeesOnly)+char(" & DelimChar & ")+convert(nvarchar,EmployeesExcluded)+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,UpdateLevel)+char(" & DelimChar & ")+convert(nvarchar,DeferCalcToEOS)+char(" & DelimChar & ")+convert(nvarchar,EveryTOD)+char(" & DelimChar & ")+convert(nvarchar,ChargebackVendorID)+char(" & DelimChar & ")+convert(nvarchar,SendIssuance)+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,ManufacturerCoupon)+char(" & DelimChar & ")+convert(nvarchar,InboundCRMEngineID)+char(" & DelimChar & ")+convert(nvarchar,EnableImpressRpt)+char(" & DelimChar & ")+isnull(ClientOfferID,'')+char(" & DelimChar & ")+isnull(VendorCouponCode,'')+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,isnull(EngineID, 2))+char(" & DelimChar & ")+convert(nvarchar,MutuallyExclusive)+char(" & DelimChar & ")+convert(nvarchar,EngineSubTypeID)+char(" & DelimChar & ")+convert(nvarchar,isnull(PromoClassID, 0))+char(" & DelimChar & ")+convert(nvarchar,isnull(RestrictedRedemption,0))+char(" & DelimChar & ")+" & _
                          "convert(nvarchar,isnull(ScorecardID, 0))+char(" & DelimChar & ")+isnull(ScorecardDesc,'')+char(" & DelimChar & ")+convert(nvarchar,isnull(PromptForReward,0)) " & _
                          "as Data " & _
                          "From CPE_ST_Incentives as I with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=I.IncentiveID;"
        TempOut = Construct_Table("Incentives", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'OfferTranslations
        'send the data from the OfferTranslations table
        Common.QueryStr = "select OT.PKID, OT.OfferID, OT.OfferName, OT.LimitScorecardDesc, OT.LanguageID " & _
                          "from CPE_ST_OfferTranslations as OT with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=OT.OfferID;"
        Construct_Table("OfferTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveDOW
        'send the data from the IncentiveDOW table
        Common.QueryStr = "select IDOW.IncentiveDOWID, IDOW.IncentiveID, IDOW.DOWID " & _
                          "from CPE_ST_IncentiveDOW as IDOW with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=IDOW.IncentiveID and IDOW.Deleted=0;"
        OutStr = OutStr & Construct_Table("IncentiveDOW", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveTOD
        'send the data from the IncentiveTOD table
        Common.QueryStr = "select ITOD.IncentiveTODID, ITOD.IncentiveID, ITOD.StartTime, ITOD.EndTime " & _
                          "  from CPE_ST_IncentiveTOD as ITOD with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=ITOD.IncentiveID;"
        OutStr = OutStr & Construct_Table("IncentiveTOD", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'TerminalTypes (in-store locations)
        Common.QueryStr = "select TerminalTypeID, Name, LayoutID, SpecificPromosOnly, FuelProcessing, AnyTerminal, LockingGroupID from TerminalTypes with (NoLock) where EngineID=2 and Deleted=0;"
        OutStr = OutStr & Construct_Table("TerminalTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'OfferTerminals (in-store locations associative table)
        'Common.QueryStr = "select OT.PKID, OT.OfferID as IncentiveID, OT.TerminalTypeID, OT.Excluded " & _
        '                  "  from CPE_ST_OfferTerminals as OT with (NoLock) " & _
        '                  " inner join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
        '                  " where T.EngineID=2 and OT.OfferID in (" & ActiveIncentiveIDs & ");"
        Common.QueryStr = "select convert(nvarchar,OT.PKID)+char(" & DelimChar & ")+convert(nvarchar,isnull(OT.OfferID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(OT.TerminalTypeID,0))+char(" & DelimChar & ")+convert(nvarchar,OT.Excluded) as Data " & _
                          "from CPE_ST_OfferTerminals as OT with (NoLock) " & _
                          "  inner join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                          "  Inner Join #ActiveIncentives as AI on AI.IncentiveID=OT.OfferID " & _
                          "where T.EngineID=2;"
        OutStr = OutStr & Construct_Table("IncentiveTerminals", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'ChargebackDepts
        Common.QueryStr = "select ChargeBackDeptID as DeptID, ExternalID as DeptNumber from ChargebackDepts with (NoLock);"
        OutStr = OutStr & Construct_Table("ChargebackDepts", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'RewardOptions
        'Common.QueryStr = "select RO.RewardOptionID, RO.Name, RO.IncentiveID, RO.Priority, RO.HHEnable, RO.TouchResponse, RO.ProductComboID, RO.ExcludedTender, " & _
        '                  "       RO.ExcludedTenderAmtRequired, RO.TierLevels, RO.AttributeComboID " & _
        '                  "  from CPE_ST_RewardOptions as RO with (NoLock) where RO.RewardOptionID in (" & ActiveROIDs & ");"
        Common.QueryStr = "select convert(nvarchar,RO.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(Name,''))+char(" & DelimChar & ")+convert(nvarchar,RO.IncentiveID)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.Priority, 0))+char(" & DelimChar & ")+convert(nvarchar,RO.HHEnable)+char(" & DelimChar & ")+convert(nvarchar,RO.TouchResponse)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.ProductComboID,0))+char(" & DelimChar & ")+convert(nvarchar,RO.ExcludedTender)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.ExcludedTenderAmtRequired,0))+char(" & DelimChar & ")+convert(nvarchar,RO.TierLevels)+char(" & DelimChar & ")+convert(nvarchar,RO.AttributeComboID)+char(" & DelimChar & ")+convert(nvarchar,RO.PreferenceComboID) as Data " & _
                          "from CPE_ST_RewardOptions as RO with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=RO.RewardOptionID and RO.Deleted=0;"
        OutStr = OutStr & Construct_Table("RewardOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")


        'Deliverables
        'Common.QueryStr = "select D.DeliverableID, D.RewardOptionID, D.RewardOptionPhase, D.DeliverableTypeID, D.OutputID, " & _
        '                  "       1 as AvailabilityTypeID, D.Priority, D.ScreenCellID " & _
        '                  "  from CPE_ST_Deliverables as D with (NoLock) where D.RewardOptionID in (" & ActiveROIDs & ") and D.Deleted=0;"
        Common.QueryStr = "select convert(nvarchar,D.DeliverableID)+char(" & DelimChar & ")+convert(nvarchar,D.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(D.RewardOptionPhase,0))+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(D.DeliverableTypeID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(D.OutputID,0))+char(" & DelimChar & ")+'1'+char(" & DelimChar & ")+convert(nvarchar,isnull(D.Priority,0))" & _
                          "+char(" & DelimChar & ")+convert(nvarchar,isnull(D.ScreenCellID,0)) as Data " & _
                          "from CPE_ST_Deliverables as D with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID and D.Deleted=0;"
        OutStr = OutStr & Construct_Table("Deliverables", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'OnScreenAds
        If Not (FailoverServer) Then
            'get the OnScreenAd data where the the corresponding record in the OnScreenAdLocUpdate table has
            'an older LastSent date or is non-existant (for that AdID and LocationID

            'the first part of this query is for reward graphics
            'the second part is for the special case of background images
            Common.QueryStr = "select distinct OSA.OnScreenAdID, OSA.Name, OSA.StoreResponse, OSA.DisplayDuration, OSA.Width, OSA.Height, OSA.ImageType, OSA.UpdateLevel " & _
                              "  from OnScreenAds as OSA with (NoLock) inner join OnScreenAdLocUpdate as OSALU with (NoLock) on OSA.OnScreenAdID=OSALU.OnScreenAdID and OSA.Deleted=0 " & _
                              " where OSALU.LocationID=" & LocationID & " " & _
                              "union " & _
                              "select distinct OSA.OnScreenAdID, OSA.Name, OSA.StoreResponse, OSA.DisplayDuration, OSA.Width, OSA.Height, OSA.ImageType, OSA.UpdateLevel " & _
                              "  from OnScreenAds as OSA with (NoLock) inner join ScreenCells as SC with (NoLock) on OSA.OnScreenAdID=SC.BackgroundImg " & _
                              " where OSA.Deleted=0 and SC.Deleted=0;"
            TempOut = Construct_Table("OnScreenAds", 5, DelimChar, LocalServerID, LocationID, "LRT")
            'get the GRAPHIC data from the OnScreenAds Table
            'the first part of this query is for reward graphics
            'the second part is for the special case of background images
            Common.QueryStr = "select OSA.OnscreenAdID, Imagetype, " & _
                              "Case ImageType " & _
                              "  When 1 then convert(varchar, OSA.OnscreenAdID)+'img.'+'jpg' " & _
                              "  When 2 then convert(varchar, OSA.OnscreenAdID)+'img.'+'gif' " & _
                              "END as Graphic, OSA.MD5sum " & _
                              "from OnScreenAds as OSA with (NoLock) Inner Join OnScreenAdLocUpdate as OSALU with (NoLock) on OSA.OnScreenAdID=OSALU.OnScreenAdID and OSA.Deleted=0 and isnull(OSA.GraphicSize, 0)>0 " & _
                              "Where OSALU.LocationID=" & LocationID & " " & _
                              "union " & _
                              "select distinct OSA.OnscreenAdID, Imagetype, " & _
                              "Case ImageType " & _
                              "  When 1 then convert(varchar, OSA.OnscreenAdID)+'img.'+'jpg' " & _
                              "  When 2 then convert(varchar, OSA.OnscreenAdID)+'img.'+'gif' " & _
                              "END as Graphic, OSA.MD5sum " & _
                              "from OnScreenAds as OSA with (NoLock) Inner Join ScreenCells as SC with (NoLock) on OSA.OnScreenAdID=SC.BackgroundImg and OSA.Deleted=0 and SC.Deleted=0 and isnull(OSA.GraphicSize, 0)>0;"
            TempOut = TempOut & Construct_Table("OnScreenAds", 3, DelimChar, LocalServerID, LocationID, "LRT")
            'update the existing OnScreenAdLocUpdate records for this location with the current time/date and set the WaitingACK bit
            Common.QueryStr = "update OnScreenAdLocUpdate with (RowLock) set LastSent=getdate() " & _
                              "Where LocationID=" & LocationID & ";"
            Common.LRT_Execute()
        Else
            Common.QueryStr = "update OnScreenAdLocUpdate with (RowLock) set LastSent='1/1/1981' where LocationID=" & LocationID & ";"
            Common.LRT_Execute()
        End If


        'PrintedMessages
        Common.QueryStr = "select isnull(PM.MessageID, 0) as MessageID, MessageTypeID as PrintZone, BodyText as TextMsg, isnull(PM.SuppressZeroBalance, 0) as SuppressZeroBalance " & _
                          "from CPE_ST_PrintedMessages as PM with (NoLock) " & _
                          "inner join CPE_ST_PrintedMessageTiers as PMT with (NoLock) on PMT.MessageID = PM.MessageID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on PMT.MessageID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where PMT.TierLevel=1 and D.DeliverableTypeID=4 and D.Deleted=0;"
        TempOut = Construct_Table("PrintedMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PrintedMessages_Tiers
        Common.QueryStr = "select PMT.PKID, isnull(PMT.MessageID, 0) as MessageID, PMT.TierLevel, PMT.BodyText " & _
                          "  from CPE_ST_PrintedMessageTiers as PMT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on PMT.MessageID=D.OutputID " & _
                          "   and D.DeliverableTypeID=4 and D.Deleted=0 " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        TempOut = Construct_Table("PrintedMessages_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PMTranslations
        Common.QueryStr = "select PMTrans.PKID, PMTrans.PMTiersID, PMTrans.LanguageID, PMTrans.BodyText " & _
                          "from CPE_ST_PMTranslations as PMTrans with (NoLock) Inner Join CPE_ST_PrintedMessageTiers as PMT with (NoLock) on PMTrans.PMTiersID=PMT.PKID " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on PMT.MessageID=D.OutputID and D.DeliverableTypeID=4 and D.Deleted=0 " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        Construct_Table("PMTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'EDiscounts
        'get the EDiscount records 
        Common.QueryStr = "select distinct ED.DiscountID as EDiscountID, ED.Name, ED.DiscountTypeID, EDT.ReceiptDescription, ED.DiscountedProductGroupID, ED.ExcludedProductGroupID, " & _
                          "       ED.BestDeal, ED.AllowNegative, ED.ComputeDiscount, EDT.DiscountAmount, ED.AmountTypeID, ED.L1Cap, ED.L2DiscountAmt, ED.L2AmountTypeID, ED.L2Cap, " & _
                          "       ED.L3DiscountAmt, ED.L3AmountTypeID, EDT.ItemLimit, EDT.WeightLimit, EDT.IsWeightTotal, EDT.DollarLimit, ED.ChargeBackDeptID, ED.DecliningBalance, " & _
                          "       ED.SVProgramID, IsNull(ED.FlexNegative, 0) As FlexNegative, ED.ScorecardID, ED.ScorecardDesc, ED.PercentFixedRounding " & _
                          "  from CPE_ST_Discounts as ED with (NoLock) " & _
                          " inner join CPE_ST_DiscountTiers as EDT with (NoLock) on ED.DiscountID=EDT.DiscountID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on ED.DiscountID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where EDT.TierLevel=1 and D.DeliverableTypeID=2 and D.Deleted=0 and ED.Deleted=0;"
        TempOut = Construct_Table("EDiscounts", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'DiscountTranslations
        Common.QueryStr = "select DT.PKID, DT.DiscountID, DT.LanguageID, isnull(DT.ScorecardDesc, '') as ScorecardDesc " & _
                          "from CPE_ST_DiscountTranslations as DT with (NoLock) inner join CPE_ST_Deliverables as D with (NoLock) on DT.DiscountID=D.OutputID " & _
                          "    and D.DeliverableTypeID=2 and D.Deleted=0 " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        Construct_Table("DiscountTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'EDiscount_Tiers
        'get the EDiscount records 
        'Common.QueryStr = "select EDT.PKID, EDT.DiscountID as EDiscountID, EDT.TierLevel, EDT.ReceiptDescription, EDT.DiscountAmount, " & _
        '                  "       EDT.ItemLimit, EDT.WeightLimit, EDT.DollarLimit, EDT.SPRepeatLevel " & _
        '                  "  from CPE_ST_DiscountTiers as EDT with (NoLock) " & _
        '                  " inner join CPE_ST_Deliverables as D with (NoLock) on EDT.DiscountID=D.OutputID " & _
        '                  "   and D.DeliverableTypeID=2 and D.RewardOptionID in (" & ActiveROIDs & ") and D.Deleted=0"
        Common.QueryStr = "select convert(nvarchar,EDT.PKID)+char(" & DelimChar & ")+convert(nvarchar,EDT.DiscountID)+char(" & DelimChar & ")+convert(nvarchar,EDT.TierLevel)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(EDT.ReceiptDescription,''))+char(" & DelimChar & ")+convert(nvarchar,EDT.DiscountAmount)+char(" & DelimChar & ")+convert(nvarchar,EDT.ItemLimit)" & _
                          "+char(" & DelimChar & ")+convert(nvarchar,EDT.WeightLimit)+char(" & DelimChar & ")+convert(nvarchar,EDT.IsWeightTotal)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,EDT.DollarLimit)+char(" & DelimChar & ")+convert(nvarchar,isnull(EDT.SPRepeatLevel,0)) as Data " & _
                          "from CPE_ST_DiscountTiers as EDT with (NoLock) " & _
                          "  inner join CPE_ST_Deliverables as D with (NoLock) on EDT.DiscountID=D.OutputID " & _
                          "    and D.DeliverableTypeID=2 and D.Deleted=0 " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        TempOut = Construct_Table("EDiscounts_Tiers", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'DiscountTiersTranslations
        Common.QueryStr = "select DTT.PKID, DTT.DiscountTiersID, DTT.LanguageID, isnull(DTT.ReceiptDesc, '') as ReceiptDescription, isnull(DTT.BuyDesc, '') as BuyDescription " & _
                          "from CPE_ST_DiscountTiersTranslations as DTT with (NoLock) Inner Join CPE_ST_DiscountTiers as EDT with (NoLock) on DTT.DiscountTiersID=EDT.PKID " & _
                          "inner join CPE_ST_Deliverables as D with (NoLock) on EDT.DiscountID=D.OutputID " & _
                          "and D.DeliverableTypeID=2 and D.Deleted=0 " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        Construct_Table("DiscountTiersTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'SpecialPricing
        'get the CPE_SpecialPricing records
        Common.QueryStr = "select SP.SpecialPricingID, SP.DiscountID, SP.DiscountTierID, SP.Value, SP.LevelID " & _
                          "  from CPE_ST_SpecialPricing AS SP with (NoLock) " & _
                          " inner join CPE_ST_DiscountTiers AS DT with (NoLock) on DT.PKID = SP.DiscountTierID " & _
                          " inner join CPE_ST_Discounts as ED with (NoLock) on SP.DiscountID=ED.DiscountID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DT.DiscountID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "where D.DeliverableTypeID=2 and D.Deleted=0 and ED.Deleted=0"
        TempOut = Construct_Table("SpecialPricing", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'FrankingMessages
        'get the FrankingMessages records 
        Common.QueryStr = "select distinct FMT.FrankID, FMT.OpenDrawer, FMT.ManagerOverride, FMT.FrankFlag, FMT.FrankingText, FMT.Line1, FMT.Line2, FMT.Beep, FMT.BeepDuration " & _
                          "  from CPE_ST_FrankingMessageTiers as FMT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on FMT.FrankID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where FMT.TierLevel=1 and D.DeliverableTypeID=10 and D.Deleted=0;"
        TempOut = Construct_Table("FrankingMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'FrankingMessages_Tiers
        Common.QueryStr = "select FMT.PKID, FMT.OpenDrawer, FMT.ManagerOverride, FMT.FrankFlag, FMT.FrankingText, FMT.Line1, FMT.Line2, FMT.Beep, FMT.BeepDuration, FMT.FrankID, FMT.TierLevel " & _
                          "  from CPE_ST_FrankingMessageTiers as FMT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on FMT.FrankID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where D.DeliverableTypeID=10 and D.Deleted=0;"
        TempOut = Construct_Table("FrankingMessages_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'DeliverableCustomerGroupTiers
        Common.QueryStr = "select DCGT.PKID, DCGT.DeliverableID, DCGT.CustomerGroupID, DCGT.TierLevel " & _
                          "  from CPE_ST_DeliverableCustomerGroupTiers as DCGT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DCGT.DeliverableID=D.DeliverableID " & _
                          "   and D.DeliverableTypeID=5 and D.Deleted=0 " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        OutStr = OutStr & Construct_Table("DeliverableCustomerGroups_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")





        'IncentiveProductGroups (keep Deleted Column)
        'get new data from the IncentiveProductGroups table (keep Deleted column)
        'Common.QueryStr = "select IPG.IncentiveProductGroupID, IPG.RewardOptionID, IPG.ProductGroupID, IPGT.Quantity QtyForIncentive, IPG.QtyUnitType, IPG.AccumMin, " & _
        '                  "       IPG.AccumLimit, IPG.AccumPeriod, IPG.ExcludedProducts, IPG.Disqualifier, IPG.UniqueProduct, IPG.Rounding, IPG.MinPurchAmt " & _
        '                  "  from CPE_ST_IncentiveProductGroups as IPG with (NoLock) " & _
        '                  "       inner join CPE_ST_IncentiveProductGroupTiers as IPGT with (NoLock) on IPG.IncentiveProductGroupID=IPGT.IncentiveProductGroupID " & _
        '                  "  where IPGT.TierLevel=1 and IPG.Deleted=0 and ExcludedProducts=0 and IPG.RewardOptionID in (" & ActiveROIDs & ")" & _
        '                  "union " & _
        '                  "select IPG.IncentiveProductGroupID, IPG.RewardOptionID, IPG.ProductGroupID, IPG.QtyForIncentive, IPG.QtyUnitType, IPG.AccumMin, " & _
        '                  "       IPG.AccumLimit, IPG.AccumPeriod, IPG.ExcludedProducts, IPG.Disqualifier, IPG.UniqueProduct, IPG.Rounding, IPG.MinPurchAmt " & _
        '                  "  from CPE_ST_IncentiveProductGroups as IPG with (NoLock) " & _
        '                  "  where IPG.Deleted=0 and IPG.ExcludedProducts=1 and IPG.RewardOptionID in (" & ActiveROIDs & ");"
        Common.QueryStr = "select convert(nvarchar,IPG.IncentiveProductGroupID)+char(" & DelimChar & ")+convert(nvarchar,IPG.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ProductGroupID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPGT.Quantity,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.QtyUnitType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.AccumMin,0))+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(IPG.AccumLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.AccumPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,IPG.ExcludedProducts)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.Disqualifier,0))+char(" & DelimChar & ")+convert(nvarchar,IPG.UniqueProduct)+char(" & DelimChar & ")+convert(nvarchar,IPG.Rounding)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.MinPurchAmt,0)) as Data " & _
                          "from CPE_ST_IncentiveProductGroups as IPG with (NoLock) " & _
                          "inner join CPE_ST_IncentiveProductGroupTiers as IPGT with (NoLock) on IPG.IncentiveProductGroupID=IPGT.IncentiveProductGroupID " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPG.RewardOptionID " & _
                          "where IPGT.TierLevel=1 and IPG.Deleted=0 and ExcludedProducts=0 " & _
                          "union " & _
                          "select convert(nvarchar,IPG.IncentiveProductGroupID)+char(" & DelimChar & ")+convert(nvarchar,IPG.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ProductGroupID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.QtyForIncentive,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.QtyUnitType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.AccumMin,0))+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(IPG.AccumLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.AccumPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,IPG.ExcludedProducts)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.Disqualifier,0))+char(" & DelimChar & ")+convert(nvarchar,IPG.UniqueProduct)+char(" & DelimChar & ")+convert(nvarchar,IPG.Rounding)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.MinPurchAmt,0)) as Data " & _
                          "from CPE_ST_IncentiveProductGroups as IPG with (NoLock) " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPG.RewardOptionID " & _
                          "where IPG.Deleted=0 and IPG.ExcludedProducts=1;"
        OutStr = OutStr & Construct_Table("IncentiveProductGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentiveProductGroupTiers (no need to send IPGT.RewardOptionID)
        'get new data from the IncentiveProductGroups table 
        'Common.QueryStr = "select IPGT.PKID, IPGT.IncentiveProductGroupID, IPGT.TierLevel, IPGT.Quantity " & _
        '                  "  from CPE_ST_IncentiveProductGroupTiers as IPGT with (NoLock) where IPGT.RewardOptionID in (" & ActiveROIDs & ")"
        Common.QueryStr = "select convert(nvarchar,IPGT.PKID)+char(" & DelimChar & ")+convert(nvarchar,IPGT.IncentiveProductGroupID)+char(" & DelimChar & ")+convert(nvarchar,IPGT.TierLevel)+char(" & DelimChar & ")+convert(nvarchar,IPGT.Quantity) as Data " & _
                          "from CPE_ST_IncentiveProductGroupTiers as IPGT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPGT.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentiveProductGroups_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentiveUserGroups (keep Deleted column) 
        'get new data from the IncentiveUserGroups table
        'Common.QueryStr = "select ICG.IncentiveCustomerID as IncentiveUserID, ICG.RewardOptionID, ICG.CustomerGroupID as UserGroupID, ICG.ExcludedUsers " & _
        '                  "  from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) " & _
        '                  " where ICG.Deleted=0 and ICG.RewardOptionID in (" & ActiveROIDs & ");"
        Common.QueryStr = "select convert(nvarchar,ICG.IncentiveCustomerID)+char(" & DelimChar & ")+convert(nvarchar,ICG.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(ICG.CustomerGroupID,0))+char(" & DelimChar & ")+convert(nvarchar,ICG.ExcludedUsers) as Data " & _
                          "from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ICG.RewardOptionID " & _
                          "where ICG.Deleted=0;"
        OutStr = OutStr & Construct_Table("IncentiveUserGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveCardTypes
        Common.QueryStr = "select convert(nvarchar,ICT.IncentiveCardTypeID)+char(" & DelimChar & ")+convert(nvarchar,ICT.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(ICT.CardTypeID,0)) as Data " & _
                          "from CPE_ST_IncentiveCardTypes as ICT with (NoLock) " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ICT.RewardOptionID " & _
                          "where ICT.Deleted=0;"
        OutStr = OutStr & Construct_Table("IncentiveCardTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentiveTenderTypes
        Common.QueryStr = "select ITT.IncentiveTenderID, ITT.RewardOptionID, ITT.TenderTypeID, ITTT.Value " & _
                          "  from CPE_ST_IncentiveTenderTypes as ITT with (NoLock) " & _
                          " inner join CPE_ST_IncentiveTenderTypeTiers as ITTT with (NoLock) on ITT.IncentiveTenderID=ITTT.IncentiveTenderID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ITT.RewardOptionID " & _
                          " where ITTT.TierLevel=1 and ITT.Deleted=0;"
        OutStr = OutStr & Construct_Table("IncentiveTenderTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveTenderType_Tiers (no need to send IPGT.RewardOptionID)
        Common.QueryStr = "select ITTT.PKID, ITTT.IncentiveTenderID, ITTT.TierLevel, ITTT.Value " & _
                          "  from CPE_ST_IncentiveTenderTypeTiers as ITTT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ITTT.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentiveTenderTypes_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentiveInstantWin
        Common.QueryStr = "select IWIN.IncentiveInstantWinID as incentiveinstantwinid, IWIN.RewardOptionID, IWIN.NumPrizesAllowed as reward, IWIN.OddsOfWinning as odds, IWIN.RandomWinners as Random " & _
                          "  from CPE_ST_IncentiveInstantWin as IWIN with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IWIN.RewardOptionID and IWIN.Deleted=0;"
        OutStr = OutStr & Construct_Table("IncentiveInstantWinPrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentivePLUs
        Common.QueryStr = "select IPLU.IncentivePLUID, IPLU.RewardOptionID, IPLU.PLU, IPLU.PerRedemption, IPLU.CashierMessage" &
                          "  from CPE_ST_IncentivePLUs as IPLU with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPLU.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentivePLUs", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentiveEIW
        Common.QueryStr = "select IEIW.IncentiveEIWID, IEIW.RewardOptionID, NumberOfPrizes, FrequencyID " & _
                          "  from CPE_ST_IncentiveEIW as IEIW with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IEIW.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentiveEIW", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'EIWTriggers
        'Common.QueryStr = "select EIWT.TriggerID, EIWT.IncentiveID, EIWT.RewardOptionID, EIWT.IncentiveEIWID, EIWT.TriggerTime, CASE ISNULL(EIWU.TriggerID, 0) WHEN 0 THEN 0 ELSE 1 END AS Consumed " & _
        '                  "  FROM CPE_EIWTriggers as EIWT with (NoLock) LEFT OUTER JOIN CPE_EIWTriggersUsed AS EIWU WITH (NoLock) ON EIWU.TriggerID = EIWT.TriggerID " & _
        '                  "  where EIWT.Removed = 0 AND EIWU.TriggerID IS NULL;"
        Common.QueryStr = "select convert(nvarchar,EIWT.TriggerID)+char(" & DelimChar & ")+convert(nvarchar,isnull(EIWT.IncentiveID,0))+char(" & DelimChar & ")+CONVERT(nvarchar,isnull(EIWT.RewardOptionID,0))+char(" & DelimChar & ")" & _
                          "+CONVERT(nvarchar,isnull(EIWT.IncentiveEIWID,0))+char(" & DelimChar & ")+CONVERT(nvarchar(45),isnull(EIWT.TriggerTime,'1/1/1980'))+char(" & DelimChar & ")+CONVERT(nvarchar,CASE ISNULL(EIWU.TriggerID, 0) WHEN 0 THEN 0 ELSE 1 END) as Data " & _
                          "FROM CPE_EIWTriggers as EIWT with (NoLock) INNER JOIN #ActiveIncentives as AI ON AI.IncentiveID = EIWT.IncentiveID LEFT OUTER JOIN CPE_EIWTriggersUsed AS EIWU WITH (NoLock) ON EIWU.TriggerID = EIWT.TriggerID " & _
                          "where EIWT.Removed = 0 AND EIWU.TriggerID IS NULL;"
        OutStr = OutStr & Construct_Table("EIWTriggers", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'Vendors
        'Common.QueryStr = "select VendorID, ExtVendorID " & _
        '                  "  FROM Vendors with (NoLock) " & _
        '                  "  WHERE Deleted=0;"
        Common.QueryStr = "select convert(nvarchar,VendorID)+char(" & DelimChar & ")+convert(nvarchar,ExtVendorID) as Data " & _
                          "  FROM Vendors with (NoLock) " & _
                          "  WHERE Deleted=0;"
        OutStr = OutStr & Construct_Table("Vendors", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'OfferCategories
        Common.QueryStr = "select OfferCategoryID, ExtCategoryID, isnull(BaseOfferID, 0) as BaseOfferID " & _
                          "  FROM OfferCategories with (NoLock) " & _
                          "  WHERE Deleted=0;"
        OutStr = OutStr & Construct_Table("OfferCategories", 5, DelimChar, LocalServerID, LocationID, "LRT")


        'IncentiveAttributes
        Common.QueryStr = "select IA.IncentiveAttributeID, IA.RewardOptionID " & _
                          "  FROM CPE_ST_IncentiveAttributes as IA with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IA.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentiveAttributes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveAttributeTiers
        Common.QueryStr = "select IAT.PKID, IAT.IncentiveAttributeID, IAT.RewardOptionID, IAT.AttributeTypeID, IAT.TierLevel, IAT.AttributeValues " & _
                          "  FROM CPE_ST_IncentiveAttributeTiers as IAT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IAT.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentiveAttributeTiers", 5, DelimChar, LocalServerID, LocationID, "LRT")



        'CPE_IncentivePrefs
        Common.QueryStr = "select IP.IncentivePrefsID, IP.RewardOptionID, IP.PreferenceID " & _
                      "  FROM CPE_ST_IncentivePrefs as IP with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IP.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentivePrefs", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_IncentivePrefTiers
        Common.QueryStr = "select IPT.IncentivePrefTiersID, IPT.IncentivePrefsID, IPT.TierLevel, IPT.ValueComboTypeID " & _
                  "  FROM CPE_ST_IncentivePrefTiers as IPT with (NoLock) Inner Join CPE_ST_IncentivePrefs as IP on IPT.IncentivePrefsID=IP.IncentivePrefsID " & _
                  "    Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IP.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentivePrefTiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_IncentivePrefTierValues
        Common.QueryStr = "select IPTV.PKID, IPTV.IncentivePrefTiersID, IPTV.Value, IPTV.OperatorTypeID, ValueTypeID, DateOperatorTypeID, ValueModifier " & _
                  "  FROM CPE_ST_IncentivePrefTierValues as IPTV with (NoLock) Inner Join CPE_ST_IncentivePrefTiers as IPT with (NoLock) on IPTV.IncentivePrefTiersID=IPT.IncentivePrefTiersID" & _
                  "  Inner Join CPE_ST_IncentivePrefs as IP on IPT.IncentivePrefsID=IP.IncentivePrefsID " & _
                  "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IP.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentivePrefTierValues", 5, DelimChar, LocalServerID, LocationID, "LRT")


        If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
            Common.QueryStr = "select ChannelID, Name, Installed, IssuanceDisabled from Channels where ChannelID=1;"
            OutStr = OutStr & Construct_Table("Channels", 1, DelimChar, LocalServerID, LocationID, "PMRT")

            Common.QueryStr = "select P.PreferenceID, P.Name, P.DataTypeID, P.MultiValue, " & _
                              "(case when P.IssuanceDisabled=1 Then 1 when isnull(PCID.PKID, 0)>0 then 1 Else 0 END) as IssuanceDisabled " & _
                              "from Preferences as P with (NoLock) Left Join PrefChannelsIssuanceDisabled as PCID with (NoLock) on PCID.PreferenceID=P.PreferenceID and PCID.ChannelID=1 " & _
                              "Inner Join PreferenceChannels as PC with (NoLock) on PC.PreferenceID=P.PreferenceID " & _
                              "where PC.ChannelID=1 and P.Deleted = 0;"
            OutStr = OutStr & Construct_Table("Preferences", 1, DelimChar, LocalServerID, LocationID, "PMRT")

            Common.QueryStr = "select PDV.PreferenceID, PDV.DefaultValue " & _
                              "from PrefDefaultValues as PDV with (NoLock) Inner Join PreferenceChannels as PC with (NoLock) on PC.PreferenceID=PDV.PreferenceID " & _
                              "Inner Join Preferences as P with (NoLock) on P.PreferenceID=PDV.PreferenceID " & _
                              "where PC.ChannelID=1 and P.Deleted = 0;"
            OutStr = OutStr & Construct_Table("PrefDefaultValues", 1, DelimChar, LocalServerID, LocationID, "PMRT")

            'MetaPrefs
            Common.QueryStr = "select PKID, Name, PreferenceID from MetaPrefs;"
            Construct_Table("MetaPrefs", 5, DelimChar, LocalServerID, LocationID, "PMRT")
        End If


        'CardTypes
        Common.QueryStr = "select CardTypeID, Description, CustTypeID, ExtCardTypeID, PaddingLength, MaxIDLength, NumericOnly " & _
                          "  FROM CardTypes with (NoLock);"
        OutStr = OutStr & Construct_Table("CardTypes", 5, DelimChar, LocalServerID, LocationID, "LXS")


        If Enterprise Then
            'send the LocationGroups table
            Common.QueryStr = "select LocationGroupID, AllLocations from LocationGroups where Deleted=0;"
            OutStr = OutStr & Construct_Table("LocationGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")

            'send the OfferLocations table
            Common.QueryStr = "select PKID, OfferID, LocationGroupID, Excluded from OfferLocations with (NoLock) where Deleted=0;"
            OutStr = OutStr & Construct_Table("OfferLocations", 5, DelimChar, LocalServerID, LocationID, "LRT")

            'send the LocGroupItems table
            Common.QueryStr = "select PKID, LocationGroupID, LocationID from LocGroupItems with (NoLock) where Deleted=0;"
            OutStr = OutStr & Construct_Table("LocGroupItems", 5, DelimChar, LocalServerID, LocationID, "LRT")
        End If


        SD("***") 'send the EOF marker

        Common.QueryStr = "drop table #ActiveROIDs; drop table #ActiveIncentives;"
        Common.LRT_Execute()

    End Sub

    ' -------------------------------------------------------------------------------------------------

    Sub Process_ACK(ByVal LocalServerID As Long, ByVal LocationID As Long)

        Common.QueryStr = "Update LocalServers with (RowLock) set IncentiveLastHeard=getdate(), MustIPL=0 where LocalServerID='" & LocalServerID & "';"
        Common.LRT_Execute()
        Send_Response_Header(ApplicationName & " - ACK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("ACK Received")
        Common.Write_Log(LogFile, "ACK Received")

    End Sub

    ' -------------------------------------------------------------------------------------------------

    Sub Process_NAK(ByVal LocalServerID As String, ByVal LocationID As String)

        Dim ErrorMsg As String
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Trim(Request.QueryString("IP"))
        MacAddress = Trim(Request.QueryString("mac"))
        If MacAddress = "" Or MacAddress = "0" Then
            MacAddress = "0"
        End If
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If

        ErrorMsg = Trim(Request.QueryString("errormsg"))
        Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server: " & Environment.MachineName & "Received NAK - ErrorMsg:" & ErrorMsg)
        Send_Response_Header(ApplicationName & " - NAK Received", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)

    End Sub

    ' -------------------------------------------------------------------------------------------------

    Function IPLStartOk(ByVal LocalServerID As Long) As String

        Dim StartResponse As String
        Dim MaxIPLs As Integer
        Dim WindowMinutes As Integer
        Dim IPLRunawayTime As Integer
        Dim NumIPLs As Integer
        Dim LastIPLAge As Long

        If Not (Request.QueryString("nothrottle") = "") Then
            StartResponse = ""
            Return StartResponse
            Exit Function
        End If

        StartResponse = ""
        'fetch the max number of concurrent IPLs
        MaxIPLs = Common.Extract_Val(Common.Fetch_CPE_SystemOption(45))
        'fetch the IPL Window time
        WindowMinutes = Common.Extract_Val(Common.Fetch_CPE_SystemOption(46))
        'fetch the IPL Runaway time
        IPLRunawayTime = Common.Extract_Val(Common.Fetch_CPE_SystemOption(47))

        Common.QueryStr = "dbo.pa_IPLOffers_StartOK"
        Common.Open_LRTsp()
        Common.LRTsp.CommandTimeout = 300
        Common.LRTsp.Parameters.Add("@WindowMinutes", SqlDbType.Int).Value = WindowMinutes
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.Parameters.Add("@NumIPLs", SqlDbType.Int).Direction = ParameterDirection.Output
        Common.LRTsp.Parameters.Add("@LastIPLAge", SqlDbType.BigInt).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        NumIPLs = Common.LRTsp.Parameters("@NumIPLs").Value
        LastIPLAge = Common.LRTsp.Parameters("@LastIPLAge").Value
        Common.Close_LRTsp()
        If (MaxIPLs > 0) And (WindowMinutes > 0) And (NumIPLs >= MaxIPLs) Then
            StartResponse = "IPL Throttle Exceeded"
        End If
        If Not (LastIPLAge = -1) And (LastIPLAge < IPLRunawayTime) Then
            StartResponse = "IPL Runaway"
        End If

        Return (StartResponse)

    End Function

</script>

<%
  Dim LocalServerID As Long
  Dim LocationID As Long
  Dim LastHeard As String
  Dim Mode As String
  Dim HistoryStartTime As DateTime
  Dim IPLTypeID As Integer
  Dim IPLStartResponse As String
  Dim IPLSessionID As String
  Dim BannerID As Integer
  Dim LSVerParts() As String
  
  IPLTypeID = 1
  ApplicationName = "IPL-Offers"
  ApplicationExtension = ".aspx"
  Common.AppName = ApplicationName & ApplicationExtension
  Response.Expires = 0
  On Error GoTo ErrorTrap
  
  StartTime = DateAndTime.Timer
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "CPE-IncentiveFetchLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
  LocalServerIP = Trim(Request.QueryString("IP"))
  MacAddress = Trim(Request.QueryString("mac"))
  If MacAddress = "" Or MacAddress = "0" Then
    MacAddress = "0"
  End If
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
    LocalServerIP = Trim(Request.UserHostAddress)
  End If  
  
  LSVersion = Trim(Request.QueryString("lsversion"))
  LSVerMajor = 0
  LSVerMinor = 0
  If InStr(LSVersion, ".", CompareMethod.Binary) > 0 Then
    LSVerParts = Split(LSVersion, ".", , CompareMethod.Binary)
    LSVerMajor = Common.Extract_Val(LSVerParts(0))
    LSVerMinor = Common.Extract_Val(LSVerParts(1))
  End If
  LSBuild = Trim(Request.QueryString("lsbuild"))
  LSBuildMajor = 0
  LSBuildMinor = 0
  If InStr(LSBuild, ".", CompareMethod.Binary) > 0 Then
    LSVerParts = Split(LSBuild, ".", , CompareMethod.Binary)
    LSBuildMajor = Common.Extract_Val(LSVerParts(0))
    LSBuildMinor = Common.Extract_Val(LSVerParts(1))
  End If
  
  LastHeard = "1/1/1980"
  
  IPLSessionID = "0"
  If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 4) Or (LSVerMajor = 5 And LSVerMinor = 4 And LSBuildMajor >= 6) Then
    IPLSessionID = IIf(Request.QueryString("sessionid") <> "", Request.QueryString("sessionid"), "0")
  End If
  
  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Open_LogixWH()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  
  If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    Common.Open_PrefManRT()
  End If
  
  BufferedRecs = 0
  
  Common.Write_Log(LogFile, "---------------------------------------------------------------------------")
  
  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, LocalServerIP)
  
  If LocationID = "0" Then
    Common.Write_Log(LogFile, Common.AppName & "    Invalid Serial Number:" & LocalServerID & " from  MacAddress: " & MacAddress & " IP:" & LocalServerIP & "  Process running on server:" & Environment.MachineName, True)
    Send_Response_Header(ApplicationName & " - Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 2)) Then
    'the location calling IPL-Offers is not associated with the CPE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than CPE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other than CPE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  Else
    Mode = UCase(Request.QueryString("mode"))
    Common.Write_Log(LogFile, "** " & Common.AppName & "   " & Microsoft.VisualBasic.DateAndTime.Now & "  CSVersion: " & Connector.CSMajorVersion & "." & Connector.CSMinorVersion & "b" & Connector.CSBuild & "r" & Connector.CSBuildRevision & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  Process running on server:" & Environment.MachineName)
    Common.Write_Log(LogFile, "SessionID=" & IPLSessionID)
    If Mode = "ACK" Then
      Process_ACK(LocalServerID, LocationID)
    ElseIf Mode = "NAK" Then
      Process_NAK(LocalServerID, LocationID)
    ElseIf Mode = "THROTTLE" Then
      IPLStartResponse = IPLStartOk(LocalServerID)
      If IPLStartResponse = "" Then
        Common.Write_Log(LogFile, "Throttle check response = OK")
        Send_Response_Header("OK", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Else
        Common.Write_Log(LogFile, "Throttle check response = " & IPLStartResponse)
        Send_Response_Header(IPLStartResponse, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      End If
    ElseIf (Mode = "IPL") Then
      'check to see if it is ok to start an IPL now
      IPLStartResponse = IPLStartOk(LocalServerID)
      If IPLStartResponse = "" Then
        
        'record the starting of the IPL
        Common.QueryStr = "dbo.pa_IPL_UpdateStartTime"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
        Common.QueryStr = "dbo.pa_IPL_HistoryStart"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
        Common.LRTsp.Parameters.Add("@IPLTypeID", SqlDbType.Int).Value = IPLTypeID
        Common.LRTsp.Parameters.Add("@StartTable", SqlDbType.VarChar, 200).Value = ""
        Common.LRTsp.Parameters.Add("@StartPK", SqlDbType.BigInt).Value = 0
        Common.LRTsp.Parameters.Add("@SessionID", SqlDbType.VarChar, 20).Value = IPLSessionID
        Common.LRTsp.Parameters.Add("@StartTime", SqlDbType.DateTime).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        HistoryStartTime = Common.LRTsp.Parameters("@StartTime").Value
        Common.Close_LRTsp()
        
        'since we are doing an IPL, get rid of any IncentiveFiles that might be waiting to go down to the Local Server
        Common.QueryStr = "update CPE_IncentiveDLBuffer with (RowLock) set WaitingACK=2 where WaitingACK>=0 and WaitingACK<2 and LocalServerID=" & LocalServerID & ";"
        Common.LRT_Execute()
        'this local server no longer needs to IPL because it's performing one right now
        Common.QueryStr = "Update LocalServers with (RowLock) set MustIPL=0 where LocalServerID='" & LocalServerID & "';"
        Common.LRT_Execute()
        'TextData = ""
        
        'since we are doing an IPL, get rid of any locks that may exist for this locationid
        Common.QueryStr = "delete from CustomerLock where LocationID=" & LocationID & ";"
        Common.LXS_Execute()
        
        Common.Write_Log(LogFile, "Removing buffered TransDownload data")
        Common.QueryStr = "dbo.pc_CPE_Gen_Purge_Output_byLoc"
        Common.Open_LXSsp()
        Common.LXSsp.CommandTimeout = 1200
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        Common.LXSsp.ExecuteNonQuery()
        Common.Close_LXSsp()
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished removing buffered TransDownload data.  Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        
        gzStream = New GZipStream(Response.OutputStream, CompressionMode.Compress, True)
        
        Construct_Output(LocalServerID, LocationID)
        
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished all queries. Closing GZip Stream ... " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        If gzStream IsNot Nothing Then
          gzStream.Close()
          gzStream.Dispose()
          gzStream = Nothing
        End If
        
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "GZip stream closed.  Flushing final records ... " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        
        FlushStartTime = DateAndTime.Timer
        Response.Flush()
        FlushTime = FlushTime + (DateAndTime.Timer - FlushStartTime)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Total time to send all records to client=" & Int(FlushTime) & Format$(FlushTime - Fix(FlushTime), ".000") & "(sec)" & "  Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        
        Common.Write_Log(LogFile, "Total uncompressed size = " & UncompressedSize)
        
        'update the history record for the IPL end time
        Common.QueryStr = "dbo.pa_IPL_HistoryEnd"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
        Common.LRTsp.Parameters.Add("@IPLTypeID", SqlDbType.Int).Value = IPLTypeID
        Common.LRTsp.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = HistoryStartTime
        Common.LRTsp.Parameters.Add("@UncompressedSize", SqlDbType.BigInt).Value = UncompressedSize
        Common.LRTsp.Parameters.Add("@CompressedSize", SqlDbType.BigInt).Value = 0
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
      Else
        Common.Write_Log(LogFile, "Throttle check response = " & IPLStartResponse)
        Send_Response_Header(IPLStartResponse, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      End If
    Else
      Send_Response_Header("Invalid Request", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    End If
  End If 'locationid="0"
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Closing database connections - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
  
  Common.Close_LogixRT()
  Common.Close_LogixXS()
  Common.Close_LogixWH()
  If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    Common.Close_PrefManRT()
  End If
  
  TotalTime = DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "RunTime=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
%>
<% 
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
  If Not (Common.LWHadoConn.State = ConnectionState.Closed) Then Common.Close_LogixWH()
  If Not (Common.PMRTadoConn.State = ConnectionState.Closed) Then Common.Close_PrefManRT()
  Common = Nothing
%>
