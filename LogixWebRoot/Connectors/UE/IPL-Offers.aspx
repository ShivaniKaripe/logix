<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO.Compression" %>
<%@ Import Namespace="CMS.AMS.Models" %>
<%@ Import Namespace="CMS.AMS.Contract" %>
<%@ Import Namespace="CMS.AMS" %>
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
  Public GZIP As New Copient.GZIPInc
  Public TextData As String
  Public IPL As Boolean
  Public LogFile As String
  Public FileStamp As String
  Public FileNum As Integer
  Public StartTime As Decimal
  Public TotalTime As Decimal
  Public ApplicationName As String
  Public ApplicationExtension As String
  Public MacAddress As String
  Public LocalServerIP As String
  Public LSVerMajor As Integer
  Public LSVerMinor As Integer
  Public LSBuildMajor As Integer
  Public LSBuildMinor As Integer
  Public GZStream As GZipStream = Nothing
  Public UncompressedSize As Long = 0
  Public BufferedRecs As Long = 0
  Public FlushTime As Decimal = 0
  Public FlushStartTime As Decimal = 0
  Public bAllowDollarTransLimit As Boolean = False
  Public bUsePromotionDisplay As Boolean = False
  Public bUseProrateonDisplay As Boolean = False

  ' -------------------------------------------------------------------------------------------------

  Sub SD(ByVal OutStr As String)
    'PrintLine(FileNum, OutStr)
    Dim Bytes As Byte()
    Bytes = Encoding.UTF8.GetBytes(OutStr & vbCrLf)
    UncompressedSize = UncompressedSize + Bytes.Length
    GZStream.Write(Bytes, 0, Bytes.Length)
    Bytes = Nothing
    BufferedRecs = BufferedRecs + 1
    If BufferedRecs >= 5000 Then
      FlushStartTime = Microsoft.VisualBasic.DateAndTime.Timer
      Response.Flush()
      FlushTime = FlushTime + (Microsoft.VisualBasic.DateAndTime.Timer - FlushStartTime)
      BufferedRecs = 0
    End If
  End Sub

  ' -------------------------------------------------------------------------------------------------

  Sub SDb(ByVal OutStr As String)
    'Print(FileNum, OutStr)
    Dim Bytes As Byte()
    Bytes = Encoding.UTF8.GetBytes(OutStr)
    UncompressedSize = UncompressedSize + Bytes.Length
    GZStream.Write(Bytes, 0, Bytes.Length)
    Bytes = Nothing
    BufferedRecs = BufferedRecs + 1
    If BufferedRecs >= 5000 Then
      FlushStartTime = Microsoft.VisualBasic.DateAndTime.Timer
      Response.Flush()
      BufferedRecs = 0
      FlushTime = FlushTime + (Microsoft.VisualBasic.DateAndTime.Timer - FlushStartTime)
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

    QueryStartTime = Microsoft.VisualBasic.DateAndTime.Timer
    If UCase(DBName) = "LXS" Then
      dst = Common.LXS_Select
    ElseIf UCase(DBName) = "LWH" Then
      dst = Common.LWH_Select
    ElseIf UCase(DBName) = "PMRT" Then
      dst = Common.PMRT_Select
    Else
      dst = Common.LRT_Select
    End If
    QueryTotalTime = Microsoft.VisualBasic.DateAndTime.Timer - QueryStartTime
    ConstructStartTime = Microsoft.VisualBasic.DateAndTime.Timer

    bUsePromotionDisplay = IIf(Common.Fetch_UE_SystemOption(145) = "1", True, False)
    bUseProrateonDisplay = IIf(Common.Fetch_UE_SystemOption(154) = "1", True, False)
    
    If dst.Rows.Count > 0 Then
      If UCase(TableName) = "INCENTIVES" And Operation = 1 Then
        FieldList = "IncentiveID" & DelimCH & "IncentiveName" & DelimCH & "Priority" & DelimCH & "StartDate" & DelimCH & "EndDate" & DelimCH & "TestingStartDate" & DelimCH & "TestingEndDate" & DelimCH & "EveryDOW" & DelimCH & "EligibilityStartDate" & DelimCH & "EligibilityEndDate" & DelimCH & "P1DistQtyLimit" & DelimCH & "P1DistPeriod" & DelimCH & "P1DistTimeType" & DelimCH & _
                    "P2DistQtyLimit" & DelimCH & "P2DistPeriod" & DelimCH & "P2DistTimeType" & DelimCH & "P3DistQtyLimit" & DelimCH & "P3DistPeriod" & DelimCH & "P3DistTimeType" & DelimCH & "Reporting" & DelimCH & "EmployeesOnly" & DelimCH & "EmployeesExcluded" & DelimCH & "UpdateLevel" & DelimCH & "DeferCalcToEOS" & DelimCH & "EveryTOD" & DelimCH & "ChargebackVendorID" & DelimCH & _
                    "SendIssuance" & DelimCH & "ManufacturerCoupon" & DelimCH & "InboundCRMEngineID" & DelimCH & "EnableImpressRpt" & DelimCH & "ClientOfferID" & DelimCH & "VendorCouponCode" & DelimCH & "EngineID" & DelimCH & "MutuallyExclusive" & DelimCH & "EngineSubTypeID" & DelimCH & "PromoClassID" & DelimCH & "DiscountEvalTypeID" & _
                    IIF(Common.Fetch_UE_SystemOption(180),	DelimCH & "PreOrderEligibility", "") & _
      IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps2297), IIF(Common.Fetch_UE_SystemOption(211), DelimCH & "DeferCalcToTotal", "") & DelimCH & "pointsprogramwatch", "")
      If bUsePromotionDisplay Then FieldList &= DelimCH & "PromotionDisplay"
      If bUseProrateonDisplay Then FieldList &= DelimCH & "ProrateonDisplay"
      FieldList &= DelimCH & "StoreCoupon"
	  FieldList &= DelimCH & "PosNotificationCheck"
      ElseIf UCase(TableName) = "INCENTIVETERMINALS" And Operation = 5 Then
        FieldList = "PKID" & DelimCH & "IncentiveID" & DelimCH & "TerminalTypeID" & DelimCH & "Excluded"
      ElseIf UCase(TableName) = "REWARDOPTIONS" And (Operation = 1 Or Operation = 5) Then
        FieldList = "RewardOptionID" & DelimCH & "Name" & DelimCH & "IncentiveID" & DelimCH & "Priority" & DelimCH & "HHEnable" & DelimCH & "TouchResponse" & DelimCH & "ProductComboID" & DelimCH & "ExcludedTender" & DelimCH & "ExcludedTenderAmtRequired" & DelimCH & "TierLevels" & DelimCH & "AttributeComboID" & DelimCH & "OfflineCustCondition" & DelimCH & "TenderComboID" & DelimCH & "PointsComboID" & DelimCH & "StoredValueComboID" & DelimCH & "PreferenceComboID" & DelimCH & "CurrencyID"
      ElseIf UCase(TableName) = "DELIVERABLES" And Operation = 5 Then
        FieldList = "DeliverableID" & DelimCH & "RewardOptionID" & DelimCH & "RewardOptionPhase" & DelimCH & "DeliverableTypeID" & DelimCH & "OutputID" & DelimCH & "AvailabilityTypeID" & DelimCH & "Priority" & DelimCH & "ScreenCellID" & DelimCH & "Required"
      ElseIf UCase(TableName) = "EDISCOUNTS_TIERS" And Operation = 1 Then
        FieldList = "PKID" & DelimCH & "EDiscountID" & DelimCH & "TierLevel" & DelimCH & "ReceiptDescription" & DelimCH & "DiscountAmount" & DelimCH & "ItemLimit" & DelimCH & "WeightLimit" & DelimCH & "DollarLimit" & DelimCH & "SPRepeatLevel" & DelimCH & "BuyDescription"
        If bAllowDollarTransLimit Then
          FieldList = FieldList & DelimCH & "RewardLimitTypeID"
        End If
      ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS" And Operation = 5 Then
        FieldList = "IncentiveProductGroupID" & DelimCH & "RewardOptionID" & DelimCH & "ProductGroupID" & DelimCH & "QtyForIncentive" & DelimCH & "QtyUnitType" & DelimCH & "ExcludedProducts" & DelimCH & "Disqualifier" & DelimCH & "UniqueProduct" & DelimCH & "Rounding" & DelimCH & _
                    "MinPurchAmt" & IIF(Common.Fetch_UE_SystemOption(210), DelimCH & "NetPriceProduct", "") & IIF(Common.Fetch_UE_SystemOption(182), DelimCH & "ReturnedItemGroup", "") & DelimCH & "MinItemPrice" & DelimCH & "FullPrice" & DelimCH & "ClearanceState" & DelimCH & "ClearanceLevel" & DelimCH & "TenderType " & DelimCH & "SameItem"
            ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS_TIERS" And Operation = 5 Then
                FieldList = "PKID" & DelimCH & "IncentiveProductGroupID" & DelimCH & "TierLevel" & DelimCH & "Quantity"
            ElseIf UCase(TableName) = "INCENTIVEUSERGROUPS" And Operation = 5 Then
                FieldList = "IncentiveUserID" & DelimCH & "RewardOptionID" & DelimCH & "UserGroupID" & DelimCH & "ExcludedUsers"
            ElseIf UCase(TableName) = "VENDORS" And Operation = 5 Then
                FieldList = "VendorID" & DelimCH & "ExtVendorID"
            ElseIf UCase(TableName) = "EIWTRIGGERS" And Operation = 5 Then
                FieldList = "TriggerID" & DelimCH & "IncentiveID" & DelimCH & "RewardOptionID" & DelimCH & "IncentiveEIWID" & DelimCH & "TriggerTime" & DelimCH & "Consumed"
            ElseIf UCase(TableName) = "DELIVERABLEPASSTHRUTIERS" And Operation = 5 Then
                FieldList = "PKID" & DelimCH & "PTPKID" & DelimCH & "TierLevel" & DelimCH & "Data" & DelimCH & "Value" & DelimCH & "LanguageID"
            ElseIf UCase(TableName) = "DELIVERABLEMONSTOREDVALUEREDEMPTION" And Operation = 5 Then
                FieldList = "PKID" & DelimCH & "DeliverableID" & DelimCH & "SVProgramID" & DelimCH & "RewardOptionID" & DelimCH & "Description"
            ElseIf UCase(TableName) = "MONSVREDEMPTIONTRANSLATIONS" And Operation = 5 Then
                FieldList = "PKID" & DelimCH & "SVProgramID" & DelimCH & "LanguageID" & DelimCH & "Description"
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

                If UCase(TableName) = "PRINTEDMESSAGES" And Operation = 5 Then '-- need to encode PrintedMessages because it allows \n
                    For Each row In dst.Rows
                        TextMsg = Common.NZ(row.Item("TextMsg"), " ")
                        SD(Common.NZ(row.Item("MessageID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("PrintZone"), 1) & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("TextMsg"), " ")) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("SuppressZeroBalance"), 0)) & Chr(DelimChar) & Common.NZ(row.Item("SortID"), 0))
                        'Common.Write_Log(LogFile, Common.NZ(row.Item("MessageID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("Name"), 0) & Chr(DelimChar) & Replace(TextMsg, vbCrLf, "|", , , vbBinaryCompare))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "PRINTEDMESSAGES_TIERS" And Operation = 5 Then '-- need to encode PrintedMessages because it allows \n
                    For Each row In dst.Rows
                        BodyText = Common.NZ(row.Item("BodyText"), " ")
                        SD(Common.NZ(row.Item("PKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("MessageID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("TierLevel"), 0) & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("BodyText"), " ")) & Chr(DelimChar) & Common.NZ(row.Item("Value"), 0))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "PMTRANSLATIONS" And Operation = 5 Then '-- need to encode PMTranslations.BodyText because it allows \r\n
                    For Each row In dst.Rows
                        SD(row.Item("PKID") & Chr(DelimChar) & row.Item("PMTiersID") & Chr(DelimChar) & row.Item("LanguageID") & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("BodyText"), " ")))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "PREDEFINEDTRIGGERCODEMESSAGES" And Operation = 5 Then '-- need to encode trigger code messages because it allows \n
                    For Each row In dst.Rows
                        SD(Common.NZ(row.Item("ReasonFlag"), -1) & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("Description"), " ")) & Chr(DelimChar) & Common.NZ(row.Item("languageID"), 0))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "PROMOGRIDOFFERS" And Operation = 5 Then '-- need to encode trigger code messages because it allows \n
                    For Each row In dst.Rows
                        SD(Common.NZ(row.Item("IncentiveID"), -1) & Chr(DelimChar) & Common.NZ(row.Item("promocategoryid"), " "))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
		
                ElseIf UCase(TableName) = "INCENTIVES" And Operation = 1 Then
                    For Each row In dst.Rows
                        SD(row.Item("IncentiveID") & Chr(DelimChar) & row.Item("IncentiveName") & Chr(DelimChar) & row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "INCENTIVETERMINALS" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "REWARDOPTIONS" And (Operation = 1 Or Operation = 5) Then
                    For Each row In dst.Rows
                        SD(row.Item("RewardOptionID") & Chr(DelimChar) & row.Item("Name") & Chr(DelimChar) & row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "DELIVERABLES" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "EDISCOUNTS_TIERS" And Operation = 1 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "INCENTIVEPRODUCTGROUPS_TIERS" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "INCENTIVEUSERGROUPS" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "VENDORS" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "EIWTRIGGERS" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(row.Item("Data"))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "DELIVERABLEPASSTHRUTIERS" And Operation = 5 Then 'Add Data Encoding - AL-4811
                    For Each row In dst.Rows
                        If (Common.NZ(row.Item("PassThruRewardID"), 0) = 11 OrElse Common.NZ(row.Item("PassThruRewardID"), 0) = 12 OrElse Common.NZ(row.Item("PassThruRewardID"), 0) = 13) Then
                            SD(Common.NZ(row.Item("PKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("PTPKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("TierLevel"), 0) & Chr(DelimChar) & Common.Comms_RemoveCRLF(Common.NZ(row.Item("Data"), " ")) & Chr(DelimChar) & Common.NZ(row.Item("Value"), 0) & Chr(DelimChar) & row.Item("LanguageID"))
                        Else
                            SD(Common.NZ(row.Item("PKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("PTPKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("TierLevel"), 0) & Chr(DelimChar) & Common.Comms_Encode(Common.NZ(row.Item("Data"), " ")) & Chr(DelimChar) & Common.NZ(row.Item("Value"), 0) & Chr(DelimChar) & row.Item("LanguageID"))
                        End If
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "EDISCOUNTS" And Operation = 1 Then
                    For Each row In dst.Rows
                        SD(Common.NZ(row.Item("EdiscountID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("Name"), 0) & Chr(DelimChar) & Common.NZ(row.Item("DiscountTypeID"), 0) & Chr(DelimChar) & _
                            Common.NZ(row.Item("ReceiptDescription"), 0) & Chr(DelimChar) & Common.NZ(row.Item("DiscountedProductGroupID"), 0) & Chr(DelimChar) & _
                            Common.NZ(row.Item("ExcludedProductGroupID"), 0) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("BestDeal"), 0)) & Chr(DelimChar) & _
                            Parse_Bit(Common.NZ(row.Item("AllowNegative"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("ComputeDiscount"), 0)) & Chr(DelimChar) & _
                            Math.Round(Common.NZ(row.Item("DiscountAmount"), 0), 3) & Chr(DelimChar) & Common.NZ(row.Item("AmountTypeID"), 0) & Chr(DelimChar) & _
                            Math.Round(Common.NZ(row.Item("L1Cap"), 0), 2) & Chr(DelimChar) & Math.Round(Common.NZ(row.Item("L2DiscountAmt"), 0), 0) & Chr(DelimChar) & _
                            Common.NZ(row.Item("L2AmountTypeID"), 0) & Chr(DelimChar) & Math.Round(Common.NZ(row.Item("L2Cap"), 0), 2) & Chr(DelimChar) & _
                            Math.Round(Common.NZ(row.Item("L3DiscountAmt"), 0), 2) & Chr(DelimChar) & Common.NZ(row.Item("L3AmountTypeID"), 0) & Chr(DelimChar) & _
                            Common.NZ(row("ItemLimit"), 1) & Chr(DelimChar) & Common.NZ(row.Item("WeightLimit"), 0) & Chr(DelimChar) & _
                            Math.Round(Common.NZ(row.Item("DollarLimit"), 0), 2) & Chr(DelimChar) & Common.NZ(row.Item("ChargeBackDeptID"), 0) & Chr(DelimChar) & _
                            Parse_Bit(Common.NZ(row.Item("DecliningBalance"), 0)) & Chr(DelimChar) & Common.NZ(row.Item("SVProgramID"), 0) & Chr(DelimChar) & _
                            Parse_Bit(Common.NZ(row.Item("FlexNegative"), 0)) & Chr(DelimChar) & Common.NZ(row.Item("ScorecardID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("ScorecardDesc"), 0) & Chr(DelimChar) & _
                            row.Item("AllowMarkup") & Chr(DelimChar) & row.Item("DiscountAtOrigPrice") & Chr(DelimChar) & row.Item("ProrationTypeID") & Chr(DelimChar) & Common.NZ(row.Item("PriceFilter"), 100) & Chr(DelimChar) & Common.NZ(row.Item("FlexOptions"), 0) & Chr(DelimChar) & _
                            Parse_Bit(Common.NZ(row.Item("GrossPrice"), 0)))
                        'Common.Write_Log(LogFile, Common.NZ(row.Item("EdiscountID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("Name"), 0) & Chr(DelimChar) & Common.NZ(row.Item("ReceiptDescription"), 0) & Chr(DelimChar) & Common.NZ(row.Item("DiscountedProductGroupID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("ExcludedProductGroupID"), 0) & Chr(DelimChar) & Math.Round(Common.NZ(row.Item("DiscountAmount"), 0), 3) & Chr(DelimChar) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("DiscountTypeID"), 0)) & Common.NZ(row.Item("AmountTypeID"), 0) & Chr(DelimChar) & Math.Round(row.Item("L1Cap"), 2) & Chr(DelimChar) & Math.Round(row.Item("L2DiscountAmt"), 0) & Chr(DelimChar) & Common.NZ(row.Item("L2AmountTypeID"), 0) & Chr(DelimChar) & Math.Round(row.Item("L2Cap"), 2) & Chr(DelimChar) & Math.Round(row.Item("L3DiscountAmt"), 2) & Chr(DelimChar) & Common.NZ(row.Item("L3AmountTypeID"), 0) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("DecliningBalance"), 0)) & Chr(DelimChar) & Common.NZ(row.Item("ChargeBackDeptID"), 0) & Chr(DelimChar) & _
                        'Common.NZ(row("ItemLimit"), 1) & Chr(DelimChar) & Common.NZ(row.Item("WeightLimit"), 0) & Chr(DelimChar) & Math.Round(row.Item("DollarLimit"), 2) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("BestDeal"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("AllowNegative"), 0)) & Chr(DelimChar) & Parse_Bit(Common.NZ(row.Item("ComputeDiscount"), 0)))
                        NumRecs = NumRecs + 1
                    Next
                ElseIf UCase(TableName) = "OFFERTRANSLATIONS" And Operation = 5 Then
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
                            ElseIf Index = 2 Then 'Offer Name
                                TempOut = TempOut & Common.NZ(row(Index), "")
                                SDb(Common.NZ(row(Index), ""))
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
                ElseIf UCase(TableName) = "DELIVERABLEMONSTOREDVALUEREDEMPTION" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(Common.NZ(row.Item("PKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("DeliverableID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("SVProgramID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("RewardOptionID"), 0) & Chr(DelimChar) & Common.Comms_RemoveCRLF(Common.NZ(row.Item("Description"), " ")))
                        NumRecs = NumRecs + 1
                    Next
                    DataBack = "Sent Data"
                ElseIf UCase(TableName) = "MONSVREDEMPTIONTRANSLATIONS" And Operation = 5 Then
                    For Each row In dst.Rows
                        SD(Common.NZ(row.Item("PKID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("SVProgramID"), 0) & Chr(DelimChar) & Common.NZ(row.Item("LanguageID"), 0) & Chr(DelimChar) & Common.Comms_RemoveCRLF(Common.NZ(row.Item("Description"), " ")))
                        NumRecs = NumRecs + 1
                    Next
                DataBack = "Sent Data"
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
                SD("###")
                Common.Write_Log(LogFile, "# records: " & NumRecs)
                ConstructTotalTime = Microsoft.VisualBasic.DateAndTime.Timer - ConstructStartTime
                TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
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
        Dim TempQuery As String
        Dim TestingLocation As Boolean
        Dim DelimChar As Integer
        'Dim ActiveROIDs As String
        'Dim ActiveIncentiveIDs As String
        Dim row As DataRow
        Dim MissingIncentives As String
        Dim OperateAtEnterprise As Boolean
        Dim ROIDs As String
        Dim m_giftcardreward As IGiftCardRewardService
        Dim m_couponreward As ICouponRewardService
        Dim m_DiscountService As IDiscountRewardService
        Dim sbExtRedemptionAuth As New StringBuilder
        Dim m_ProdCond As IProductConditionService = CurrentRequest.Resolver.Resolve(Of IProductConditionService)()
        m_giftcardreward = CurrentRequest.Resolver.Resolve(Of IGiftCardRewardService)()
        m_couponreward = CurrentRequest.Resolver.Resolve(Of ICouponRewardService)()
        Dim m_PMReward As IProximityMessageRewardService
        m_PMReward = CurrentRequest.Resolver.Resolve(Of IProximityMessageRewardService)()
        m_DiscountService = CurrentRequest.Resolver.Resolve(Of IDiscountRewardService)()
        Dim m_CustCondService As ICustomerGroupCondition
        m_CustCondService = CurrentRequest.Resolver.Resolve(Of ICustomerGroupCondition)()
        DelimChar = 30
        TempOut = ""

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

        MissingIncentives = ""
        'Common.QueryStr = "select IncentiveID from CPE_IncentiveLoc_Func(" & LocationID & ");"

        Common.QueryStr = "dbo.pa_UE_CheckEnterprise"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
        Common.LRTsp.Parameters.Add("@Enterprise", SqlDbType.Bit).Direction = ParameterDirection.Output
        Common.LRTsp.ExecuteNonQuery()
        OperateAtEnterprise = Common.NZ(Common.LRTsp.Parameters("@Enterprise").Value, False)
        Common.Close_LRTsp()

        If OperateAtEnterprise Then
            If Common.Fetch_UE_SystemOption(80) = 0 Then  'lock offers after expiration is turned off
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "  from UE_OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  " inner join UE_IncentiveLocationsView_Enterprise as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID " & _
                                  "  left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                          " where(STI.EngineID=9) " & _
                                  " order by OfferID;"
            Else  'lock offers after expiration is turned on
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "from UE_OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  "inner join UE_IncentiveLocationsView_Enterprise as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID " & _
                                  "left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                          "where(STI.EngineID=9) and (dateadd(d, 1, STI.EndDate)>getdate() or dateadd(d, 1, STI.EligibilityEndDate)>getdate() or  dateadd(d, 1, STI.TestingEndDate)>getdate()) " & _
                                  "order by OfferID;"
            End If
        Else
            If Common.Fetch_UE_SystemOption(80) = 0 Then   'lock offers after expiration is turned off
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "  from UE_OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  " inner join UE_IncentiveLocationsView as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID and ILV.LocationID=" & LocationID & " " & _
                                  "  left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                  " where(OLU.LocationID=" & LocationID & " and STI.EngineID=9) " & _
                                  " order by OfferID;"
            Else   'lock offers after expiration is turned on
                Common.QueryStr = "Insert into #ActiveIncentives (IncentiveID, STIncentiveID) " & _
                                  "select distinct OLU.OfferID as IncentiveID, isnull(STI.IncentiveID, -1) as STIncentiveID " & _
                                  "from UE_OfferLocUpdate as OLU with (NoLock) Inner Join CPE_Incentives as I with (NoLock) on I.IncentiveID=OLU.OfferID and I.Deleted=0 " & _
                                  "inner join UE_IncentiveLocationsView as ILV with (NoLock) on I.IncentiveID=ILV.IncentiveID and ILV.LocationID=" & LocationID & " " & _
                                  "left join CPE_ST_Incentives as STI with (NoLock) on STI.IncentiveID=I.IncentiveID and STI.Deleted=0 " & _
                                  "where(OLU.LocationID=" & LocationID & " and STI.EngineID=9) and (dateadd(d, 1, STI.EndDate)>getdate() or dateadd(d, 1, STI.EligibilityEndDate)>getdate() or  dateadd(d, 1, STI.TestingEndDate)>getdate()) " & _
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
            Common.Write_Log(LogFile, "Serial= " & LocalServerID & " IPAddress=" & Trim(Request.UserHostAddress) & " The following offers can not be deployed during an IPL of LocationID (" & LocationID & ") on because they are missing from the shadow tables.  These offers will have to be re-deployed: " & MissingIncentives & " Serial= " & LocalServerID & "; Mac IPAddress=" & (Trim(Request.UserHostAddress)) & " Server = " & Environment.MachineName, True)
            Common.LastErrorTime = "1/1/1980 00:00:00"
            Common.Error_Processor("Serial=" & LocalServerID & " IPAddress=" & Trim(Request.UserHostAddress) & " The following offers can not be deployed during an IPL of LocationID (" & LocationID & ") because they are missing from the shadow tables.  These offers will have to be re-deployed: " & MissingIncentives, "The following offers can not be deployed during an IPL of LocationID (" & LocationID & ") because they are missing from the shadow tables.  These offers will have to be re-deployed: " & MissingIncentives, , , 0)
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
        If transNumSeedResult.Rows.Count <> 1 Then
            Throw New Exception("Found unexpected condition")
        End If
        Dim TransNumSeed As Integer = transNumSeedResult.Rows(0).Item("LogixTransNumSeed")
        Common.QueryStr = "SELECT IPLSequenceNum FROM [dbo].[LocationSeqNum] WHERE LocationID=" & LocationID
        Dim IPLSequenceNumResult As DataTable = Common.LXS_Select()
        Dim IPLSequenceNum As Integer = 1
        If IPLSequenceNumResult.Rows.Count <> 1 Then
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
        Construct_Table("LocalServers", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'Locations
        'send the data from the Locations table
        If Not (OperateAtEnterprise) Then
            'send only the one row for the location performing the IPL
            Common.QueryStr = "select LocationID, LocationName, Address1, Address2, City, State, Zip, ExtLocationCode as ClientLocationCode, TestingLocation, LocationTypeID, CurrencyID from Locations with (NoLock) where Deleted=0 and EngineID=9 and LocationID=" & LocationID & ";"
            Construct_Table("Locations", 5, DelimChar, LocalServerID, LocationID, "LRT")
            Common.QueryStr = "select PKID, LocationID, LanguageID, Required from LocationLanguages with (NoLock) where Deleted=0 and LocationID=" & LocationID & ";"
            Construct_Table("LocationLanguages", 5, DelimChar, LocalServerID, LocationID, "LRT")
        Else
            'send all of the rows from the Locations table
            Common.QueryStr = "select LocationID, LocationName, Address1, Address2, City, State, Zip, ExtLocationCode as ClientLocationCode, TestingLocation, LocationTypeID, TimeZone, CurrencyID from Locations with (NoLock) where Deleted=0 and EngineID=9;"
            Construct_Table("Locations", 5, DelimChar, LocalServerID, LocationID, "LRT")
            Common.QueryStr = "select PKID, LocationID, LanguageID, Required from LocationLanguages with (NoLock) where Deleted=0;"
            Construct_Table("LocationLanguages", 5, DelimChar, LocalServerID, LocationID, "LRT")
        End If

        'LocalServer_Seeds
        'send the data to indicate the LocalID seed numbers for RewardAccumulation, RewardDistribution, and StoredValue
        Common.QueryStr = "exec pa_CPE_LocalID_Seeds " & LocalServerID
        Construct_Table("LocalIDSeeds", 5, DelimChar, LocalServerID, LocationID, "LXS")

        'PromoEngines
        Common.QueryStr = "select EngineID, Description, DefaultEngine, Installed from PromoEngines with (NoLock)"
        Construct_Table("PromoEngines", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'PromoEngineSubTypes
        Common.QueryStr = "select PKID, PromoEngineID, SubTypeID, SubTypeName, Installed, ReplayEnabled from PromoEngineSubTypes with (NoLock)"
        Construct_Table("PromoEngineSubTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'ExtCRMInterfaces
        Common.QueryStr = "select ExtInterfaceID, Name, Description from ExtCRMInterfaces where Deleted=0;"
        Construct_Table("ExtCRMInterfaces", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'Integrations
        Common.QueryStr = "select IntegrationID, Name, Installed from Integrations;"
        OutStr = OutStr & Construct_Table("Integrations", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'TerminalSetItems
        Common.QueryStr = "select TSI.PKID, isnull(TSI.TerminalSetID, 0) as TerminalSetID, isnull(TSI.TerminalID, 0) as TerminalID, isnull(TSI.TerminalTypeID, 0) as TerminalTypeID, isnull(TSI.PrinterTypeID, 0) as PrinterTypeID, isnull(TSI.OpDisplayTypeID, 0) as OpDisplayTypeID " & _
                          "from TerminalSetItems as TSI with (NoLock);"
        Construct_Table("TerminalSetItems", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'LocationTerminals
        If Not (OperateAtEnterprise) Then
            Common.QueryStr = "select PKID, LT.LocationID, LT.TerminalSetID " & _
                              "from LocationTerminals as LT with (NoLock) " & _
                              "where LT.LocationID=" & LocationID & ";"
        Else
            Common.QueryStr = "select PKID, LT.LocationID, LT.TerminalSetID " & _
                              "from LocationTerminals as LT with (NoLock);"
        End If
        Construct_Table("LocationTerminals", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'OpDisplaytypes
        'send the date from the OpDisplayTypes Table
        Common.QueryStr = "select OpDisplayTypeID, Name from CPE_OpDisplayTypes ODT with (NoLock);"
        Construct_Table("OpDisplayTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'Printtypes
        Common.QueryStr = "Select PrintTypeID,TypeDescription from PrintTypes with (NoLock);"
        Construct_Table("PrintTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'PrintSubTypes
        Common.QueryStr = "Select PrintSubTypeID,TypeDescription,PrintTypeID from printsubtypes with (NoLock); "
        Construct_Table("PrintSubTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'TrackableCouponDeliveryTypes
        Common.QueryStr = "Select TCDeliveryTypeID,TypeDescription from TrackableCouponDeliveryTypes with (NoLock);"
        Construct_Table("TrackableCouponDeliveryTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'RemoteDataOptions
        'This query will return tags for both EngineID 2 and 9 for the time being - not sure at this point if RemoteDataOptions will ever be different between UE and CPE
        Common.QueryStr = "select RDO.PKID, RDO.RemoteDataTypeID, RDO.StyleID, RDO.Enabled " & _
                          "from RemoteDataOptions as RDO with (NoLock) Inner Join RemoteDataTypes as RDT with (NoLock) on RDO.RemoteDataTypeID=RDT.RemoteDataTypeID and RDT.EngineID in (2, 9);"
        Construct_Table("RemoteDataOptions", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'SystemOptions
        Common.QueryStr = "select OptionID, OptionName, OptionValue from SystemOptions with (NoLock);"
        Construct_Table("SystemOptions", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'ProductTypes
        Common.QueryStr = "Select ProductTypeID,PaddingLength,MaxLength,IsNumeric from ProductTypes with (NoLock);"
        Construct_Table("ProductTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")
        'Languages
        Common.QueryStr = "select LanguageID, Name, MSNetCode, JavaLocaleCode, RightToLeftText, AvailableForCustFacing " & _
                          "from Languages with (NoLock);"
        Construct_Table("Languages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'Currencies
        Common.QueryStr = "select CurrencyID, Name, Abbreviation, Symbol, Precision from Currencies with (NoLock);"
        Construct_Table("Currencies", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'UOMTypes
        Common.QueryStr = "select UOMTypeID, Name, UnitTypeID from UOMTypes with (NoLock);"
        Construct_Table("UOMTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'UOMSubTypes
        Common.QueryStr = "select UOMSubTypeID, UOMTypeID, Name, Abbreviation from UOMSubTypes with (NoLock);"
        Construct_Table("UOMSubTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'UOMAmountTypes
        Common.QueryStr = "select AmountTypeID, UOMTypeID from UOMAmountTypes with (NoLock);"
        Construct_Table("UOMAmountTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'UnitTypes
        Common.QueryStr = "select UnitTypeID, Description from CPE_UnitTypes with (NoLock);"
        Construct_Table("UnitTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'AmountTypes
        Common.QueryStr = "select AmountTypeID, Name from CPE_AmountTypes with (NoLock);"
        Construct_Table("AmountTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'CouponTypes 
        Common.QueryStr = "select CouponTypeID, Description FROM CouponTypes with (NoLock);"
        Construct_Table("CouponTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_TenderTypes
        Common.QueryStr = "select TenderTypeID, Name, ExtTenderType, ExtVariety, ExtBinNum from CPE_TenderTypes where Deleted=0;"
        Construct_Table("TenderTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'PrinterTypes
        Common.QueryStr = "select PrinterTypeID, PageWidth, Name, MaxLines from PrinterTypes with (NoLock);"
        Construct_Table("PrinterTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'TrackableCouponProgram
        Common.QueryStr = "select ProgramID,ExtProgramID,Name,CreatedDate,LastUpdate,LastLoaded,LastLoadMsg,Deleted,Description,MaxRedeemCount,ExpireDate from TrackableCouponProgram with (NoLock);"
        Construct_Table("TrackableCouponProgram", 1, DelimChar, LocalServerID, LocationID, "LRT")
        'MarkupTags
        'This query will return tags for both EngineID 2 and 9 for the time being - not sure at this point if these tags will ever be different between UE and CPE
        Common.QueryStr = "select distinct MT.MarkupID, case when MT.NumParams>0 then '|'+Tag+'[' else '|'+Tag+'|' end as Tag from Markuptags as MT with (NoLock) Inner Join MarkupTagUsage as MTU with (NoLock) on MT.MarkupID=MTU.MarkupID where EngineID in (2, 9)"
        Construct_Table("MarkupTags", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'MarkupTagUsage
        Common.QueryStr = "select MarkupID, RewardTypeID, EngineID " & _
                        "from MarkupTagUsage with (NoLock) " & _
                        "where EngineID=9;"
        Construct_Table("MarkupTagUsage", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'PrinterTranslation
        'This query will return tags for both EngineID 2 and 9 for the time being - not sure at this point if these translation rows will ever be different between UE and CPE
        Common.QueryStr = "select distinct PT.TranslationID, PT.PrinterTypeID, PT.MarkupID, ControlChars from PrinterTranslation as PT with (NoLock) Inner Join MarkupTagUsage as MTU with (NoLock) on PT.MarkupID=MTU.MarkUpID where EngineID in (2, 9);"
        Construct_Table("PrinterTranslation", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'ScreenLayouts
        'Common.QueryStr = "select LayoutID, Name, Width, Height from ScreenLayouts with (NoLock) where Deleted=0;"
        'Construct_Table("ScreenLayouts", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'PassThruInterfaceTypes
        Common.QueryStr = "select ProrationTypeId, Description from UE_ProrationTypes;"
        Construct_Table("UEProrationTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'CashierMessages -- leave fields for backward compatibility
        Common.QueryStr = "select CMT.MessageID PKID, " & IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps396), "CMT.Line1, CMT.Line2, CMT.Line3, CMT.Line4, CMT.Line5, CMT.Line6, CMT.Line7, CMT.Line8, CMT.Line9, CMT.Line10,", "CMT.Line1, CMT.Line2, ") & "CMT.Beep, CMT.BeepDuration, CM.PLU " & _
                          "  from CPE_ST_CashierMessageTiers as CMT with (NoLock) " & _
                          " Inner Join CPE_ST_Deliverables as D with (NoLock) on CMT.MessageID=D.OutputID " & _
                          " Inner Join CPE_ST_CashierMessages as CM with (NoLock) on CMT.MessageID=CM.MessageID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "   and D.DeliverableTypeID=9 and D.Deleted=0 and CM.PLU=0;"
        Construct_Table("CashierMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CashierMessageTiers
        Common.QueryStr = "select CMT.PKID, " & IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps396), "CMT.Line1, CMT.Line2, CMT.Line3, CMT.Line4, CMT.Line5, CMT.Line6, CMT.Line7, CMT.Line8, CMT.Line9, CMT.Line10,", "CMT.Line1, CMT.Line2,") & " CMT.Beep, CMT.BeepDuration, CMT.MessageID, CMT.TierLevel, CMT.DisplayImmediate, CM.PLU, CMT.Value " & _
                          "  from CPE_ST_CashierMessageTiers as CMT with (NoLock) " & _
                          " Inner Join CPE_ST_Deliverables as D with (NoLock) on CMT.MessageID=D.OutputID " & _
                          " Inner Join CPE_ST_CashierMessages as CM with (NoLock) on CMT.MessageID=CM.MessageID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "   and D.DeliverableTypeID=9 and D.Deleted=0 and CM.PLU=0;"
        Construct_Table("CashierMessages_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CashierMsgTranslations
        Common.QueryStr = "select CMTrans.PKID, CMTrans.CashierMsgTierID, CMTrans.LanguageID, " & IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps396), "isnull(CMTrans.Line1, '') as Line1, isnull(CMTrans.Line2, '') as Line2, isnull(CMTrans.Line3, '') as Line3, isnull(CMTrans.Line4, '') as Line4, isnull(CMTrans.Line5, '') as Line5, isnull(CMTrans.Line6, '') as Line6, isnull(CMTrans.Line7, '') as Line7, isnull(CMTrans.Line8, '') as Line8, isnull(CMTrans.Line9, '') as Line9, isnull(CMTrans.Line10, '') as Line10 ", "isnull(CMTrans.Line1, '') as Line1, isnull(CMTrans.Line2, '') as Line2 ") & _
                          "from CPE_ST_CashierMsgTranslations as CMTrans with (NoLock) Inner Join CPE_ST_CashierMessageTiers as CMT with (NoLock) on CMTrans.CashierMsgTierID=CMT.PKID " & _
                          "Inner Join CPE_ST_Deliverables as D with (NoLock) on CMT.MessageID=D.OutputID and D.DeliverableTypeID=9 and D.Deleted=0 " & _
                          "Inner Join CPE_ST_CashierMessages as CM with (NoLock) on CMT.MessageID=CM.MessageID and CM.PLU=0 " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        Construct_Table("cashiermessages_tiers_translation", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'Send the special "PLU Not Used" (return to customer) cashier message
        Common.QueryStr = "select CMT.MessageID as PKID, " & IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps396), "CMT.Line1, CMT.Line2, CMT.Line3, CMT.Line4, CMT.Line5, CMT.Line6, CMT.Line7, CMT.Line8, CMT.Line9, CMT.Line10,", "CMT.Line1, CMT.Line2,") & " CMT.Beep, CMT.BeepDuration, CM.PLU " & _
                          "  from CPE_CashierMessageTiers as CMT with (NoLock) " & _
                          " inner join CPE_CashierMessages as CM with (NoLock) on CM.MessageID=CMT.MessageID " & _
                          " where CMT.TierLevel=1 and CM.PLU=1;"
        Construct_Table("CashierMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")
        Common.QueryStr = "select CMT.PKID, " & IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps396), "CMT.Line1, CMT.Line2, CMT.Line3, CMT.Line4, CMT.Line5, CMT.Line6, CMT.Line7, CMT.Line8, CMT.Line9, CMT.Line10,", "CMT.Line1, CMT.Line2,") & " CMT.Beep, CMT.BeepDuration, CMT.MessageID, CMT.TierLevel, CMT.DisplayImmediate, CM.PLU, CMT.Value " & _
                          "  from CPE_CashierMessageTiers as CMT with (NoLock) " & _
                          " inner join CPE_CashierMessages as CM with (NoLock) on CM.MessageID=CMT.MessageID " & _
                          " where CM.PLU=1;"
        Construct_Table("CashierMessages_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")
        Common.QueryStr = "select CMTrans.PKID, CMTrans.CashierMsgTierID, CMTrans.LanguageID, " & IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps396), "isnull(CMTrans.Line1, '') as Line1, isnull(CMTrans.Line2, '') as Line2, isnull(CMTrans.Line3, '') as Line3, isnull(CMTrans.Line4, '') as Line4, isnull(CMTrans.Line5, '') as Line5, isnull(CMTrans.Line6, '') as Line6, isnull(CMTrans.Line7, '') as Line7, isnull(CMTrans.Line8, '') as Line8, isnull(CMTrans.Line9, '') as Line9, isnull(CMTrans.Line10, '') as Line10 ", "isnull(CMTrans.Line1, '') as Line1, isnull(CMTrans.Line2, '') as Line2 ") & _
                          "from CPE_CashierMsgTranslations as CMTrans with (NoLock) Inner Join CPE_CashierMessageTiers as CMT with (NoLock) on CMTrans.CashierMsgTierID=CMT.PKID " & _
                          "Inner Join CPE_CashierMessages as CM with (NoLock) on CMT.MessageID=CM.MessageID " & _
                          "where CM.PLU=1;"
        Construct_Table("cashiermessages_tiers_translation", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PassThrus
        Common.QueryStr = "select DPT.PKID, DPT.DeliverableID, DPT.PassThruRewardID, D.RewardOptionID, DPT.LSInterfaceID, DPT.ActionTypeID " & _
                          "  from CPE_ST_PassThrus as DPT with (NoLock) Inner Join CPE_ST_Deliverables as D with (NoLock) on DPT.PKID=D.OutputID " & _
                          "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "  where D.DeliverableTypeID=12 and D.Deleted=0;"
        Construct_Table("DeliverablePassThrus", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PassThruTiers
        Common.QueryStr = "select DPTT.PKID, DPTT.PTPKID, DPTT.TierLevel, DPTT.Data, DPTT.Value, DPTT.LanguageID,DPT.PassThruRewardID " & _
                          "  from CPE_ST_Deliverables as D with (NoLock) Inner Join CPE_ST_PassThrus as DPT with (NoLock) on DPT.PKID=D.OutputID " & _
                          "    Inner Join CPE_ST_PassThruTiers as DPTT with (NoLock) on DPTT.PTPKID=DPT.PKID " & _
                          "    Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "  where D.DeliverableTypeID=12 and D.Deleted=0;"
        Construct_Table("DeliverablePassThruTiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'ScreenCells
        'Common.QueryStr = "select CellID, LayoutID, ContentsID, X, Y, Width, Height, BackgroundImg from ScreenCells with (NoLock) where Deleted=0;"
        'Construct_Table("ScreenCells", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'ScreenCellContents
        'Common.QueryStr = "select ContentsID, Name from ScreenCellContents with (NoLock) where Deleted=0;"
        'Construct_Table("ScreenCellContents", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'TouchAreas
        'Common.QueryStr = "select AreaID, TA.OnScreenAdID, Name, X, Y, Width, Height " & _
        '                  "from TouchAreas as TA with (NoLock) Inner Join UE_OnScreenAdLocUpdate as OSALU with (NoLock) on TA.OnScreenAdID=OSALU.OnScreenAdID and TA.Deleted=0 " & _
        '                  "Where OSALU.LocationID=" & LocationID & ";"
        'Construct_Table("TouchAreas", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'Scorecards
        'This query will return tags for EngineID 2, 6 and 9 for the time being - not sure at this point if scorecards will ever be different between UE and CPE
        Common.QueryStr = "select ScorecardID, ScorecardTypeID, Description, Priority, Bold, PrintTotalLine, TotalLinePosition from ScoreCards where Deleted=0 and EngineID in (9);"
        Construct_Table("ScoreCards", 1, DelimChar, LocalServerID, LocationID, "LRT")
        
        'ScoreCardConfiguration
        'This Query will return Scorecard Configurations 
        Common.QueryStr = "select PKID, Configuration, Value, ScorecardTypeID from ScoreCardConfiguration;"
        Construct_Table("ScoreCardConfiguration", 1, DelimChar, LocalServerID, LocationID, "LRT")
        
        'ScoreCardConfigurationTranslations
        'This Query will return Scorecard Configurations Translations
        Common.QueryStr = "select PKID, ConfigurationID, LanguageId, PhraseText from ScoreCardConfigurationTranslations;"
        Construct_Table("ScoreCardConfigurationTranslations", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverableROIDs
        Common.QueryStr = "select DR.PKID, DR.DeliverableID, DR.AreaID, DR.RewardOptionID, DR.IncentiveID " & _
                          "  from CPE_ST_DeliverableROIDs DR with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DR.DeliverableID=D.DeliverableID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where D.DeliverableTypeID=1 and D.Deleted=0 and DR.Deleted=0;"
        Construct_Table("DeliverableROIDs", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PointsPrograms
        Common.QueryStr = "select PP.ProgramID as ProgramID, ProgramName, ScorecardID, ScorecardDesc, ScorecardBold, AdjustmentUPC, isnull(CAMProgram, 0) as CAMProgram, ExtHostTypeID, ExtHostProgramID, " & _
                          "isnull(ReturnHandlingTypeID, 1) as ReturnHandlingTypeID, isnull(DisallowRedeemInEarnTrans, 0) as DisallowRedeemInEarnTrans, isnull(AllowNegativeBal, 0) as AllowNegativeBal, isnull(AllowAnyCustomer,0)  as AllowAnyCustomer" & _
                          " from PointsPrograms PP Left join dbo.PointsProgramsPromoEngineSettings as PEPP  with (NoLock) on PP.ProgramID =PEPP.ProgramID  where Deleted=0;"
        Construct_Table("PointsPrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        Common.QueryStr = "select SVProgramID, Name, Value, OneUnitPerRec, SVExpireType, SVExpirePeriodType, ExpirePeriod, ExpireTOD, ExpireDate, " & _
                          " ScorecardID, ScorecardDesc, ScorecardBold, AdjustmentUPC, " & _
                          "isnull(ReturnHandlingTypeID, 1) as ReturnHandlingTypeID, isnull(DisallowRedeemInEarnTrans, 0) as DisallowRedeemInEarnTrans, isnull(AllowNegativeBal, 0) as AllowNegativeBal " & _
                          "  from StoredValuePrograms with (NoLock) where Deleted=0;"

        Construct_Table("StoredValuePrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentivePointsGroups
	
        Dim bEnablePointsCondition As Boolean = IIf(Common.Fetch_UE_SystemOption(181) = "1", True, False)
        Common.QueryStr = "select IPtG.IncentivePointsID, IPtG.RewardOptionID, IPtG.ProgramID, IPtGT.Quantity QtyForIncentive "
        If (bEnablePointsCondition) Then Common.QueryStr &= " , isnull(IPtG.PointsAlgorithmTypeID,3) as PointsAlgorithmTypeID,isnull(IPtG.RewardGrantConditionID,2) as RewardGrantConditionID  "
        Common.QueryStr &= "  from CPE_ST_IncentivePointsGroups as IPtG with (NoLock) " & _
                          " inner join CPE_ST_IncentivePointsGroupTiers as IPtGT with (NoLock) on IPtG.IncentivePointsID=IPtGT.IncentivePointsID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPtG.RewardOptionID " & _
                          " where IPtGT.TierLevel=1 and IPtG.Deleted=0;"
        Construct_Table("IncentivePointsGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentivePointsGroups_Tiers
        Common.QueryStr = "select IPtGT.PKID, IPtGT.IncentivePointsID, IPtGT.TierLevel, IPtGT.Quantity " & _
                          "  from CPE_ST_IncentivePointsGroupTiers as IPtGT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPtGT.RewardOptionID;"
        Construct_Table("IncentivePointsGroups_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT") '-- no need RewardOptionID

        'DeliverablePoints
        Common.QueryStr = "select DP.PKID, DP.DeliverableID, DP.ProgramID, DP.RewardOptionID, DPT.Quantity, DP.ChargebackDeptID, DP.ScorecardID, " & _
                          "       DP.ScorecardDesc, DP.ScorecardBold " & _
                          "  from CPE_ST_DeliverablePoints as DP with (NoLock) " & _
                          " inner join CPE_ST_DeliverablePointTiers as DPT on DP.PKID=DPT.DPPKID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DP.PKID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where DPT.TierLevel=1 and D.DeliverableTypeID=8 and D.Deleted=0 and DP.Deleted=0;"
        Construct_Table("DeliverablePoints", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverablePoints_Tiers
        Common.QueryStr = "select DPT.PKID, DPT.DPPKID, DPT.TierLevel, DPT.Quantity " & _
                          "  from CPE_ST_DeliverablePointTiers as DPT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DPT.DPPKID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "  and D.DeliverableTypeID=8 and D.Deleted=0;"
        Construct_Table("DeliverablePoints_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        ROIDs = "-77"
        Common.QueryStr = "select RewardOptionID from #ActiveROIDs with (NoLock);"
        dst = Common.LRT_Select
        For Each row In dst.Rows
            ROIDs = ROIDs & "," & row.Item("RewardOptionID")
        Next

        'PreferenceRewards
        ' Old Query with In 
        'Common.QueryStr = "Select PreferenceRewardID,DeliverableID,PreferenceID,RewardOptionID from ST_PreferenceRewards " & _
        '                  "where PreferenceRewardID IN (select OutputID from CPE_ST_Deliverables with (NoLock) where DeliverableTypeID=15 and RewardOptionID in (" & ROIDs & ")) and Deleted=0"
        
        Common.QueryStr = "Select SPR.PreferenceRewardID,SPR.DeliverableID,SPR.PreferenceID,SPR.RewardOptionID from ST_PreferenceRewards as SPR with (NoLock)" & _
                     "Inner Join  CPE_ST_Deliverables as CSD with (NoLock) On SPR.PreferenceRewardID = CSD.OutputID " & _
                     "Inner Join  #ActiveROIDs as AR on AR.RewardOptionID =CSD.RewardOptionID and CSD.DeliverableTypeID = 15 and CSD.DELETED =0 where SPR.Deleted =0"
        Construct_Table("PreferenceRewards", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PreferenceRewardTiers
        'Common.QueryStr = "select PreferenceRewardTierID,PreferenceRewardID,TierLevel from  ST_PreferenceRewardTiers where PreferenceRewardID IN " & _
        '                  " (select OutputID from CPE_ST_Deliverables with (NoLock) where DeliverableTypeID=15 and Deleted=0 and RewardOptionID in (" & ROIDs & "))"
        Common.QueryStr = "select SPRT.PreferenceRewardTierID,SPRT.PreferenceRewardID,SPRT.TierLevel from  ST_PreferenceRewardTiers as SPRT with (NoLock) " & _
                       " Inner Join CPE_ST_Deliverables  as CSD with (NoLock) on SPRT.PreferenceRewardID = CSD.OutputID " & _
                       " Inner Join  #ActiveROIDs as AR on AR.RewardOptionID =CSD.RewardOptionID and CSD.DeliverableTypeID = 15 and CSD.DELETED =0 "
        Construct_Table("PreferenceRewardTiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PreferenceRewardTierValues
        Common.QueryStr = "select PTV.PreferenceRewardTierValueID, PTV.PreferenceRewardTierID, PreferenceValue from ST_PreferenceRewardTierValues as PTV  with (NoLock) " & _
                          " Inner Join ST_PreferenceRewardTiers as PT with (NoLock) on PTV.PreferenceRewardTierID=PT.PreferenceRewardTierID " & _
                          " Inner Join ST_PreferenceRewards  as P with (NoLock) on PT.PreferenceRewardID=P.PreferenceRewardID  " & _
                          " Inner Join CPE_ST_Deliverables  as CSD with (NoLock) on P.PreferenceRewardID = CSD.OutputID" & _
                          " Inner Join  #ActiveROIDs as AR on AR.RewardOptionID =CSD.RewardOptionID and CSD.DeliverableTypeID = 15 and CSD.DELETED =0"
        ' " Where P.PreferenceRewardID In (select OutputID from CPE_ST_Deliverables with (NoLock) where DeliverableTypeID=15  and Deleted=0 and RewardOptionID  IN (" & ROIDs & "))"
        Construct_Table("PreferenceRewardTierValues", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveStoredValuePrograms
        Common.QueryStr = "select ISVP.IncentiveStoredValueID, ISVP.RewardOptionID, ISVP.SVProgramID, ISVPT.Quantity QtyForIncentive " & _
                          "  from CPE_ST_IncentiveStoredValuePrograms as ISVP with (NoLock) " & _
                          " inner join CPE_ST_IncentiveStoredValueProgramTiers as ISVPT with (NoLock) on ISVP.IncentiveStoredValueID=ISVPT.IncentiveStoredValueID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ISVP.RewardOptionID " & _
                          " where ISVPT.TierLevel=1 and Deleted=0;"
        Construct_Table("IncentiveStoredValuePrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveStoredValuePrograms_Tiers (no need to send IPGT.RewardOptionID)
        Common.QueryStr = "select ISVPT.PKID, ISVPT.IncentiveStoredValueID, ISVPT.TierLevel, ISVPT.Quantity " & _
                          "  from CPE_ST_IncentiveStoredValueProgramTiers as ISVPT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ISVPT.RewardOptionID;"
        Construct_Table("IncentiveStoredValuePrograms_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverableStoredValue
        Common.QueryStr = "select DSV.PKID, DSV.DeliverableID, DSV.SVProgramID, DSV.RewardOptionID, DSV.Quantity, DSV.ScorecardID, DSV.ScorecardDesc, DSV.ScorecardBold " & _
                          "  from CPE_ST_DeliverableStoredValue as DSV with (NoLock) " & _
                          " inner join CPE_ST_DeliverableStoredValueTiers as DSVT with (NoLock) on DSV.PKID=DSVT.DSVPKID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DSV.PKID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where DSVT.TierLevel=1 and D.DeliverableTypeID=11 and D.Deleted=0 and DSV.Deleted=0;"
        Construct_Table("DeliverableStoredValue", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverableStoredValueTiers
        Common.QueryStr = "select DSVT.PKID, DSVT.DSVPKID, DSVT.TierLevel, DSVT.Quantity " & _
                          "  from CPE_ST_DeliverableStoredValueTiers as DSVT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DSVT.DSVPKID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where D.DeliverableTypeID=11 and D.Deleted=0;"
        Construct_Table("DeliverableStoredValue_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverableMonStoredValue
        Common.QueryStr = "select convert(nvarchar,DSV.PKID,0) as PKID, convert(nvarchar,DSV.DeliverableID,0) as DeliverableID, convert(nvarchar,DSV.SVProgramID,0) as SVProgramID," & _
                          "  convert(nvarchar,DSV.RewardOptionID,0) as RewardOptionID, convert(nvarchar(1000),isnull(RTRIM(LTRIM(SVP.Description)),''),0) as Description" & _
                      "  from CPE_ST_DeliverableMonStoredValue as DSV with (NoLock) " & _
                      " inner join CPE_ST_Deliverables as D with (NoLock) on DSV.PKID=D.OutputID " & _
                      " inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                      " inner join StoredValuePrograms as SVP with (NoLock) on DSV.SVProgramID = SVP.SVProgramID  " & _
                      " where D.DeliverableTypeID=16 and DSV.Deleted=0 and D.Deleted=0 and SVP.SVTypeID=2;"
        Construct_Table("DELIVERABLEMONSTOREDVALUEREDEMPTION", 5, DelimChar, LocalServerID, LocationID, "LRT")
      
        'Monetary Stored Value Translations       
        Common.QueryStr = "select MSVT.PKID, MSVT.SVProgramID,MSVT.LanguageID, convert(nvarchar(1000),isnull(RTRIM(LTRIM(MSVT.Description)),''),0) as Description " & _
                  "from CPE_ST_DeliverableMonSVTranslations as MSVT with (NoLock) " & _
                  "inner join CPE_ST_DeliverableMonStoredValue as MSV with (NoLock) on MSVT.SVProgramID = MSV.SVProgramID " & _
                  "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=MSV.RewardOptionID " & _
                  "and MSV.Deleted=0;"
        Construct_Table("MONSVREDEMPTIONTRANSLATIONS", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'IncentiveTrackableCouponPrograms
        Common.QueryStr = "select TCC.TrackableCouponConditionId, AR.RewardOptionID, TCC.ProgramId, ORC.JoinTypeId from Conditions_ST CON with (NoLock) " & _
                          " inner join OfferRegularConditions_ST ORC with (NoLock) on CON.ConditionId=ORC.ConditionId " & _
                          " inner join TrackableCouponsCondition_ST TCC with (NoLock) on TCC.ConditionId=CON.ConditionId " & _
                          " inner join CPE_ST_RewardOptions RO with (NoLock) on RO.IncentiveId=ORC.OfferId " & _
                          " inner join #ActiveROIDs AR on AR.RewardOptionId=RO.RewardOptionId " & _
                          " where CON.Deleted=0 AND CON.ConditionTypeId=15 AND CON.EngineId=9 AND RO.Deleted=0;"

        Construct_Table("IncentiveTrackableCouponPrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'LocationOptions
        If Not (OperateAtEnterprise) Then
            Common.QueryStr = "select PKID, LocationID, OptionID, OptionValue from LocationOptions with (NoLock) where Deleted=0 and LocationID=" & LocationID & ";"
        Else
            Common.QueryStr = "select PKID, LocationID, OptionID, OptionValue from LocationOptions with (NoLock) where Deleted=0;"
        End If
        Construct_Table("LocationOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'SiteSpecificOptions
        Common.QueryStr = "select SSO.OptionID, OptionName from SiteSpecificOptions SSO with (NoLock);"
        Construct_Table("SiteSpecificOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'SiteSpecificOptionValues
        Common.QueryStr = "select PKID, OptionID, ValueDescription, OptionValue, DefaultVal " & _
                          "from SiteSpecificOptionValues SSOV with (NoLock);"
        Construct_Table("SiteSpecificOptionValues", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'Incentives
        'Common.QueryStr = "select distinct I.IncentiveID, I.IncentiveName, I.Priority, I.StartDate, I.EndDate, I.TestingStartDate, I.TestingEndDate, isnull(EveryDOW, 1) as EveryDOW, isnull(I.EligibilityStartDate, I.StartDate) as EligibilityStartDate, isnull(I.EligibilityEndDate, I.EndDate) as EligibilityEndDate, P1DistQtyLimit, P1DistPeriod, P1DistTimeType, " & _
        '                  "P2DistQtyLimit, P2DistPeriod, P2DistTimeType, P3DistQtyLimit, P3DistPeriod, P3DistTimeType, isnull(EnableRedeemRpt, 1) as Reporting, EmployeesOnly, EmployeesExcluded, UpdateLevel, DeferCalcToEOS, EveryTOD, ChargebackVendorID, SendIssuance, ManufacturerCoupon, InboundCRMEngineID, EnableImpressRpt, ClientOfferID, VendorCouponCode, isnull(EngineID, 2) as EngineID, MutuallyExclusive, EngineSubTypeID, isnull(PromoClassID, 0) as PromoClassID " & _
        '                  "From CPE_ST_Incentives as I with (NoLock) where I.IncentiveID in (" & ActiveIncentiveIDs & ");"

        If (Common.IsEngineInstalled(9) AndAlso (Common.Fetch_UE_SystemOption(146) = "1")) Then
            Common.QueryStr = "select convert(nvarchar,I.IncentiveID) as IncentiveID,I.IncentiveName,convert(nvarchar,isnull((100 - I.Priority),0))+char(" & DelimChar & ")+convert(nvarchar,I.StartDate,120)+char(" & DelimChar & ")+convert(nvarchar,I.EndDate,120)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,I.TestingStartDate,120)+char(" & DelimChar & ")+convert(nvarchar,I.TestingEndDate,120)+char(" & DelimChar & ")+convert(nvarchar,isnull(EveryDOW, 1))+char(" & DelimChar & ")+convert(nvarchar,isnull(I.EligibilityStartDate, I.StartDate),120)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(I.EligibilityEndDate, I.EndDate),120)+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistQtyLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistTimeType,0))+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(P2DistQtyLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P2DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P2DistTimeType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P3DistQtyLimit,0))+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(P3DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P3DistTimeType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(EnableRedeemRpt, 1))+char(" & DelimChar & ")+convert(nvarchar,EmployeesOnly)+char(" & DelimChar & ")+convert(nvarchar,EmployeesExcluded)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,UpdateLevel)+char(" & DelimChar & ")+convert(nvarchar,DeferCalcToEOS)+char(" & DelimChar & ")+convert(nvarchar,EveryTOD)+char(" & DelimChar & ")+convert(nvarchar,ChargebackVendorID)+char(" & DelimChar & ")+convert(nvarchar,SendIssuance)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,ManufacturerCoupon)+char(" & DelimChar & ")+convert(nvarchar,InboundCRMEngineID)+char(" & DelimChar & ")+convert(nvarchar,EnableImpressRpt)+char(" & DelimChar & ")+isnull(ClientOfferID,'')+char(" & DelimChar & ")+isnull(VendorCouponCode,'')+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(EngineID, 9))+char(" & DelimChar & ")+convert(nvarchar,MutuallyExclusive)+char(" & DelimChar & ")+convert(nvarchar,EngineSubTypeID)+char(" & DelimChar & ")+convert(nvarchar,isnull(PromoClassID, 0))+char(" & DelimChar & ")+convert(nvarchar,isnull(DiscountEvalTypeID, 0))" & _
                              IIf(Common.Fetch_UE_SystemOption(180) = "1", "+char(" & DelimChar & ")+convert(nvarchar,PreOrderEligibility)", "") & _
                IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps2297), IIF(Common.Fetch_UE_SystemOption(211) = "1", "+char(" & DelimChar & ")+convert(nvarchar,isnull(DeferCalcToTotal, 0))", "") & "+char(" & DelimChar & ")+convert(nvarchar,isnull(pointsprogramwatch,0)) ", " ")
                If bUsePromotionDisplay Then Common.QueryStr &= "+char(" & DelimChar & ")+convert(nvarchar,isnull(PromotionDisplay, 0)) "
                If bUseProrateonDisplay Then Common.QueryStr &= "+char(" & DelimChar & ")+convert(nvarchar,isnull(ProrateonDisplay, 0)) "
                Common.QueryStr &= "+char(" & DelimChar & ")+convert(nvarchar,isnull(StoreCoupon, 0))+char(" & DelimChar & ")+convert(nvarchar,isnull(PosNotificationCheck, 0)) "
                Common.QueryStr &= " as Data " & _
                              "From CPE_ST_Incentives as I with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=I.IncentiveID;"
        Else
            Common.QueryStr = "select convert(nvarchar,I.IncentiveID) as IncentiveID,I.IncentiveName,convert(nvarchar,isnull(I.Priority,0))+char(" & DelimChar & ")+convert(nvarchar,I.StartDate,120)+char(" & DelimChar & ")+convert(nvarchar,I.EndDate,120)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,I.TestingStartDate,120)+char(" & DelimChar & ")+convert(nvarchar,I.TestingEndDate,120)+char(" & DelimChar & ")+convert(nvarchar,isnull(EveryDOW, 1))+char(" & DelimChar & ")+convert(nvarchar,isnull(I.EligibilityStartDate, I.StartDate),120)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(I.EligibilityEndDate, I.EndDate),120)+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistQtyLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P1DistTimeType,0))+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(P2DistQtyLimit,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P2DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P2DistTimeType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P3DistQtyLimit,0))+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(P3DistPeriod,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(P3DistTimeType,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(EnableRedeemRpt, 1))+char(" & DelimChar & ")+convert(nvarchar,EmployeesOnly)+char(" & DelimChar & ")+convert(nvarchar,EmployeesExcluded)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,UpdateLevel)+char(" & DelimChar & ")+convert(nvarchar,DeferCalcToEOS)+char(" & DelimChar & ")+convert(nvarchar,EveryTOD)+char(" & DelimChar & ")+convert(nvarchar,ChargebackVendorID)+char(" & DelimChar & ")+convert(nvarchar,SendIssuance)+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,ManufacturerCoupon)+char(" & DelimChar & ")+convert(nvarchar,InboundCRMEngineID)+char(" & DelimChar & ")+convert(nvarchar,EnableImpressRpt)+char(" & DelimChar & ")+isnull(ClientOfferID,'')+char(" & DelimChar & ")+isnull(VendorCouponCode,'')+char(" & DelimChar & ")+" & _
                              "convert(nvarchar,isnull(EngineID, 9))+char(" & DelimChar & ")+convert(nvarchar,MutuallyExclusive)+char(" & DelimChar & ")+convert(nvarchar,EngineSubTypeID)+char(" & DelimChar & ")+convert(nvarchar,isnull(PromoClassID, 0))+char(" & DelimChar & ")+convert(nvarchar,isnull(DiscountEvalTypeID, 0))" & _
                              IIf(Common.Fetch_UE_SystemOption(180) = "1", "+char(" & DelimChar & ")+convert(nvarchar,PreOrderEligibility)", "") & _
                IIf(Common.use_development_feature(Copient.CommonIncConfigurable.DevFeatureSwitches.amsps2297), IIF(Common.Fetch_UE_SystemOption(211) = "1", "+char(" & DelimChar & ")+convert(nvarchar,isnull(DeferCalcToTotal, 0))", "") & "+char(" & DelimChar & ")+convert(nvarchar,isnull(pointsprogramwatch,0)) ", " ")
                If bUsePromotionDisplay Then Common.QueryStr &= "+char(" & DelimChar & ")+convert(nvarchar,isnull(PromotionDisplay, 0)) "
                If bUseProrateonDisplay Then Common.QueryStr &= "+char(" & DelimChar & ")+convert(nvarchar,isnull(ProrateonDisplay, 0)) "
                Common.QueryStr &= "+char(" & DelimChar & ")+convert(nvarchar,isnull(StoreCoupon, 0))+char(" & DelimChar & ")+convert(nvarchar,isnull(PosNotificationCheck, 0)) "
                Common.QueryStr &= " as Data " & _
                              "From CPE_ST_Incentives as I with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=I.IncentiveID;"
        End If
        Construct_Table("Incentives", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'OfferTranslations
        'send the data from the OfferTranslations table
        Common.QueryStr = "select PKID, OfferID, OfferName, LanguageID " & _
                          "from CPE_ST_OfferTranslations as OT with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=OT.OfferID;"
        Construct_Table("OfferTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveDOW
        'send the data from the IncentiveDOW table
        Common.QueryStr = "select IDOW.IncentiveDOWID, IDOW.IncentiveID, IDOW.DOWID " & _
                          "from CPE_ST_IncentiveDOW as IDOW with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=IDOW.IncentiveID and IDOW.Deleted=0;"
        Construct_Table("IncentiveDOW", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveTOD
        'send the data from the IncentiveTOD table
        Common.QueryStr = "select ITOD.IncentiveTODID, ITOD.IncentiveID, ITOD.StartTime, ITOD.EndTime " & _
                          "  from CPE_ST_IncentiveTOD as ITOD with (NoLock) Inner Join #ActiveIncentives as AI on AI.IncentiveID=ITOD.IncentiveID;"
        Construct_Table("IncentiveTOD", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'TerminalTypes (in-store locations)
        Common.QueryStr = "select TerminalTypeID, Name, LayoutID, SpecificPromosOnly, FuelProcessing, AnyTerminal, LockingGroupID from TerminalTypes with (NoLock) where EngineID=9 and Deleted=0;"
        Construct_Table("TerminalTypes", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'OfferTerminals (in-store locations associative table)
        'Common.QueryStr = "select OT.PKID, OT.OfferID as IncentiveID, OT.TerminalTypeID, OT.Excluded " & _
        '                  "  from CPE_ST_OfferTerminals as OT with (NoLock) " & _
        '                  " inner join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
        '                  " where T.EngineID=9 and OT.OfferID in (" & ActiveIncentiveIDs & ");"
        Common.QueryStr = "select convert(nvarchar,OT.PKID)+char(" & DelimChar & ")+convert(nvarchar,isnull(OT.OfferID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(OT.TerminalTypeID,0))+char(" & DelimChar & ")+convert(nvarchar,OT.Excluded) as Data " & _
                          "from CPE_ST_OfferTerminals as OT with (NoLock) " & _
                          "inner join TerminalTypes as T with (NoLock) on OT.TerminalTypeID=T.TerminalTypeID " & _
                          "Inner Join #ActiveIncentives as AI on AI.IncentiveID=OT.OfferID " & _
                          "where T.EngineID=9;"
        Construct_Table("IncentiveTerminals", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'ChargebackDepts
        Common.QueryStr = "select ChargeBackDeptID as DeptID, ExternalID as DeptNumber from ChargebackDepts with (NoLock);"
        Construct_Table("ChargebackDepts", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'RewardOptions
        'Common.QueryStr = "select RO.RewardOptionID, RO.Name, RO.IncentiveID, RO.Priority, RO.HHEnable, RO.TouchResponse, RO.ProductComboID, RO.ExcludedTender, " & _
        '                  "       RO.ExcludedTenderAmtRequired, RO.TierLevels, RO.AttributeComboID " & _
        '                  "  from CPE_ST_RewardOptions as RO with (NoLock) where RO.RewardOptionID in (" & ActiveROIDs & ");"
        Common.QueryStr = "select convert(nvarchar,RO.RewardOptionID) as 'RewardOptionID',convert(nvarchar,isnull(Name,'')) as 'Name',convert(nvarchar,RO.IncentiveID)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.Priority, 0))+char(" & DelimChar & ")+convert(nvarchar,RO.HHEnable)+char(" & DelimChar & ")+convert(nvarchar,RO.TouchResponse)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.ProductComboID,0))+char(" & DelimChar & ")+convert(nvarchar,RO.ExcludedTender)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.ExcludedTenderAmtRequired,0))+char(" & DelimChar & ")+convert(nvarchar,RO.TierLevels)+char(" & DelimChar & ")+convert(nvarchar,RO.AttributeComboID)+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.OfflineCustCondition,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(RO.TenderComboID,1))+char(" & DelimChar & ")+convert(nvarchar,isnull(RO.PointsComboID,1))+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(RO.StoredValueComboID,1))+char(" & DelimChar & ")+convert(nvarchar,RO.PreferenceComboID)+char(" & DelimChar & ")+convert(nvarchar,RO.CurrencyID) as Data " & _
                          "from CPE_ST_RewardOptions as RO with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=RO.RewardOptionID;"
        Construct_Table("RewardOptions", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'RewardOptionUOMs
        Common.QueryStr = "select ROUOM.PKID, ROUOM.RewardOptionID, ROUOM.UOMTypeID, ROUOM.UOMSubTypeID " & _
                          "from CPE_ST_RewardOptionUOMs as ROUOM with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ROUOM.RewardOptionID;"
        Construct_Table("RewardOptionUOMs", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'Deliverables
        'Common.QueryStr = "select D.DeliverableID, D.RewardOptionID, D.RewardOptionPhase, D.DeliverableTypeID, D.OutputID, " & _
        '                  "       1 as AvailabilityTypeID, D.Priority, D.ScreenCellID " & _
        '                  "  from CPE_ST_Deliverables as D with (NoLock) where D.RewardOptionID in (" & ActiveROIDs & ") and D.Deleted=0;"
        Common.QueryStr = "select convert(nvarchar,D.DeliverableID)+char(" & DelimChar & ")+convert(nvarchar,D.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(D.RewardOptionPhase,0))+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,isnull(D.DeliverableTypeID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(D.OutputID,0))+char(" & DelimChar & ")+'1'+char(" & DelimChar & ")+convert(nvarchar,isnull(D.Priority,0))" & _
                          "+char(" & DelimChar & ")+convert(nvarchar,isnull(D.ScreenCellID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(D.Required,1)) as Data " & _
                          "from CPE_ST_Deliverables as D with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID and D.Deleted=0;"
        Construct_Table("Deliverables", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'OnScreenAds
        If Not (FailoverServer) Then
            '-------------------------------------------------------
            'Disabling this section since graphics are not currently used by UE
            '-------------------------------------------------------
            'get the OnScreenAd data where the the corresponding record in the OnScreenAdLocUpdate table has
            'an older LastSent date or is non-existent (for that AdID and LocationID

            'the first part of this query is for reward graphics
            'the second part is for the special case of background images
            'Common.QueryStr = "select distinct OSA.OnScreenAdID, OSA.Name, OSA.StoreResponse, OSA.DisplayDuration, OSA.Width, OSA.Height, OSA.ImageType, OSA.UpdateLevel " & _
            '                  "  from OnScreenAds as OSA with (NoLock) inner join UE_OnScreenAdLocUpdate as OSALU with (NoLock) on OSA.OnScreenAdID=OSALU.OnScreenAdID and OSA.Deleted=0 " & _
            '                  " where OSALU.LocationID=" & LocationID & " " & _
            '                  "union " & _
            '                  "select distinct OSA.OnScreenAdID, OSA.Name, OSA.StoreResponse, OSA.DisplayDuration, OSA.Width, OSA.Height, OSA.ImageType, OSA.UpdateLevel " & _
            '                  "  from OnScreenAds as OSA with (NoLock) inner join ScreenCells as SC with (NoLock) on OSA.OnScreenAdID=SC.BackgroundImg " & _
            '                  " where OSA.Deleted=0 and SC.Deleted=0;"
            'Construct_Table("OnScreenAds", 5, DelimChar, LocalServerID, LocationID, "LRT")

            'get the GRAPHIC data from the OnScreenAds Table
            'the first part of this query is for reward graphics
            'the second part is for the special case of background images
            'Common.QueryStr = "select OSA.OnscreenAdID, Imagetype, " & _
            '                  "Case ImageType " & _
            '                  "  When 1 then convert(varchar, OSA.OnscreenAdID)+'img.'+'jpg' " & _
            '                  "  When 2 then convert(varchar, OSA.OnscreenAdID)+'img.'+'gif' " & _
            '                  "END as Graphic, OSA.MD5sum " & _
            '                  "from OnScreenAds as OSA with (NoLock) Inner Join UE_OnScreenAdLocUpdate as OSALU with (NoLock) on OSA.OnScreenAdID=OSALU.OnScreenAdID and OSA.Deleted=0 and isnull(OSA.GraphicSize, 0)>0 " & _
            '                  "Where OSALU.LocationID=" & LocationID & " " & _
            '                  "union " & _
            '                  "select distinct OSA.OnscreenAdID, Imagetype, " & _
            '                  "Case ImageType " & _
            '                  "  When 1 then convert(varchar, OSA.OnscreenAdID)+'img.'+'jpg' " & _
            '                  "  When 2 then convert(varchar, OSA.OnscreenAdID)+'img.'+'gif' " & _
            '                  "END as Graphic, OSA.MD5sum " & _
            '                  "from OnScreenAds as OSA with (NoLock) Inner Join ScreenCells as SC with (NoLock) on OSA.OnScreenAdID=SC.BackgroundImg and OSA.Deleted=0 and SC.Deleted=0 and isnull(OSA.GraphicSize, 0)>0;"
            'Construct_Table("OnScreenAds", 3, DelimChar, LocalServerID, LocationID, "LRT")
            'update the existing OnScreenAdLocUpdate records for this location with the current time/date and set the WaitingACK bit
            'Common.QueryStr = "update UE_OnScreenAdLocUpdate with (RowLock) set LastSent=getdate() " & _
            '                  "Where LocationID=" & LocationID & ";"
            'Common.LRT_Execute()
            'Else
            'Common.QueryStr = "update UE_OnScreenAdLocUpdate with (RowLock) set LastSent='1/1/1981' where LocationID=" & LocationID & ";"
            'Common.LRT_Execute()
        End If

        'PrintedMessages
        Common.QueryStr = "select isnull(PM.MessageID, 0) as MessageID, MessageTypeID as PrintZone, BodyText as TextMsg, " & _
                          "  isnull(PM.SuppressZeroBalance, 0) as SuppressZeroBalance, isnull(PM.SortID,0) as SortID " & _
                          "from CPE_ST_PrintedMessages as PM with (NoLock) " & _
                          "inner join CPE_ST_PrintedMessageTiers as PMT with (NoLock) on PMT.MessageID = PM.MessageID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on PMT.MessageID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where PMT.TierLevel=1 and D.DeliverableTypeID=4 and D.Deleted=0;"
        Construct_Table("PrintedMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PrintedMessages_Tiers
        Common.QueryStr = "select PMT.PKID, isnull(PMT.MessageID, 0) as MessageID, PMT.TierLevel, PMT.BodyText, PMT.Value " & _
                          "  from CPE_ST_PrintedMessageTiers as PMT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on PMT.MessageID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "   and D.DeliverableTypeID=4 and D.Deleted=0;"
        Construct_Table("PrintedMessages_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'PMTranslations
        Common.QueryStr = "select PMTrans.PKID, PMTrans.PMTiersID, PMTrans.LanguageID, PMTrans.BodyText " & _
                          "from CPE_ST_PMTranslations as PMTrans with (NoLock) Inner Join CPE_ST_PrintedMessageTiers as PMT with (NoLock) on PMTrans.PMTiersID=PMT.PKID " & _
                          "inner join CPE_ST_Deliverables as D with (NoLock) on PMT.MessageID=D.OutputID and D.DeliverableTypeID=4 and D.Deleted=0 " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        Construct_Table("PMTranslations", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'EDiscounts
        'get the EDiscount records
        Common.QueryStr = "select distinct ED.DiscountID as EDiscountID, ED.Name, ED.DiscountTypeID, EDT.ReceiptDescription, ED.DiscountedProductGroupID, ED.ExcludedProductGroupID, " & _
                          "       ED.BestDeal, ED.AllowNegative, ED.ComputeDiscount, EDT.DiscountAmount, ED.AmountTypeID, ED.L1Cap, ED.L2DiscountAmt, ED.L2AmountTypeID, ED.L2Cap, " & _
                          "       ED.L3DiscountAmt, ED.L3AmountTypeID, EDT.ItemLimit, EDT.WeightLimit, EDT.DollarLimit, ED.ChargeBackDeptID, ED.DecliningBalance, ED.SVProgramID, " & _
                          "       IsNull(ED.FlexNegative, 0) As FlexNegative, ED.ScorecardID, ED.ScorecardDesc, isnull(ED.AllowMarkup, 0) as AllowMarkup, isnull(ED.DiscountAtOrigPrice, 0) as DiscountAtOrigPrice, " & _
                          "       isnull(ED.ProrationTypeID, 0) as ProrationTypeID, isnull(ED.PriceFilter, 100) as PriceFilter, IsNull(ED.FlexOptions, 0) As FlexOptions, IsNull(ED.GrossPrice, 0) As GrossPrice " & _
                          "  from CPE_ST_Discounts as ED with (NoLock) " & _
                          " inner join CPE_ST_DiscountTiers as EDT with (NoLock) on ED.DiscountID=EDT.DiscountID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on ED.DiscountID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where EDT.TierLevel=1 and D.DeliverableTypeID=2 and D.Deleted=0 and ED.Deleted=0;"
        Construct_Table("EDiscounts", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'EDiscounts_Translations
        ' get the EDiscounts_Translations records
        Common.QueryStr = "select DT.PKID, DT.DiscountID, DT.LanguageID, DT.ScorecardDesc  from CPE_ST_DiscountTranslations DT" & _
                          " Inner Join CPE_ST_Deliverables D ON D.OutputID = DT.DiscountID " & _
                          " where D.DeliverableTypeID = 2;"
        Construct_Table("EDiscounts_Translations", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        '`
        'get the EDiscount records
        'Common.QueryStr = "select EDT.PKID, EDT.DiscountID as EDiscountID, EDT.TierLevel, EDT.ReceiptDescription, EDT.DiscountAmount, " & _
        '                  "       EDT.ItemLimit, EDT.WeightLimit, EDT.DollarLimit, EDT.SPRepeatLevel " & _
        '                  "  from CPE_ST_DiscountTiers as EDT with (NoLock) " & _
        '                  " inner join CPE_ST_Deliverables as D with (NoLock) on EDT.DiscountID=D.OutputID " & _
        '                  "   and D.DeliverableTypeID=2 and D.RewardOptionID in (" & ActiveROIDs & ") and D.Deleted=0"
    
        bAllowDollarTransLimit = (Common.Fetch_UE_SystemOption(187) = "1")
        If bAllowDollarTransLimit Then
            Common.QueryStr = "select convert(nvarchar,EDT.PKID)+char(" & DelimChar & ")+convert(nvarchar,EDT.DiscountID)+char(" & DelimChar & ")+convert(nvarchar,EDT.TierLevel)+char(" & DelimChar & ")" & _
                              "+convert(nvarchar,isnull(EDT.ReceiptDescription,''))+char(" & DelimChar & ")+convert(nvarchar,EDT.DiscountAmount)+char(" & DelimChar & ")+convert(nvarchar,EDT.ItemLimit)" & _
                              "+char(" & DelimChar & ")+convert(nvarchar,EDT.WeightLimit)+char(" & DelimChar & ")+convert(nvarchar,EDT.DollarLimit)+char(" & DelimChar & ")+convert(nvarchar,isnull(EDT.SPRepeatLevel,0))" & _
                              "+char(" & DelimChar & ")+convert(nvarchar,isnull(EDT.BuyDescription, ''))+char(" & DelimChar & ")+convert(nvarchar,isnull(EDT.RewardLimitTypeID,0)) as Data " & _
                              "from CPE_ST_DiscountTiers as EDT with (NoLock) " & _
                              "inner join CPE_ST_Deliverables as D with (NoLock) on EDT.DiscountID=D.OutputID " & _
                              "and D.DeliverableTypeID=2 and D.Deleted=0 " & _
                              "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        Else
            Common.QueryStr = "select convert(nvarchar,EDT.PKID)+char(" & DelimChar & ")+convert(nvarchar,EDT.DiscountID)+char(" & DelimChar & ")+convert(nvarchar,EDT.TierLevel)+char(" & DelimChar & ")" & _
                              "+convert(nvarchar,isnull(EDT.ReceiptDescription,''))+char(" & DelimChar & ")+convert(nvarchar,EDT.DiscountAmount)+char(" & DelimChar & ")+convert(nvarchar,EDT.ItemLimit)" & _
                              "+char(" & DelimChar & ")+convert(nvarchar,EDT.WeightLimit)+char(" & DelimChar & ")+convert(nvarchar,EDT.DollarLimit)+char(" & DelimChar & ")+convert(nvarchar,isnull(EDT.SPRepeatLevel,0))" & _
                              "+char(" & DelimChar & ")+convert(nvarchar,isnull(EDT.BuyDescription, '')) as Data " & _
                              "from CPE_ST_DiscountTiers as EDT with (NoLock) " & _
                              "inner join CPE_ST_Deliverables as D with (NoLock) on EDT.DiscountID=D.OutputID " & _
                              "and D.DeliverableTypeID=2 and D.Deleted=0 " & _
                              "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID;"
        End If
        Construct_Table("EDiscounts_Tiers", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'DiscountTiersTranslations
        Common.QueryStr = "select DTT.PKID, DTT.DiscountTiersID, DTT.LanguageID, isnull(DTT.ReceiptDesc, '') as ReceiptDescription, isnull(DTT.BuyDesc, '') as BuyDescription " & _
                          "from CPE_ST_DiscountTiersTranslations as DTT with (NoLock) Inner Join CPE_ST_DiscountTiers as EDT with (NoLock) on DTT.DiscountTiersID=EDT.PKID " & _
                          "inner join CPE_ST_Deliverables as D with (NoLock) on EDT.DiscountID=D.OutputID " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "and D.DeliverableTypeID=2 and D.Deleted=0;"
        Construct_Table("ediscounts_tiers_translation", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'Discount productgroup exclusions
        Dim resultDiscountExGroups As AMSResult(Of String)
        resultDiscountExGroups = m_DiscountService.GetExclusionGroupDeploymentData(dst, Chr(DelimChar).ToString())
        If resultDiscountExGroups.ResultType = AMSResultType.Success AndAlso Not String.IsNullOrWhiteSpace(resultDiscountExGroups.Result) Then
            SDb(resultDiscountExGroups.Result)
        End If

        'GiftCard
        Dim result As AMSResult(Of String)
        result = m_giftcardreward.ConstructGiftCardDataForEngine(ROIDs, Chr(DelimChar).ToString())
        result.Result += Environment.NewLine
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SDb(result.Result)
        End If
        'GiftCardTier
        result = m_giftcardreward.ConstructGiftCardTierDataForEngine(ROIDs, Chr(DelimChar).ToString())
        result.Result += Environment.NewLine
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SDb(result.Result)
        End If
        'GiftCardTierTranslation
        result = m_giftcardreward.ConstructGiftCardTierTranslationDataForEngine(ROIDs, Chr(DelimChar).ToString())
        result.Result += Environment.NewLine
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SDb(result.Result)
        End If
        'coupon reward
        result = m_couponreward.ConstructCouponRewardDataForEngine(ROIDs, Chr(DelimChar).ToString())
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SD(result.Result)
        End If
        'coiponrewardtierdata
        result = m_couponreward.ConstructCouponRewardTierDataForEngine(ROIDs, Chr(DelimChar).ToString())
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SD(result.Result)
        End If
        'couponrewardtranslation
        result = m_couponreward.ConstructCouponRewardTranslationDataForEngine(ROIDs, Chr(DelimChar).ToString())
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SD(result.Result)
        End If
        'Proximity Message
        result = m_PMReward.ConstructProximityMessageDataForEngine(ROIDs, Chr(DelimChar).ToString())
        result.Result += Environment.NewLine
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SDb(result.Result)
        End If
        ' Proximity Message Tier
        result = m_PMReward.ConstructProximityMessageTierDataForEngine(ROIDs, Chr(DelimChar).ToString())
        result.Result += Environment.NewLine
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SDb(result.Result)
        End If
        'Proximity Message Tier Translation
        result = m_PMReward.ConstructProximityMessageTierTranslationDataForEngine(ROIDs, Chr(DelimChar).ToString())
        result.Result += Environment.NewLine
        If (result.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(result.Result)) Then SDb(result.Result)
        End If
        'SpecialPricing
        'get the CPE_SpecialPricing records
        Common.QueryStr = "select SP.SpecialPricingID, SP.DiscountID, SP.DiscountTierID, SP.Value, SP.LevelID " & _
                          "  from CPE_ST_SpecialPricing AS SP with (NoLock) " & _
                          " inner join CPE_ST_DiscountTiers AS DT with (NoLock) on DT.PKID = SP.DiscountTierID " & _
                          " inner join CPE_ST_Discounts as ED with (NoLock) on SP.DiscountID=ED.DiscountID " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DT.DiscountID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "where D.DeliverableTypeID=2 and D.Deleted=0 and ED.Deleted=0"
        Construct_Table("SpecialPricing", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'FrankingMessages
        'get the FrankingMessages records
        Common.QueryStr = "select distinct FMT.FrankID, FMT.OpenDrawer, FMT.ManagerOverride, FMT.FrankFlag, FMT.FrankingText, FMT.Line1, FMT.Line2, FMT.Beep, FMT.BeepDuration " & _
                          "  from CPE_ST_FrankingMessageTiers as FMT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on FMT.FrankID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where FMT.TierLevel=1 and D.DeliverableTypeID=10 and D.Deleted=0;"
        Construct_Table("FrankingMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'FrankingMessages_Tiers
        Common.QueryStr = "select FMT.PKID, FMT.OpenDrawer, FMT.ManagerOverride, FMT.FrankFlag, FMT.FrankingText, FMT.Line1, FMT.Line2, FMT.Beep, FMT.BeepDuration, FMT.FrankID, FMT.TierLevel, FMT.Value " & _
                          "  from CPE_ST_FrankingMessageTiers as FMT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on FMT.FrankID=D.OutputID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          " where D.DeliverableTypeID=10 and D.Deleted=0;"
        Construct_Table("FrankingMessages_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'DeliverableCustomerGroupTiers
        Common.QueryStr = "select DCGT.PKID, DCGT.DeliverableID, DCGT.CustomerGroupID, DCGT.TierLevel, DCGT.Value " & _
                          "  from CPE_ST_DeliverableCustomerGroupTiers as DCGT with (NoLock) " & _
                          " inner join CPE_ST_Deliverables as D with (NoLock) on DCGT.DeliverableID=D.DeliverableID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=D.RewardOptionID " & _
                          "   and D.DeliverableTypeID=5 and D.Deleted=0;"
        Construct_Table("DeliverableCustomerGroups_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveProductGroups (keep Deleted Column)
        'get new data from the IncentiveProductGroups table (keep Deleted column)
        TempQuery = "select convert(nvarchar,IPG.IncentiveProductGroupID)+char(" & DelimChar & ")+convert(nvarchar,IPG.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ProductGroupID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPGT.Quantity,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.QtyUnitType,0))+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,IPG.ExcludedProducts)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.Disqualifier,0))+char(" & DelimChar & ")+convert(nvarchar,IPG.UniqueProduct)+char(" & DelimChar & ")+convert(nvarchar,IPG.Rounding)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.MinPurchAmt,0))" & _
                          IIf(Common.Fetch_UE_SystemOption(210) = "1", "+char(" & DelimChar & ")+convert(nvarchar,IPG.NetPriceProduct)", "")
        If Common.Fetch_UE_SystemOption(182) = "1" Then
            TempQuery = TempQuery & "+char(" & DelimChar & ")+ convert(nvarchar,IPG.ReturnedItemGroup)"
        End If
        TempQuery = TempQuery & "+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.MinItemPrice,0))" & _
                          "+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.FullPrice,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ClearanceState,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ClearanceLevel,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.TenderType,0))+char(" & DelimChar & ")+ convert(nvarchar,isnull(IPG.SameItem,0)) as Data " & _
                          "from CPE_ST_IncentiveProductGroups as IPG with (NoLock) " & _
                          "inner join CPE_ST_IncentiveProductGroupTiers as IPGT with (NoLock) on IPG.IncentiveProductGroupID=IPGT.IncentiveProductGroupID " & _
                          "Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPG.RewardOptionID " & _
                          "where IPGT.TierLevel=1 and IPG.Deleted=0 and ExcludedProducts=0 " & _
                          "union " & _
                          "select convert(nvarchar,IPG.IncentiveProductGroupID)+char(" & DelimChar & ")+convert(nvarchar,IPG.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ProductGroupID,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.QtyForIncentive,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.QtyUnitType,0))+char(" & DelimChar & ")" & _
                          "+convert(nvarchar,IPG.ExcludedProducts)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.Disqualifier,0))+char(" & DelimChar & ")+convert(nvarchar,IPG.UniqueProduct)+char(" & DelimChar & ")+convert(nvarchar,IPG.Rounding)+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.MinPurchAmt,0))" & _
                          IIf(Common.Fetch_UE_SystemOption(210) = "1", "+char(" & DelimChar & ")+convert(nvarchar,IPG.NetPriceProduct)", "")
        If Common.Fetch_UE_SystemOption(182) = "1" Then
            TempQuery = TempQuery & "+char(" & DelimChar & ")+ convert(nvarchar,IPG.ReturnedItemGroup)"
        End If
        Common.QueryStr = TempQuery & "+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.MinItemPrice,0))" & _
                  "+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.FullPrice,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ClearanceState,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.ClearanceLevel,0))+char(" & DelimChar & ")+convert(nvarchar,isnull(IPG.TenderType,0)) +char(" & DelimChar & ")+ convert(nvarchar,isnull(IPG.SameItem,0))as Data " & _
                          "from CPE_ST_IncentiveProductGroups as IPG with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPG.RewardOptionID " & _
                          "where IPG.Deleted=0 and IPG.ExcludedProducts=1;"
        Construct_Table("IncentiveProductGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")
                
        'IncentiveProductGroupTiers (no need to send IPGT.RewardOptionID)
        'get new data from the IncentiveProductGroups table
        Common.QueryStr = "select convert(nvarchar,IPGT.PKID)+char(" & DelimChar & ")+convert(nvarchar,IPGT.IncentiveProductGroupID)+char(" & DelimChar & ")+convert(nvarchar,IPGT.TierLevel)+char(" & DelimChar & ")+convert(nvarchar,IPGT.Quantity) as Data " & _
                          "from CPE_ST_IncentiveProductGroupTiers as IPGT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPGT.RewardOptionID;"
        Construct_Table("IncentiveProductGroups_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'AMS-684 Multiple exclusion groups
        Dim resultExclusionGroups As AMSResult(Of String) = m_ProdCond.GetExclusionGroupDeploymentData(dst, Chr(DelimChar))
        If resultExclusionGroups.ResultType = AMSResultType.Success Then
            If (Not String.IsNullOrWhiteSpace(resultExclusionGroups.Result)) Then SDb(resultExclusionGroups.Result)
        End If
        'IncentiveUserGroups (keep Deleted column)
        'get new data from the IncentiveUserGroups table
        'Common.QueryStr = "select ICG.IncentiveCustomerID as IncentiveUserID, ICG.RewardOptionID, ICG.CustomerGroupID as UserGroupID, ICG.ExcludedUsers " & _
        '                  "  from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) " & _
        '                  " where ICG.Deleted=0 and ICG.RewardOptionID in (" & ActiveROIDs & ");"
        Common.QueryStr = "select convert(nvarchar,ICG.IncentiveCustomerID)+char(" & DelimChar & ")+convert(nvarchar,ICG.RewardOptionID)+char(" & DelimChar & ")+convert(nvarchar,isnull(ICG.CustomerGroupID,0))+char(" & DelimChar & ")+convert(nvarchar,ICG.ExcludedUsers) as Data " & _
                          "from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ICG.RewardOptionID " & _
                          "where ICG.Deleted=0;"
        Construct_Table("IncentiveUserGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")
        
        'IncentiveUserGroupCardTypes
        Dim resultCustCondCardTypes As AMSResult(Of String)
        resultCustCondCardTypes = m_CustCondService.ConstructCustomerConditionCardTypesForEngine(dst, Chr(DelimChar).ToString())
        If resultCustCondCardTypes.ResultType = AMSResultType.Success AndAlso Not String.IsNullOrWhiteSpace(resultCustCondCardTypes.Result) Then
            SDb(resultCustCondCardTypes.Result)
        End If

        'IncentiveCustomerApproval
        Dim resultCustomerApproval As AMSResult(Of String) = m_CustCondService.ConstructCustomerApprovalDataForEngine(ROIDs, Chr(DelimChar).ToString())
        If (resultCustomerApproval.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(resultCustomerApproval.Result)) Then SDb(resultCustomerApproval.Result)
        End If

        'IncentiveCustomerApprovalTranslation
        Dim resultCustomerApprovalTrans As AMSResult(Of String) = m_CustCondService.ConstructCustomerApprovalTranslationDataForEngine(ROIDs, Chr(DelimChar).ToString())
        If (resultCustomerApprovalTrans.ResultType = AMSResultType.Success) Then
            If (Not String.IsNullOrWhiteSpace(resultCustomerApprovalTrans.Result)) Then SDb(resultCustomerApprovalTrans.Result)
        End If

        'IncentiveTenderTypes
        Common.QueryStr = "select ITT.IncentiveTenderID, ITT.RewardOptionID, ITT.TenderTypeID, ITTT.Value " & _
                          "  from CPE_ST_IncentiveTenderTypes as ITT with (NoLock) " & _
                          " inner join CPE_ST_IncentiveTenderTypeTiers as ITTT with (NoLock) on ITT.IncentiveTenderID=ITTT.IncentiveTenderID " & _
                          " Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ITT.RewardOptionID " & _
                          " where ITTT.TierLevel=1 and ITT.Deleted=0;"
        Construct_Table("IncentiveTenderTypes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveTenderType_Tiers (no need to send IPGT.RewardOptionID)
        Common.QueryStr = "select ITTT.PKID, ITTT.IncentiveTenderID, ITTT.TierLevel, ITTT.Value " & _
                          "  from CPE_ST_IncentiveTenderTypeTiers as ITTT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=ITTT.RewardOptionID;"
        Construct_Table("IncentiveTenderTypes_Tiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveInstantWin
        Common.QueryStr = "select IWIN.IncentiveInstantWinID as incentiveinstantwinid, IWIN.RewardOptionID, IWIN.NumPrizesAllowed as reward, IWIN.OddsOfWinning as odds, IWIN.RandomWinners as Random, IWIN.Unlimited as Unlimited, AwardLimitEnterprise as awardlimitenterprise, ChanceOfWinningEnterprise as chanceofwinningenterprise " & _
                          "  from CPE_ST_IncentiveInstantWin as IWIN with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IWIN.RewardOptionID " & _
                          "where IWIN.Deleted=0;"
        Construct_Table("IncentiveInstantWinPrograms", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentivePLUs
        sbExtRedemptionAuth.Append("select IPLU.IncentivePLUID, IPLU.RewardOptionID, IPLU.PLU, IPLU.PerRedemption, IPLU.CashierMessage, IPLU.PLUQuantity")
        If Common.Fetch_UE_SystemOption(172) Then
            sbExtRedemptionAuth.Append(", IPLU.ExternalRedemptionAuthorization")
        End If
        Common.QueryStr = sbExtRedemptionAuth.ToString() & "  from CPE_ST_IncentivePLUs as IPLU with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IPLU.RewardOptionID;"
        Construct_Table("IncentivePLUs", 5, DelimChar, LocalServerID, LocationID, "LRT")
        'From 6.1, we dont have seperate tables for Enterprise IW and triggers concept
        'IncentiveEIW
        'Common.QueryStr = "select IEIW.IncentiveEIWID, IEIW.RewardOptionID, NumberOfPrizes, FrequencyID " & _
        '                 "  from CPE_ST_IncentiveEIW as IEIW with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IEIW.RewardOptionID;"
        'Construct_Table("IncentiveEIW", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'EIWTriggers
        'Common.QueryStr = "select EIWT.TriggerID, EIWT.IncentiveID, EIWT.RewardOptionID, EIWT.IncentiveEIWID, EIWT.TriggerTime, CASE ISNULL(EIWU.TriggerID, 0) WHEN 0 THEN 0 ELSE 1 END AS Consumed " & _
        '                  "  FROM CPE_EIWTriggers as EIWT with (NoLock) LEFT OUTER JOIN CPE_EIWTriggersUsed AS EIWU WITH (NoLock) ON EIWU.TriggerID = EIWT.TriggerID " & _
        '                  "  where EIWT.Removed = 0 AND EIWU.TriggerID IS NULL;"
        'Common.QueryStr = "select convert(nvarchar,EIWT.TriggerID)+char(" & DelimChar & ")+convert(nvarchar,isnull(EIWT.IncentiveID,0))+char(" & DelimChar & ")+CONVERT(nvarchar,isnull(EIWT.RewardOptionID,0))+char(" & DelimChar & ")" & _
        '                 "+CONVERT(nvarchar,isnull(EIWT.IncentiveEIWID,0))+char(" & DelimChar & ")+CONVERT(nvarchar,isnull(EIWT.TriggerTime,'1/1/1980'))+char(" & DelimChar & ")+CONVERT(nvarchar,CASE ISNULL(EIWU.TriggerID, 0) WHEN 0 THEN 0 ELSE 1 END) as Data " & _
        '                "FROM CPE_EIWTriggers as EIWT with (NoLock) LEFT OUTER JOIN CPE_EIWTriggersUsed AS EIWU WITH (NoLock) ON EIWU.TriggerID = EIWT.TriggerID " & _
        '               "where EIWT.Removed = 0 AND EIWU.TriggerID IS NULL;"
        'Construct_Table("EIWTriggers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'Vendors
        'Common.QueryStr = "select VendorID, ExtVendorID " & _
        '                  "  FROM Vendors with (NoLock) " & _
        '                  "  WHERE Deleted=0;"
        Common.QueryStr = "select convert(nvarchar,VendorID)+char(" & DelimChar & ")+convert(nvarchar,ExtVendorID) as Data " & _
                          "  FROM Vendors with (NoLock) " & _
                          "  WHERE Deleted=0;"
        Construct_Table("Vendors", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'OfferCategories
        Common.QueryStr = "select OfferCategoryID, ExtCategoryID, isnull(BaseOfferID, 0) as BaseOfferID " & _
                          "  FROM OfferCategories with (NoLock) " & _
                          "  WHERE Deleted=0;"
        Construct_Table("OfferCategories", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveAttributes
        Common.QueryStr = "select IA.IncentiveAttributeID, IA.RewardOptionID " & _
                          "  FROM CPE_ST_IncentiveAttributes as IA with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IA.RewardOptionID;"
        Construct_Table("IncentiveAttributes", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'IncentiveAttributeTiers
        Common.QueryStr = "select IAT.PKID, IAT.IncentiveAttributeID, IAT.RewardOptionID, IAT.AttributeTypeID, IAT.TierLevel, IAT.AttributeValues " & _
                          "  FROM CPE_ST_IncentiveAttributeTiers as IAT with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IAT.RewardOptionID;"
        Construct_Table("IncentiveAttributeTiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_IncentivePrefs
        Common.QueryStr = "select IP.IncentivePrefsID, IP.RewardOptionID, IP.PreferenceID " & _
                      "  FROM CPE_ST_IncentivePrefs as IP with (NoLock) Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IP.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentivePrefs", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_IncentivePrefTiers
        Common.QueryStr = "select IPT.IncentivePrefTiersID, IPT.IncentivePrefsID, IPT.TierLevel, IPT.ValueComboTypeID " & _
                  "  FROM CPE_ST_IncentivePrefTiers as IPT with (NoLock) Inner Join CPE_ST_IncentivePrefs as IP on IPT.IncentivePrefsID=IP.IncentivePrefsID " & _
                  "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IP.RewardOptionID;"
        OutStr = OutStr & Construct_Table("IncentivePrefTiers", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'CPE_IncentivePrefTierValues
        Common.QueryStr = "select IPTV.PKID, IPTV.IncentivePrefTiersID, IPTV.Value, IPTV.OperatorTypeID, ValueTypeID, DateOperatorTypeID, ValueModifier " & _
                  "  FROM CPE_ST_IncentivePrefTierValues as IPTV with (NoLock) Inner Join CPE_ST_IncentivePrefTiers as IPT with (NoLock) on IPTV.IncentivePrefTiersID=IPT.IncentivePrefTiersID" & _
                  "  Inner Join CPE_ST_IncentivePrefs as IP on IPT.IncentivePrefsID=IP.IncentivePrefsID " & _
                  "  Inner Join #ActiveROIDs as AR on AR.RewardOptionID=IP.RewardOptionID;"
        Construct_Table("IncentivePrefTierValues", 5, DelimChar, LocalServerID, LocationID, "LRT")

        If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
            Common.QueryStr = "select ChannelID, Name, Installed, IssuanceDisabled from Channels where ChannelID=1;"
            Construct_Table("Channels", 5, DelimChar, LocalServerID, LocationID, "PMRT")

            Common.QueryStr = "select P.PreferenceID, P.Name, P.DataTypeID, P.MultiValue, " & _
                              "(case when P.IssuanceDisabled=1 Then 1 when isnull(PCID.PKID, 0)>0 then 1 Else 0 END) as IssuanceDisabled " & _
                              "from Preferences as P with (NoLock) Left Join PrefChannelsIssuanceDisabled as PCID with (NoLock) on PCID.PreferenceID=P.PreferenceID and PCID.ChannelID=1 " & _
                              "Inner Join PreferenceChannels as PC with (NoLock) on PC.PreferenceID=P.PreferenceID " & _
                              "where PC.ChannelID=1 and P.Deleted = 0;"
            Construct_Table("Preferences", 5, DelimChar, LocalServerID, LocationID, "PMRT")

            Common.QueryStr = "select PDV.PreferenceID, PDV.DefaultValue " & _
                              "from PrefDefaultValues as PDV with (NoLock) Inner Join PreferenceChannels as PC with (NoLock) on PC.PreferenceID=PDV.PreferenceID " & _
                              "Inner Join Preferences as P with (NoLock) on P.PreferenceID=PDV.PreferenceID " & _
                              "where PC.ChannelID=1 and P.Deleted = 0;"
            Construct_Table("PrefDefaultValues", 5, DelimChar, LocalServerID, LocationID, "PMRT")

            Common.QueryStr = "select PKID,PC.PreferenceID,ChannelID,ExternalID from PreferenceChannels as PC  " & _
                              "Inner Join Preferences as P with (NoLock) on P.PreferenceID=PC.PreferenceID " & _
                              "where PC.ChannelID=1 and P.Deleted = 0;"
            Construct_Table("PreferenceChannels", 5, DelimChar, LocalServerID, LocationID, "PMRT")

            'MetaPrefs
            Common.QueryStr = "select PKID, Name, PreferenceID from MetaPrefs;"
            Construct_Table("MetaPrefs", 5, DelimChar, LocalServerID, LocationID, "PMRT")
        End If

        'CardTypes
        Common.QueryStr = "select CardTypeID, Description, CustTypeID, ExtCardTypeID, PaddingLength, NumericOnly  " & _
                          "  FROM CardTypes with (NoLock);"
        Construct_Table("CardTypes", 5, DelimChar, LocalServerID, LocationID, "LXS")

        'Predefined Trigger Code Messages
        Common.QueryStr = "select ReasonFlag, Description, languageID from PredefinedTriggerCodeMessages with (NoLock) where Deleted = 0;"
        Construct_Table("PredefinedTriggerCodeMessages", 5, DelimChar, LocalServerID, LocationID, "LRT")
	
        If Common.Fetch_UE_SystemOption(175) = "1" Then
            Common.QueryStr = "select pgo.IncentiveID, PromoCategoryID from promogridoffers pgo inner join UE_IncentiveLocationsView loc on pgo.IncentiveID = loc.IncentiveID and loc.locationid = " & LocationID
            Construct_Table("PromoGridOffers", 5, DelimChar, LocalServerID, LocationID, "LRT")
        End If

        ' UEPriorityExclusions
        If (Common.IsEngineInstalled(9) AndAlso (Common.Fetch_UE_SystemOption(146) = "1")) Then
            Common.QueryStr = "select (100 - PriorityID) as PriorityID,ExcludedBy = CASE ExcludedBy WHEN -1 THEN -1 ELSE (100 - ExcludedBy) END from UE_PriorityExclusions with (NoLock);"
        Else
            Common.QueryStr = "select PriorityID, ExcludedBy " & _
                   "  FROM UE_PriorityExclusions with (NoLock);"
        End If

        Construct_Table("PriorityExclusions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        'UE_SystemOptions
        Common.QueryStr = "select OptionID, OptionName, OptionValue from UE_SystemOptions with (NoLock);"
        Construct_Table("UE_SystemOptions", 1, DelimChar, LocalServerID, LocationID, "LRT")

        If OperateAtEnterprise Then
            'send the LocationGroups table
            Common.QueryStr = "select LocationGroupID, AllLocations from LocationGroups where Deleted=0;"
            Construct_Table("LocationGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")

            'send the OfferLocations table
            Common.QueryStr = "select PKID, OfferID, LocationGroupID, Excluded from OfferLocations with (NoLock) where Deleted=0;"
            Construct_Table("OfferLocations", 5, DelimChar, LocalServerID, LocationID, "LRT")

            'send the LocGroupItems table
            Common.QueryStr = "select PKID, LocationGroupID, LocationID from LocGroupItems with (NoLock) where Deleted=0;"
            Construct_Table("LocGroupItems", 5, DelimChar, LocalServerID, LocationID, "LRT")
        End If

        'MutualExclusionGroups
        Common.QueryStr = "select MutualExclusionGroupID, Name, ItemLevel " & _
                          " FROM MutualExclusionGroups with (NoLock) " & _
                          " WHERE Deleted=0;"
        Construct_Table("MutualExclusionGroups", 5, DelimChar, LocalServerID, LocationID, "LRT")

        'MutualExclusionGroupOffers
        Common.QueryStr = "select MutualExclusionGroupID, OfferID " & _
                          " FROM MutualExclusionGroupOffers with (NoLock)"
        Construct_Table("MutualExclusionGroupOffers", 5, DelimChar, LocalServerID, LocationID, "LRT")

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
    MaxIPLs = Common.Extract_Val(Common.Fetch_UE_SystemOption(45))
    'fetch the IPL Window time
    WindowMinutes = Common.Extract_Val(Common.Fetch_UE_SystemOption(46))
    'fetch the IPL Runaway time
    IPLRunawayTime = Common.Extract_Val(Common.Fetch_UE_SystemOption(47))

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
  Dim ZipOutput As Boolean
  Dim Mode As String
  Dim HistoryStartTime As DateTime
  Dim IPLTypeID As Integer
  Dim IPLStartResponse As String
  Dim IPLSessionID As String
  Dim BannerID As Integer
  Dim LSVerParts() As String

  IPLTypeID = 1
  CurrentRequest.Resolver.AppName = "IPL-Offers"
  ApplicationName = "IPL-Offers"
  ApplicationExtension = ".aspx"
  Common.AppName = ApplicationName & ApplicationExtension
  Response.Expires = 0
  On Error GoTo ErrorTrap

  StartTime = Microsoft.VisualBasic.DateAndTime.Timer

  Common.Open_LogixXS()
  Common.Open_LogixWH()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)
  If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    Common.Open_PrefManRT()
  End If
  LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
  LogFile = "UE-IncentiveFetchLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
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

  Common.Write_Log(LogFile, "---------------------------------------------------------------------------")

  ZipOutput = True
  Response.ContentType = "application/x-gzip"

  Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, Request.UserHostAddress)

  If LocationID = "0" Then
    Common.Write_Log(LogFile, Common.AppName & "    Invalid Serial Number:" & LocalServerID & " from  MacAddress: " & MacAddress & " IP:" & LocalServerIP & "  Process running on server:" & Environment.MachineName, True)
    Send_Response_Header(ApplicationName & " - Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocationID, 9)) Then
    'the location calling TransDownload is not associated with the UE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than UE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other than UE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
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
        Common.QueryStr = "update UE_IncentiveDLBuffer with (RowLock) set WaitingACK=2 where WaitingACK>=0 and WaitingACK<2 and LocalServerID=" & LocalServerID & ";"
        Common.LRT_Execute()
        'this local server no longer needs to IPL because it is performing one right now
        Common.QueryStr = "Update LocalServers with (RowLock) set MustIPL=0 where LocalServerID='" & LocalServerID & "';"
        Common.LRT_Execute()
        'TextData = ""

        'since we are doing an IPL, get rid of any locks that may exist for this LocationID
                Common.QueryStr = "delete from CustomerLock where LocationID=" & LocationID & ";"
        Common.LXS_Execute()

        Common.Write_Log(LogFile, "Removing buffered TransDownload data")
        Common.QueryStr = "dbo.pc_CPE_Gen_Purge_Output_byLoc"
        Common.Open_LXSsp()
        Common.LXSsp.CommandTimeout = 1200
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
        Common.LXSsp.ExecuteNonQuery()
        Common.Close_LXSsp()
        TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished removing buffered TransDownload data.  Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")

        GZStream = New GZipStream(Response.OutputStream, CompressionMode.Compress, True)
        Construct_Output(LocalServerID, LocationID)

        TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Finished All Queries=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        If GZStream IsNot Nothing Then
          GZStream.Close()
          GZStream.Dispose()
          GZStream = Nothing
        End If
        TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "GZip stream closed.  Flushing final records ... " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")
        FlushStartTime = Microsoft.VisualBasic.DateAndTime.Timer
        Response.Flush()
        FlushTime = FlushTime + (Microsoft.VisualBasic.DateAndTime.Timer - FlushStartTime)
        TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
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
        Common.LRTsp.Parameters.Add("@CompressedSize", SqlDbType.BigInt).Value = 0 'unknown
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
  TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
  Common.Write_Log(LogFile, "Closing database connections - Total elapsed time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".000") & "(sec)")

  Common.Close_LogixRT()
  Common.Close_LogixXS()
  Common.Close_LogixWH()
  If Common.IsIntegrationInstalled(Copient.CommonInc.Integrations.PREFERENCE_MANAGER) Then
    Common.Close_PrefManRT()
  End If

  TotalTime = Microsoft.VisualBasic.DateAndTime.Timer - StartTime
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