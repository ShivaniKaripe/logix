<%@ WebService Language="VB" Class="Service" %>

Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.IO

Imports Copient.CommonInc
Imports Copient.IdNotFoundException
Imports Copient
Imports Copient.CryptLib

<WebService(Namespace:="http://www.copienttech.com/CustomerInquiry/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService
    ' version:7.3.1.138972.Official Build (SUSDAY10202)
    ' $Id: CustomerInquiry.asmx 126078 2018-07-10 06:59:32Z ma185300 $
    Private MyCommon As New Copient.CommonInc
    Private LogFile As String = "CustomerInquiryWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private MyCryptlib As New Copient.CryptLib
    Public Enum StatusCodes As Integer
        SUCCESS = 0
        INVALID_GUID = 1
        INVALID_CUSTOMERID = 2
        INVALID_CUSTOMERTYPEID = 3
        INVALID_CUSTOMERGROUPID = 4
        INVALID_MODE = 5
        NOTFOUND_CUSTOMER = 6
        NOTFOUND_HOUSEHOLD = 7
        NOTFOUND_CAM = 8
        FAILED_OPTIN = 9
        FAILED_OPTOUT = 10
        FAILED_BALANCE_LOOKUP = 11
        INVALID_POINTS_PROGRAM = 12
        FAILED_CUSTOMER_NOTE = 13
        NO_ACTIVITY_FOUND_FOR_SESSION_ID = 14
        NON_POSITIVE_SV_ADJUST_AMOUNT = 15
        INVALID_SV_MULITPLE = 16
        INVALID_STORED_VALUE_PROGRAM = 17
        INVALID_CRITERIA_XML = 18
        ADJUSTMENT_FAILED = 19
        INVALID_AMOUNT = 20
        INVALID_XML_DOCUMENT = 21
        MALFORMED_XML_FOR_OPERATION = 22
        OPERATION_TAG_LIMIT_EXCEEDED = 23
        CUST_NOMATCH_PASSWORD = 24
        CUST_MULTIPLE_EMAIL = 25
        CUST_INVALID_PASSWORD = 26
        INVALID_ADMINID = 27
        UNABLE_TO_DISPLAY_YTDSAVINGS = 28
        INVALID_ATTRIBUTETYPE = 29
        INVALID_ATTRIBUTEVALUE = 30
        INVALID_LOCATIONCODE = 31
        INVALID_STARTDATE = 32
        INVALID_ENDDATE = 33
        INVALID_NOTE = 34
        NOTFOUND_RECORDS = 35
        INVALID_SORTORDER = 36
        FAILED_UPDATE_SIGN = 37
        INVALID_POSDATETIME = 38
        INVALID_EMAILID = 39
        INVALID_PROGRAMID = 40
        INVALID_PROMOVARID = 41
        NOTFOUND_PROGRAMID = 42
        NOTFOUND_PROMOVARID = 43
        PROGRAMID_PROMOVARID_MISMATCH = 44
        INVALID_INCLUDEPENDING = 45
        APPLICATION_EXCEPTION = 9999
        INVALID_ROWNUM = 46 'added
        INVALID_SOURCEID = 47
        INVALID_TYPEID = 48
        INVALID_REASONID = 49
        INVALID_REASONTEXT = 50
        PROVIDE_PROGRAMID_OR_PROMOVARID = 51
    End Enum

    Public Const OPERATION_TAG_LIMIT As Integer = 1
    <WebMethod()> _
    Public Function GetTransactionTotalAmount(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String) As DataSet
        Dim iCardTypeID As Integer = -1
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        Return _GetTransactionTotalAmount(GUID, ExtCardID, iCardTypeID)
    End Function

    Private Function _GetTransactionTotalAmount(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("TransactionTotal")
        Dim dtStatus As DataTable
        Dim dtTransactionTotal As DataTable = Nothing
        Dim dr, row As DataRow
        Dim dt As DataTable
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim LogixTransNum As String
        Dim TransactionTotal As Decimal
        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "GetCustomerPointsHistory_SW") Then
                ' Lookup the customer
                If IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then

                    'Ignore
                End If
                RetCode = RetCode
                RetMsg = RetMsg

                If RetCode = StatusCodes.SUCCESS Then
                    'Create a new datatable to hold the results we'll be assembling
                    dtTransactionTotal = New DataTable("PointsHistory")
                    dtTransactionTotal.Columns.Add("LogixTransNum", System.Type.GetType("System.String"))
                    dtTransactionTotal.Columns.Add("TransactionTotal", System.Type.GetType("System.Decimal"))
                    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)

                    MyCommon.QueryStr = "Select TransTotal,LogixTransNum from TransHist" & _
                                        " with (NoLock)where CustomerPrimaryExtID= @ExtCardID and CustomerTypeID = @CardTypeID "
                    'No Encryption in WH DB
                    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = ExtCardID
                    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    dt = MyCommon.ExecuteQuery(DataBases.LogixWH)
                    If dt.Rows.Count > 0 Then
                        For Each dr In dt.Rows
                            LogixTransNum = dr.Item("LogixTransNum")
                            TransactionTotal = MyCommon.NZ(dr.Item("TransTotal"), 0.0)

                            row = dtTransactionTotal.NewRow()
                            row.Item("LogixTransNum") = LogixTransNum
                            row.Item("TransactionTotal") = TransactionTotal
                            dtTransactionTotal.Rows.Add(row)
                        Next
                    End If

                    If dtTransactionTotal.Rows.Count > 0 Then
                        dtTransactionTotal.AcceptChanges()
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Success."
                        dtStatus.Rows.Add(row)
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Records Not Found"
                        dtStatus.Rows.Add(row)
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    ResultDataSet.Tables.Add(dtStatus.Copy())
                End If
                If dtTransactionTotal IsNot Nothing Then ResultDataSet.Tables.Add(dtTransactionTotal)

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return ResultDataSet
    End Function

    <WebMethod()> _
    Public Function GetCustomerPointsHistory(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, _
                                        ByVal StartDate As String, ByVal EndDate As String) As DataSet
        Dim iCardTypeID As Integer = -1
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        If IsValidDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If IsValidDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)
        Return _GetCustomerPointsHistory(GUID, ExtCardID, iCardTypeID, sStartDate, sEndDate)
    End Function

    Private Function _GetCustomerPointsHistory(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, _
                                                ByVal StartDate As Date, ByVal EndDate As Date) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("PointHistory")
        Dim dtStatus As DataTable
        Dim dtPointsHistory As DataTable = Nothing
        Dim row, dr, drTransRedemption As DataRow
        Dim dt, dtTransRedemption As DataTable
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
        Const StoreLocationType As Integer = 1
        Const AdjustLocationID As Long = -9
        Dim Programid As Integer
        Dim LocId As Long
        Dim ProgramName As String = "null"
        Dim TransactionDate As Date
        Dim TransDate As String
        Dim LogixTransnum As String
        Dim extlocationcode As String = String.Empty
        Dim SourceID As Integer = 0
        Dim TypeID As Integer = 0
        Dim ReasonID As Integer = 0
        Dim SourceName As String = "null"
        Dim TypeName As String = "null"
        Dim ReasonDesc As String = "null"

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()

            If IsValidGUID(GUID, "GetCustomerPointsHistory") Then
                ' Lookup the customer
                If IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then
                    'Ignore
                End If
                RetCode = RetCode
                RetMsg = RetMsg

                If (StartDate = "01-01-1900") Then
                    'Bad Start Date
                    RetCode = StatusCodes.INVALID_STARTDATE
                    RetMsg = "Failure: Invalid StartDate"
                ElseIf (EndDate = "01-01-1900") Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_ENDDATE
                    RetMsg = "Failure: Invalid EndDate"
                End If
                If RetCode = StatusCodes.SUCCESS Then

                    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                    CustomerPK = LookupCustomerPK(ExtCardID, CardTypeID, RetCode, RetMsg)
                    If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                        'Create a new datatable to hold the results we'll be assembling
                        dtPointsHistory = New DataTable("PointsHistory")
                        dtPointsHistory.Columns.Add("TransactionDate", System.Type.GetType("System.String"))
                        dtPointsHistory.Columns.Add("ProgramID", System.Type.GetType("System.Int32"))
                        dtPointsHistory.Columns.Add("AdjustmentAmount", System.Type.GetType("System.Int64"))
                        dtPointsHistory.Columns.Add("LocationID", System.Type.GetType("System.Int64"))
                        dtPointsHistory.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                        dtPointsHistory.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
                        dtPointsHistory.Columns.Add("AdjustmentSourceName", System.Type.GetType("System.String"))
                        dtPointsHistory.Columns.Add("AdjustmentTypeName", System.Type.GetType("System.String"))
                        dtPointsHistory.Columns.Add("AdjustmentReasonDesc", System.Type.GetType("System.String"))
                        dtPointsHistory.Columns.Add("AdjustmentReasonText", System.Type.GetType("System.String"))

                        MyCommon.QueryStr = "Select DISTINCT ProgramID,AdjAmount,LocationID,LogixTransNum, " & _
                        "AdjustmentSourceID, AdjustmentTypeID, AdjustmentReasonID, AdjustmentReasonText from " & _
                                           " PointsHistory WITH (NOLOCK) where customerpk = @CustomerPK"
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.Int).Value = CustomerPK
                        dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dt.Rows.Count > 0 Then
                            For Each dr In dt.Rows
                                Programid = dr.Item("ProgramID")
                                LogixTransnum = dr.Item("LogixTransNum")
                                LocId = dr.Item("LocationID")

                                If GetExtLocationCode(MyCommon, LocId).Rows.Count > 0 Then
                                    MyCommon.QueryStr = "select ProgramName from PointsPrograms WITH (NOLOCK) where ProgramID = @Programid"
                                    MyCommon.DBParameters.Add("@Programid", SqlDbType.Int).Value = Programid
                                    Dim dtProgramName As New DataTable
                                    dtProgramName = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                    If dtProgramName.Rows.Count > 0 Then
                                        ProgramName = MyCommon.NZ(dtProgramName.Rows(0).Item("ProgramName"), "")
                                    End If

                                    SourceID = MyCommon.NZ(dr.Item("AdjustmentSourceID"), 0)
                                    If (SourceID > 0) Then
                                        MyCommon.QueryStr = "select Name from AdjustmentSources WITH (NOLOCK) where AdjustmentSourceID= @SourceID "
                                        MyCommon.DBParameters.Add("@SourceID", SqlDbType.Int).Value = SourceID
                                        Dim dtSourceName As New DataTable
                                        dtSourceName = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                        If (dtSourceName.Rows.Count > 0) Then
                                            SourceName = MyCommon.NZ(dtSourceName.Rows(0).Item("Name"), "")
                                        End If
                                    End If
                                    TypeID = MyCommon.NZ(dr.Item("AdjustmentTypeID"), 0)
                                    If (SourceID > 0) Then
                                        MyCommon.QueryStr = "select Name from AdjustmentTypes WITH (NOLOCK) where AdjustmentTypeID= @TypeID "
                                        MyCommon.DBParameters.Add("@TypeID", SqlDbType.SmallInt).Value = TypeID
                                        Dim dtTypeName As New DataTable
                                        dtTypeName = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                        If (dtTypeName.Rows.Count > 0) Then
                                            TypeName = MyCommon.NZ(dtTypeName.Rows(0).Item("Name"), "")
                                        End If
                                    End If
                                    ReasonID = MyCommon.NZ(dr.Item("AdjustmentReasonID"), 0)
                                    If (ReasonID > 0) Then
                                        MyCommon.QueryStr = "select Description from AdjustmentReasons WITH (NOLOCK) where AdjustmentReasonID= @ReasonID "
                                        MyCommon.DBParameters.Add("@ReasonID", SqlDbType.Int).Value = ReasonID
                                        Dim dtReasonDesc As New DataTable
                                        dtReasonDesc = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                        If (dtReasonDesc.Rows.Count > 0) Then
                                            ReasonDesc = MyCommon.NZ(dtReasonDesc.Rows(0).Item("Description"), "")
                                        End If
                                    End If

                                    MyCommon.QueryStr = "select Distinct TransDate from TransRedemption WITH (NOLOCK) where LogixTransNum= @LogixTransnum "
                                    MyCommon.DBParameters.Add("@LogixTransnum", SqlDbType.Char).Value = LogixTransnum
                                    dtTransRedemption = MyCommon.ExecuteQuery(DataBases.LogixWH)
                                    If dtTransRedemption.Rows.Count > 0 Then
                                        For Each drTransRedemption In dtTransRedemption.Rows
                                            TransactionDate = MyCommon.NZ(drTransRedemption.Item("TransDate"), "01-01-1900")
                                            MyCommon.QueryStr = " select Distinct TransDate from TransRedemption WITH (NOLOCK) " & _
                                                                " where ( @TransactionDate >= @StartDate) " & _
                                                                " AND ( @TransactionDate <= @EndDate)"
                                            MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                                            MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                                            MyCommon.DBParameters.Add("@TransactionDate", SqlDbType.DateTime).Value = TransactionDate

                                            If MyCommon.ExecuteQuery(DataBases.LogixWH).Rows.Count > 0 Then
                                                TransDate = TransactionDate.ToString("yyyy-MM-dd HH:mm:ss")
                                                MyCommon.QueryStr = " select ExtLocationCode from Locations WITH (NOLOCK) where " & _
                                                                    " LocationID = @LocationID And LocationTypeID = @LocationTypeID "
                                                MyCommon.DBParameters.Add("@LocationID", SqlDbType.BigInt).Value = LocId
                                                MyCommon.DBParameters.Add("@LocationTypeID", SqlDbType.Int).Value = StoreLocationType
                                                Dim dtExtLocation As New DataTable
                                                dtExtLocation = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                                If (dtExtLocation.Rows.Count > 0) Then
                                                    extlocationcode = MyCommon.NZ(dtExtLocation.Rows(0).Item("ExtLocationCode"), "")
                                                End If
                                                row = dtPointsHistory.NewRow()
                                                row.Item("TransactionDate") = TransDate
                                                row.Item("ProgramID") = Programid
                                                row.Item("AdjustmentAmount") = dr.Item("AdjAmount")
                                                row.Item("ProgramName") = ProgramName
                                                row.Item("LocationID") = LocId
                                                row.Item("ExtLocationCode") = extlocationcode
                                                row.Item("AdjustmentSourceName") = SourceName
                                                row.Item("AdjustmentTypeName") = TypeName
                                                row.Item("AdjustmentReasonDesc") = ReasonDesc
                                                row.Item("AdjustmentReasonText") = MyCommon.NZ(dr.Item("AdjustmentReasonText"), "")
                                                dtPointsHistory.Rows.Add(row)
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If

                        MyCommon.QueryStr = "Select DISTINCT LastUpdate,ProgramID,AdjAmount,LocationID, " & _
                        "AdjustmentSourceID, AdjustmentTypeID, AdjustmentReasonID, AdjustmentReasonText from " & _
                               " PointsHistory where CustomerPK = @CustomerPK and " & _
                               " LocationID = @AdjustLocationID and (LastUpdate >= @StartDate) AND " & _
                               " (LastUpdate <= @EndDate)"
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.Int).Value = CustomerPK
                        MyCommon.DBParameters.Add("@AdjustLocationID", SqlDbType.BigInt).Value = AdjustLocationID
                        MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                        MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                        dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dt.Rows.Count > 0 Then
                            For Each dr In dt.Rows
                                Programid = dr.Item("ProgramID")
                                MyCommon.QueryStr = "select ProgramName from PointsPrograms WITH (NOLOCK) where ProgramID= @Programid "
                                MyCommon.DBParameters.Add("@Programid", SqlDbType.BigInt).Value = Programid
                                Dim dtProgramName As New DataTable
                                dtProgramName = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                If (dtProgramName.Rows.Count > 0) Then
                                    ProgramName = MyCommon.NZ(dtProgramName.Rows(0).Item("ProgramName"), "")
                                End If

                                SourceID = MyCommon.NZ(dr.Item("AdjustmentSourceID"), 0)
                                If (SourceID > 0) Then
                                    MyCommon.QueryStr = "select Name from AdjustmentSources WITH (NOLOCK) where AdjustmentSourceID= @SourceID "
                                    MyCommon.DBParameters.Add("@SourceID", SqlDbType.Int).Value = SourceID
                                    Dim dtSourceName As New DataTable
                                    dtSourceName = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                    If (dtSourceName.Rows.Count > 0) Then
                                        SourceName = MyCommon.NZ(dtSourceName.Rows(0).Item("Name"), "")
                                    End If
                                End If
                                TypeID = MyCommon.NZ(dr.Item("AdjustmentTypeID"), 0)
                                If (TypeID > 0) Then
                                    MyCommon.QueryStr = "select Name from AdjustmentTypes WITH (NOLOCK) where AdjustmentTypeID= @TypeID "
                                    MyCommon.DBParameters.Add("@TypeID", SqlDbType.SmallInt).Value = TypeID
                                    Dim dtTypeName As New DataTable
                                    dtTypeName = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                    If (dtTypeName.Rows.Count > 0) Then
                                        TypeName = MyCommon.NZ(dtTypeName.Rows(0).Item("Name"), "")
                                    End If
                                End If
                                ReasonID = MyCommon.NZ(dr.Item("AdjustmentReasonID"), 0)
                                If (ReasonID > 0) Then
                                    MyCommon.QueryStr = "select Description from AdjustmentReasons WITH (NOLOCK) where AdjustmentReasonID= @ReasonID "
                                    MyCommon.DBParameters.Add("@ReasonID", SqlDbType.Int).Value = ReasonID
                                    Dim dtReasonDesc As New DataTable
                                    dtReasonDesc = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                    If (dtReasonDesc.Rows.Count > 0) Then
                                        ReasonDesc = MyCommon.NZ(dtReasonDesc.Rows(0).Item("Description"), "")
                                    End If
                                End If

                                row = dtPointsHistory.NewRow()
                                row.Item("TransactionDate") = dr.Item("LastUpdate")
                                row.Item("ProgramID") = Programid
                                row.Item("AdjustmentAmount") = dr.Item("AdjAmount")
                                row.Item("ProgramName") = ProgramName
                                row.Item("LocationID") = dr.Item("LocationID")
                                row.Item("ExtLocationCode") = AdjustLocationID
                                row.Item("AdjustmentSourceName") = SourceName
                                row.Item("AdjustmentTypeName") = TypeName
                                row.Item("AdjustmentReasonDesc") = ReasonDesc
                                row.Item("AdjustmentReasonText") = MyCommon.NZ(dr.Item("AdjustmentReasonText"), "")
                                dtPointsHistory.Rows.Add(row)
                            Next
                        End If

                        If dtPointsHistory.Rows.Count > 0 Then
                            dtPointsHistory.AcceptChanges()
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success."
                            dtStatus.Rows.Add(row)
                        Else
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Records Not Found"
                            dtStatus.Rows.Add(row)
                        End If
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = RetCode
                        row.Item("Description") = "CardID: " & ExtCardID & " with CardTypeID: " & CardTypeID & " not found."
                        dtStatus.Rows.Add(row)
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    ResultDataSet.Tables.Add(dtStatus.Copy())
                End If
                If dtPointsHistory IsNot Nothing Then ResultDataSet.Tables.Add(dtPointsHistory)

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
        End Try

        Return ResultDataSet
    End Function
    <WebMethod()> _
    Public Function GetCustomerOfferBalances(ByVal GUID As String, ByVal CardPrimaryExtID As String, ByVal CardTypeID As String) As DataSet
        Dim iCardTypeID As Integer = -1
        Try
            iCardTypeID = CInt(CardTypeID)
        Catch ex As Exception
            iCardTypeID = -1
        End Try
        Return _GetCustomerOfferBalances(GUID, CardPrimaryExtID, iCardTypeID)
    End Function

    Private Function _GetCustomerOfferBalances(ByVal GUID As String, ByVal CardPrimaryExtID As String, ByVal CardTypeID As Integer) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("PointHistory")
        Dim dtStatus As System.Data.DataTable
        Dim dtOfferBalance As DataTable = Nothing
        Dim dt As DataTable
        Dim row, dr As DataRow
        Dim OfferID As Long
        Dim P3DistQtyLimit As String = "Null"
        Dim objErrorResponse As CardValidationResponse

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()

            If Not IsValidGUID(GUID, "GetCustomerOfferBalances") Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)

            ElseIf (MyCommon.AllowToProcessCustomerCard(CardPrimaryExtID, CardTypeID, objErrorResponse) = False) Then
                If CardPrimaryExtID Is Nothing OrElse CardPrimaryExtID.Trim = "" Then
                    'Bad customer Id
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                    row.Item("Description") = "Failure: CardPrimaryExtID is not provided"
                    dtStatus.Rows.Add(row)
                ElseIf Not String.IsNullOrEmpty(CardPrimaryExtID) AndAlso Not CardTypeID = -1 Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                    row.Item("Description") = "CardPrimaryExtID: " & CardPrimaryExtID & " with CustomerTypeID: " & CardTypeID & " not found."
                    dtStatus.Rows.Add(row)
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                    row.Item("Description") = "Failure: CustomerTypeID is not provided"
                    dtStatus.Rows.Add(row)
                End If
            Else
                'Create a new datatable to hold the results we'll be assembling
                dtOfferBalance = New DataTable("OfferBalance")
                dtOfferBalance.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                dtOfferBalance.Columns.Add("RedemptionCount", System.Type.GetType("System.Int32"))
                dtOfferBalance.Columns.Add("RedemptionAmount", System.Type.GetType("System.Decimal"))
                dtOfferBalance.Columns.Add("TransactionDate", System.Type.GetType("System.String"))
                dtOfferBalance.Columns.Add("P3DistQtyLimit", System.Type.GetType("System.String"))
                CardPrimaryExtID = MyCommon.Pad_ExtCardID(CardPrimaryExtID, CardTypeID)

                MyCommon.QueryStr = " SELECT DISTINCT offerid,SUM(RedemptionCount) RedemptionCount,SUM(RedemptionAmount) RedemptionAmount,MAX(TransDate) TransactionDate " & _
                                   " From TransRedemptionView with (NoLock) where CustomerPrimaryExtID = @CardPrimaryExtID AND " & _
                                   " PresentedCardTypeID = @PresentedCardTypeID GROUP BY OfferID"
                MyCommon.DBParameters.Add("@CardPrimaryExtID", SqlDbType.NVarChar).Value = CardPrimaryExtID.ConvertBlankIfNothing()
                MyCommon.DBParameters.Add("@PresentedCardTypeID", SqlDbType.Int).Value = CardTypeID
                dt = MyCommon.ExecuteQuery(DataBases.LogixWH)
                If dt.Rows.Count > 0 Then
                    For Each dr In dt.Rows
                        OfferID = dr.Item("offerid")
                        MyCommon.QueryStr = "select P3DistQtyLimit from CPE_Incentives with (NoLock) where IncentiveID= @OfferID "
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                        Dim dtDistQtyLimit As New DataTable
                        dtDistQtyLimit = MyCommon.ExecuteQuery(DataBases.LogixRT)
                        If dtDistQtyLimit.Rows.Count > 0 Then
                            P3DistQtyLimit = MyCommon.NZ(dtDistQtyLimit.Rows(0).Item("P3DistQtyLimit"), 0)
                        End If
                        row = dtOfferBalance.NewRow()
                        row.Item("OfferID") = OfferID
                        row.Item("RedemptionCount") = dr.Item("RedemptionCount")
                        row.Item("RedemptionAmount") = dr.Item("RedemptionAmount")
                        row.Item("TransactionDate") = dr.Item("TransactionDate")
                        row.Item("P3DistQtyLimit") = P3DistQtyLimit
                        dtOfferBalance.Rows.Add(row)
                    Next
                End If
                If dtOfferBalance.Rows.Count > 0 Then
                    dtOfferBalance.AcceptChanges()
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.SUCCESS
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.SUCCESS
                    row.Item("Description") = "Records Not Found."
                    dtStatus.Rows.Add(row)
                End If
            End If

            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If
            If dtOfferBalance IsNot Nothing Then ResultDataSet.Tables.Add(dtOfferBalance)

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
        End Try
        Return ResultDataSet
    End Function

    <WebMethod()> _
    Public Function GetClubBalancesByCustomerPK(ByVal GUID As String, ByVal CustomerPK As String) As DataSet
        Dim lCustomerPK As Long = -1
        Try
            lCustomerPK = Convert.ToInt64(CustomerPK)
        Catch ex As Exception
            lCustomerPK = -1
        End Try
        Return _GetClubBalancesByCustomerPK(GUID, lCustomerPK)
    End Function

    Private Function _GetClubBalancesByCustomerPK(ByVal GUID As String, ByVal CustomerPK As Long) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("ClubBalances")
        Dim dtStatus As System.Data.DataTable
        Dim dtClubBalance As DataTable = Nothing
        Dim row, dr, PointBalanceRow As DataRow
        Dim dt As DataTable
        Dim PointsProgram, Balance As Integer

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID, "GetClubBalancesByCustomerPK") Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
            ElseIf (CustomerPK = -1) Then
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                row.Item("Description") = "Failure: CustomerPK not provided"
                dtStatus.Rows.Add(row)
            Else
                'Create a new datatable to hold the results we'll be assembling
                dtClubBalance = New DataTable("OfferBalance")
                dtClubBalance.Columns.Add("Name", System.Type.GetType("System.String"))
                dtClubBalance.Columns.Add("Description", System.Type.GetType("System.String"))
                dtClubBalance.Columns.Add("PointsProgram", System.Type.GetType("System.Int32"))
                dtClubBalance.Columns.Add("Threshold", System.Type.GetType("System.Int32"))
                dtClubBalance.Columns.Add("Balance", System.Type.GetType("System.Int32"))

                MyCommon.QueryStr = "SELECT i.IncentiveName AS 'Name', i.Description, ip.ProgramID AS 'PointsProgram', ip.QtyForIncentive AS 'Threshold' " & _
                                    "FROM CPE_ST_Incentives AS i INNER JOIN CPE_ST_RewardOptions AS ro ON i.IncentiveID = ro.IncentiveID " & _
                                    "INNER JOIN CPE_ST_IncentivePointsGroups AS ip ON ro.RewardOptionID = ip.RewardOptionID WHERE CONVERT (DATE, GETDATE()) BETWEEN i.StartDate AND i.EndDate"
                dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
                If dt.Rows.Count > 0 Then
                    For Each dr In dt.Rows
                        PointsProgram = dr.Item("PointsProgram")
                        MyCommon.QueryStr = "Select cast(Amount as int) AS 'Balance' From Points Where ProgramID= @PointsProgram And CustomerPK = @CustomerPK "
                        MyCommon.DBParameters.Add("@PointsProgram", SqlDbType.BigInt).Value = PointsProgram
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        Dim dtBalance As New DataTable
                        dtBalance = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dtBalance.Rows.Count > 0 Then
                            For Each PointBalanceRow In dtBalance.Rows
                                Balance = MyCommon.NZ(PointBalanceRow.Item("Balance"), 0)
                                row = dtClubBalance.NewRow()
                                row.Item("Name") = dr.Item("Name")
                                row.Item("Description") = dr.Item("Description")
                                row.Item("PointsProgram") = PointsProgram
                                row.Item("Threshold") = dr.Item("Threshold")
                                row.Item("Balance") = Balance
                                dtClubBalance.Rows.Add(row)
                            Next
                        End If
                    Next
                End If

                If dtClubBalance.Rows.Count > 0 Then
                    dtClubBalance.AcceptChanges()
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.SUCCESS
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.SUCCESS
                    row.Item("Description") = "Records Not Found."
                    dtStatus.Rows.Add(row)
                End If
            End If

            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If
            If dtClubBalance IsNot Nothing Then ResultDataSet.Tables.Add(dtClubBalance)

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
        End Try
        Return ResultDataSet
    End Function

    <WebMethod()> _
    Public Function AddCustomerNotesByExtID(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, _
                                          ByVal NoteTypeID As String, ByVal AdminUserID As String, ByVal Note As String) As DataSet
        Dim iCardTypeID As Integer = -1
        Dim iNoteTypeID As Integer = -1
        Dim iAdminUserID As Integer = -1
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        If Not String.IsNullOrEmpty(NoteTypeID) AndAlso Not NoteTypeID.Trim = String.Empty AndAlso IsNumeric(NoteTypeID) Then iNoteTypeID = Convert.ToInt32(NoteTypeID)
        If Not String.IsNullOrEmpty(AdminUserID) AndAlso Not AdminUserID.Trim = String.Empty AndAlso IsNumeric(AdminUserID) Then iAdminUserID = Convert.ToInt32(AdminUserID)
        Return _AddCustomerNotesByExtID(GUID, ExtCardID, iCardTypeID, iNoteTypeID, iAdminUserID, Note)
    End Function

    Private Function _AddCustomerNotesByExtID(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, _
                                              ByVal NoteTypeID As Integer, ByVal AdminUserID As Integer, ByVal Note As String) As DataSet
        Dim ResultDataSet As New System.Data.DataSet
        Dim dtStatus As DataTable
        Dim dt As DataTable
        Dim row As DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim FirstName As String = ""
        Dim LastName As String = ""
        Dim ActivityText As String = ""

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "AddCustomerNotesByExtID") Then
                If IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then

                    'Ignore
                End If
                RetCode = RetCode
                RetMsg = RetMsg
                If (NoteTypeID = -1) Then
                    'Bad Start Date
                    RetCode = StatusCodes.INVALID_NOTE
                    RetMsg = "Failure: NoteTypeId is not provided"
                ElseIf Not IsValidCustNoteTypeID(NoteTypeID) Then
                    'Bad CustNoteTypeID
                    RetCode = StatusCodes.INVALID_NOTE
                    RetMsg = "Failure: Invalid Note Type Id"
                ElseIf (AdminUserID = -1) Then
                    'Bad AdminID
                    RetCode = StatusCodes.INVALID_ADMINID
                    RetMsg = "Failure: Invalid AdminUserId"
                ElseIf (Not IsValidAdminUserID(AdminUserID.ToString, RetMsg)) Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_ADMINID
                    RetMsg = "Failure: Invalid AdminUserID" & RetMsg
                ElseIf Note Is Nothing OrElse Note.Trim = "" Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_NOTE
                    RetMsg = "Failure: Note is not provided"
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)

                    CustomerPK = LookupCustomerPK(ExtCardID, CardTypeID, RetCode, RetMsg)
                    If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                        'Create a new datatable to hold the results we'll be assembling
                        MyCommon.QueryStr = "SELECT FirstName, LastName from AdminUsers WITH (NOLOCK)WHERE AdminUserID = @AdminUserID "
                        MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                        dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
                        If dt.Rows.Count > 0 Then
                            FirstName = MyCommon.NZ(dt.Rows(0).Item("FirstName"), "")
                            LastName = MyCommon.NZ(dt.Rows(0).Item("LastName"), "")
                        End If
                        MyCommon.QueryStr = "INSERT INTO CustomerNotes with (RowLock) (CustomerPK, NoteTypeID,AdminUserID, FirstName, LastName, CreatedDate, Note) " & _
                               " VALUES (@CustomerPK , @NoteTypeID ,@AdminUserID ,@FirstName, @LastName , getdate(), @Note)"
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        MyCommon.DBParameters.Add("@NoteTypeID", SqlDbType.SmallInt).Value = NoteTypeID
                        MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                        MyCommon.DBParameters.Add("@FirstName", SqlDbType.NVarChar).Value = FirstName
                        MyCommon.DBParameters.Add("@LastName", SqlDbType.NVarChar).Value = LastName
                        MyCommon.DBParameters.Add("@Note", SqlDbType.NVarChar).Value = MyCommon.Parse_Quotes(Note.ConvertBlankIfNothing())
                        MyCommon.ExecuteNonQuery(DataBases.LogixXS)
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Customer Notes added successfully"
                        dtStatus.Rows.Add(row)
                        ActivityText = Copient.PhraseLib.Lookup("history.customer-added-note", 1) & ": " & Note
                        If ActivityText.Length > 1000 Then
                            ActivityText = Left(ActivityText, 997) & "..."
                        End If
                        MyCommon.Activity_Log2(25, 8, CustomerPK, AdminUserID, ActivityText)
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = RetCode
                        row.Item("Description") = "CardID: " & ExtCardID & " with CardTypeID: " & CardTypeID & " not found."
                        dtStatus.Rows.Add(row)
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    ResultDataSet.Tables.Add(dtStatus.Copy())
                End If
            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return ResultDataSet
    End Function

    'function to check whether Sorted Order is asc or desc
    Private Function IsStringSortedOrder(ByVal strInputText As String) As Boolean
        Dim IsAlpha As Boolean = False
        If System.Text.RegularExpressions.Regex.IsMatch(strInputText, "^(asc|desc|ASC|Asc|DESC|Desc)$") Then
            IsAlpha = True
        Else
            IsAlpha = False
        End If
        Return IsAlpha
    End Function

    'function to check whether RowNumber is digit
    Private Function IsNumericRowNum(ByVal strInputText As String) As Boolean
        Dim IsAlpha As Boolean = False
        If System.Text.RegularExpressions.Regex.IsMatch(strInputText, "^[1-9][0-9]{1,4}$") Then
            IsAlpha = True
        Else
            IsAlpha = False
        End If
        Return IsAlpha
    End Function

    <WebMethod()> _
    Public Function GetCustomerPointsHistoryWithIssuance(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, _
                                                        ByVal StartDate As String, ByVal EndDate As String) As DataSet
        Dim iCardTypeID As Integer = -1 '0
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        If IsValidDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If IsValidDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)

        Return _GetCustomerPointsHistoryWithIssuance(GUID, ExtCardID, iCardTypeID, sStartDate, sEndDate)
    End Function

    Private Function _GetCustomerPointsHistoryWithIssuance(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, _
                                                            ByVal StartDate As Date, ByVal EndDate As Date) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("PointHistory")
        Dim dtStatus As DataTable
        Dim dtPointsHistoryWithIssuance As DataTable = Nothing
        Dim row, dr As DataRow
        Dim dt As DataTable
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim Programid As Integer
        Dim LocId As Long
        Dim ProgramName As String = "null"
        Dim extlocationcode As String
        Dim TableDatePart As Date = Date.Now
        ' Dim objErrorResponse As CardValidationResponse

        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()
            If IsValidGUID(GUID, "GetCustomerPointsHistoryWithIssuance") Then
                If IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then

                    'Ignore
                End If
                RetCode = RetCode
                RetMsg = RetMsg
                If (StartDate = "01-01-1900") Then
                    'Bad Start Date
                    RetCode = StatusCodes.INVALID_STARTDATE
                    RetMsg = "Failure: Invalid StartDate"
                ElseIf (EndDate = "01-01-1900") Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_ENDDATE
                    RetMsg = "Failure: Invalid EndDate"
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                    CustomerPK = LookupCustomerPK(ExtCardID, CardTypeID, RetCode, RetMsg)
                    If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                        'Create a new datatable to hold the results we'll be assembling
                        dtPointsHistoryWithIssuance = New DataTable("PointsHistory")
                        dtPointsHistoryWithIssuance.Columns.Add("POSTimeStamp", System.Type.GetType("System.String"))
                        dtPointsHistoryWithIssuance.Columns.Add("ProgramID", System.Type.GetType("System.Int32"))
                        dtPointsHistoryWithIssuance.Columns.Add("RewardQty", System.Type.GetType("System.Int32"))
                        dtPointsHistoryWithIssuance.Columns.Add("LocationID", System.Type.GetType("System.Int32"))
                        dtPointsHistoryWithIssuance.Columns.Add("Net", System.Type.GetType("System.Int32"))
                        dtPointsHistoryWithIssuance.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                        dtPointsHistoryWithIssuance.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
                        Dim strExtCardID As String = String.Empty
                        Dim strExtCardIDOriginal As String = String.Empty
                        MyCommon.QueryStr = " Select ExtCardID,ExtCardIDOriginal from CardIDs where ExtCardID = @ExtCardID"
                        MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                        dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dt.Rows.Count > 0 Then
                            strExtCardID = dt.Rows(0)("ExtCardID").ToString()
                            strExtCardIDOriginal = dt.Rows(0)("ExtCardIDOriginal").ToString()
                        End If
                        MyCommon.QueryStr = " SELECT IHT1.POSTimeStamp, IHT1.ProgramID, IHT1.RewardQty, IHT1.LocationID, CAST((ROUND(IHT2.Net, 2, 1) * 100) as INTEGER) as 'Net' " & _
                                            " FROM Issuance" & TableDatePart.ToString("yyyyMMdd") & " as IHT1 INNER JOIN " & _
                                            " Issuance" & TableDatePart.ToString("yyyyMMdd") & " as IHT2 ON IHT1.LogixTransNum = IHT2.LogixTransNum" & _
                                            " Where (IHT1.PrimaryExtID = @ExtCardID OR IHT1.PrimaryExtID = @ExtCardIDOriginal) AND IHT1.CardTypeID = @CardTypeID " & _
                                            " AND (IHT1.POSTimeStamp BETWEEN @StartDate AND @EndDate) " & _
                                            " AND IHT1.ProgramID > 0 AND IHT2.ProgramID = 0 AND IHT2.Void = 0"
                        
                        MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = strExtCardID
                        MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                        MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                        MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                        MyCommon.DBParameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar).Value = strExtCardIDOriginal
                        dt = MyCommon.ExecuteQuery(DataBases.LogixEX)
                        If dt.Rows.Count > 0 Then
                            For Each dr In dt.Rows
                                Programid = dr.Item("ProgramID")
                                LocId = dr.Item("LocationID")
                                Dim dtExtLocationCode As New DataTable
                                dtExtLocationCode = GetExtLocationCode(MyCommon, LocId)
                                If dtExtLocationCode.Rows.Count > 0 Then
                                    extlocationcode = MyCommon.NZ(dtExtLocationCode.Rows(0).Item("ExtLocationCode"), "")
                                    MyCommon.QueryStr = "select ProgramName from PointsPrograms WITH (NOLOCK) where ProgramID= @Programid "
                                    MyCommon.DBParameters.Add("@Programid", SqlDbType.BigInt).Value = Programid
                                    Dim dtProgramName As New DataTable
                                    dtProgramName = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                    If dtProgramName.Rows.Count > 0 Then
                                        ProgramName = MyCommon.NZ(dtProgramName.Rows(0).Item("ProgramName"), "")
                                    End If
                                    row = dtPointsHistoryWithIssuance.NewRow()
                                    row.Item("POSTimeStamp") = dr.Item("POSTimeStamp")
                                    row.Item("ProgramID") = Programid
                                    row.Item("RewardQty") = dr.Item("RewardQty")
                                    row.Item("LocationID") = LocId
                                    row.Item("Net") = dr.Item("Net")
                                    row.Item("ProgramName") = ProgramName
                                    row.Item("ExtLocationCode") = extlocationcode
                                    dtPointsHistoryWithIssuance.Rows.Add(row)
                                End If
                            Next
                        End If
                        MyCommon.QueryStr = "Select PH.POSTimeStamp, PH.ProgramID, PH.AdjAmount as 'RewardQty', PH.LocationID, cast(0 as INTEGER) as 'Net', " & _
                                           " cast(9999 as INTEGER) as 'ExtLocationCode' FROM PointsHistory as PH WITH (NOLOCK) " & _
                                           " INNER JOIN CardIDs as CID WITH(NOLOCK) ON PH.CustomerPK = CID.CustomerPK " & _
                                           " WHERE CID.ExtCardID = @ExtCardID AND CID.CardTypeID = @CardTypeID " & _
                                           " AND (PH.POSTimeStamp BETWEEN @StartDate AND @EndDate) " & _
                                           " AND cast(PH.LogixTransNum as nvarchar(MAX)) = '0' ORDER BY POSTimeStamp DESC"
                        MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                        MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                        MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                        MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                        dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dt.Rows.Count > 0 Then
                            For Each dr In dt.Rows
                                Programid = dr.Item("ProgramID")
                                MyCommon.QueryStr = "select ProgramName from PointsPrograms WITH (NOLOCK) where ProgramID= @Programid "
                                MyCommon.DBParameters.Add("@Programid", SqlDbType.BigInt).Value = Programid
                                Dim dtProgramName As DataTable
                                dtProgramName = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                If (dtProgramName.Rows.Count > 0) Then
                                    ProgramName = MyCommon.NZ(dtProgramName.Rows(0).Item("ProgramName"), "")
                                End If
                                row = dtPointsHistoryWithIssuance.NewRow()
                                row.Item("POSTimeStamp") = dr.Item("POSTimeStamp")
                                row.Item("ProgramID") = Programid
                                row.Item("RewardQty") = dr.Item("RewardQty")
                                row.Item("LocationID") = dr.Item("LocationID")
                                row.Item("Net") = dr.Item("Net")
                                row.Item("ProgramName") = ProgramName
                                row.Item("ExtLocationCode") = dr.Item("Extlocationcode")
                                dtPointsHistoryWithIssuance.Rows.Add(row)
                            Next
                        End If

                        If dtPointsHistoryWithIssuance.Rows.Count > 0 Then
                            dtPointsHistoryWithIssuance.AcceptChanges()
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success."
                            dtStatus.Rows.Add(row)
                        Else
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Records Not Found"
                            dtStatus.Rows.Add(row)
                        End If
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = RetCode
                        row.Item("Description") = "CardID: " & ExtCardID & " with CardTypeID: " & Convert.ToString(CardTypeID) & " not found."
                        dtStatus.Rows.Add(row)
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    ResultDataSet.Tables.Add(dtStatus.Copy())
                End If
                If dtPointsHistoryWithIssuance IsNot Nothing Then ResultDataSet.Tables.Add(dtPointsHistoryWithIssuance)

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())

            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            If MyCommon.LEXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixEX()
        End Try

        Return ResultDataSet
    End Function

    <WebMethod()> _
    Public Function GetCustomerPointsHistoryTransRedempt(ByVal GUID As String, ByVal ExtCardID As String, ByVal StartDate As String, ByVal EndDate As String) As DataSet
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If Not String.IsNullOrEmpty(StartDate) AndAlso Not StartDate.Trim = String.Empty AndAlso IsDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If Not String.IsNullOrEmpty(EndDate) AndAlso Not EndDate.Trim = String.Empty AndAlso IsDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)
        Return _GetCustomerPointsHistoryTransRedempt(GUID, ExtCardID, sStartDate, sEndDate)
    End Function

    <WebMethod()> _
    Public Function GetCustomerPointsHistoryTransRedempt_ByCardID(ByVal GUID As String, ByVal ExtCardID As String, ByVal StartDate As String, ByVal EndDate As String, ByVal CardTypeID As Integer) As DataSet
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If Not String.IsNullOrEmpty(StartDate) AndAlso Not StartDate.Trim = String.Empty AndAlso IsDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If Not String.IsNullOrEmpty(EndDate) AndAlso Not EndDate.Trim = String.Empty AndAlso IsDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)
        Return _GetCustomerPointsHistoryTransRedempt(GUID, ExtCardID, sStartDate, sEndDate, CardTypeID)
    End Function

    Private Function _GetCustomerPointsHistoryTransRedempt(ByVal GUID As String, ByVal ExtCardID As String, ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal CardTypeID As Integer = 0) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("PointHistory")
        Dim dtStatus As DataTable
        Dim dtPointsHistory_TransRedempt As DataTable = Nothing
        Dim row As DataRow
        Dim dr As DataRow = Nothing

        Dim dt As DataTable
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim Programid As Integer
        Dim LocId As Long
        Dim ProgramName As String = "null"
        Dim extlocationcode As String

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()

            If IsValidGUID(GUID, "GetCustomerPointsHistoryTransRedempt") Then
                ' Lookup the customer
                If ExtCardID Is Nothing OrElse ExtCardID.Trim = "" Then
                    'Bad customer ID
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Failure: ExtCardID is not provided"
                ElseIf (StartDate = "01-01-1900") Then
                    'Bad Start Date
                    RetCode = StatusCodes.INVALID_STARTDATE
                    RetMsg = "Failure: Invalid StartDate"
                ElseIf (EndDate = "01-01-1900") Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_ENDDATE
                    RetMsg = "Failure: Invalid EndDate"
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    'Create a new datatable to hold the results we'll be assembling

                    dtPointsHistory_TransRedempt = New DataTable("PointsHistory")
                    dtPointsHistory_TransRedempt.Columns.Add("POSTimeStamp", System.Type.GetType("System.String"))
                    dtPointsHistory_TransRedempt.Columns.Add("PointsProgramID", System.Type.GetType("System.Int32"))
                    dtPointsHistory_TransRedempt.Columns.Add("PointsAmount", System.Type.GetType("System.Int64"))
                    dtPointsHistory_TransRedempt.Columns.Add("LocationID", System.Type.GetType("System.Int64"))
                    dtPointsHistory_TransRedempt.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                    dtPointsHistory_TransRedempt.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))

                    ' If (CardTypeID <> -1) Then
                    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                    'End If
                    MyCommon.QueryStr = " SELECT TR.POSTimeStamp,TR.PointsProgramID,TR.PointsAmount,TR.ExtLocationCode,TR.ID FROM TransRedemption as TR WITH(NOLOCK) " & _
                                    " WHERE TR.CustomerPrimaryExtID = @CustomerPrimaryExtID " & _
                                    " AND (TR.POSTimeStamp BETWEEN @StartDate AND @EndDate) " & _
                                    " AND TR.PointsProgramID > 0"
                    'No Encryption in WH DB
                    MyCommon.DBParameters.Add("@CustomerPrimaryExtID", SqlDbType.NVarChar).Value = ExtCardID.ConvertBlankIfNothing()
                    MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                    MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                    dt = MyCommon.ExecuteQuery(DataBases.LogixWH)
                    If dt.Rows.Count > 0 Then
                        For Each dr In dt.Rows
                            ' Next
                            Programid = dr.Item("PointsProgramID")
                            LocId = dr.Item("ID") 'LocationID
                            Dim dtExtLocationCode As New DataTable
                            dtExtLocationCode = GetExtLocationCode(MyCommon, LocId)
                            If dtExtLocationCode.Rows.Count > 0 Then
                                extlocationcode = MyCommon.NZ(dtExtLocationCode.Rows(0).Item("ExtLocationCode"), "")
                                MyCommon.QueryStr = "select ProgramName from PointsPrograms WITH (NOLOCK) where ProgramID= @ProgramID "
                                MyCommon.DBParameters.Add("@ProgramID", SqlDbType.BigInt).Value = Programid
                                Dim dtProgramName As New DataTable
                                dtProgramName = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                If dtProgramName.Rows.Count > 0 Then
                                    ProgramName = MyCommon.NZ(dtProgramName.Rows(0).Item("ProgramName"), "")
                                End If

                                row = dtPointsHistory_TransRedempt.NewRow()
                                row.Item("POSTimeStamp") = dr.Item("POSTimeStamp")
                                row.Item("PointsProgramID") = Programid
                                row.Item("PointsAmount") = dr.Item("PointsAmount")
                                row.Item("ProgramName") = ProgramName
                                row.Item("LocationID") = LocId
                                row.Item("ExtLocationCode") = extlocationcode
                                dtPointsHistory_TransRedempt.Rows.Add(row)

                            End If
                        Next
                    End If

                    If dtPointsHistory_TransRedempt.Rows.Count > 0 Then
                        dtPointsHistory_TransRedempt.AcceptChanges()
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Success."
                        dtStatus.Rows.Add(row)
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Records Not Found"
                        dtStatus.Rows.Add(row)
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    ResultDataSet.Tables.Add(dtStatus.Copy())
                End If
                If dtPointsHistory_TransRedempt IsNot Nothing Then ResultDataSet.Tables.Add(dtPointsHistory_TransRedempt)
            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
        End Try
        Return ResultDataSet
    End Function

    <WebMethod()> _
    Public Function GetCustomerPointsHistoryTopRows(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, _
                                                    ByVal SortOrder As String, ByVal NumRows As String) As DataSet
        Dim iCardTypeID As Integer = -1 '0
        Dim iNumRows As Integer = 0
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        If Not String.IsNullOrEmpty(NumRows) AndAlso Not NumRows.Trim = String.Empty AndAlso IsNumeric(NumRows) Then iNumRows = Convert.ToInt32(NumRows)
        Return _GetCustomerPointsHistoryTopRows(GUID, ExtCardID, SortOrder, iCardTypeID, iNumRows)
    End Function

    Private Function _GetCustomerPointsHistoryTopRows(ByVal GUID As String, ByVal ExtCardID As String, ByVal SortOrder As String, ByVal CardTypeID As Integer, _
                                                      ByVal NumRows As Integer) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("PointHistory")
        Dim dtStatus As DataTable
        Dim dtPointsHistoryWithIssuance As DataTable = Nothing
        Dim row, dr As DataRow
        Dim dt As DataTable
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim Programid As Integer
        Dim LocId As Long
        Dim ProgramName As String = "null"
        Dim extlocationcode As String
        Dim TableDatePart As Date = Date.Now

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            If MyCommon.LEXadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixEX()

            If IsValidGUID(GUID, "GetCustomerPointsHistoryTopRows") Then
                If IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then
                    RetCode = RetCode
                    RetMsg = RetMsg
                End If
                RetCode = RetCode
                RetMsg = RetMsg
                If IsStringSortedOrder(SortOrder) Then
                    If IsNumericRowNum(NumRows) Then
                        If RetCode = StatusCodes.SUCCESS Then

                            ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                            CustomerPK = LookupCustomerPK(ExtCardID, CardTypeID, RetCode, RetMsg)
                            If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                                'Create a new datatable to hold the results we'll be assembling
                                dtPointsHistoryWithIssuance = New DataTable("PointsHistory")
                                dtPointsHistoryWithIssuance.Columns.Add("POSTimeStamp", System.Type.GetType("System.String"))
                                dtPointsHistoryWithIssuance.Columns.Add("ProgramID", System.Type.GetType("System.Int32"))
                                dtPointsHistoryWithIssuance.Columns.Add("RewardQty", System.Type.GetType("System.Int32"))
                                dtPointsHistoryWithIssuance.Columns.Add("LocationID", System.Type.GetType("System.Int32"))
                                dtPointsHistoryWithIssuance.Columns.Add("Net", System.Type.GetType("System.Int32"))
                                dtPointsHistoryWithIssuance.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                                dtPointsHistoryWithIssuance.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
                                Dim strExtCardID As String = String.Empty
                                Dim strExtCardIDOriginal As String = String.Empty
                                MyCommon.QueryStr = " Select ExtCardID,ExtCardIDOriginal from CardIDs where ExtCardID = @ExtCardID"
                                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                                dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                If dt.Rows.Count > 0 Then
                                    strExtCardID = dt.Rows(0)("ExtCardID").ToString()
                                    strExtCardIDOriginal = dt.Rows(0)("ExtCardIDOriginal").ToString()
                                End If
                                
                                MyCommon.QueryStr = " Select * from (SELECT Top " & NumRows & "  IHT1.POSTimeStamp, IHT1.ProgramID, IHT1.RewardQty, IHT1.LocationID, " & _
                                                    " CAST((ROUND(IHT2.Net, 2, 1) * 100) as INTEGER) as 'Net' " & _
                                                    " FROM Issuance" & TableDatePart.ToString("yyyyMMdd") & " as IHT1 " & _
                                                    " INNER JOIN Issuance" & TableDatePart.ToString("yyyyMMdd") & " as IHT2 ON IHT1.LogixTransNum = IHT2.LogixTransNum" & _
                                                    " Where (IHT1.PrimaryExtID = @ExtCardID OR IHT1.PrimaryExtID = @ExtCardIDOriginal) AND IHT1.CardTypeID = @CardTypeID AND IHT1.ProgramID > 0 " & _
                                                    " AND IHT2.ProgramID > 0 AND IHT2.Void = 0)  AS CustHistData "
                                'No Encryption in EX DB
                                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = strExtCardID
                                MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                                MyCommon.DBParameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar).Value = strExtCardIDOriginal
                                Select Case SortOrder
                                    Case "ASC"
                                        MyCommon.QueryStr &= " Order By POSTimeStamp Asc"
                                    Case "DESC"
                                        MyCommon.QueryStr &= " Order By POSTimeStamp Desc"
                                    Case "asc"
                                        MyCommon.QueryStr &= " Order By POSTimeStamp Asc"
                                    Case "desc"
                                        MyCommon.QueryStr &= " Order By POSTimeStamp Desc"

                                End Select
                                dt = MyCommon.ExecuteQuery(DataBases.LogixEX)
                                If dt.Rows.Count > 0 Then
                                    For Each dr In dt.Rows
                                        Programid = dr.Item("ProgramID")
                                        LocId = dr.Item("LocationID")
                                        Dim dtExtLocationCode As New DataTable
                                        dtExtLocationCode = GetExtLocationCode(MyCommon, LocId)
                                        If dtExtLocationCode.Rows.Count > 0 Then
                                            extlocationcode = MyCommon.NZ(dtExtLocationCode.Rows(0).Item("ExtLocationCode"), "")
                                            MyCommon.QueryStr = "select ProgramName from PointsPrograms WITH (NOLOCK) where ProgramID= @Programid "
                                            MyCommon.DBParameters.Add("@Programid", SqlDbType.BigInt).Value = Programid
                                            Dim dtProgramName As New DataTable
                                            dtProgramName = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                            If dtProgramName.Rows.Count > 0 Then
                                                ProgramName = MyCommon.NZ(dtProgramName.Rows(0).Item("ProgramName"), "")
                                            End If
                                            row = dtPointsHistoryWithIssuance.NewRow()
                                            row.Item("POSTimeStamp") = dr.Item("POSTimeStamp")
                                            row.Item("ProgramID") = Programid
                                            row.Item("RewardQty") = dr.Item("RewardQty")
                                            row.Item("LocationID") = LocId
                                            row.Item("Net") = dr.Item("Net")
                                            row.Item("ProgramName") = ProgramName
                                            row.Item("ExtLocationCode") = extlocationcode
                                            dtPointsHistoryWithIssuance.Rows.Add(row)
                                        End If
                                    Next
                                End If

                                If dtPointsHistoryWithIssuance.Rows.Count > 0 Then
                                    dtPointsHistoryWithIssuance.AcceptChanges()
                                    row = dtStatus.NewRow()
                                    row.Item("StatusCode") = StatusCodes.SUCCESS
                                    row.Item("Description") = "Success."
                                    dtStatus.Rows.Add(row)
                                Else
                                    row = dtStatus.NewRow()
                                    row.Item("StatusCode") = StatusCodes.SUCCESS
                                    row.Item("Description") = "Records Not Found"
                                    dtStatus.Rows.Add(row)
                                End If
                            Else
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = RetCode
                                row.Item("Description") = "CardID: " & ExtCardID & " with CardTypeID: " & CardTypeID & " not found."
                                dtStatus.Rows.Add(row)
                            End If
                        Else
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = RetCode
                            row.Item("Description") = RetMsg
                            dtStatus.Rows.Add(row)
                        End If
                        If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                            dtStatus.AcceptChanges()
                            ResultDataSet.Tables.Add(dtStatus.Copy())
                        End If
                        If dtPointsHistoryWithIssuance IsNot Nothing Then ResultDataSet.Tables.Add(dtPointsHistoryWithIssuance)

                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_ROWNUM
                        row.Item("Description") = "Error:Row Number Must be Numeric Only"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultDataSet.Tables.Add(dtStatus.Copy())
                    End If

                Else
                    row = dtStatus.NewRow()
                    If SortOrder Is Nothing OrElse SortOrder.Trim = "" Then
                        'Bad customer ID
                        row.Item("StatusCode") = StatusCodes.INVALID_SORTORDER
                        row.Item("Description") = "Failure:Sorted Oredr is not Provided"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultDataSet.Tables.Add(dtStatus.Copy())
                    Else
                        row.Item("StatusCode") = StatusCodes.INVALID_SORTORDER
                        row.Item("Description") = "Error:Pls Provide Sorted Oredr As either asc or desc"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultDataSet.Tables.Add(dtStatus.Copy())
                    End If
                End If

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            If MyCommon.LEXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixEX()
        End Try
        Return ResultDataSet
    End Function

    <WebMethod()> _
    Public Function UpdateSign(ByVal GUID As String, ByVal XmlDoc As String) As DataSet
        Dim ResultSet As New System.Data.DataSet("UpdateSigns")
        Dim MethodName As String
        Dim dtstatus As DataTable
        Dim dtUpdateSignsStatus As DataTable = Nothing
        Dim RetMsg As String = ""
        Dim DeptNo As String
        Dim SignNumber As String = ""
        Dim SignDefault As String = ""
        Dim SignDescription As String
        Dim SignsXML As New XmlDocument
        Dim StoredProcStatus As Integer = 0
        Dim row As DataRow
        Dim i As Integer
        Dim ConnInc As New Copient.ConnectorInc

        'SignsXML.Load(XmlDoc)
        ConnInc.ConvertStringToXML(XmlDoc, SignsXML)
        dtstatus = New DataTable
        dtstatus.TableName = "Status"
        dtstatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtstatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MethodName = "UpdateSign"
            If Not IsValidGUID(GUID, MethodName) Then
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Invalid GUID sent. " & GUID
                dtstatus.Rows.Add(row)
            ElseIf Not IsValidXmlDocument("SignsUpdate.xsd", SignsXML) Then
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CRITERIA_XML
                row.Item("Description") = "XML document sent does not conform to the SignsUpdate.xsd schema."
                dtstatus.Rows.Add(row)
            Else
                'First delete all the data from HierarchySigns table
                MyCommon.QueryStr = "Delete From HierarchySigns with (rowlock);"
                MyCommon.ExecuteNonQuery(DataBases.LogixRT)
                'Now Update the fresh signs
                dtUpdateSignsStatus = New DataTable("CustomerGroups")
                dtUpdateSignsStatus.Columns.Add("Status", System.Type.GetType("System.String"))
                For Each Dept As System.Xml.XmlElement In SignsXML.DocumentElement.ChildNodes
                    For i = 0 To Dept.ChildNodes.Count - 1
                        Select Case Dept.ChildNodes(i).Name
                            Case "DeptNo"
                                DeptNo = Dept.ChildNodes(i).InnerText
                            Case "SignNo"
                                SignNumber = Dept.ChildNodes(i).InnerText
                            Case "SignDescription"
                                SignDescription = Dept.ChildNodes(i).InnerText
                            Case "Default"
                                SignDefault = Dept.ChildNodes(i).InnerText
                        End Select
                    Next
                    ' immediately update the signs
                    MyCommon.QueryStr = "dbo.pt_Signs_Update"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@DeptNo", SqlDbType.NVarChar).Value = DeptNo
                    MyCommon.LRTsp.Parameters.Add("@SignNumber", SqlDbType.NVarChar).Value = SignNumber
                    MyCommon.LRTsp.Parameters.Add("@SignDescription", SqlDbType.Char).Value = IIf(SignDescription = "", Convert.DBNull, SignDescription)
                    MyCommon.LRTsp.Parameters.Add("@SignDefault", SqlDbType.NVarChar).Value = IIf(SignDefault = "", Convert.DBNull, SignDefault)
                    MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()
                    StoredProcStatus = MyCommon.LRTsp.Parameters("@Status").Value
                    MyCommon.Close_LRTsp()

                    If StoredProcStatus = 1 Then
                        row = dtstatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Successfully Updated " & SignNumber & " Sign In AMS."
                        dtstatus.Rows.Add(row)

                    ElseIf StoredProcStatus = 2 Then
                        row = dtstatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.FAILED_UPDATE_SIGN
                        row.Item("Description") = "Provided Department " & DeptNo & " not found."
                        dtstatus.Rows.Add(row)

                    Else
                        row = dtstatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.FAILED_UPDATE_SIGN
                        row.Item("Description") = "Failed to Update " & SignNumber & " Sign in In AMS."
                        dtstatus.Rows.Add(row)

                    End If
                Next
            End If
            If dtstatus.Rows.Count > 0 Then
                dtstatus.AcceptChanges()
                ResultSet.Tables.Add(dtstatus.Copy())
            End If
        Catch ex As Exception
            row = dtstatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Error encountered: " & ex.ToString
            dtstatus.Rows.Add(row)
            dtstatus.AcceptChanges()
            ResultSet.Tables.Add(dtstatus.Copy())
            MyCommon.Write_Log(LogFile, RetMsg, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
        Return ResultSet
    End Function

    Private Function GetExtLocationCode(ByVal MyCommon As Copient.CommonInc, ByVal LocID As Integer) As DataTable
        Dim dtExtLocationCode As New DataTable
        MyCommon.QueryStr = "SELECT ExtLocationCode FROM Locations WHERE LocationID= @LocId "
        MyCommon.DBParameters.Add("@LocId", SqlDbType.BigInt).Value = LocID
        dtExtLocationCode = MyCommon.ExecuteQuery(DataBases.LogixRT)
        Return dtExtLocationCode
    End Function

    '"Ahold AMS Integration- Product Enhancements FIS I" -<2.3.22>
    <WebMethod()> _
    Public Function DisplayYTDSavings(ByVal GUID As String, ByVal CardID As String) As DataSet
        Return _DisplayYTDSavings(GUID, CardID)
    End Function

    '"Ahold AMS Integration- Product Enhancements FIS I" -<2.3.22>
    Private Function _DisplayYTDSavings(ByVal GUID As String, ByVal CardID As String) As DataSet
        Dim ResultSet As New System.Data.DataSet("YTDSavings")
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If GUID = "" Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DisplayYTDSavings = ResultSet
            Exit Function
        End If
        If CardID = "" Then
            'CardID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DisplayYTDSavings = ResultSet
            Exit Function
        End If
        If GUID.Contains("'") = True Or GUID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DisplayYTDSavings = ResultSet
            Exit Function
        End If
        If CardID.Contains("'") = True Or CardID.Contains(Chr(34)) = True Then
            'CardID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DisplayYTDSavings = ResultSet
            Exit Function
        End If
        Dim dt As DataTable
        Dim CustomerPK As Long = 0
        Dim dtYTDSavings As System.Data.DataTable

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(118)) = 1) Then
                'Support to display YTD Savings
                'Check and validate the GUID
                'Check if the GUID is valid for Customer Inquiry
                If Not IsValidGUID(GUID, "YTDSavings") Then
                    'Wrong GUID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_GUID
                    row.Item("Description") = "Failure: Invalid GUID."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                ElseIf (CardID.Length < 1) Then
                    'Bad customer ID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                    row.Item("Description") = "Failure: Invalid customer ID."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                Else

                    If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(6)) >= 2) Then
                        'Customer's Householdcard YTD Savings
                        CardID = MyCommon.Pad_ExtCardID(CardID, 1)
                        'Find the Customer PK
                        MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) " & _
                                            "where ExtCardID= @CardID and CardTypeID=1 ;"
                        MyCommon.DBParameters.Add("@CardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(CardID, True)
                    Else
                        'Customer's Customercard YTD Savings
                        CardID = MyCommon.Pad_ExtCardID(CardID, 0)
                        'Find the Customer PK
                        MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) " & _
                                            "where ExtCardID= @CardID and CardTypeID=0 ;"
                        MyCommon.DBParameters.Add("@CardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(CardID, True)
                    End If
                    'dt = MyCommon.LXS_Select
                    dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                    If dt.Rows.Count = 0 Then
                        'Customer PK Not found
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                        row.Item("Description") = "Failure: Invalid customer ID."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else ' To display the YTD Savings based on the parameter(System option 6)
                        'Customer PK found
                        CustomerPK = dt.Rows(0)(0)
                        MyCommon.QueryStr = " select CustomerPK, FirstName, LastName, CurrYearSTD, LastYearSTD " & _
                                            " from Customers with (NoLock) where customerpk= @CustomerPK"
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        dtYTDSavings = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dtYTDSavings.Rows.Count > 0 Then
                            'Customer PK Not found
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success:"
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                            dtYTDSavings.TableName = "YTDSavings"
                            dtYTDSavings.AcceptChanges()
                            ResultSet.Tables.Add(dtYTDSavings.Copy())
                        Else
                            'No records found for the customer
                            'Customer PK Not found
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                            row.Item("Description") = "Failure: Invalid customer ID."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If
                    End If

                End If
            Else
                ' Display YTD Settings value is <> 1
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.UNABLE_TO_DISPLAY_YTDSAVINGS
                row.Item("Description") = "Failure: Unable to display YTD Savings"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Return ResultSet
    End Function

    'Requirement 2.3.4(Ahold Product Enhancement FIS)
    <WebMethod()> _
    Public Function GetcardsByAlternateID(ByVal GUID As String, ByVal AlternateID As String, ByVal Pin As String) As DataSet
        Return _GetcardsByAlternateID(GUID, AlternateID, Pin)
    End Function

    Private Function _GetcardsByAlternateID(ByVal GUID As String, ByVal AlternateID As String, Optional ByVal Pin As String = "") As DataSet
        Dim dtStatus As DataTable
        Dim row, CustomerPKrow, row1 As DataRow
        Dim dt As DataTable
        Dim dtloyaltycards As DataTable = Nothing
        Dim LoyaltyCardsDataSet As New System.Data.DataSet
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim FinalAltID As String = MyCommon.Parse_Quotes(AlternateID & Pin)     ''Concatenating altid and pin to make it final altid
        Dim CardTypeID As Integer = 3 ''card type will always be 3 for alternative id
        Dim CustomerPK As New DataTable
        Dim CustomerPKtemp As Long
        Dim RetMsg As String = ""

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            ''''for opening the database connections
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            'for validating the GUID
            If IsValidGUID(GUID, "GetCards") Then
                ' validate the AltID

                If AlternateID Is Nothing OrElse AlternateID.Trim = "" Then
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Failure.AlternateID is not Provided"
                ElseIf AlternateID.Length < 10 Then
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Failure.Invalid AlternateId"
                End If

                If RetCode = StatusCodes.SUCCESS Then
                    FinalAltID = MyCommon.Pad_ExtCardID(FinalAltID, CardTypeID)
                    MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) " & _
                      "where ExtCardID= @FinalAltID and CardTypeID= @CardTypeID"
                    MyCommon.DBParameters.Add("@FinalAltID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(FinalAltID, True)
                    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                    If dt.Rows.Count > 0 Then
                        CustomerPK = dt
                    Else
                        RetCode = StatusCodes.INVALID_CUSTOMERID
                        RetMsg = "Failure.AlternateID not found"
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = RetCode
                        row.Item("Description") = RetMsg
                        dtStatus.Rows.Add(row)
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    dtloyaltycards = New DataTable("LoyalityCards")
                    dtloyaltycards.Columns.Add("LoyaltyCardId", System.Type.GetType("System.String"))
                    For Each CustomerPKrow In CustomerPK.Rows
                        CustomerPKtemp = MyCommon.NZ(CustomerPKrow.Item("CustomerPK"), 0)
                        MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where customerpk= @CustomerPKtemp and CardTypeId=0;"
                        MyCommon.DBParameters.Add("@CustomerPKtemp", SqlDbType.BigInt).Value = CustomerPKtemp
                        dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        For Each row1 In dt.Rows
                            dtloyaltycards.Rows.Add(MyCryptlib.SQL_StringDecrypt(row1.Item(0).ToString()))
                        Next
                    Next
                    If dtloyaltycards.Rows.Count > 0 Then dtloyaltycards.AcceptChanges()

                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.SUCCESS
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)

                End If
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    LoyaltyCardsDataSet.Tables.Add(dtStatus.Copy())
                End If
                If dtloyaltycards IsNot Nothing Then LoyaltyCardsDataSet.Tables.Add(dtloyaltycards.Copy())

            Else    'if invalid GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                LoyaltyCardsDataSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            LoyaltyCardsDataSet.Tables.Add(dtStatus.Copy())
        Finally

            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return LoyaltyCardsDataSet
    End Function

    <WebMethod()> _
    Public Function GetPointsBalances(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As DataSet
        Dim PointsDataSet As New System.Data.DataSet("PointsBalances")
        Dim dtStatus As DataTable
        Dim dtBalances As DataTable = Nothing
        Dim row As DataRow
        Dim MyLookup As New Copient.CustomerLookup
        Dim LookupRetCode As Copient.CustomerLookup.RETURN_CODE
        Dim Balances(-1) As Copient.CustomerLookup.PointsBalance
        Dim i As Integer
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim PointsTable As Hashtable

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "GetPointsBalances") Then

                ' Lookup the customer
                CustomerPK = LookupCustomerPK(CustomerID, CustomerTypeID, RetCode, RetMsg)

                If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                    'Create a new datatable to hold the results we'll be assembling
                    dtBalances = New DataTable("PointsProgram")
                    dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtBalances.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                    dtBalances.Columns.Add("Balance", System.Type.GetType("System.Int32"))

                    LookupRetCode = Copient.CustomerLookup.RETURN_CODE.OK
                    Balances = MyLookup.GetCustomerPointsBalances(CustomerPK, False, LookupRetCode)

                    If LookupRetCode = Copient.CustomerLookup.RETURN_CODE.OK Then
                        ' load all points programs for lookup
                        PointsTable = GetPointsProgramTable()

                        For i = 0 To Balances.GetUpperBound(0)
                            row = dtBalances.NewRow()
                            row.Item("ProgramID") = Balances(i).ProgramID
                            row.Item("Balance") = Balances(i).Balance

                            If PointsTable.ContainsKey("ID:" & Balances(i).ProgramID) Then
                                row.Item("ProgramName") = PointsTable.Item("ID:" & Balances(i).ProgramID).ToString
                            Else
                                row.Item("ProgramName") = ""
                            End If

                            dtBalances.Rows.Add(row)
                        Next
                        If dtBalances.Rows.Count > 0 Then dtBalances.AcceptChanges()

                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Success."
                        dtStatus.Rows.Add(row)
                    Else
                        RetCode = StatusCodes.FAILED_BALANCE_LOOKUP
                        RetMsg = "Failed to load customer's points program balances.  Check log for details on this exception."
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If

                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus)
                End If
                If dtBalances IsNot Nothing Then PointsDataSet.Tables.Add(dtBalances)

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                PointsDataSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return PointsDataSet
    End Function

    <WebMethod()> _
    Public Function GetProgramBalancesIncludePending(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String, _
                                                   ByVal ProgramID As String, ByVal PromovarID As String, ByVal IncludePending As String) As DataSet
        Return _GetPointsBalancesForProgram(GUID, CardID, CardTypeID, ProgramID, PromovarID, IncludePending)
    End Function

    <WebMethod()> _
    Public Function GetPointsBalancesForProgram(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String, ByVal ProgramID As String, ByVal PromovarID As String) As DataSet
        Return _GetPointsBalancesForProgram(GUID, CardID, CardTypeID, ProgramID, PromovarID)
    End Function

    Public Function _GetPointsBalancesForProgram(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal ProgramID As String, ByVal PromovarID As String, Optional ByVal IncludePending As String = "false") As DataSet
        Dim PointsDataSet As New System.Data.DataSet("PointsBalances")
        Dim dtStatus As DataTable
        Dim dtBalances As DataTable = Nothing
        Dim row As DataRow
        Dim MyLookup As New Copient.CustomerLookup
        Dim LookupRetCode As Copient.CustomerLookup.RETURN_CODE
        Dim Balances(-1) As Copient.CustomerLookup.CustomerPointsBalance
        Dim i As Integer
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim progID As Integer = 0
        Dim PromotionVarId As Integer = 0
        Dim PointsTable As Hashtable
        Dim iCustomerTypeID As Integer = -1

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If Not IsValidGUID(GUID, "GetPointsBalancesForProgram") Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
            _GetPointsBalancesForProgram = PointsDataSet
            Exit Function
        End If
        If String.IsNullOrEmpty(CustomerID) Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "The Customer ID is invalid."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
            _GetPointsBalancesForProgram = PointsDataSet
            Exit Function
        End If
        'Check to see if either ProgramID or PromoVarID are provided.
        If String.IsNullOrEmpty(ProgramID) AndAlso String.IsNullOrEmpty(PromovarID) Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.PROVIDE_PROGRAMID_OR_PROMOVARID
            row.Item("Description") = "Either ProgramID or PromovarID should be provided."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
            _GetPointsBalancesForProgram = PointsDataSet
            Exit Function
        End If
        If String.IsNullOrEmpty(CustomerTypeID) Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "The Customer Type ID is invalid."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
            _GetPointsBalancesForProgram = PointsDataSet
            Exit Function
        End If
        If Not Integer.TryParse(CustomerTypeID, iCustomerTypeID) Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "The Customer Type ID is invalid."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
            _GetPointsBalancesForProgram = PointsDataSet
            Exit Function
        End If
        If (LookupCustomerPK(CustomerID, iCustomerTypeID, RetCode, RetMsg) = 0 AndAlso RetCode <> StatusCodes.SUCCESS) Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
            _GetPointsBalancesForProgram = PointsDataSet
            Exit Function
        End If
        If (Not String.IsNullOrEmpty(ProgramID) Or Not String.IsNullOrEmpty(PromovarID)) Then
            If Not String.IsNullOrEmpty(ProgramID) Then
                progID = LookupProgramID(ProgramID, RetCode, RetMsg)
                If (progID = 0 AndAlso RetCode <> StatusCodes.SUCCESS) Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    _GetPointsBalancesForProgram = PointsDataSet
                    Exit Function
                End If
            End If
            If Not String.IsNullOrEmpty(PromovarID) Then
                PromotionVarId = LookupPromoVarID(PromovarID, RetCode, RetMsg)
                If (PromotionVarId = 0 AndAlso RetCode <> StatusCodes.SUCCESS) Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    _GetPointsBalancesForProgram = PointsDataSet
                    Exit Function
                End If
            End If
        End If
        If (progID > 0 AndAlso PromotionVarId > 0) Then
            If (MyLookup.DoesPromoVarIDExistForProgramID(ProgramID, PromovarID) = False) Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.PROGRAMID_PROMOVARID_MISMATCH
                row.Item("Description") = "Program ID and Promovar ID does not match."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                PointsDataSet.Tables.Add(dtStatus.Copy())
                _GetPointsBalancesForProgram = PointsDataSet
                Exit Function
            End If
        End If
        If String.IsNullOrEmpty(IncludePending) Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_INCLUDEPENDING
            row.Item("Description") = "Failure: Invalid IncludePending."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
            _GetPointsBalancesForProgram = PointsDataSet
            Exit Function
        End If
        If IncludePending <> "" Then
            IncludePending = IncludePending.Trim.ToLower()
            If (IncludePending <> "false" And IncludePending <> "true") And (IncludePending <> "0" And IncludePending <> "1") Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_INCLUDEPENDING
                row.Item("Description") = "Failure: Invalid IncludePending."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                PointsDataSet.Tables.Add(dtStatus.Copy())
                _GetPointsBalancesForProgram = PointsDataSet
                Exit Function
            End If
            If IncludePending = "0" Then IncludePending = "false"
            If IncludePending = "1" Then IncludePending = "true"
        End If
        Try

            ' Lookup the customer
            CustomerPK = LookupCustomerPK(CustomerID, iCustomerTypeID, RetCode, RetMsg)

            If CustomerPK > 0 Then

                'Create a new datatable to hold the results we'll be assembling
                dtBalances = New DataTable("PointsProgram")
                dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                dtBalances.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                dtBalances.Columns.Add("PromoVarID", System.Type.GetType("System.Int64"))
                dtBalances.Columns.Add("Balance", System.Type.GetType("System.Int32"))

                LookupRetCode = Copient.CustomerLookup.RETURN_CODE.OK
                Balances = MyLookup.GetCustomerPointsBalancesForProgram(CustomerPK, True, LookupRetCode, progID, PromotionVarId, IncludePending)

                If LookupRetCode = Copient.CustomerLookup.RETURN_CODE.OK Then
                    ' load all points programs for lookup
                    PointsTable = GetPointsProgramTable()

                    For i = 0 To Balances.GetUpperBound(0)
                        row = dtBalances.NewRow()
                        row.Item("ProgramID") = IIf(Balances(i).ProgramID > 0, Balances(i).ProgramID, "")
                        row.Item("PromoVarID") = IIf(Balances(i).PromovarID > 0, Balances(i).PromovarID, "")
                        row.Item("Balance") = Balances(i).Balance

                        If Balances(i).ProgramID > 0 AndAlso PointsTable.ContainsKey("ID:" & Balances(i).ProgramID) Then
                            row.Item("ProgramName") = PointsTable.Item("ID:" & Balances(i).ProgramID).ToString
                        Else
                            row.Item("ProgramName") = ""
                        End If

                        dtBalances.Rows.Add(row)
                    Next
                    If dtBalances.Rows.Count > 0 Then dtBalances.AcceptChanges()

                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.SUCCESS
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.FAILED_BALANCE_LOOKUP
                    row.Item("Description") = "Failed to load customer's points program balances.  Check log for details on this exception."
                    dtStatus.Rows.Add(row)
                End If
            End If

            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                dtStatus.AcceptChanges()
                PointsDataSet.Tables.Add(dtStatus)
            End If

            If dtBalances IsNot Nothing Then PointsDataSet.Tables.Add(dtBalances)

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return PointsDataSet
    End Function

    <WebMethod()> _
    Public Function AdjustPointsBalanceLocationOffer(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
          ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String, ByVal SessionID As String, _
          ByVal SelectedOfferID As String, ByVal ExtLocationCode As String, ByVal ReasonID As String, ByVal ReasonText As String, _
          ByVal ExternalUserName As String, ByVal PosDateTime As String) As DataSet
        Dim lSelectedOfferID As Long = 0
        Try
            If IsNumeric(SelectedOfferID) = True Then
                lSelectedOfferID = Convert.ToInt64(SelectedOfferID)
            Else
                lSelectedOfferID = 0
            End If
        Catch ex As Exception
            lSelectedOfferID = 0
        End Try
        Dim iReasonId As Integer = 0
        Try
            If IsNumeric(ReasonID) = True Then
                iReasonId = CInt(ReasonID)
            Else
                iReasonId = 0
            End If
        Catch ex As Exception
            iReasonId = 0
        End Try
        If PosDateTime <> "" Then
            Try
                PosDateTime = Convert.ToDateTime(PosDateTime).ToString
            Catch ex As Exception
                PosDateTime = ""
            End Try
        End If
        Return AdjustPointsBalanceLocationOfferDetails(GUID, CustomerID, CustomerTypeID, AdminUserID, ProgramID, AdjustAmount, "", _
            SessionID, lSelectedOfferID, ExtLocationCode, iReasonId, ReasonText, ExternalUserName, PosDateTime, 0, 0)
    End Function

    <WebMethod()> _
    Public Function AdjustPointsBalanceLocationOfferBySource(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
          ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String, ByVal SessionID As String, _
          ByVal SelectedOfferID As String, ByVal ExtLocationCode As String, ByVal ReasonID As String, ByVal ReasonText As String, _
          ByVal ExternalUserName As String, ByVal PosDateTime As String, _
          ByVal AdjustmentSourceID As String, ByVal AdjustmentTypeID As String) As DataSet
        Dim lSelectedOfferID As Long = 0
        Try
            If IsNumeric(SelectedOfferID) = True Then
                lSelectedOfferID = Convert.ToInt64(SelectedOfferID)
            Else
                lSelectedOfferID = 0
            End If
        Catch ex As Exception
            lSelectedOfferID = 0
        End Try
        Dim iReasonId As Integer = 0
        Try
            If IsNumeric(ReasonID) = True Then
                iReasonId = CInt(ReasonID)
            Else
                iReasonId = 0
            End If
        Catch ex As Exception
            iReasonId = 0
        End Try
        If PosDateTime <> "" Then
            Try
                PosDateTime = Convert.ToDateTime(PosDateTime).ToString
            Catch ex As Exception
                PosDateTime = ""
            End Try
        End If
        Dim iSourceId As Integer = 0
        Try
            If IsNumeric(AdjustmentSourceID) = True Then
                iSourceId = CInt(AdjustmentSourceID)
            Else
                iSourceId = 0
            End If
        Catch ex As Exception
            iSourceId = 0
        End Try
        Dim iTypeId As Integer = 0
        Try
            If IsNumeric(AdjustmentTypeID) = True Then
                iTypeId = CInt(AdjustmentTypeID)
            Else
                iTypeId = 0
            End If
        Catch ex As Exception
            iTypeId = 0
        End Try
        Return AdjustPointsBalanceLocationOfferDetails(GUID, CustomerID, CustomerTypeID, AdminUserID, ProgramID, AdjustAmount, "", _
                SessionID, lSelectedOfferID, ExtLocationCode, iReasonId, ReasonText, ExternalUserName, PosDateTime, iSourceId, iTypeId)
    End Function

    Public Function AdjustPointsBalanceLocationOfferDetails(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
            ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String, ByVal Comments As String, ByVal SessionID As String, _
            ByVal SelectedOfferID As Long, ByVal ExtLocationCode As String, ByVal ReasonID As Integer, ByVal ReasonText As String, _
              ByVal ExternalUserName As String, ByVal POSdateTime As String, _
              ByVal AdjustmentSourceID As Integer, ByVal AdjustmentTypeID As Integer) As DataSet

        'Dim PointsDataSet As New System.Data.DataSet("PointsBalances")
        Dim PointsDS As New System.Data.DataSet
        Dim dtStatus As DataTable
        Dim row As DataRow
        Dim CustomerPK As Long
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim MyPoints As New Copient.Points
        Dim NoteResponse As New Copient.CustomerLookup.RETURN_CODE
        Const WEB_SERVICE_NOTE_TYPE As Integer = 2
        'Initialize the status table, which will report the success or failure of the operation
        Dim bOpenedRTConnection As Boolean = False
        Dim bOpenedXSConnection As Boolean = False
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        'Validations
        Using PointsDataSet As Data.DataSet = New DataSet("PointsBalances")
            Try
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                    bOpenedRTConnection = True
                End If

                If GUID.Trim = "" Then
                    'Invalid GUID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_GUID
                    row.Item("Description") = "Invalid GUID"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If CustomerID.Trim = "" Then
                    'Invalid Customer ID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                    row.Item("Description") = "Invalid CustomerID"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If CustomerTypeID.Trim = "" Then
                    'Invalid CustomerTypeID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                    row.Item("Description") = "Invalid Customer Type ID"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If Not IsValidAdminUserID(AdminUserID, RetMsg) Then
                    'Invalid AdminUserID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_ADMINID
                    row.Item("Description") = "The Admin User ID is invalid."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If ProgramID.Trim = "" Then
                    'Invalid ProgramID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_POINTS_PROGRAM
                    row.Item("Description") = "The Program ID is invalid."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If AdjustAmount.Trim = "" Then
                    'Invalid AdjustAmount
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_AMOUNT
                    row.Item("Description") = "The Adjust Amount is invalid."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If ExtLocationCode.Trim = "" Then
                    'Invalid AdjustAmount
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_LOCATIONCODE
                    row.Item("Description") = "Invalid ExtLocationCode."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If POSdateTime.Trim = "" Then
                    'Invalid AdjustAmount
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_POSDATETIME
                    row.Item("Description") = "The POS Date Time is invalid."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If GUID.Contains("'") Or GUID.Contains(Chr(34)) = True Then
                    'Invalid GUID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_GUID
                    row.Item("Description") = "Invalid GUID"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If SessionID.Contains("'") Or SessionID.Contains(Chr(34)) = True Then
                    'Invalid SessionID
                    SessionID = ""
                End If
                If ReasonText.Contains("'") Or ReasonText.Contains(Chr(34)) = True Then
                    ReasonText.Replace("'", "`")
                    ReasonText.Replace(Chr(34), "`")
                End If
                If ExtLocationCode.Contains("'") Or ExtLocationCode.Contains(Chr(34)) = True Then
                    'Invalid GUID
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_LOCATIONCODE
                    row.Item("Description") = "Invalid ExtLocationCode."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                    Return PointsDataSet
                    Exit Function
                End If
                If ExternalUserName.Contains("'") Or ExternalUserName.Contains(Chr(34)) = True Then
                    ExternalUserName.Replace("'", "`")
                    ExternalUserName.Replace(Chr(34), "`")
                End If
            Catch ex As Exception
            Finally
                If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            End Try

            Try
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                    bOpenedRTConnection = True
                End If
                If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixXS()
                    bOpenedXSConnection = True
                End If

                If IsValidGUID(GUID, "AdjustPointsBalance" & IIf(Comments = "", "", "Commented")) Then
                    If Not IsNumeric(CustomerTypeID) Then
                        RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                        RetMsg = "The Customer Type ID is invalid."
                    ElseIf Not IsNumeric(AdjustAmount) Then
                        RetCode = StatusCodes.INVALID_AMOUNT
                        RetMsg = "The Adjust Amount is invalid."
                    ElseIf Not IsNumeric(ProgramID) Then
                        RetCode = StatusCodes.INVALID_POINTS_PROGRAM
                        RetMsg = "The Program ID is invalid."
                    ElseIf Not IsValidAdminUserID(AdminUserID.ToString, RetMsg) Then
                        RetCode = StatusCodes.INVALID_ADMINID
                        RetMsg = "The Admin User ID is invalid. " & RetMsg
                    Else
                        ' Lookup the customer
                        Dim MyLookup As New Copient.CustomerLookup(AdminUserID, 1)

                        CustomerPK = LookupCustomerPK(CustomerID, CustomerTypeID, RetCode, RetMsg)

                        If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                            Try
                                ' create a customer note for reporting purposes and then make the adjustment
                                ' only if the note is successfully created.
                                If Comments Is Nothing Then Comments = ""
                                If Comments.Trim = "" OrElse MyLookup.AddCustomerNote(CustomerPK, Comments, WEB_SERVICE_NOTE_TYPE, NoteResponse) Then
                                    Dim LocationID As Integer = -9
                                    Dim dt1 As DataTable
                                    Try
                                        MyCommon.QueryStr = "select LocationID from dbo.Locations with (noLock) where ExtLocationCode = @ExtLocationCode and Deleted=0"
                                        MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = ExtLocationCode.ConvertBlankIfNothing()
                                        dt1 = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                        If dt1.Rows.Count > 0 Then
                                            LocationID = dt1.Rows(0)(0)
                                        End If
                                    Catch ex As Exception
                                    End Try

                                    If LocationID = -9 Then
                                        'Invalid ExtLocationID
                                        RetCode = StatusCodes.INVALID_LOCATIONCODE
                                        RetMsg = "Invalid ExtLocationCode."
                                    Else
                                        MyPoints.AdjustPoints(AdminUserID, ProgramID, CustomerPK, AdjustAmount, 0, 0, SessionID, SelectedOfferID, _
                    LocationID, ReasonID, ReasonText, ExternalUserName, POSdateTime, AdjustmentSourceID, AdjustmentTypeID)
                                        RetCode = StatusCodes.SUCCESS
                                        RetMsg = AdjustAmount & " points submitted for adjustment to Program ID " & ProgramID & "."
                                    End If
                                Else
                                    RetCode = StatusCodes.FAILED_CUSTOMER_NOTE
                                    RetMsg = "Unable to create the customer note: " & Comments
                                End If
                            Catch idEx As Copient.IdNotFoundException
                                Select Case idEx.GetExceptionType
                                    Case ID_TYPES.CUSTOMER_PK
                                        RetCode = StatusCodes.NOTFOUND_CUSTOMER
                                    Case ID_TYPES.POINTS_PROGRAM_ID
                                        RetCode = StatusCodes.INVALID_POINTS_PROGRAM
                                End Select
                                RetMsg = idEx.GetExceptionMessage
                            Catch ex As Exception
                                Throw ex
                            End Try
                        End If
                    End If
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                Else
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Invalid GUID."
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    PointsDataSet.Tables.Add(dtStatus.Copy())
                End If
            Catch ex As Exception
                If ex.Message.ToString.Contains("The timeout period elapsed prior to obtaining a connection from the pool.") Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                        SqlClient.SqlConnection.ClearPool(MyCommon.LRTadoConn)
                    End If
                    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                        SqlClient.SqlConnection.ClearPool(MyCommon.LXSadoConn)
                    End If
                End If
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
                row.Item("Description") = "Failure: Application " & ex.ToString
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                PointsDataSet.Tables.Add(dtStatus.Copy())
            Finally
                If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
                If bOpenedXSConnection = True And MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            End Try
            PointsDS = PointsDataSet
        End Using
        Return PointsDS
    End Function

    <WebMethod()> _
    Public Function AdjustPointsBalance(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                  ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String) As DataSet
        Return _AdjustPointsBalanceCommented(GUID, CustomerID, CustomerTypeID, AdminUserID, ProgramID, AdjustAmount, "", 0, 0, 0, "")
    End Function

    <WebMethod()> _
    Public Function AdjustPointsBalanceBySource(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                  ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String, ByVal AdjustmentSourceID As String, ByVal AdjustmentTypeID As String, ByVal ReasonID As String, ByVal ReasonText As String) As DataSet
        Return _AdjustPointsBalanceCommented(GUID, CustomerID, CustomerTypeID, AdminUserID, ProgramID, AdjustAmount, "", AdjustmentSourceID, AdjustmentTypeID, ReasonID, ReasonText)
    End Function

    <WebMethod()> _
    Public Function AdjustPointsBalanceCommented(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                    ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String, _
                                    ByVal Comments As String) As DataSet
        Return _AdjustPointsBalanceCommented(GUID, CustomerID, CustomerTypeID, AdminUserID, ProgramID, AdjustAmount, Comments, 0, 0, 0, "")
    End Function
    <WebMethod()> _
    Public Function AdjustPointsBalanceCommentedBySource(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                    ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String, _
                                    ByVal Comments As String, ByVal AdjustmentSourceID As String, ByVal AdjustmentTypeID As String, ByVal ReasonID As String, ByVal ReasonText As String) As DataSet
        Return _AdjustPointsBalanceCommented(GUID, CustomerID, CustomerTypeID, AdminUserID, ProgramID, AdjustAmount, Comments, AdjustmentSourceID, AdjustmentTypeID, ReasonID, ReasonText)
    End Function

    <WebMethod()> _
    Public Function GetStoredValueBalances(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As DataSet
        Dim SVDataSet As New System.Data.DataSet("SVBalances")
        Dim dtStatus As DataTable
        Dim dtBalances As DataTable = Nothing
        Dim row As DataRow
        Dim MyLookup As New Copient.CustomerLookup
        Dim LookupRetCode As Copient.CustomerLookup.RETURN_CODE
        Dim Balances(-1) As Copient.CustomerLookup.StoredValueBalance
        Dim i As Integer
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim SVTable As Hashtable

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "GetStoredValueBalances") Then

                ' Lookup the customer
                CustomerPK = LookupCustomerPK(CustomerID, CustomerTypeID, RetCode, RetMsg)

                If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                    'Create a new datatable to hold the results we'll be assembling
                    dtBalances = New DataTable("StoredValueProgram")
                    dtBalances.Columns.Add("SVProgramID", System.Type.GetType("System.Int64"))
                    dtBalances.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                    dtBalances.Columns.Add("Units", System.Type.GetType("System.Int32"))
                    dtBalances.Columns.Add("Balance", System.Type.GetType("System.Decimal"))

                    LookupRetCode = Copient.CustomerLookup.RETURN_CODE.OK
                    Balances = MyLookup.GetCustomerSVBalances(CustomerPK, False, LookupRetCode)

                    If LookupRetCode = Copient.CustomerLookup.RETURN_CODE.OK Then
                        ' load all stored value programs for lookup
                        SVTable = GetSVProgramTable()

                        For i = 0 To Balances.GetUpperBound(0)
                            row = dtBalances.NewRow()
                            row.Item("SVProgramID") = Balances(i).SVProgramID
                            row.Item("Units") = Balances(i).Units
                            row.Item("Balance") = Balances(i).Balance

                            If SVTable.ContainsKey("ID:" & Balances(i).SVProgramID) Then
                                row.Item("ProgramName") = SVTable.Item("ID:" & Balances(i).SVProgramID).ToString
                            Else
                                row.Item("ProgramName") = ""
                            End If

                            dtBalances.Rows.Add(row)
                        Next
                        If dtBalances.Rows.Count > 0 Then dtBalances.AcceptChanges()

                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Success."
                        dtStatus.Rows.Add(row)
                    Else
                        RetCode = StatusCodes.FAILED_BALANCE_LOOKUP
                        RetMsg = "Failed to load customer's stored value program balances.  Check log for details on this exception."
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If

                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    SVDataSet.Tables.Add(dtStatus)
                End If
                If dtBalances IsNot Nothing Then SVDataSet.Tables.Add(dtBalances)

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                SVDataSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            SVDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SVDataSet
    End Function

    <WebMethod()> _
    Public Function AdjustStoredValueBalance(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                           ByVal AdminUserID As String, ByVal SVProgramID As String, ByVal AdjustAmount As String) As DataSet
        Dim iAdminUserID As Long = -1
        Dim iSVProgramID As Long = -1
        Dim iAdjustAmount As Decimal = -1

        If Not String.IsNullOrEmpty(AdminUserID) AndAlso Not AdminUserID.Trim = String.Empty AndAlso IsNumeric(AdminUserID) Then iAdminUserID = Convert.ToInt64(AdminUserID)
        If Not String.IsNullOrEmpty(SVProgramID) AndAlso Not SVProgramID.Trim = String.Empty AndAlso IsNumeric(SVProgramID) Then iSVProgramID = Convert.ToInt64(SVProgramID)
        If Not String.IsNullOrEmpty(AdjustAmount) AndAlso Not AdjustAmount.Trim = String.Empty AndAlso IsNumeric(AdjustAmount) Then iAdjustAmount = Convert.ToDecimal(AdjustAmount)

        Return _AdjustStoredValueBalanceCommented(GUID, CustomerID, CustomerTypeID, iAdminUserID, iSVProgramID, iAdjustAmount, "", _
        0, 0, 0, "")
    End Function

    <WebMethod()> _
    Public Function AdjustStoredValueBalanceBySource(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                           ByVal AdminUserID As String, ByVal SVProgramID As String, ByVal AdjustAmount As String, _
                                           ByVal AdjustmentSourceID As String, ByVal AdjustmentTypeID As String, ByVal ReasonID As String, ByVal ReasonText As String) As DataSet
        Dim iAdminUserID As Long = -1
        Dim iSVProgramID As Long = -1
        Dim iAdjustAmount As Decimal = -1

        If Not String.IsNullOrEmpty(AdminUserID) AndAlso Not AdminUserID.Trim = String.Empty AndAlso IsNumeric(AdminUserID) Then iAdminUserID = Convert.ToInt64(AdminUserID)
        If Not String.IsNullOrEmpty(SVProgramID) AndAlso Not SVProgramID.Trim = String.Empty AndAlso IsNumeric(SVProgramID) Then iSVProgramID = Convert.ToInt64(SVProgramID)
        If Not String.IsNullOrEmpty(AdjustAmount) AndAlso Not AdjustAmount.Trim = String.Empty AndAlso IsNumeric(AdjustAmount) Then iAdjustAmount = Convert.ToDecimal(AdjustAmount)

        Return _AdjustStoredValueBalanceCommented(GUID, CustomerID, CustomerTypeID, iAdminUserID, iSVProgramID, iAdjustAmount, "", _
            AdjustmentSourceID, AdjustmentTypeID, ReasonID, ReasonText)
    End Function

    <WebMethod()> _
    Public Function AdjustStoredValueBalanceCommented(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                    ByVal AdminUserID As Long, ByVal SVProgramID As Long, ByVal AdjustAmount As Decimal, _
                                    ByVal Comments As String) As DataSet

        Return _AdjustStoredValueBalanceCommented(GUID, CustomerID, CustomerTypeID, AdminUserID, SVProgramID, AdjustAmount, Comments, _
      0, 0, 0, "")

    End Function
    <WebMethod()> _
    Public Function AdjustStoredValueBalanceCommentedBySource(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                    ByVal AdminUserID As Long, ByVal SVProgramID As Long, ByVal AdjustAmount As Decimal, _
                                    ByVal Comments As String, ByVal AdjustmentSourceID As String, ByVal AdjustmentTypeID As String, ByVal ReasonID As String, ByVal ReasonText As String) As DataSet

        Return _AdjustStoredValueBalanceCommented(GUID, CustomerID, CustomerTypeID, AdminUserID, SVProgramID, AdjustAmount, Comments, _
        AdjustmentSourceID, AdjustmentTypeID, ReasonID, ReasonText)

    End Function

    <WebMethod()> _
    Public Function GetSessionActivities(ByVal GUID As String, ByVal SessionID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim MyLookup As New Copient.CustomerLookup
        Dim Customer As Copient.Customer
        Dim dt As DataTable
        Dim row As DataRow
        Dim RetCode As Copient.CustomerAbstract.RETURN_CODE = Copient.CustomerAbstract.RETURN_CODE.OK
        Dim ErrorXml As String = ""
        Dim UseActivityExtLog As Boolean

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            UseActivityExtLog = IIf(MyCommon.Fetch_SystemOption(320) = "0", True, False)

            If IsValidGUID(GUID, "GetSessionActivities") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("CustomerContact")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))

                Writer.WriteStartElement("Contact")
                Writer.WriteElementString("ID", SessionID)

                ' Lookup the activity information for this session
                MyCommon.QueryStr = "dbo.pa_ActivityLog_GetBySession"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@SessionID", SqlDbType.NVarChar, 50).Value = SessionID
                dt = MyCommon.LRTsp_select
                MyCommon.Close_LRTsp()

                If dt.Rows.Count > 0 Then
                    ' load the customer data
                    Customer = MyLookup.FindCustomerInfo(MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0), RetCode)
                    If Customer IsNot Nothing AndAlso RetCode = Copient.CustomerAbstract.RETURN_CODE.OK Then
                        ' write the customer tag
                        Writer.WriteStartElement("Customer")

                        If Customer.GetCards.Length > 0 Then
                            Writer.WriteElementString("CustomerID", Customer.GetCards(0).GetExtCardID)
                        Else
                            Writer.WriteElementString("CustomerID", Customer.GetInitialExtCardID)
                        End If

                        Writer.WriteElementString("CustomerTypeID", Customer.GetCustomerTypeID)
                        Writer.WriteEndElement() ' end customer

                        ' write user tag representing who performed the activities in this session
                        Writer.WriteStartElement("User")
                        Writer.WriteElementString("ID", MyCommon.NZ(dt.Rows(0).Item("AdminID"), 0))
                        Writer.WriteElementString("UserName", MyCommon.NZ(dt.Rows(0).Item("UserName"), ""))
                        Writer.WriteElementString("FirstName", MyCommon.NZ(dt.Rows(0).Item("FirstName"), ""))
                        Writer.WriteElementString("LastName", MyCommon.NZ(dt.Rows(0).Item("LastName"), ""))
                        Writer.WriteEndElement() ' end user

                        Writer.WriteEndElement() ' end contact

                        Writer.WriteStartElement("Activities")
                        If UseActivityExtLog Then
                            For Each row In dt.Rows
                                WriteActivity(Writer, row, MyCommon.NZ(row.Item("ActivitySubTypeID"), 0))
                            Next
                        End If
                        Writer.WriteEndElement() ' end activities
                    Else
                        ' customer was not found so send back error
                        ErrorXml = GetErrorXML(StatusCodes.INVALID_CUSTOMERID, SessionID)
                    End If
                Else
                    ' no activity found for the Session ID
                    ErrorXml = GetErrorXML(StatusCodes.NO_ACTIVITY_FOUND_FOR_SESSION_ID, SessionID)
                End If

                Writer.WriteEndElement() ' end customercontact
                Writer.WriteEndDocument()
                Writer.Flush()

            Else
                ' Send back Invalid GUID return code
                ErrorXml = GetErrorXML(StatusCodes.INVALID_GUID, SessionID)
            End If

            If ErrorXml = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXml)
            End If

        Catch ex As Exception
            ' send the application exception return code
            SessionXml = New XmlDocument()
            SessionXml.LoadXml(GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, SessionID))
            Try
                MyCommon.Write_Log(LogFile, "Error encountering during GetSessionActivities call for SessionID " & SessionID & _
                                            ControlChars.CrLf & " Reported Exception: " & ex.ToString, True)
            Catch ex2 As Exception
                ' ignore
            End Try
        Finally
            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetPointsAdjustReport(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date) As DataSet
        Return GeneratePointsAdjustReport(GUID, StartDate, EndDate, "", -1, 0, 0)
    End Function

    <WebMethod()> _
    Public Function GetPointsAdjustReportByUser(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date, _
                                            ByVal AdminUserID As String) As DataSet
        Return GeneratePointsAdjustReport(GUID, StartDate, EndDate, "", -1, AdminUserID, 1)
    End Function

    <WebMethod()> _
    Public Function GetPointsAdjustReportByCustomer(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date, _
                                              ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As DataSet
        Return GeneratePointsAdjustReport(GUID, StartDate, EndDate, CustomerID, CustomerTypeID, 0, 2)
    End Function

    Private Function _AdjustPointsBalanceCommented(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                      ByVal AdminUserID As String, ByVal ProgramID As String, ByVal AdjustAmount As String, _
                                      ByVal Comments As String, ByVal AdjustmentSourceID As String, ByVal AdjustmentTypeID As String, ByVal ReasonID As String, ByVal ReasonText As String) As DataSet
        Dim PointsDataSet As New System.Data.DataSet("PointsBalances")
        Dim dtStatus As DataTable
        Dim row As DataRow
        Dim CustomerPK As Long
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim MyPoints As New Copient.Points
        Dim NoteResponse As New Copient.CustomerLookup.RETURN_CODE
        Const WEB_SERVICE_NOTE_TYPE As Integer = 2
        Dim iAdjustmentSourceID As Integer = 0
        Dim iAdjustmentTypeID As Integer = 0
        Dim iReasonID As Integer = 0

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "AdjustPointsBalance" & IIf(Comments = "", "", "Commented")) Then
                If CustomerID Is Nothing OrElse CustomerID.Trim = "" Then
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Invalid Customer ID. "
                ElseIf CustomerID.Contains("'") = True Or CustomerID.Contains(Chr(34)) = True Then
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Invalid Customer ID. "
                ElseIf Not IsNumeric(CustomerTypeID) Then
                    RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                    RetMsg = "The Customer Type ID is invalid."
                ElseIf Not IsNumeric(AdjustAmount) Then
                    RetCode = StatusCodes.INVALID_AMOUNT
                    RetMsg = "The Adjust Amount is invalid."
                ElseIf Not IsNumeric(ProgramID) Then
                    RetCode = StatusCodes.INVALID_POINTS_PROGRAM
                    RetMsg = "The Program ID is invalid."
                ElseIf Not IsValidAdminUserID(AdminUserID.ToString, RetMsg) Then
                    RetCode = StatusCodes.INVALID_ADMINID
                    RetMsg = "The Admin User ID is invalid. " & RetMsg
                Else
                    If Not String.IsNullOrEmpty(AdjustmentSourceID) AndAlso Not AdjustmentSourceID.Trim = String.Empty AndAlso IsNumeric(AdjustmentSourceID) Then iAdjustmentSourceID = Convert.ToInt32(AdjustmentSourceID)
                    If Not String.IsNullOrEmpty(AdjustmentTypeID) AndAlso Not AdjustmentTypeID.Trim = String.Empty AndAlso IsNumeric(AdjustmentTypeID) Then iAdjustmentTypeID = Convert.ToInt32(AdjustmentTypeID)
                    If Not String.IsNullOrEmpty(ReasonID) AndAlso Not ReasonID.Trim = String.Empty AndAlso IsNumeric(ReasonID) Then iReasonID = Convert.ToInt32(ReasonID)
                    ' Lookup the customer
                    Dim MyLookup As New Copient.CustomerLookup(AdminUserID, 1)

                    CustomerPK = LookupCustomerPK(CustomerID, CustomerTypeID, RetCode, RetMsg)

                    If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                        Try
                            ' create a customer note for reporting purposes and then make the adjustment
                            ' only if the note is successfully created.
                            If Comments Is Nothing Then Comments = ""
                            If Comments.Trim = "" OrElse MyLookup.AddCustomerNote(CustomerPK, Comments, WEB_SERVICE_NOTE_TYPE, NoteResponse) Then
                                MyPoints.AdjustPoints(AdminUserID, ProgramID, CustomerPK, AdjustAmount, 0, 0, _
                                "", 0, -9, iReasonID, ReasonText, "", "", iAdjustmentSourceID, iAdjustmentTypeID)
                                RetCode = StatusCodes.SUCCESS
                                RetMsg = AdjustAmount & " points submitted for adjustment to Program ID " & ProgramID & "."
                            Else
                                RetCode = StatusCodes.FAILED_CUSTOMER_NOTE
                                RetMsg = "Unable to create the customer note: " & Comments
                            End If
                        Catch idEx As Copient.IdNotFoundException
                            Select Case idEx.GetExceptionType
                                Case ID_TYPES.CUSTOMER_PK
                                    RetCode = StatusCodes.NOTFOUND_CUSTOMER
                                    RetMsg = "The Customer ID is not found."
                                Case ID_TYPES.POINTS_PROGRAM_ID
                                    RetCode = StatusCodes.INVALID_POINTS_PROGRAM
                                    RetMsg = "The Program ID is invalid."
                            End Select
                            RetMsg = idEx.GetExceptionMessage
                        Catch ex As Exception
                            Throw ex
                        End Try
                    End If
                End If

                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()

                PointsDataSet.Tables.Add(dtStatus.Copy())
            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                PointsDataSet.Tables.Add(dtStatus.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return PointsDataSet
    End Function
    Private Function _AdjustStoredValueBalanceCommented(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, _
                                      ByVal AdminUserID As Long, ByVal SVProgramID As Long, ByVal AdjustAmount As Decimal, _
                                      ByVal Comments As String, ByVal AdjustmentSourceID As String, ByVal AdjustmentTypeID As String, ByVal ReasonID As String, ByVal ReasonText As String) As DataSet
        Dim SVDataSet As New System.Data.DataSet("SVBalances")
        Dim dtStatus As DataTable
        Dim row As DataRow
        Dim CustomerPK As Long
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim MySV As New Copient.StoredValue
        Dim MyLookup As New Copient.CustomerLookup(AdminUserID, 1)
        Dim NoteResponse As New Copient.CustomerLookup.RETURN_CODE
        Dim RawAdjustment As Decimal = AdjustAmount
        Dim iCustomerTypeID As Integer = -1
        Dim iAdjustmentSourceID As Integer = 0
        Dim iAdjustmentTypeID As Integer = 0
        Dim iReasonID As Integer = 0
        Const WEB_SERVICE_NOTE_TYPE As Integer = 3

        If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
        If Not String.IsNullOrEmpty(AdjustmentSourceID) AndAlso Not AdjustmentSourceID.Trim = String.Empty AndAlso IsNumeric(AdjustmentSourceID) Then iAdjustmentSourceID = Convert.ToInt32(AdjustmentSourceID)
        If Not String.IsNullOrEmpty(AdjustmentTypeID) AndAlso Not AdjustmentTypeID.Trim = String.Empty AndAlso IsNumeric(AdjustmentTypeID) Then iAdjustmentTypeID = Convert.ToInt32(AdjustmentTypeID)
        If Not String.IsNullOrEmpty(ReasonID) AndAlso Not ReasonID.Trim = String.Empty AndAlso IsNumeric(ReasonID) Then iReasonID = Convert.ToInt32(ReasonID)

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "AdjustStoredValueBalance" & IIf(Comments = "", "", "Commented")) Then

                If (AdminUserID = -1) Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_ADMINID
                    row.Item("Description") = "Failure : Invalid(AdminUserID)"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    SVDataSet.Tables.Add(dtStatus.Copy())
                ElseIf IsValidAdminUserID(AdminUserID, RetMsg) Then

                    ' Lookup the customer
                    CustomerPK = LookupCustomerPK(CustomerID, iCustomerTypeID, RetCode, RetMsg)

                    If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
                        Try
                            ' Validate Adjustment Amount
                            If IsValidSVAdjustAmount(SVProgramID, AdjustAmount, RetCode, RetMsg) Then
                                ' create a customer note for reporting purposes and then make the adjustment
                                ' only if the note is successfully created.
                                If Comments Is Nothing Then Comments = ""
                                If Comments.Trim = "" OrElse MyLookup.AddCustomerNote(CustomerPK, Comments, WEB_SERVICE_NOTE_TYPE, NoteResponse) Then
                                    Dim Parms As New StoredValue.AdjustmentData
                                    Parms.AdminUserID = AdminUserID
                                    Parms.ProgramID = SVProgramID
                                    Parms.CustomerPK = CustomerPK
                                    Parms.Adjust = AdjustAmount
                                    Parms.RevokeLocalID = 0
                                    Parms.SessionID = 0
                                    Parms.SelectedOfferID = 0
                                    Parms.LocationID = 0
                                    Parms.IgnoreHouseholding = False
                                    Parms.AdjustmentSourceID = iAdjustmentSourceID
                                    Parms.AdjustmentTypeID = iAdjustmentTypeID
                                    Parms.ReasonID = iReasonID
                                    Parms.ReasonText = ReasonText
                                    Parms.Comments = Comments
                                    
                                    RetMsg = MySV.AdjustStoredValue(Parms)
                                    If RetMsg = "" Then
                                        RetCode = StatusCodes.SUCCESS
                                        RetMsg = RawAdjustment & " stored value submitted for adjustment to SVProgram ID " & SVProgramID & "."
                                    Else
                                        RetCode = StatusCodes.APPLICATION_EXCEPTION
                                    End If
                                Else
                                    RetCode = StatusCodes.FAILED_CUSTOMER_NOTE
                                    RetMsg = "Unable to create the customer note: " & Comments
                                End If
                            End If

                        Catch idEx As Copient.IdNotFoundException
                            Select Case idEx.GetExceptionType
                                Case ID_TYPES.CUSTOMER_PK
                                    RetCode = StatusCodes.NOTFOUND_CUSTOMER
                                Case ID_TYPES.POINTS_PROGRAM_ID
                                    RetCode = StatusCodes.INVALID_STORED_VALUE_PROGRAM
                            End Select
                            RetMsg = idEx.GetExceptionMessage
                        Catch ex As Exception
                            Throw ex
                        End Try
                    End If
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    SVDataSet.Tables.Add(dtStatus.Copy())
                Else
                    'invalid admin userid
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_ADMINID
                    row.Item("Description") = "Failure : Invalid(AdminUserID)"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    SVDataSet.Tables.Add(dtStatus.Copy())
                End If
            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                SVDataSet.Tables.Add(dtStatus.Copy())
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            SVDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SVDataSet
    End Function

    Private Function GeneratePointsAdjustReport(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date, _
                                                ByVal CustomerID As String, ByVal CustomerTypeID As Integer, _
                                                ByVal AdminUserID As String, ByVal Caller As Integer) As DataSet
        Dim PointsDataSet As New System.Data.DataSet("PointsAdjustReport")
        Dim dtStatus As DataTable
        Dim dtReports As DataTable = Nothing
        Dim row As DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = "Success."
        Dim MyLookup As New Copient.CustomerLookup()
        Dim LookupCode As New Copient.CustomerLookup.RETURN_CODE
        Dim CustomerPK As Long = 0
        Dim QueryBuf As New StringBuilder()
        Dim MethodName As String
        Dim validationRespCode As CardValidationResponse = CardValidationResponse.SUCCESS

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable("Status")
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ' determine which web service call was made for logging purposes
            MethodName = "GetPointsAdjustReport"
            If Caller = 2 Then
                MethodName &= "ByCustomer"
            ElseIf Caller = 1 Then
                MethodName &= "ByUser"
            End If

            If IsValidGUID(GUID, MethodName) Then
                If Not IsValidDate(StartDate) Then
                    'Bad Start Date
                    RetCode = StatusCodes.INVALID_STARTDATE
                    RetMsg = "Failure: Invalid StartDate"
                ElseIf Not IsValidDate(EndDate) Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_ENDDATE
                    RetMsg = "Failure: Invalid EndDate"
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    QueryBuf.Append("select CustomerPK, AdminUserID, IsNull(FirstName,'') + ' ' +  IsNull(LastName,'') as AdjustersName, CreatedDate, Note as Comments " & _
                           "from CustomerNotes with (NoLock) " & _
                           "where NoteTypeID=2 and CreatedDate between @StartDate " & _
                           "  and @EndDate ")
                    ' add clause to restrict only to include a specific customer
                    If Caller = 2 AndAlso MyCommon.AllowToProcessCustomerCard(CustomerID, CustomerTypeID, validationRespCode) Then
                        ' handle customer ID padding
                        CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CustomerTypeID)
                        CustomerPK = MyLookup.GetCustomerPK(CustomerID, CustomerTypeID, LookupCode)
                        If CustomerPK > 0 AndAlso LookupCode = Copient.CustomerAbstract.RETURN_CODE.OK Then
                            QueryBuf.Append(" and CustomerPK = @CustomerPK")
                            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        Else
                            RetCode = StatusCodes.NOTFOUND_CUSTOMER
                            RetMsg = "CustomerID: " & CustomerID & " with CustomerTypeID: " & CustomerTypeID & " not found."
                        End If
                    ElseIf Caller = 2 AndAlso validationRespCode <> CardValidationResponse.SUCCESS Then
                        If validationRespCode = CardValidationResponse.CARDIDNOTNUMERIC OrElse validationRespCode = CardValidationResponse.INVALIDCARDFORMAT Then
                            RetCode = StatusCodes.INVALID_CUSTOMERID
                        ElseIf validationRespCode = CardValidationResponse.CARDTYPENOTFOUND OrElse validationRespCode = CardValidationResponse.INVALIDCARDTYPEFORMAT Then
                            RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                        ElseIf validationRespCode = CardValidationResponse.ERROR_APPLICATION Then
                            RetCode = StatusCodes.APPLICATION_EXCEPTION
                        End If
                        RetMsg = MyCommon.CardValidationResponseMessage(CustomerID, CustomerTypeID, validationRespCode)
                    End If

                    ' add clause to restrict to only a certain AdminUserID
                    If Caller = 1 Then
                        If Not IsValidAdminUserID(AdminUserID, RetMsg) Then
                            RetCode = StatusCodes.INVALID_ADMINID
                            RetMsg = "Invalid AdminUserID."
                        Else
                            QueryBuf.Append(" and AdminUserID = @AdminUserID")
                            MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                        End If
                    End If

                    If RetCode = StatusCodes.SUCCESS Then
                        MyCommon.QueryStr = QueryBuf.ToString
                        MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                        MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                        dtReports = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        dtReports.TableName = "PointsAdjustments"
                        RetMsg &= dtReports.Rows.Count & " rows in report."
                    End If

                End If
            Else
                'Wrong GUID
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "Failure: Invalid GUID."
            End If
            row = dtStatus.NewRow()
            row.Item("StatusCode") = RetCode
            row.Item("Description") = RetMsg
            dtStatus.Rows.Add(row)

            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                dtStatus.AcceptChanges()
                PointsDataSet.Tables.Add(dtStatus.Copy)
            End If
            If dtReports IsNot Nothing Then PointsDataSet.Tables.Add(dtReports.Copy)

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            PointsDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return PointsDataSet
    End Function

    Private Function IsValidDate(ByVal inputDate As String) As Boolean
        Dim dt As DateTime
        Dim IsDate As Boolean = False
        Try
            If DateTime.TryParse(inputDate, dt) Then
                If dt >= SqlTypes.SqlDateTime.MinValue And dt <= SqlTypes.SqlDateTime.MaxValue Then
                    IsDate = True
                End If
            End If
        Catch
        End Try
        Return IsDate
    End Function

    <WebMethod()> _
    Public Function CustomerCardMove(ByVal GUID As String, ByVal FromExtCardID As String, ByVal FromCardTypeID As String, ByVal ToExtCardID As String, ByVal ToCardTypeID As String, ByVal DeleteFromCustomer As Boolean) As DataSet

        Dim iFromCardTypeID As Integer = -1
        Dim iToCardTypeID As Integer = -1
        If Not String.IsNullOrEmpty(FromCardTypeID) AndAlso Not FromCardTypeID.Trim = String.Empty AndAlso IsNumeric(FromCardTypeID) Then iFromCardTypeID = Convert.ToInt32(FromCardTypeID)
        If Not String.IsNullOrEmpty(ToCardTypeID) AndAlso Not ToCardTypeID.Trim = String.Empty AndAlso IsNumeric(ToCardTypeID) Then iToCardTypeID = Convert.ToInt32(ToCardTypeID)
        Return GenerateCustomerCardMove(GUID, FromExtCardID, iFromCardTypeID, ToExtCardID, iToCardTypeID, DeleteFromCustomer)
    End Function

    ' This method moves a card from one customer to another.It gets the customer pk of the cardid and then associates the card with a
    ' different customerpk. It does not do any transfer of points
    ' Input params :  GUID - The identifier for the web service
    '                 FromExtCardID - The Card ID for the card that we are trying to move to a different customerpk
    '                 FromCardTypeID - The card type of the card that we are trying to move
    '                 ToExtCardID - The Card ID for the card that we is associated with the destination customerpk
    '                 DeleteFromcustomer - The flag that indicates whether to delete the "from" customer record after moving the card associated with it

    Private Function GenerateCustomerCardMove(ByVal GUID As String, ByVal FromExtCardID As String, ByVal FromCardTypeID As Integer, ByVal ToExtCardID As String, ByVal ToCardTypeID As Integer, ByVal DeleteFromCustomer As Boolean) As DataSet
        Dim MoveDataSet As New System.Data.DataSet("CustomerCardMove")
        Dim dtStatus As DataTable
        Dim dt As DataTable
        Dim row As DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = "Success."
        Dim MyLookup As New Copient.CustomerLookup()
        Dim LookupCode As New Copient.CustomerLookup.RETURN_CODE
        Dim FromCustomerPK As Long = 0
        Dim ToCustPK As Long = 0
        Dim CardIDPadding As Integer = 0
        Dim Cust As New Copient.Customer
        Dim MethodName As String

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable("Status")
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            MethodName = "GenerateCustomerCardMove"

            If IsValidGUID(GUID, MethodName) Then

                If IsValidCustomerCard(FromExtCardID, FromCardTypeID, RetCode, RetMsg) Then

                    FromExtCardID = MyCommon.Pad_ExtCardID(FromExtCardID, FromCardTypeID)

                    If IsValidCustomerCard(ToExtCardID, ToCardTypeID, RetCode, RetMsg) Then
                        ToExtCardID = MyCommon.Pad_ExtCardID(ToExtCardID, ToCardTypeID)

                        MyCommon.QueryStr = "SELECT CustomerPK FROM Customers WITH (NoLock) WHERE CustomerPK IN " & _
                                           "  (SELECT CustomerPK from CardIDs WITH (NoLock) WHERE ExtCardID= @FromExtCardID and CardTypeID= @FromCardTypeID )"
                        MyCommon.DBParameters.Add("@FromExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(FromExtCardID, True)
                        MyCommon.DBParameters.Add("@FromCardTypeID", SqlDbType.Int).Value = FromCardTypeID
                        dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dt.Rows.Count = 1 Then
                            FromCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                            'Check to see if the To CustPK exists
                            ToCustPK = MyLookup.GetCustomerPK(ToExtCardID, ToCardTypeID, LookupCode)
                            If ToCustPK > 0 Then
                                If (ToCustPK <> FromCustomerPK) Then 'Take further action only if source and destination customerpks are not identical
                                    MyCommon.QueryStr = "UPDATE CardIDS SET CustomerPK = @ToCustPK WHERE " & _
                                                      "CustomerPK = @FromCustomerPK AND ExtCardID = @FromExtCardID AND CardTypeID = @FromCardTypeID "
                                    MyCommon.DBParameters.Add("@ToCustPK", SqlDbType.BigInt).Value = ToCustPK
                                    MyCommon.DBParameters.Add("@FromCustomerPK", SqlDbType.BigInt).Value = FromCustomerPK
                                    MyCommon.DBParameters.Add("@FromExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(FromExtCardID, True)
                                    MyCommon.DBParameters.Add("@FromCardTypeID", SqlDbType.Int).Value = FromCardTypeID
                                    MyCommon.ExecuteNonQuery(DataBases.LogixXS)
                                    If DeleteFromCustomer Then
                                        MyCommon.QueryStr = "pt_CustomerRemovalQueue_Insert"
                                        MyCommon.Open_LXSsp()
                                        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = FromCustomerPK
                                        MyCommon.LXSsp.ExecuteNonQuery()
                                    End If
                                    RetCode = StatusCodes.SUCCESS
                                    RetMsg = "Success."
                                Else ' If source and destination cards have the same customerpk, no action to take
                                    RetCode = StatusCodes.INVALID_CUSTOMERID
                                    RetMsg = "Customer record already same for both cards. "
                                End If

                            Else
                                RetCode = StatusCodes.INVALID_CUSTOMERID
                                RetMsg = "Customer record for Card: " & ToExtCardID & " with CardTypeID: " & ToCardTypeID & " not found."
                            End If
                        ElseIf dt.Rows.Count > 1 Then 'more than one
                            RetCode = StatusCodes.INVALID_CUSTOMERID
                            RetMsg = "CardID: " & FromExtCardID & " with CardTypeID: " & FromCardTypeID & " is associated with more than one customer record."
                        ElseIf dt.Rows.Count < 1 Then 'no customer with that number
                            RetCode = StatusCodes.INVALID_CUSTOMERID
                            RetMsg = "Customer record for: " & FromExtCardID & " with CardTypeID: " & FromCardTypeID & " not found."
                        Else

                            RetCode = RetCode
                            RetMsg = RetMsg
                        End If

                    Else

                        RetCode = RetCode
                        RetMsg = RetMsg
                    End If

                End If

                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
            End If

            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                dtStatus.AcceptChanges()
                MoveDataSet.Tables.Add(dtStatus.Copy)
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application exception " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            MoveDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return MoveDataSet
    End Function

    <WebMethod()> _
    Public Function TransferProgramPoints(ByVal GUID As String, ByVal TransferXmlDoc As XmlDocument) As XmlDocument
        Dim MethodName As String
        Dim MyLookup As New Copient.CustomerLookup()
        Dim MyPoints As New Copient.Points
        Dim Program As Copient.PointsProgram
        Dim LookupRetCode As Copient.CustomerAbstract.RETURN_CODE = Copient.CustomerAbstract.RETURN_CODE.OK
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim SourceCustPK, DestCustPK As Long
        Dim CurrentBalance As Long = 0
        Dim RetMsg As String = ""
        Dim OperateAtEnterprise As Boolean = False
        Dim SourceCard As String = ""
        Dim DestinationCard As String = ""
        Dim ProgramID, PointsToTransfer As Long
        Dim SourceCardTypeID, DestCardTypeID As Integer
        Dim ht As New Hashtable(10)
        Dim TransferAll As Boolean = True
        Dim StoredProcStatus As Integer = 0
        Dim TransferAmtReduced As Boolean = False
        Dim TransferCode As String = ""
        Dim HHPK As Long = 0

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            MethodName = "TransferProgramPoints"

            If Not IsValidGUID(GUID, MethodName) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "Invalid GUID sent. " & GUID
            ElseIf Not IsValidXmlDocument("PointsTransfer.xsd", TransferXmlDoc) Then
                RetCode = StatusCodes.INVALID_CRITERIA_XML
                RetMsg = "XML document sent does not conform to the PointsTransfer.xsd schema.  Check log file " & LogFile & " for specific details."
            Else
                ' Parse the values from the XML Document into variables
                ht = ParseTransferCriteria(TransferXmlDoc)
                SourceCard = ht.Item("SourceCard")
                DestinationCard = ht.Item("DestinationCard")
                Integer.TryParse(ht.Item("SourceCardTypeID"), SourceCardTypeID)
                Integer.TryParse(ht.Item("DestinationCardTypeID"), DestCardTypeID)
                Long.TryParse(ht.Item("PointsProgramID"), ProgramID)
                Long.TryParse(ht.Item("TransferValue"), PointsToTransfer)
                TransferAll = (ht.Item("TransferAll").ToString = "1")

                SourceCard = MyCommon.Pad_ExtCardID(SourceCard, commonShared.CardTypes.CUSTOMER)
                DestinationCard = MyCommon.Pad_ExtCardID(DestinationCard, commonShared.CardTypes.CUSTOMER)

                ' validate the cards - switch to the household PK when the customer is householded.
                SourceCustPK = MyLookup.GetCustomerPK(SourceCard, SourceCardTypeID, LookupRetCode)
                HHPK = MyLookup.GetCustomerHHPK(SourceCustPK)
                If HHPK > 0 Then
                    SourceCustPK = HHPK
                    MyCommon.Write_Log(LogFile, "Source card " & MaskHelper.MaskCard(SourceCard, commonShared.CardTypes.HOUSEHOLD) & " is householded.  Using household account (" & HHPK & ") to transfer points.", True)
                End If
                DestCustPK = MyLookup.GetCustomerPK(DestinationCard, DestCardTypeID, LookupRetCode)
                HHPK = MyLookup.GetCustomerHHPK(DestCustPK)
                If HHPK > 0 Then
                    DestCustPK = HHPK
                    MyCommon.Write_Log(LogFile, "Destination card " & MaskHelper.MaskCard(SourceCard, commonShared.CardTypes.HOUSEHOLD) & " is householded.  Using household account (" & HHPK & ") to receive points transfer.", True)
                End If

                ' validate the program
                Program = MyPoints.GetPointsProgram(ProgramID)

                If SourceCustPK = 0 Then
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Source card number " & SourceCard & " was not found."
                ElseIf DestCustPK = 0 Then
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Destination card number " & DestinationCard & " was not found."
                ElseIf Program Is Nothing OrElse (Program IsNot Nothing AndAlso Program.GetProgramID = 0) Then
                    RetCode = StatusCodes.INVALID_POINTS_PROGRAM
                    RetMsg = "Program ID " & ProgramID & " was not found."
                Else

                    ' determine what portion of the balance needs transferred
                    CurrentBalance = MyPoints.GetBalance(SourceCustPK, ProgramID)
                    If TransferAll Then PointsToTransfer = CurrentBalance
                    If PointsToTransfer > CurrentBalance Then
                        TransferAmtReduced = True
                        PointsToTransfer = CurrentBalance
                    End If

                    If PointsToTransfer > 0 Then
                        OperateAtEnterprise = (MyCommon.Fetch_CPE_SystemOption(91) = "1")
                        If OperateAtEnterprise Then
                            ' immediately apply the points adjustment
                            MyCommon.QueryStr = "dbo.pa_CPE_CI_PointsAdj"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@SourceCustPK", SqlDbType.BigInt).Value = SourceCustPK
                            MyCommon.LXSsp.Parameters.Add("@DestCustPK", SqlDbType.BigInt).Value = DestCustPK
                            MyCommon.LXSsp.Parameters.Add("@PointsProgramID", SqlDbType.BigInt).Value = ProgramID
                            MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Value = Program.GetPromoVarID
                            MyCommon.LXSsp.Parameters.Add("@TransferAmount", SqlDbType.Decimal, 15).Value = New Decimal(PointsToTransfer)
                            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LXSsp.ExecuteNonQuery()
                            StoredProcStatus = MyCommon.LXSsp.Parameters("@Status").Value
                            MyCommon.Close_LXSsp()

                            If StoredProcStatus = 1 Then
                                RetMsg = "Successfully transferred " & PointsToTransfer & " points in program " & ProgramID & " from card " & SourceCard & " to card " & DestinationCard & "."
                            Else
                                RetMsg = "Failed to transfer " & PointsToTransfer & " points in program " & ProgramID & " from card " & SourceCard & " to card " & DestinationCard & ". " & _
                                          "Return Status Code = " & StoredProcStatus
                            End If
                        Else
                            ' queue the adjustment for standard points adjustment processing by CPETransUpdateAgent_PA agent
                            MyPoints.AdjustPoints(1, ProgramID, SourceCustPK, -PointsToTransfer, 0, 0)
                            MyPoints.AdjustPoints(1, ProgramID, DestCustPK, PointsToTransfer, 0, 0)
                            RetMsg = "Successfully queued the transfer of " & PointsToTransfer & " points in program " & ProgramID & " from card " & SourceCard & _
                                      " to card " & DestinationCard & ".  Adjustment will be processed with the next run of the " & _
                                      "Points Adjustment agent."
                        End If
                    Else
                        RetCode = StatusCodes.INVALID_AMOUNT
                        RetMsg = "Source card " & SourceCard & " has only " & CurrentBalance & " points available in program " & ProgramID & " for transfer to card " & DestinationCard & "."
                    End If
                End If
            End If

        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Error encountered: " & ex.ToString
            MyCommon.Write_Log(LogFile, RetMsg, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        ' populate the results in the XML document returned in the response
        If RetCode = StatusCodes.SUCCESS AndAlso TransferAmtReduced Then
            TransferCode = "TRANSFER_PARTIALLY_ADJUSTED"
        ElseIf RetCode = StatusCodes.SUCCESS AndAlso Not TransferAmtReduced Then
            TransferCode = "TRANSFER_FULLY_ADJUSTED"
        Else
            TransferCode = "TRANSFER_FAILED"
            PointsToTransfer = 0
        End If
        PopulateTransferResults(TransferXmlDoc, TransferCode, RetMsg, PointsToTransfer)

        Return TransferXmlDoc
    End Function

    Private Function ParseTransferCriteria(ByVal TransferXmlDoc As XmlDocument) As Hashtable
        Dim ht As New Hashtable(10)
        Dim TempNode As XmlNode
        Dim DefaultCardTypeID As String
        Dim CardDirections() As String = {"Source", "Destination"}

        DefaultCardTypeID = MyCommon.Fetch_SystemOption(30)

        ' parse the source and destination card values
        For Each Direction As String In CardDirections
            ht.Add(Direction & "Card", "")
            ht.Add(Direction & "CardTypeID", DefaultCardTypeID)
            TempNode = TransferXmlDoc.SelectSingleNode("/PointsTransfer/Criteria/" & Direction & "Card")
            If TempNode IsNot Nothing Then
                ht.Item(Direction & "Card") = TempNode.Attributes("cardNumber").InnerText
                If TempNode.Attributes("cardType") IsNot Nothing Then
                    ht.Item(Direction & "CardTypeID") = TempNode.Attributes("cardType").InnerText
                End If
            End If
        Next

        ' parse the points program ID
        ht.Add("PointsProgramID", "-1")
        TempNode = TransferXmlDoc.SelectSingleNode("/PointsTransfer/Criteria/PointsProgram")
        If TempNode IsNot Nothing Then
            If TempNode.Attributes("logixID") IsNot Nothing Then
                ht.Item("PointsProgramID") = TempNode.Attributes("logixID").InnerText
            End If
        End If

        ' parse the transfer amount criteria
        ht.Add("TransferAll", "1")
        ht.Add("TransferValue", "0")
        TempNode = TransferXmlDoc.SelectSingleNode("/PointsTransfer/Criteria/Transfer")
        If TempNode IsNot Nothing Then
            If TempNode.Attributes("value") IsNot Nothing Then
                ht.Item("TransferValue") = TempNode.Attributes("value").InnerText
                ht.Item("TransferAll") = "0"
            End If
        End If

        Return ht
    End Function

    <WebMethod()> _
    Public Function GetStoredValueAdjustReport(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date) As DataSet
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If IsValidDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If IsValidDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)
        Return GenerateStoredValueAdjustReport(GUID, sStartDate, sEndDate, 0)
    End Function

    <WebMethod()> _
    Public Function GetStoredValueAdjustReportByUser(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date, _
                                            ByVal AdminUserID As Integer) As DataSet
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If IsValidDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If IsValidDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)
        Return GenerateStoredValueAdjustReport(GUID, sStartDate, sEndDate, 1, , , AdminUserID)
    End Function

    <WebMethod()> _
    Public Function GetStoredValueAdjustReportByCustomer(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date, _
                                              ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As DataSet
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If IsValidDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If IsValidDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)
        Return GenerateStoredValueAdjustReport(GUID, sStartDate, sEndDate, 2, CustomerID, CustomerTypeID)
    End Function

    Private Function GenerateStoredValueAdjustReport(ByVal GUID As String, ByVal StartDate As Date, ByVal EndDate As Date, ByVal Caller As Integer, _
                                                Optional ByVal CustomerID As String = "", Optional ByVal CustomerTypeID As Integer = -1, _
                                                Optional ByVal AdminUserID As Integer = 0) As DataSet
        Dim SVDataSet As New System.Data.DataSet("StoredValueAdjustReport")
        Dim dtStatus As DataTable
        Dim dtReports As DataTable = Nothing
        Dim row As DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = "Success."
        Dim MyLookup As New Copient.CustomerLookup(AdminUserID, 1)
        Dim LookupCode As New Copient.CustomerLookup.RETURN_CODE
        Dim CustomerPK As Long = 0
        Dim QueryBuf As New StringBuilder()
        Dim MethodName As String
        Dim validationRespCode As CardValidationResponse = CardValidationResponse.SUCCESS

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable("Status")
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            ' determine which web service call was made for logging purposes
            MethodName = "GetStoredValueAdjustReport"
            If Caller = 2 Then
                MethodName &= "ByCustomer"
            ElseIf Caller = 1 Then
                MethodName &= "ByUser"
            End If

            If IsValidGUID(GUID, MethodName) Then
                If (StartDate = "01-01-1900") Then
                    'Bad Start Date
                    RetCode = StatusCodes.INVALID_STARTDATE
                    RetMsg = "Failure: Invalid StartDate"
                ElseIf (EndDate = "01-01-1900") Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_ENDDATE
                    RetMsg = "Failure: Invalid EndDate"
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    QueryBuf.Append("select CustomerPK, AdminUserID, IsNull(FirstName,'') + ' ' +  IsNull(LastName,'') as AdjustersName, CreatedDate, Note as Comments " & _
                                   "from CustomerNotes with (NoLock) " & _
                                   "where NoteTypeID=3 and CreatedDate between @StartDate " & _
                                   "  and @EndDate ")
                    ' add clause to restrict only to include a specific customer
                    If Caller = 2 AndAlso MyCommon.AllowToProcessCustomerCard(CustomerID, CustomerTypeID, validationRespCode) Then
                        ' handle customer ID padding
                        CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CustomerTypeID)
                        CustomerPK = MyLookup.GetCustomerPK(CustomerID, CustomerTypeID, LookupCode)
                        If CustomerPK > 0 AndAlso LookupCode = Copient.CustomerAbstract.RETURN_CODE.OK Then
                            QueryBuf.Append(" and CustomerPK = @CustomerPK")
                            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        Else
                            RetCode = StatusCodes.NOTFOUND_CUSTOMER
                            RetMsg = "CustomerID: " & CustomerID & " with CustomerTypeID: " & CustomerTypeID & " not found."
                        End If
                    ElseIf Caller = 2 AndAlso validationRespCode <> CardValidationResponse.SUCCESS Then
                        If validationRespCode = CardValidationResponse.CARDIDNOTNUMERIC OrElse validationRespCode = CardValidationResponse.INVALIDCARDFORMAT Then
                            RetCode = StatusCodes.INVALID_CUSTOMERID
                        ElseIf validationRespCode = CardValidationResponse.CARDTYPENOTFOUND OrElse validationRespCode = CardValidationResponse.INVALIDCARDTYPEFORMAT Then
                            RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                        ElseIf validationRespCode = CardValidationResponse.ERROR_APPLICATION Then
                            RetCode = StatusCodes.APPLICATION_EXCEPTION
                        End If
                        RetMsg = MyCommon.CardValidationResponseMessage(CustomerID, CustomerTypeID, validationRespCode)
                    End If

                    ' add clause to restrict to only a certain AdminUserID
                    If Caller = 1 AndAlso AdminUserID > 0 Then
                        Dim dt As DataTable
                        MyCommon.QueryStr = "Select AdminUserID from AdminUsers where AdminUserID = @AdminUserID"
                        MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                        dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
                        If dt.Rows.Count = 0 Then
                            RetCode = StatusCodes.INVALID_ADMINID
                            RetMsg = "Invalid AdminUserID."
                        Else
                            QueryBuf.Append(" and AdminUserID = @AdminUserID")
                            MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
                        End If
                    ElseIf Caller = 1 AndAlso AdminUserID <= 0 Then
                        RetCode = StatusCodes.INVALID_ADMINID
                        RetMsg = "Invalid AdminUserID."
                    End If

                    If RetCode = StatusCodes.SUCCESS Then
                        MyCommon.QueryStr = QueryBuf.ToString
                        MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                        MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                        dtReports = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        dtReports.TableName = "StoredValueAdjustments"
                        RetMsg &= dtReports.Rows.Count & " rows in report."
                    End If
                End If
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
            End If

            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                dtStatus.AcceptChanges()
                SVDataSet.Tables.Add(dtStatus.Copy)
            End If
            If dtReports IsNot Nothing Then SVDataSet.Tables.Add(dtReports.Copy)

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            SVDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SVDataSet
    End Function

    <WebMethod()> _
    Public Function GetCustomerInfoChanges(ByVal GUID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim ds As New DataSet()
        Dim dtCustInfo As New DataTable
        Dim dtCardInfo As New DataTable
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim row As DataRow
        Dim RetCode As Copient.CustomerAbstract.RETURN_CODE = Copient.CustomerAbstract.RETURN_CODE.OK
        Dim ErrorXml As String = ""
        Dim RecCt As Integer
        Dim BatchGUID As String = System.Guid.NewGuid().ToString

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "GetCustomerInfoChanges") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("Customers")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))

                ' rollback this batch and any previous records that were abandoned in the in-process state
                MyCommon.QueryStr = "dbo.pa_CustomerInquiry_MarkInfoAsReady"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@BatchGUID", SqlDbType.NVarChar, 36).Value = Left(BatchGUID, 36)
                MyCommon.LXSsp.Parameters.Add("@RecsRolledback", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                RecCt = MyCommon.LXSsp.Parameters("@RecsRolledback").Value
                MyCommon.Close_LXSsp()
                If RecCt > 0 Then
                    MyCommon.Write_Log(LogFile, "Records rolled back during pre-processing: " & RecCt, True)
                End If

                ' mark a batch of customer changes into in-process status
                MyCommon.QueryStr = "dbo.pa_CustomerInquiry_MarkInfoInProcess"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@BatchGUID", SqlDbType.NVarChar, 36).Value = Left(BatchGUID, 36)
                MyCommon.LXSsp.Parameters.Add("@RemainingRecs", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                RecCt = MyCommon.LXSsp.Parameters("@RemainingRecs").Value
                MyCommon.Close_LXSsp()

                ' load the marked batch of in-process records
                MyCommon.QueryStr = "dbo.pa_CustomerInquiry_InfoChanges"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@BatchGUID", SqlDbType.NVarChar, 36).Value = Left(BatchGUID, 36)
                reader = MyCommon.LXSsp.ExecuteReader

                ds.Tables.Add(dtCustInfo)
                ds.Tables.Add(dtCardInfo)

                ds.Load(reader, LoadOption.OverwriteChanges, Nothing, New DataTable() {dtCustInfo, dtCardInfo})

                MyCommon.Close_LXSsp()
                reader.Close()

                Writer.WriteAttributeString("recordCount", dtCustInfo.Rows.Count)
                Writer.WriteAttributeString("hasMore", IIf(RecCt > 0, "true", "false"))
                Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

                If dtCustInfo.Rows.Count > 0 Then
                    ' load the customer data
                    For Each row In dtCustInfo.Rows
                        WriteCustomerRecord(Writer, row, dtCardInfo, False)
                    Next
                End If

                Writer.WriteEndElement() ' end customers
                Writer.WriteEndDocument()
                Writer.Flush()

                ' mark the in-process records as processed
                MyCommon.QueryStr = "dbo.pa_CustomerInquiry_MarkInfoProcessed"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@BatchGUID", SqlDbType.NVarChar, 36).Value = Left(BatchGUID, 36)
                MyCommon.LXSsp.Parameters.Add("@RecordsMarked", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                MyCommon.Close_LXSsp()

            Else
                ' Send back Invalid GUID return code
                ErrorXml = GetCustomerChangesErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXml = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXml)
            End If

        Catch ex As Exception

            MyCommon.Write_Log(LogFile, String.Format("Caught Exception: {0}", ex), True)

            ' rollback this batch and any previous records that were abandoned in the in-process state
            Try
                MyCommon.QueryStr = "dbo.pa_CustomerInquiry_MarkInfoAsReady"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@BatchGUID", SqlDbType.NVarChar, 36).Value = Left(BatchGUID, 36)
                MyCommon.LXSsp.Parameters.Add("@RecsRolledback", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                RecCt = MyCommon.LXSsp.Parameters("@RecsRolledback").Value
                MyCommon.Close_LXSsp()
                If RecCt > 0 Then
                    MyCommon.Write_Log(LogFile, "Error encountered during processing " & RecCt & " records were rolled back for future processing.", True)
                End If
            Catch ex2 As Exception
                MyCommon.Write_Log(LogFile, "Unable to roll back records for BatchGUID " & BatchGUID & vbCrLf & ex2.ToString, True)
            End Try

            ' send the application exception return code
            SessionXml = New XmlDocument()
            SessionXml.LoadXml(GetCustomerChangesErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString))
            Try
                MyCommon.Write_Log(LogFile, "Error encountering during GetCustomerUIChanges call " & _
                                            ControlChars.CrLf & " Reported Exception: " & ex.ToString, True)
            Catch ex2 As Exception
                'ignore
            End Try
        Finally
            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetCustomerRecordByCard(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim MyLookup As New Copient.CustomerLookup(1, 1)
        Dim CustomerPK As Long
        Dim CustRetCode As Copient.CustomerAbstract.RETURN_CODE = Copient.CustomerAbstract.RETURN_CODE.OK
        Dim ds As New DataSet()
        Dim dtCustInfo As New DataTable
        Dim dtCardInfo As New DataTable
        Dim row As DataRow
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim ErrorXML As String = ""
        Dim bReturnSupps As Boolean = False
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        Dim iCardTypeID As Integer = -1
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "GetCustomerRecord") Then
                If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(266)) = 1) Then
                    bReturnSupps = True
                End If

                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("Customers")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
        
                If IsValidCustomerCard(CardID, iCardTypeID, RetCode, RetMsg) Then
                    ' Padding the card id based on the cardtype from CardTypes database table 
                    CardID = MyCommon.Pad_ExtCardID(CardID, iCardTypeID)
                    ' load the marked batch of in-process records
                    CustomerPK = MyLookup.GetCustomerPK(CardID, iCardTypeID, CustRetCode)
        
                    If CustomerPK > 0 Then
                        MyCommon.QueryStr = "dbo.pa_CustomerInquiry_FetchCustRecord"
                        MyCommon.Open_LXSsp()
                        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
                        reader = MyCommon.LXSsp.ExecuteReader

                        ds.Tables.Add(dtCustInfo)
                        ds.Tables.Add(dtCardInfo)

                        ds.Load(reader, LoadOption.OverwriteChanges, Nothing, New DataTable() {dtCustInfo, dtCardInfo})

                        MyCommon.Close_LXSsp()
                        reader.Close()

                        Writer.WriteAttributeString("recordCount", dtCustInfo.Rows.Count)
                        Writer.WriteAttributeString("hasMore", "True")
                        Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

                        If dtCustInfo.Rows.Count > 0 Then
                            ' load the customer data
                            For Each row In dtCustInfo.Rows
                                WriteCustomerRecord(Writer, row, dtCardInfo, bReturnSupps)
                            Next
                        End If

                        Writer.WriteEndElement() ' end customers
                        Writer.WriteEndDocument()
                        Writer.Flush()
                    Else
                        ' customer not found
                        ErrorXML = GetCustomerChangesErrorXML(StatusCodes.NOTFOUND_CUSTOMER, "CardID: " & CardID & " With CardTypeID: " & CardTypeID & " not found.")
                    End If
                Else
                    ErrorXML = GetCustomerChangesErrorXML(RetCode, RetMsg)
                End If
            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetCustomerChangesErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception

            ' send the application exception return code
            SessionXml = New XmlDocument()
            SessionXml.LoadXml(GetCustomerChangesErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString))
            Try
                MyCommon.Write_Log(LogFile, "Error encountering during GetCustomerRecordByCard call " & _
                                            ControlChars.CrLf & " Reported Exception: " & ex.ToString, True)
            Catch ex2 As Exception
                'ignore
            End Try
        Finally
            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    <WebMethod()> _
    Public Function GetCustomerRecordWithCardTypes(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String) As XmlDocument
        Dim SessionXml As New XmlDocument()
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim MyLookup As New Copient.CustomerLookup(1, 1)
        Dim CustomerPK As Long
        Dim CustRetCode As Copient.CustomerAbstract.RETURN_CODE = Copient.CustomerAbstract.RETURN_CODE.OK
        Dim ds As New DataSet()
        Dim dtCustInfo As New DataTable
        Dim dtCardInfo As New DataTable
        Dim row As DataRow
        Dim reader As System.Data.SqlClient.SqlDataReader
        Dim ErrorXML As String = ""
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        Dim iCardTypeID As Integer = -1
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If IsValidGUID(GUID, "GetCustomerRecord") Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4

                Writer.WriteStartDocument()
                Writer.WriteStartElement("Customers")
                Writer.WriteAttributeString("returnCode", "SUCCESS")
                Writer.WriteAttributeString("responseTime", Date.Now.ToString("yyyy-MM-ddTHH:mm:ss"))
                
                CardID = AntiXss.AntiXssEncoder.HtmlEncode(CardID, True)
                CardTypeID = AntiXss.AntiXssEncoder.HtmlEncode(CardTypeID, True)
                If IsValidCustomerCard(CardID, iCardTypeID, RetCode, RetMsg) Then
                    ' Padding the card id based on the cardtype from CardTypes database table 
                    CardID = MyCommon.Pad_ExtCardID(CardID, CardTypeID)
                    ' load the marked batch of in-process records
                    CustomerPK = MyLookup.GetCustomerPK(CardID, CardTypeID, CustRetCode)
        
                    If CustomerPK > 0 Then
                        MyCommon.QueryStr = "dbo.pa_CustomerInquiry_FetchCustRecordWithCardtypes"
                        MyCommon.Open_LXSsp()
                        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
                        reader = MyCommon.LXSsp.ExecuteReader

                        ds.Tables.Add(dtCustInfo)
                        ds.Tables.Add(dtCardInfo)

                        ds.Load(reader, LoadOption.OverwriteChanges, Nothing, New DataTable() {dtCustInfo, dtCardInfo})

                        MyCommon.Close_LXSsp()
                        reader.Close()

                        Writer.WriteAttributeString("recordCount", dtCustInfo.Rows.Count)
                        Writer.WriteAttributeString("hasMore", "true")
                        Writer.WriteAttributeString("xmlns", "xsi", Nothing, "http://www.w3.org/2001/XMLSchema-instance")

                        If dtCustInfo.Rows.Count > 0 Then
                            ' load the customer data
                            For Each row In dtCustInfo.Rows
                                WriteCustomerRecord(Writer, row, dtCardInfo, False)
                            Next
                        End If

                        Writer.WriteEndElement() ' end customers
                        Writer.WriteEndDocument()
                        Writer.Flush()
                    Else
                        ' customer not found
                        ErrorXML = GetCustomerChangesErrorXML(StatusCodes.NOTFOUND_CUSTOMER, "Card Number: " & CardID & " Of Card Type: " & CardTypeID & " not found.")
                    End If
                Else
                    ErrorXML = GetCustomerChangesErrorXML(RetCode, RetMsg)
                End If
            Else
                ' Send back Invalid GUID return code
                ErrorXML = GetCustomerChangesErrorXML(StatusCodes.INVALID_GUID, "GUID: " & GUID & " is invalid")
            End If

            If ErrorXML = "" Then
                ms.Seek(0, SeekOrigin.Begin)
                SessionXml.Load(ms)
            Else
                SessionXml.LoadXml(ErrorXML)
            End If

        Catch ex As Exception

            ' send the application exception return code
            SessionXml = New XmlDocument()
            SessionXml.LoadXml(GetCustomerChangesErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString))
            Try
                MyCommon.Write_Log(LogFile, "Error encountering during GetCustomerRecordByCard call " & _
                                            ControlChars.CrLf & " Reported Exception: " & ex.ToString, True)
            Catch ex2 As Exception
                'ignore
            End Try
        Finally
            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return SessionXml
    End Function

    Private Sub writeCardElement(ByRef Writer As XmlTextWriter, ByRef card_row As DataRow)

        Writer.WriteStartElement("Card")
        Writer.WriteAttributeString("id", MyCryptlib.SQL_StringDecrypt(card_row.Item("ExtCardID").ToString()))
        Writer.WriteAttributeString("type", MyCommon.NZ(card_row.Item("CardType"), ""))
        Writer.WriteAttributeString("status", MyCommon.NZ(card_row.Item("CardStatus"), ""))
        Writer.WriteEndElement() ' Card

    End Sub

    Private Sub WriteCustomerRecord_CardsElement(ByRef Writer As XmlTextWriter, ByRef dtCardInfo As DataTable, ByRef customerPK As String)

        Dim cardSelectString As String = String.Format(" CustomerPK = {0}", customerPK)
        Dim cardRows() As DataRow = dtCardInfo.Select(cardSelectString)
        If cardRows.Length > 0 Then
            Writer.WriteStartElement("Cards")
            For Each r As DataRow In cardRows
                writeCardElement(Writer, r)
            Next
            Writer.WriteEndElement() ' Cards
        End If

    End Sub

    Private Sub WriteCustomerRecord_PasswordElement(ByRef Writer As XmlTextWriter, ByVal cust_password As String)

        If cust_password.Length > 0 Then
            Dim MyCryptLib As New Copient.CryptLib
            cust_password = MyCryptLib.SQL_StringDecrypt(cust_password)
        End If
        Writer.WriteElementString("Password", cust_password) ' WRITE THE USER'S CLEAR-TEXT PASSWORD!!!

    End Sub

    Private Sub WriteCustomerRecord(ByRef Writer As XmlTextWriter, ByVal row As DataRow, ByRef dtCardInfo As DataTable, ByVal bReturnSupps As Boolean)

        Dim MyLookup As New Copient.CustomerLookup
        Dim dt As DataTable
        Dim suppRow As Datarow

        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        Writer.WriteStartElement("Customer")
        Writer.WriteStartElement("GeneralInformation")
        Writer.WriteElementString("CustomerPK", MyCommon.NZ(row.Item("CustomerPK"), 0))
        Writer.WriteElementString("FirstName", MyCommon.NZ(row.Item("FirstName"), ""))
        Writer.WriteElementString("MiddleName", MyCommon.NZ(row.Item("MiddleName"), ""))
        Writer.WriteElementString("LastName", MyCommon.NZ(row.Item("LastName"), ""))
        Writer.WriteElementString("AltID", MyCommon.NZ(row.Item("AltID"), ""))
        Writer.WriteElementString("Verifier", MyCommon.NZ(row.Item("Verifier"), ""))
        Writer.WriteElementString("EmployeeID", MyCommon.NZ(row.Item("EmployeeID"), 0))
        Writer.WriteElementString("Employee", MyCommon.NZ(row.Item("Employee"), "false").ToString.ToLower)
        Writer.WriteElementString("TestCard", MyCommon.NZ(row.Item("TestCard"), "false").ToString.ToLower)
        Writer.WriteElementString("CustomerStatus", MyCommon.NZ(row.Item("CustomerStatus"), ""))
        Writer.WriteElementString("Household", IIf(MyCommon.NZ(row.Item("Household"), False), "true", "false"))

        Dim enroll As Date = MyLookup.ParseEnrollmentDate(MyCommon.NZ(row.Item("EnrollmentDate"), ""))
        If enroll = Nothing Then
            Writer.WriteStartElement("EnrollmentDate")
            Writer.WriteAttributeString("xsi", "nil", "http://www.w3.org/2001/XMLSchema-instance", "true")
            Writer.WriteEndElement() ' EnrollmentDate
        Else
            Writer.WriteElementString("EnrollmentDate", enroll.ToString("yyyy-MM-dd"))
        End If
        Writer.WriteElementString("Address", MyCommon.NZ(row.Item("Address"), ""))
        Writer.WriteElementString("City", MyCommon.NZ(row.Item("City"), ""))
        Writer.WriteElementString("State", MyCommon.NZ(row.Item("State"), ""))
        Writer.WriteElementString("ZIP", MyCommon.NZ(row.Item("Zip"), ""))
        Writer.WriteElementString("Country", MyCommon.NZ(row.Item("Country"), ""))
        Writer.WriteElementString("Phone", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("Phone"), "")))
        Writer.WriteElementString("MobilePhone", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("MobilePhone"), "")))
        Writer.WriteElementString("Email", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("email"), "")))

        Dim dob As Date = MyLookup.ParseDateOfBirth(MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("DOB"), "")))
        If dob = Nothing Then
            Writer.WriteStartElement("DateOfBirth")
            Writer.WriteAttributeString("xsi", "nil", "http://www.w3.org/2001/XMLSchema-instance", "true")
            Writer.WriteEndElement() ' DateOfBirth
        Else
            Writer.WriteElementString("DateOfBirth", dob.ToString("yyyy-MM-dd"))
        End If

        WriteCustomerRecord_PasswordElement(Writer, MyCommon.NZ(row.Item("Password"), ""))

        Writer.WriteElementString("DriverLicenseID", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("DriverLicenseID"), "")))
        Writer.WriteElementString("TaxExemptID", MyCryptlib.SQL_StringDecrypt(MyCommon.NZ(row.Item("TaxExemptID"), "")))
        Dim dateopened As Date = MyLookup.ParseDateOpened(MyCommon.NZ(row.Item("DateOpened"), ""))
        If dateopened = Nothing Then
            Writer.WriteStartElement("DateOpened")
            Writer.WriteAttributeString("xsi", "nil", "http://www.w3.org/2001/XMLSchema-instance", "true")
            Writer.WriteEndElement() ' DateOpened
        Else
            Writer.WriteElementString("DateOpened", dateopened.ToString("yyyy-MM-dd"))
        End If

        Writer.WriteElementString("ARCustomer", MyCommon.NZ(row.Item("ARCustomer"), "false").ToString.ToLower)
        Writer.WriteElementString("CompoundCharge", MyCommon.NZ(row.Item("CompoundCharge"), "false").ToString.ToLower)
        Writer.WriteElementString("FinanceCharge", MyCommon.NZ(row.Item("FinanceCharge"), "false").ToString.ToLower)
        Writer.WriteElementString("CreditLimit", MyCommon.NZ(row.Item("CreditLimit"), 0))
        Writer.WriteElementString("APR", MyCommon.NZ(row.Item("APR"), 0))

        'Writer.WriteElementString("Prefix", MyCommon.NZ(row.Item("Prefix"), ""))

        Writer.WriteEndElement() ' end GeneralInformation
        If bReturnSupps Then
            Writer.WriteStartElement("Supplementals")
            MyCommon.QueryStr = "dbo.pa_CustomerInquiry_FetchCustSupplementals"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 36).Value = row.Item("CustomerPK")
            MyCommon.LXSsp.Parameters.Add("@sortMethod", SqlDbType.Int).Value = MyCommon.Fetch_SystemOption(265)
            dt = MyCommon.LXSsp_select
            For Each suppRow In dt.Rows
                Writer.WriteStartElement("Supplemental")
                Writer.WriteAttributeString("Val", suppRow.Item("Value"))
                Writer.WriteAttributeString("ExtID", suppRow.Item("ExtFieldID"))
                Writer.WriteEndElement() ' Supplemental
            Next
            Writer.WriteEndElement() ' end Supplementals
        End If

        WriteCustomerRecord_CardsElement(Writer, dtCardInfo, MyCommon.NZ(row.Item("CustomerPK"), 0))

        Writer.WriteEndElement() ' end customer

    End Sub

    Private Sub WriteActivity(ByRef Writer As XmlTextWriter, ByVal row As DataRow, ByVal ActivitySubTypeID As Integer)
        Dim TempDate As New Date()
        Dim dt As DataTable

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

        MyCommon.QueryStr = "dbo.pa_ActivityExtLog_GetByActivity"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ActivityID", SqlDbType.NVarChar, 50).Value = MyCommon.NZ(row.Item("ActivityID"), 0)
        dt = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()

        If dt.Rows.Count > 0 Then
            Writer.WriteStartElement(GetTagName(ActivitySubTypeID, 0))

            If Not IsDBNull(row.Item("ActivityDate")) Then
                TempDate = row.Item("ActivityDate")
            End If
            Writer.WriteElementString("ActivityTime", TempDate.ToString("yyyy-MM-ddTHH:mm:ss"))

            ' the first returned row will be the offer the customer care agent clicked in Customer Inquiry
            Writer.WriteStartElement("SelectedOffer")
            Writer.WriteElementString("ID", MyCommon.NZ(dt.Rows(0).Item("OfferID"), 0))
            Writer.WriteElementString("Name", MyCommon.NZ(dt.Rows(0).Item("OfferName"), ""))
            Writer.WriteEndElement() ' end SelectedOffer

            Select Case ActivitySubTypeID
                Case 12, 13 ' points/sv adjust
                    WriteProgramTags(Writer, row, dt, ActivitySubTypeID)
                Case 14 ' accum adjust
                    WriteAdjustTags(Writer, row)
                Case 15, 16 ' add/remove offer
                    WriteCustGroupTags(Writer, row, dt, ActivitySubTypeID)
            End Select

            Writer.WriteEndElement()
        End If

    End Sub

    Private Sub PopulateTransferResults(ByRef TransferXmlDoc As XmlDocument, ByVal TransferCode As String, _
                                        ByVal RetMsg As String, ByVal TransferAmount As Long)
        Dim ResultsNode, TempNode, RootNode As XmlNode
        Dim ElemResults, ElemRetCode, ElemMsg, ElemTransferAmt As XmlElement

        ' build the Results tag, if it hasn't already been created
        ResultsNode = TransferXmlDoc.SelectSingleNode("/PointsTransfer/Results")
        If ResultsNode Is Nothing Then
            ' create the XML tags
            ElemResults = TransferXmlDoc.CreateElement("Results")
            ElemRetCode = TransferXmlDoc.CreateElement("ReturnCode")
            ElemMsg = TransferXmlDoc.CreateElement("Message")
            ' assign the tags to their parent
            ElemTransferAmt = TransferXmlDoc.CreateElement("TransferAmount")
            ElemResults.AppendChild(ElemRetCode)
            ElemResults.AppendChild(ElemMsg)
            ElemResults.AppendChild(ElemTransferAmt)

            RootNode = TransferXmlDoc.SelectSingleNode("/PointsTransfer")
            RootNode.AppendChild(ElemResults)
        End If

        ResultsNode = TransferXmlDoc.SelectSingleNode("/PointsTransfer/Results")
        If ResultsNode IsNot Nothing Then
            TempNode = ResultsNode.SelectSingleNode("ReturnCode")
            If TempNode IsNot Nothing Then TempNode.InnerText = TransferCode
            TempNode = ResultsNode.SelectSingleNode("Message")
            If TempNode IsNot Nothing Then TempNode.InnerText = RetMsg
            TempNode = ResultsNode.SelectSingleNode("TransferAmount")
            If TempNode IsNot Nothing Then TempNode.InnerText = TransferAmount
        End If

    End Sub

    Private Function GetTagName(ByVal ActivitySubTypeID As Integer, ByVal TagID As Integer) As String
        Dim TagName As String = ""
        Const TagUpperBound As Integer = 1
        Dim AddOfferTags As String() = {"AddToOffer", "CustomerGroup"}
        Dim RemoveOfferTags As String() = {"RemoveFromOffer", "CustomerGroup"}
        Dim PointsAdjustTags As String() = {"PointsAdjustment", "PointsProgram"}
        Dim SVAdjustTags As String() = {"StoredValueAdjustment", "StoredValueProgram"}
        Dim AccumAdjustTags As String() = {"AccumulationAdjustment", ""}

        If TagID >= 0 AndAlso TagID <= TagUpperBound Then
            Select Case ActivitySubTypeID
                Case 12 ' points adjust
                    TagName = PointsAdjustTags(TagID)
                Case 13 ' sv adjust
                    TagName = SVAdjustTags(TagID)
                Case 14 ' accum adjust
                    TagName = AccumAdjustTags(TagID)
                Case 15 ' add offer
                    TagName = AddOfferTags(TagID)
                Case 16 ' remove offer
                    TagName = RemoveOfferTags(TagID)
            End Select
        End If

        Return TagName
    End Function

    Private Sub WriteAdjustTags(ByRef Writer As XmlTextWriter, ByVal row As DataRow)
        If Writer IsNot Nothing AndAlso row IsNot Nothing Then
            Writer.WriteElementString("PreAdjustBalance", MyCommon.NZ(row.Item("PreAdjustBalance"), 0))
            Writer.WriteElementString("Adjustment", MyCommon.NZ(row.Item("Adjustment"), 0))
            Writer.WriteElementString("PostAdjustBalance", MyCommon.NZ(row.Item("PostAdjustBalance"), 0))
        End If
    End Sub

    Private Sub WriteCustGroupTags(ByRef Writer As XmlTextWriter, ByVal row As DataRow, ByVal dtAssoc As DataTable, ByVal ActivitySubTypeID As Integer)
        Dim dt As DataTable
        Dim CgName As String = ""

        If Writer IsNot Nothing AndAlso row IsNot Nothing Then
            Writer.WriteStartElement(GetTagName(ActivitySubTypeID, 1))

            ' get the customer groups name
            MyCommon.QueryStr = "select Name from CustomerGroups with (NoLock) where CustomerGroupID= @CustomerGroupId"
            MyCommon.DBParameters.Add("@CustomerGroupId", SqlDbType.BigInt).Value = MyCommon.NZ(row.Item("LinkID2"), 0)
            dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                CgName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
            End If

            Writer.WriteElementString("ID", MyCommon.NZ(row.Item("LinkID2"), 0))
            Writer.WriteElementString("Name", CgName)
            WriteAssocOfferTags(Writer, dtAssoc, ActivitySubTypeID)

            Writer.WriteEndElement() ' ends the TagName element
        End If

    End Sub

    Private Sub WriteProgramTags(ByRef Writer As XmlTextWriter, ByVal row As DataRow, ByVal dtAssoc As DataTable, ByVal ActivitySubTypeID As Integer)
        Dim dt As DataTable
        Dim ProgName As String = ""

        If Writer IsNot Nothing AndAlso row IsNot Nothing Then
            Writer.WriteStartElement(GetTagName(ActivitySubTypeID, 1))

            ' get the program name
            If ActivitySubTypeID = 12 Then
                MyCommon.QueryStr = "select ProgramName from PointsPrograms with (NoLock) where ProgramID= @ProgramId"
                MyCommon.DBParameters.Add("@ProgramId", SqlDbType.BigInt).Value = MyCommon.NZ(row.Item("LinkID2"), 0)
            Else
                MyCommon.QueryStr = "select Name as ProgramName from StoredValuePrograms with (NoLock) where SVProgramID= @ProgramId"
                MyCommon.DBParameters.Add("@ProgramId", SqlDbType.BigInt).Value = MyCommon.NZ(row.Item("LinkID2"), 0)
            End If
            dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                ProgName = MyCommon.NZ(dt.Rows(0).Item("ProgramName"), "")
            End If

            Writer.WriteElementString("ID", MyCommon.NZ(row.Item("LinkID2"), 0))
            Writer.WriteElementString("Name", ProgName)
            WriteAdjustTags(Writer, row)
            WriteAssocOfferTags(Writer, dtAssoc, ActivitySubTypeID)

            Writer.WriteEndElement() ' ends the TagName element
        End If

    End Sub

    Private Sub WriteAssocOfferTags(ByRef Writer As XmlTextWriter, ByVal dtAssoc As DataTable, ByVal ActivitySubTypeID As Integer)
        Dim rowOffer As DataRow

        If Writer IsNot Nothing AndAlso dtAssoc IsNot Nothing Then
            Writer.WriteStartElement("AssociatedOffers")

            For Each rowOffer In dtAssoc.Rows
                Writer.WriteStartElement("Offer")
                Writer.WriteElementString("ID", MyCommon.NZ(rowOffer.Item("OfferID"), 0))
                Writer.WriteElementString("Name", MyCommon.NZ(rowOffer.Item("OfferName"), ""))
                Writer.WriteEndElement() ' end offer
            Next

            Writer.WriteEndElement() ' end AssociatedOffers
        End If
    End Sub

    Private Function GetErrorXML(ByVal Code As StatusCodes, ByVal SessionID As String, Optional ByVal CustomerPK As Long = 0) As String
        Dim ErrorXml As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")

        ErrorXml.Append("<CustomerContact returnCode=""")
        Select Case Code
            Case StatusCodes.INVALID_GUID
                ErrorXml.Append("INVALID_GUID")
            Case StatusCodes.INVALID_CUSTOMERID
                ErrorXml.Append("INVALID_CUSTOMERID")
            Case StatusCodes.NO_ACTIVITY_FOUND_FOR_SESSION_ID
                ErrorXml.Append("NO_ACTIVITY_FOUND_FOR_SESSION_ID")
            Case Else
                ' treat everything else as an application exception
                ErrorXml.Append("APPLICATION_EXCEPTION")
        End Select
        ErrorXml.Append(""" responseTime=""" & Date.Now.ToString("yyyy-MM-ddTHH:mm:ss") & """>")
        ErrorXml.Append(" <Contact>")
        ErrorXml.Append("    <ID>" & SessionID & "</ID>")
        ErrorXml.Append("    <Customer>")
        ErrorXml.Append("      <CustomerID>" & IIf(CustomerPK > 0, "PK: " & CustomerPK, "") & "</CustomerID>")
        ErrorXml.Append("      <CustomerTypeID>0</CustomerTypeID>")
        ErrorXml.Append("    </Customer>")
        ErrorXml.Append("    <User>")
        ErrorXml.Append("      <ID></ID>")
        ErrorXml.Append("      <UserName></UserName>")
        ErrorXml.Append("      <FirstName></FirstName>")
        ErrorXml.Append("      <LastName></LastName>")
        ErrorXml.Append("    </User>")
        ErrorXml.Append("  </Contact>")
        ErrorXml.Append("</CustomerContact>")

        Return ErrorXml.ToString
    End Function

    Private Function GetCustomerChangesErrorXML(ByVal Code As StatusCodes, Optional ByVal ErrorMsg As String = "") As String
        Dim ErrorXml As New StringBuilder("<?xml version=""1.0"" encoding=""utf-8""?>")

        ErrorXml.Append("<Customers returnCode=""")
        Select Case Code
            Case StatusCodes.INVALID_GUID
                ErrorXml.Append("INVALID_GUID")
            Case StatusCodes.NOTFOUND_CUSTOMER
                ErrorXml.Append("CUSTOMER_NOT_FOUND")
            Case StatusCodes.INVALID_CUSTOMERTYPEID
                ErrorXml.Append("CUSTOMER_NOT_FOUND")
            Case Else
                ' treat everything else as an application exception
                ErrorXml.Append("APPLICATION_EXCEPTION")
        End Select
        ErrorXml.Append(""" responseTime=""" & Date.Now.ToString("yyyy-MM-ddTHH:mm:ss") & """ " & _
                        " recordCount=""0"" hasMore=""false"">")
        ErrorXml.Append("  <ErrorMessage>" & ErrorMsg & "</ErrorMessage>")
        ErrorXml.Append("</Customers>")

        Return ErrorXml.ToString
    End Function

    Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String) As Boolean
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc
        Dim MsgBuf As New StringBuilder()

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 2, GUID)
        Catch ex As Exception
            IsValid = False
        End Try

        ' Log the call
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MsgBuf.Append(IIf(IsValid, "Validated call to ", "Invalid call to "))
            MsgBuf.Append(MethodName)
            MsgBuf.Append(" from GUID: ")
            MsgBuf.Append(GUID)
            MsgBuf.Append(" and IP: " & HttpContext.Current.Request.UserHostAddress)
            MyCommon.Write_Log(LogFile, MsgBuf.ToString, True)
        Catch ex As Exception
            ' ignore
        End Try

        Return IsValid
    End Function

    Private Function IsValidCustomerCard(ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByRef RetCode As StatusCodes, ByRef RetMsg As String) As Boolean
        Dim IsValid As Boolean = False
        Dim objErrorResponse As CardValidationResponse
        RetCode = StatusCodes.SUCCESS
        RetMsg = ""
        Try
            If (MyCommon.AllowToProcessCustomerCard(ExtCardID, CardTypeID, objErrorResponse) = False) Then
                If ExtCardID Is Nothing OrElse ExtCardID.Trim = "" Then
                    'Bad customer ID
                    RetCode = StatusCodes.INVALID_CUSTOMERID
                    RetMsg = "Failure: ExtCardID is not provided"
                ElseIf Not String.IsNullOrEmpty(ExtCardID) AndAlso Not CardTypeID = -1 Then
                    RetCode = StatusCodes.NOTFOUND_CUSTOMER
                    RetMsg = "CardID: " & ExtCardID & " with CardTypeID: " & CardTypeID & " not found."
                Else
                    'Bad customer type ID
                    RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                    RetMsg = "Failure: Invalid CardTypeID"
                End If
            Else
                IsValid = True
            End If

        Catch ex As Exception
            IsValid = False
        End Try
        Return IsValid
    End Function

    Private Function LookupProgramID(ByVal ProgramID As String, _
                                      ByRef RetCode As StatusCodes, ByRef RetMsg As String) As Long
        Dim ProgID As Integer = 0
        Dim LookupRetCode As Copient.CustomerAbstract.RETURN_CODE
        Dim MyLookup As New Copient.CustomerLookup

        ProgID = MyLookup.LookupProgramID(ProgramID, LookupRetCode)
        Select Case LookupRetCode
            Case Copient.CustomerLookup.RETURN_CODE.NOTFOUND_PROGRAMID
                RetCode = StatusCodes.NOTFOUND_PROGRAMID
                RetMsg = "The Program ID is not found."
            Case Copient.CustomerLookup.RETURN_CODE.INVALID_PROGRAMID
                RetCode = StatusCodes.INVALID_PROGRAMID
                RetMsg = "The Program ID is invalid."
            Case Copient.CustomerLookup.RETURN_CODE.APPLICATION_ERROR
                RetCode = StatusCodes.APPLICATION_EXCEPTION
                RetMsg = "Failure: Application Error."
        End Select
        Return ProgID
    End Function

    Private Function LookupPromoVarID(ByVal PromovarID As String, _
                                     ByRef RetCode As StatusCodes, ByRef RetMsg As String) As Long
        Dim PromotionVarID As Integer = 0
        Dim LookupRetCode As Copient.CustomerAbstract.RETURN_CODE
        Dim MyLookup As New Copient.CustomerLookup

        PromotionVarID = MyLookup.LookupPromoVarID(PromovarID, LookupRetCode)
        Select Case LookupRetCode
            Case Copient.CustomerLookup.RETURN_CODE.NOTFOUND_PROMOVARID
                RetCode = StatusCodes.NOTFOUND_PROMOVARID
                RetMsg = "The Promovar ID is not found."
            Case Copient.CustomerLookup.RETURN_CODE.INVALID_PROMOVARID
                RetCode = StatusCodes.INVALID_PROMOVARID
                RetMsg = "The Promovar ID is invalid."
            Case Copient.CustomerLookup.RETURN_CODE.APPLICATION_ERROR
                RetCode = StatusCodes.APPLICATION_EXCEPTION
                RetMsg = "Failure: Application Error."
        End Select
        Return PromotionVarID
    End Function

    Private Function LookupCustomerPK(ByVal CustomerID As String, ByVal CustomerTypeID As Integer, _
                                      ByRef RetCode As StatusCodes, ByRef RetMsg As String) As Long
        Dim CustomerPK As Long
        Dim LookupRetCode As Copient.CustomerAbstract.RETURN_CODE
        Dim MyLookup As New Copient.CustomerLookup

        CustomerPK = MyLookup.GetCustomerPK(CustomerID, CustomerTypeID, LookupRetCode)

        ' map the appropriate return code from the customer lookup to this web service's return code
        Select Case LookupRetCode
            Case Copient.CustomerLookup.RETURN_CODE.INVALID_CUSTOMERID
                RetCode = StatusCodes.INVALID_CUSTOMERID
                RetMsg = "The Customer ID is invalid."
            Case Copient.CustomerLookup.RETURN_CODE.INVALID_CUSTOMERTYPEID
                RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                RetMsg = "The Customer Type ID is invalid."
            Case Copient.CustomerLookup.RETURN_CODE.NOTFOUND_CUSTOMER
                RetCode = StatusCodes.NOTFOUND_CUSTOMER
                RetMsg = "The Customer ID is not found."
            Case Copient.CustomerLookup.RETURN_CODE.NOTFOUND_HOUSEHOLD
                RetCode = StatusCodes.NOTFOUND_HOUSEHOLD
                RetMsg = "The Household ID is not found."
            Case Copient.CustomerLookup.RETURN_CODE.NOTFOUND_CAM
                RetCode = StatusCodes.NOTFOUND_CAM
        End Select

        Return CustomerPK
    End Function

    Private Function GetPointsProgramTable() As Hashtable
        Dim PointsTable As New Hashtable
        Dim MyPoints As New Copient.Points
        Dim Criteria As New Copient.ListCriteria()
        Dim TotalRows As Integer
        Dim dt As DataTable
        Dim row As DataRow
        Dim ProgramID As Long
        Dim ProgramName As String

        dt = MyPoints.GetPointsProgramList(Criteria, TotalRows)
        If TotalRows > 0 Then
            For Each row In dt.Rows
                ProgramID = MyCommon.NZ(row.Item("ProgramID"), 0)
                ProgramName = MyCommon.NZ(row.Item("ProgramName"), "")
                If Not PointsTable.ContainsKey("ID:" & ProgramID) Then
                    PointsTable.Add("ID:" & ProgramID, ProgramName)
                End If
            Next
        End If

        Return PointsTable
    End Function

    Private Function GetSVProgramTable() As Hashtable
        Dim SVTable As New Hashtable
        Dim MySV As New Copient.StoredValue
        Dim Criteria As New Copient.ListCriteria()
        Dim TotalRows As Integer
        Dim dt As DataTable
        Dim row As DataRow
        Dim SVProgramID As Long
        Dim SVProgramName As String

        dt = MySV.GetSVProgramList(Criteria, TotalRows)
        If TotalRows > 0 Then
            For Each row In dt.Rows
                SVProgramID = MyCommon.NZ(row.Item("SVProgramID"), 0)
                SVProgramName = MyCommon.NZ(row.Item("Name"), "")
                If Not SVTable.ContainsKey("ID:" & SVProgramID) Then
                    SVTable.Add("ID:" & SVProgramID, SVProgramName)
                End If
            Next
        End If

        Return SVTable
    End Function

    Private Function IsValidSVAdjustAmount(ByVal SVProgramID As Long, ByRef AdjustAmount As Decimal, _
                                            ByRef RetCode As StatusCodes, ByRef RetMsg As String) As Boolean
        Dim ValidAmt As Boolean = False
        Dim UnitValue As Decimal
        Dim UnitType As Integer
        Dim SVProg As New Copient.StoredValueProgram
        Dim MySV As New Copient.StoredValue()

        SVProg = MySV.GetStoredValueProgram(SVProgramID)

        If SVProg IsNot Nothing Then
            UnitValue = SVProg.GetValue
            UnitType = SVProg.GetSVTypeID

            If AdjustAmount = 0 Then
                RetCode = StatusCodes.NON_POSITIVE_SV_ADJUST_AMOUNT
                RetMsg = Copient.PhraseLib.Lookup("sv-adjust.nozeroadjustments", 1)
            ElseIf MyCommon.Fetch_SystemOption(100) <> "1" AndAlso AdjustAmount < 0 Then
                RetCode = StatusCodes.NON_POSITIVE_SV_ADJUST_AMOUNT
                RetMsg = Copient.PhraseLib.Lookup("sv-adjust.mustrevoke", 1)
            ElseIf UnitType = 1 AndAlso (Decimal.ToInt32(AdjustAmount) Mod Decimal.ToInt32(UnitValue)) <> 0 Then
                RetCode = StatusCodes.INVALID_SV_MULITPLE
                RetMsg = "Adjustment amount parameter value of " & AdjustAmount & " is not an even multiple of the unit value (" & UnitValue & ")."
            ElseIf UnitType <> 1 AndAlso (Math.Abs(AdjustAmount) Mod UnitValue) <> 0 Then
                RetCode = StatusCodes.INVALID_SV_MULITPLE
                RetMsg = "Adjustment amount parameter value of " & AdjustAmount & " is not an even multiple of the unit value (" & UnitValue & ")."
            ElseIf Math.Abs(AdjustAmount) < UnitValue Then
                RetCode = StatusCodes.INVALID_SV_MULITPLE
                RetMsg = "Adjustment amount parameter value of " & AdjustAmount & " is not an even multiple of the unit value (" & UnitValue & ")."
            Else
                ' convert the adjustment amount into units
                AdjustAmount = CInt(AdjustAmount / UnitValue)
                ValidAmt = True
            End If
        Else
            RetCode = StatusCodes.INVALID_STORED_VALUE_PROGRAM
            RetMsg = "Stored Value program ID " & SVProgramID & " was not found."
        End If

        Return ValidAmt
    End Function

    Private Function IsValidXmlDocument(ByVal sXsdFileName As String, ByVal XmlDoc As XmlDocument) As Boolean
        Dim Settings As XmlReaderSettings
        Dim xr As XmlReader = Nothing
        Dim ms As New MemoryStream()
        Dim sMsg As String = ""
        Dim bValid As Boolean = True
        Dim xsdPath As String

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            xsdPath = MyCommon.Get_Install_Path & "AgentFiles\" & sXsdFileName
            If sXsdFileName <> "" AndAlso XmlDoc IsNot Nothing Then
                Settings = New XmlReaderSettings()
                Settings.Schemas.Add(Nothing, xsdPath)
                Settings.ValidationType = ValidationType.Schema
                Settings.IgnoreComments = True
                Settings.IgnoreProcessingInstructions = True
                Settings.IgnoreWhitespace = True

                XmlDoc.Save(ms)
                ms.Seek(0, SeekOrigin.Begin)

                xr = XmlReader.Create(ms, Settings)
                Do While (xr.Read())
                    'Console.WriteLine("NodeType: " & xr.NodeType.ToString & " - " & xr.LocalName & " Depth: " & xr.Depth.ToString)
                Loop
                bValid = True
            Else
                bValid = False
            End If
        Catch eXmlSch As XmlSchemaException
            sMsg = "(Xml Schema Validation Error Line: " & eXmlSch.LineNumber.ToString & " - Col: " & eXmlSch.LinePosition.ToString & ") " & eXmlSch.Message
            bValid = False
        Catch eXml As XmlException
            sMsg = "(Xml Error Line: " & eXml.LineNumber.ToString & " - Col: " & eXml.LinePosition.ToString & ") " & eXml.Message
            bValid = False
        Catch exApp As ApplicationException
            sMsg = "Application Error: " & exApp.ToString
            bValid = False
        Catch ex As Exception
            sMsg = "Error: " & ex.ToString
            bValid = False
        Finally
            If Not xr Is Nothing Then
                xr.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If
        End Try

        ' Log Error if one exists
        If (sMsg <> "") Then
            MyCommon.Write_Log(LogFile, sMsg, True)
        End If

        Return bValid
    End Function

    <WebMethod()> _
    Public Function CustomerUpdate(ByVal GUID As String, ByVal CashierID As String, ByVal StoreID As String, ByVal CustomerXML As String) As String
        Dim MyCommon As New Copient.CommonInc
        Dim MyMassUpdate As Copient.MassUpdate
        Dim MyLookup As New Copient.CustomerLookup()
        Dim Parameters As New Copient.MassUpdate.ProcessingInstructions
        Dim RetXmlDoc As New XmlDocument
        Dim RetXmlStr As String = ""
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerXmlDoc As New XmlDocument()
        Dim ConnInc As New Copient.ConnectorInc
        Dim TagCt As Integer = 0
        Dim CustomerPK As Integer = 0
        Dim ExtCardID As String = ""
        Dim ExtCardTypeID As String = ""
        Dim CardTypeID As Integer = -1
        Dim Fields As New Copient.CommonInc.ActivityLogFields
        Dim objErrorResponse As CardValidationResponse

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            MyMassUpdate = New Copient.MassUpdate(MyCommon, Copient.MassUpdate.CALLER.CUSTOMER_INQUIRY_WS)
            MyMassUpdate.EnableLogging(LogFile)

            ' validate the request
            If Not IsValidGUID(GUID, "Update") Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "GUID " & GUID & " is not valid for the Customer Inquiry web service."
            ElseIf Not ConnInc.ConvertStringToXML(CustomerXML, CustomerXmlDoc) Then
                RetCode = StatusCodes.INVALID_XML_DOCUMENT
                RetMsg = "CustomerXML parameter is not a valid XML Document"
            ElseIf Not IsValidTagCount(CustomerXmlDoc, TagCt) Then
                RetCode = StatusCodes.OPERATION_TAG_LIMIT_EXCEEDED
                RetMsg = "Tag limit of " & OPERATION_TAG_LIMIT & " was exceeded. This request includes " & TagCt & " operational tags."
            End If
            MyCommon.Write_Log(LogFile, "TagCt: " & TagCt, True)

            If RetCode = StatusCodes.SUCCESS Then
                Parameters.Mode = Copient.MassUpdate.PROCESS_MODE.XML_STRING
                Parameters.CustomerXML = CustomerXmlDoc
                Parameters.Caller = Copient.MassUpdate.CALLER.CUSTOMER_INQUIRY_WS

                'Get the ExtCardID and ExtCardTypeID from the XML, then use them to find the CustomerPK.
                TryParseElementValue(Parameters.CustomerXML, "//CardID", ExtCardID) '//CustomerInquiry/UpdateCustomer/CardID
                TryParseElementValue(Parameters.CustomerXML, "//ExtCardTypeID", ExtCardTypeID) '//CustomerInquiry/UpdateCustomer/ExtCardTypeID
                If (MyCommon.AllowToProcessCustomerCard(ExtCardID, ExtCardTypeID, objErrorResponse) = False) Then
                    RetXmlStr = GetErrorXML(StatusCodes.INVALID_CUSTOMERID, MyCommon.CardValidationResponseMessage(ExtCardID, ExtCardTypeID, objErrorResponse))
                    Return RetXmlStr
                End If
                CardTypeID = MyLookup.FindCardTypeID(ExtCardTypeID)
                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                CustomerPK = MyLookup.FindCustomerPKFromExtID(ExtCardID, ExtCardTypeID)

                'Process the XML
                If (CustomerPK > 0) OrElse (CustomerPK = 0 AndAlso MyCommon.Fetch_SystemOption(103)) Then
                    RetXmlDoc = MyMassUpdate.ProcessCustomerXML(Parameters)
                    If RetXmlDoc IsNot Nothing Then
                        RetXmlStr = RetXmlDoc.OuterXml
                        'Log activity
                        Fields.ActivityTypeID = 25
                        Fields.ActivitySubTypeID = 15
                        Fields.LinkID = CustomerPK
                        Fields.AdminUserID = 1
                        Fields.Description = Copient.PhraseLib.Lookup("history.customer-edited-info", 1)
                        Fields.LinkID3 = CashierID
                        Fields.LinkID4 = StoreID
                        Fields.LinkID5 = ExtCardID
                        Fields.LinkID6 = ExtCardTypeID
                        MyCommon.Activity_Log3(Fields)
                    Else
                        RetXmlStr = GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, "Error encountered during processing customer XML")
                    End If
                Else
                    RetXmlStr = GetErrorXML(StatusCodes.INVALID_CUSTOMERID, "Cannot find the specified customer.")
                End If
            Else
                RetXmlStr = GetErrorXML(RetCode, RetMsg)
                MyCommon.Write_Log(LogFile, "Error:  " & RetCode.ToString & " - " & RetMsg, True)
            End If
        Catch argEx As ArgumentException
            Select Case argEx.ParamName.ToUpper
                Case "CUSTOMERXML"
                    RetXmlStr = GetErrorXML(StatusCodes.INVALID_XML_DOCUMENT, argEx.Message)
                Case Else
                    RetXmlStr = GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, argEx.ToString)
            End Select
        Catch ex As Exception
            RetXmlStr = GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString)
            MyCommon.Write_Log(LogFile, "Error:  " & StatusCodes.APPLICATION_EXCEPTION.ToString & " - " & ex.ToString, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return RetXmlStr
    End Function

    Private Function IsValidTagCount(ByRef CustomerXMLDoc As XmlDocument, ByRef TagCount As Integer) As Boolean
        Dim RootNode As XmlNode = Nothing
        Dim ValidCount As Boolean = False

        TagCount = 0
        If CustomerXMLDoc IsNot Nothing Then
            RootNode = CustomerXMLDoc.SelectSingleNode("/CustomerUpdate")

            If RootNode IsNot Nothing Then
                TagCount = RootNode.ChildNodes.Count
            End If
        End If

        ValidCount = (TagCount <= OPERATION_TAG_LIMIT)

        Return ValidCount
    End Function

    Private Function TryParseElementValue(ByVal OfferXmlDoc As XmlDocument, ByVal XPath As String, ByRef ElementValue As String) As Boolean
        Dim SingleNode As XmlNode = Nothing
        Dim IsParsed As Boolean = False

        SingleNode = OfferXmlDoc.SelectSingleNode(XPath)
        If (SingleNode IsNot Nothing) Then
            ElementValue = SingleNode.InnerText
            IsParsed = True
        End If

        Return IsParsed
    End Function

    Private Function TryParseAttributeValue(ByVal OfferXmlDoc As XmlDocument, ByVal ElementName As String, _
                                            ByVal AttributeName As String, ByRef AttributeValue As String) As Boolean
        Dim SingleNode As XmlNode = Nothing
        Dim Attrib As XmlAttribute = Nothing
        Dim IsParsed As Boolean = False

        SingleNode = OfferXmlDoc.SelectSingleNode(ElementName)
        If (SingleNode IsNot Nothing) Then
            Attrib = SingleNode.Attributes(AttributeName)
            If Attrib IsNot Nothing Then
                AttributeValue = Attrib.InnerText
                IsParsed = True
            End If
        End If

        Return IsParsed
    End Function

    <WebMethod()> _
    Public Function ReportCardLostStolen(ByVal GUID As String, ByVal CardTypeID As String, ByVal ExtCardID As String) As String
        If IsNumeric(CardTypeID) And (CardTypeID <> "") Then
            CardTypeID = Integer.Parse(CardTypeID)
        Else
            CardTypeID = -1
        End If
        Return MarkCustomerCardLostStolen(GUID, CardTypeID, ExtCardID)
    End Function

    Private Function MarkCustomerCardLostStolen(ByVal GUID As String, ByVal CardTypeID As Integer, ByVal ExtCardID As String) As String
        Dim RetMsg As String = "Success"
        Dim MethodName As String
        Dim dt As DataTable
        Try
            'Establish a connection to the LogixXS database
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MethodName = "MarkCustomerCardLostStolen"
            'Check if the GUID is valid for Customer Inquiry
            If IsValidGUID(GUID, MethodName) And (CardTypeID <> -1) Then
                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                'Check if Card exists
                MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID= @ExtCardID AND CardTypeID = @CardTypeID "
                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                If dt.Rows.Count > 0 Then
                    'Set the new status in the CardIDs table
                    MyCommon.QueryStr = "Update CardIDs with (RowLock) SET CardStatusID = 5 WHERE ExtCardID = @ExtCardID  AND CardTypeID = @CardTypeID "
                    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    MyCommon.ExecuteNonQuery(DataBases.LogixXS)
                    RetMsg = "Success"
                Else 'If row count is not greater than 0, it means that the record was not found.
                    RetMsg = "Card not found"
                End If
            Else
                If (CardTypeID = -1) Then
                    RetMsg = "Invalid CardTypeID"
                Else
                    'Wrong GUID
                    RetMsg = "Invalid GUID"
                End If
            End If
        Catch ex As Exception
        Finally
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Return RetMsg
    End Function

    <WebMethod()> _
    Public Function BlockCard(ByVal GUID As String, ByVal CardTypeID As Integer, ByVal ExtCardID As String) As String
        Return MarkCustomerCardInactive(GUID, CardTypeID, ExtCardID)
    End Function

    'updated the method to add 14digit validation for AltIDs [AMS - Ahold enhancement]
    Private Function MarkCustomerCardInactive(ByVal GUID As String, ByVal CardTypeID As Integer, ByVal ExtCardID As String) As String

        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = "Success"
        Dim MethodName As String
        Dim dt As DataTable

        Try
            'Establish a connection to the LogixXS database
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MethodName = "MarkCustomerCardInactive"
            'Check if the GUID is valid for Customer Inquiry
            If IsValidGUID(GUID, MethodName) Then

                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                'Check if Card exists
                MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID = @ExtCardID AND CardTypeID = @CardTypeID "
                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                If dt.Rows.Count > 0 Then
                    'Set the new status in the CardIDs table
                    MyCommon.QueryStr = "Update CardIDs with (RowLock) SET CardStatusID=2 WHERE ExtCardID = @ExtCardID AND CardTypeID = @CardTypeID "
                    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    MyCommon.ExecuteNonQuery(DataBases.LogixXS)
                    RetMsg = "Success"
                Else 'If row count is not greater than 0, it means that the record was not found.
                    RetMsg = "Card not found"
                End If
            Else
                'Wrong GUID
                RetMsg = "Invalid GUID"
            End If
        Catch ex As Exception
        Finally
            'Close the connection to the database
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Return RetMsg
    End Function

    '"Ahold AMS Integration-PayVantage FIS Rev B" - <2.3.3> : Modify the output type as dataset
    <WebMethod()> _
    Public Function UpdateCustomerAttributeValue(ByVal GUID As String, ByVal CardTypeID As String, ByVal ExtCardID As String, _
                                              ByVal AttributeTypeDesc As String, ByVal AttributeValueExtID As String) As DataSet
        Dim iCardTypeId As Integer = -1
        Try
            iCardTypeId = CInt(CardTypeID)
        Catch ex As Exception
            iCardTypeId = -1
        End Try
        Return UpdateMemberAttributeValue(GUID, iCardTypeId, ExtCardID, AttributeTypeDesc, AttributeValueExtID)
    End Function

    Private Function UpdateMemberAttributeValue(ByVal GUID As String, ByVal CardTypeID As Integer, ByVal ExtCardID As String, _
                                                ByVal AttributeTypeDesc As String, ByVal AttributeValExtID As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UpdateMemberAttributeValue") 'To return output as dataset
        Dim dtStatus As System.Data.DataTable
        Dim dt As DataTable
        Dim dt_oldValue As DataTable
        Dim dt_oldDesc As DataTable
        Dim MethodName As String
        Dim CustomerPK As Long = 0
        Dim MyCustAttrib As New Copient.CustomerAttribute
        Dim AttributeTypeID As Integer = 0
        Dim AttributeValueID As Integer = 0
        Dim row As System.Data.DataRow
        Dim bOpenedRTConnection As Boolean = False
        Dim bOpenedXSConnection As Boolean = False
        Dim MyMassUpdate As Copient.MassUpdate
        Dim bUpdated As Boolean = False
        Dim AttributeValueDescription As String
        Dim Prev_AttributeValueDescription As String = ""

        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bOpenedRTConnection = True
            End If
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixXS()
                bOpenedXSConnection = True
            End If
            If CardTypeID = -1 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                row.Item("Description") = "Failure: Invalid CardTypeID"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                UpdateMemberAttributeValue = ResultSet
                Exit Function
            End If
            If ExtCardID = "" Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: CustomerPK not found"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                UpdateMemberAttributeValue = ResultSet
                Exit Function
            End If
            If AttributeTypeDesc = "" Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_ATTRIBUTETYPE
                row.Item("Description") = "Failure: Attribute Type not found"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                UpdateMemberAttributeValue = ResultSet
                Exit Function
            End If
            If AttributeValExtID = "" Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_ATTRIBUTEVALUE
                row.Item("Description") = "Failure: Attribute Value not found"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                UpdateMemberAttributeValue = ResultSet
                Exit Function
            End If
            MethodName = "UpdateMemberAttributeValue"
            '        'Check if the GUID is valid for Customer Inquiry
            If IsValidGUID(GUID, MethodName) Then
                'Pad the customer ID

                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                MyCommon.QueryStr = "select top 1 CustomerPK from CardIDs with (NoLock) where ExtCardID= @ExtCardID AND CardTypeID = @CardTypeID "
                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
                MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                If dt.Rows.Count = 1 Then 'if CustomerPK exists
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    MyCommon.QueryStr = "select top 1 AttributeTypeID from AttributeTypes with (NoLock) where Description= @AttributeTypeDesc and Deleted=0 ;"
                    MyCommon.DBParameters.Add("@AttributeTypeDesc", SqlDbType.NVarChar).Value = (AttributeTypeDesc.ConvertBlankIfNothing()).Trim()
                    dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
                    If dt.Rows.Count = 1 Then 'Attrib Type Found
                        'Get the AttributeTypeID
                        AttributeTypeID = MyCommon.NZ(dt.Rows(0).Item("AttributeTypeID"), 0)
                        'Get the ID for the AttributeValue ExtID that was passed in
                        MyCommon.QueryStr = "select top 1 AttributeValueID,Description from AttributeValues AV with (NoLock) where AV.ExtID= @AttributeValExtID AND " & _
                                                        " AV.AttributeTypeID = @AttributeTypeID and AV.Deleted=0 ;"
                        MyCommon.DBParameters.Add("@AttributeValExtID", SqlDbType.NVarChar).Value = (AttributeValExtID.ConvertBlankIfNothing()).Trim()
                        MyCommon.DBParameters.Add("@AttributeTypeID", SqlDbType.Int).Value = AttributeTypeID
                        dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
                        If dt.Rows.Count = 1 Then 'Attrib Val found
                            'Get the AttributeTypeID. The method EditCustomerAttribute is not handling updating attributes using the
                            'external type and value. Hence the SQL selects above.
                            AttributeValueID = MyCommon.NZ(dt.Rows(0).Item("AttributeValueID"), 0)
                            AttributeValueDescription = MyCommon.NZ(dt.Rows(0).Item("Description"), "")
                            MyCommon.QueryStr = "select top 1 AttributeValueID from CustomerAttributes with (NoLock) where CustomerPK = @CustomerPK and " & _
                                                " AttributeTypeID = @AttributeTypeID and Deleted=0 "
                            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                            MyCommon.DBParameters.Add("@AttributeTypeID", SqlDbType.Int).Value = AttributeTypeID
                            dt_oldValue = MyCommon.ExecuteQuery(DataBases.LogixXS)
                            If dt_oldValue.Rows.Count = 1 Then
                                MyCommon.QueryStr = "select description from AttributeValues where AttributeValueID = @dt_oldValue_Rows"
                                MyCommon.DBParameters.Add("@dt_oldValue_Rows", SqlDbType.Int).Value = dt_oldValue.Rows(0)(0)
                                dt_oldDesc = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                If dt_oldDesc.Rows.Count > 0 Then
                                    Prev_AttributeValueDescription = dt_oldDesc.Rows(0)(0)
                                End If
                            End If
                            MyMassUpdate = New Copient.MassUpdate(MyCommon, Copient.MassUpdate.CALLER.CUSTOMER_INQUIRY_WS)
                            MyCommon.Write_Log("AAA.txt", "Setting attributes ... CustomerPK=" & CustomerPK & " ExtCardID='" & MaskHelper.MaskCard(ExtCardID, CardTypeID) & "'  CardTypeID='" & CardTypeID & "'", True)
                            MyCustAttrib.SetExtCardID(ExtCardID)
                            MyCustAttrib.SetCardTypeID(CardTypeID)
                            MyCustAttrib.SetAttributeTypeID(AttributeTypeID)
                            MyCustAttrib.SetAttributeValueID(AttributeValueID)
                            bUpdated = MyMassUpdate.EditCustomerAtribute(MyCustAttrib)
                            If bUpdated Then
                                ' Customer attribute value activity log
                                Dim ActivityText As String = ""
                                ActivityText = "Edited customer attribute values," & AttributeTypeDesc.Trim & " from " & Chr(34) & _
                                                    "" & Prev_AttributeValueDescription & Chr(34) & " to " & Chr(34) & AttributeValueDescription & Chr(34)
                                If ActivityText <> "" Then MyCommon.Activity_Log2(25, 11, CustomerPK, 1, ActivityText)
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.SUCCESS
                                row.Item("Description") = "Success: Updated customer attributes"
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                            Else
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
                                row.Item("Description") = "Failed to Update customer attributes"
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                            End If
                        Else ' Invalid attribute value
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.INVALID_ATTRIBUTEVALUE
                            row.Item("Description") = "Failure: Attribute Value not found"
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If
                    Else 'Invalid attribute type
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_ATTRIBUTETYPE
                        row.Item("Description") = "Failure: Attribute Type not found"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else 'Customer pk not found
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                    row.Item("Description") = "Failure: Customer Not Found"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                End If
            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If bOpenedRTConnection And (MyCommon.LRTadoConn.State <> ConnectionState.Closed) Then MyCommon.Close_LogixRT()
            If bOpenedXSConnection And (MyCommon.LXSadoConn.State <> ConnectionState.Closed) Then MyCommon.Close_LogixXS()
        End Try

        Return ResultSet
    End Function

    <WebMethod()> _
    Public Function UpdateCustomerPassword(ByVal GUID As String, ByVal CardTypeID As String, ByVal ExtCardID As String, ByVal NewPassword As String) As String
        'Dim strCardTypeID as string = CardTypeID
        If IsNumeric(CardTypeID) And (CardTypeID <> "") Then
            CardTypeID = Integer.Parse(CardTypeID)
        Else
            CardTypeID = -1
        End If
        Return UpdateMemberPassword(GUID, CardTypeID, ExtCardID, NewPassword)
    End Function

    'updated the method to add 14digit validation for AltIDs - AMS - Ahold enhancement
    Private Function UpdateMemberPassword(ByVal GUID As String, ByVal CardTypeID As Integer, ByVal ExtCardID As String, ByVal NewPassword As String) As String

        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = "Success"
        Dim MethodName As String
        Dim dt As DataTable
        Dim CustomerPK As Long = 0
        Dim MyCryptLib As New Copient.CryptLib
        Dim EncryptedPassword As String
        Try
            'Establish a connection to the LogixXS and LogixRT databases
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MethodName = "UpdateMemberPassword"
            EncryptedPassword = MyCryptLib.SQL_StringEncrypt(NewPassword)
            'Check if the GUID is valid for Customer Inquiry
            If IsValidGUID(GUID, MethodName) And (CardTypeID <> -1) Then

                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)

                'Check if Card exists and get the CustomerPK, which will be used in subsequent queries
                MyCommon.QueryStr = "select top 1 CustomerPK from CardIDs with (NoLock) where ExtCardID= @ExtCardID AND CardTypeID = @CardTypeID "
                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID, True)
                MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                If dt.Rows.Count = 1 Then 'CustomerPK exists
                    'Get the CustomerPK
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    MyCommon.QueryStr = "Update Customers with (RowLock) SET Password= @EncryptedPassword WHERE CustomerPK= @CustomerPK "
                    MyCommon.DBParameters.Add("@EncryptedPassword", SqlDbType.NVarChar).Value = EncryptedPassword.ConvertBlankIfNothing()
                    MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                    MyCommon.ExecuteNonQuery(DataBases.LogixXS)
                    RetMsg = "Updated customer Password"

                Else 'If row count is not 1, CustomerPK not found
                    RetMsg = "CustomerPK not found"
                End If
            Else
                If (CardTypeID = -1) Then
                    RetMsg = "Invalid CardTypeID"
                Else
                    'Wrong GUID
                    RetMsg = "Invalid GUID"
                End If
            End If

        Catch ex As Exception
            RetCode = StatusCodes.CUST_INVALID_PASSWORD
            RetMsg = "Error encountered: " & ex.ToString
            MyCommon.Write_Log(LogFile, RetMsg, True)
        Finally
            'Close the connection to the database
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Return RetMsg
    End Function

    <WebMethod()> _
    Public Function AuthenticateCardByEmailAndPassword(ByVal GUID As String, ByVal Email As String, ByVal Password As String) As DataSet
        Return AuthenticateMember(GUID, Email, Password)
    End Function

    Private Function AuthenticateMember(ByVal GUID As String, ByVal Email As String, ByVal Password As String) As DataSet

        Dim RetMsg As String = "Success"
        Dim MethodName As String
        Dim dt As DataTable
        Dim CustomerPK As Long = 0
        Dim MyCryptLib As New Copient.CryptLib
        Dim DecryptedPassword As String
        Dim DBPassword As String
        Dim ExtCardID As String = ""
        Dim CardDataSet As New System.Data.DataSet("Cards")
        Dim dtCards As DataTable
        Dim dtPass As DataTable
        Dim dtFinal As DataTable
        Dim dtStatus As DataTable
        Dim rowS As DataRow
        Dim CardCount As Integer
        Dim sBaseProgramID As String = ""
        Dim dtCustDetails As DataTable
        Dim dtPointBalances As DataTable
        Dim sFirstName As String = ""
        Dim sLastName As String = ""
        Dim lPoints As Long = 0
        Dim row1 As DataRow
        Dim CustCount As Integer = 0
        Dim sCustPKList As String = ""
        Dim dt1 As DataTable
        Dim rowCust As DataRow
        Dim IsAuthenticated As Boolean = False
        Dim IsValidEmail As Boolean = True

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
        dt1 = New DataTable
        dt1.TableName = "CustomerPKTable"
        dt1.Columns.Add("CustomerPK", System.Type.GetType("System.Int32"))

        Try
            'Establish a connection to the LogixXS and LogixRT databases
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MethodName = "AuthenticateMember"
            dtFinal = New DataTable
            dtFinal.TableName = "CardTable"
            dtFinal.Columns.Add("ExtCardID", System.Type.GetType("System.String"))
            dtFinal.Columns.Add("FirstName", System.Type.GetType("System.String"))
            dtFinal.Columns.Add("LastName", System.Type.GetType("System.String"))
            dtFinal.Columns.Add("BasePoints", System.Type.GetType("System.Int32"))

            'Check if the GUID is valid for Customer Inquiry
            If IsValidGUID(GUID, MethodName) Then 'Match GUID
                If Email Is Nothing OrElse Email.Trim = "" Then
                    RetMsg = "Invalid Email ID."
                    rowS = dtStatus.NewRow()
                    rowS.Item("StatusCode") = StatusCodes.INVALID_EMAILID
                    rowS.Item("Description") = "Failure: Invalid Email ID."
                    dtStatus.Rows.Add(rowS)
                    IsValidEmail = False
                ElseIf Email.Contains("'") = True Or Email.Contains(Chr(34)) = True Then
                    RetMsg = "Invalid Email ID"
                    rowS = dtStatus.NewRow()
                    rowS.Item("StatusCode") = StatusCodes.INVALID_EMAILID
                    rowS.Item("Description") = "Failure: Invalid Email ID."
                    dtStatus.Rows.Add(rowS)
                    IsValidEmail = False
                End If
                'Check if Card exists and get the CustomerPK, which will be used in subsequent queries
                If IsValidEmail Then
                    MyCommon.QueryStr = "select CustomerPK from CustomerExt with (NoLock) where Email= @Email"
                    MyCommon.DBParameters.Add("@Email", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringEncrypt(Email)
                    dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                    'More than one record found with matching email
                    If (dt.Rows.Count > 1) Then
                        For CustCount = 0 To (dt.Rows.Count) - 1
                            sCustPKList = sCustPKList & MyCommon.NZ(dt.Rows(CustCount).Item("CustomerPK"), 0) & ";"
                            rowCust = dt1.NewRow()
                            rowCust.Item("CustomerPK") = MyCommon.NZ(dt.Rows(CustCount).Item("CustomerPK"), 0)
                            dt1.Rows.Add(rowCust)
                        Next
                        CardDataSet.Tables.Add(dt1.Copy)
                        RetMsg = " Multiple customers found with matching email"
                        rowS = dtStatus.NewRow()
                        rowS.Item("StatusCode") = StatusCodes.CUST_MULTIPLE_EMAIL
                        rowS.Item("Description") = "Failure: Multiple customer matches found with specified email " & sCustPKList
                        dtStatus.Rows.Add(rowS)

                        'One record found with matching email
                    ElseIf (dt.Rows.Count = 1) Then 'CustomerPK exists with email
                        'Get the CustomerPK
                        CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), "")
                        MyCommon.QueryStr = "select Password  from Customers with (NoLock) WHERE CustomerPK= @CustomerPK"
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        dtPass = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dtPass.Rows.Count = 1 Then 'Entry exists in Customers table for the CustomerPK so password can be matched
                            'Found customer Password
                            DBPassword = MyCommon.NZ(dtPass.Rows(0).Item("Password"), "")

                            If Not DBPassword = "" Then
                                DecryptedPassword = MyCryptLib.SQL_StringDecrypt(DBPassword)
                            Else
                                'Blank DBPassword
                                DecryptedPassword = ""
                                RetMsg = "Decrypted DB Password" & DecryptedPassword
                            End If

                            If IsNumeric(DecryptedPassword) And IsNumeric(Password) Then
                                If Integer.Parse(DecryptedPassword) = Integer.Parse(Password) Then
                                    IsAuthenticated = True
                                End If
                            Else
                                If DecryptedPassword = Password Then
                                    IsAuthenticated = True
                                End If
                            End If

                            If IsAuthenticated Then
                                'Password match occurred
                                'Get the base points program
                                sBaseProgramID = MyCommon.Fetch_CPE_SystemOption(104).Trim
                                If (sBaseProgramID.IndexOf(";") > -1) Then
                                    sBaseProgramID = sBaseProgramID.Substring(sBaseProgramID.IndexOf(";") + 1)
                                Else
                                    sBaseProgramID = 0
                                End If
                                MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from Cardids with (NoLock) where CardTypeID=0 AND CardStatusID IN (1,5,6) AND CustomerPK= @CustomerPK "
                                MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                                dtCards = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                If (dtCards.Rows.Count >= 1) Then
                                    For CardCount = 0 To (dtCards.Rows.Count) - 1
                                        ExtCardID = ExtCardID & MyCommon.NZ(MyCryptLib.SQL_StringDecrypt(dtCards.Rows(CardCount).Item("ExtCardID").ToString()), 0) & ";"
                                    Next
                                    MyCommon.QueryStr = "select FirstName, LastName from Customers with (NoLock) WHERE CustomerPK = @CustomerPK "
                                    MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                                    dtCustDetails = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                    If (dtCustDetails.Rows.Count = 1) Then
                                        sFirstName = MyCommon.NZ(dtCustDetails.Rows(0).Item("FirstName"), "")
                                        sLastName = MyCommon.NZ(dtCustDetails.Rows(0).Item("LastName"), "")
                                    End If
                                    MyCommon.QueryStr = "select AdjAmount from PointsHistory with (NoLock) WHERE ProgramID =  @sBaseProgramID   AND CustomerPK=  @CustomerPK"
                                    MyCommon.DBParameters.Add("@sBaseProgramID", SqlDbType.Int).Value = sBaseProgramID
                                    MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.Int).Value = CustomerPK
                                    dtPointBalances = MyCommon.ExecuteQuery(DataBases.LogixXS)
                                    If (dtPointBalances.Rows.Count = 1) Then
                                        lPoints = MyCommon.NZ(dtPointBalances.Rows(0).Item("AdjAmount"), 0)
                                    End If
                                    row1 = dtFinal.NewRow()
                                    row1.Item("ExtCardID") = ExtCardID
                                    row1.Item("FirstName") = sFirstName
                                    row1.Item("LastName") = sLastName
                                    row1.Item("BasePoints") = lPoints
                                    dtFinal.Rows.Add(row1)
                                    rowS = dtStatus.NewRow()
                                    rowS.Item("StatusCode") = StatusCodes.SUCCESS
                                    rowS.Item("Description") = "Success"
                                    dtStatus.Rows.Add(rowS)
                                Else
                                    RetMsg = "Cannot find card record for customer password match"
                                    rowS = dtStatus.NewRow()
                                    rowS.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                                    rowS.Item("Description") = "Failure: Card not found for password match "
                                    dtStatus.Rows.Add(rowS)
                                End If

                            Else
                                RetMsg = "Passwords don't match"
                                rowS = dtStatus.NewRow()
                                rowS.Item("StatusCode") = StatusCodes.CUST_INVALID_PASSWORD
                                rowS.Item("Description") = "Failure: Password does not match value in database for customer " & CustomerPK
                                dtStatus.Rows.Add(rowS)
                            End If
                            'No records found with matching email
                        Else 'CustomerPK not found in customers table for password match
                            RetMsg = "Cannot find customer for password match"
                            rowS = dtStatus.NewRow()
                            rowS.Item("StatusCode") = StatusCodes.CUST_NOMATCH_PASSWORD
                            rowS.Item("Description") = "Failure: Customer not found for password match "
                            dtStatus.Rows.Add(rowS)
                        End If

                    Else 'If row count is not 1, CustomerPK not found in CustomerExt table with email
                        RetMsg = "CustomerPK not found for email"
                        rowS = dtStatus.NewRow()
                        rowS.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                        rowS.Item("Description") = "Failure: Customer Not found with specified email "
                        dtStatus.Rows.Add(rowS)
                    End If
                End If

                'Add data table containing customer data to the data set
                If dtFinal.Rows.Count >= 1 Then
                    dtFinal.AcceptChanges()
                    CardDataSet.Tables.Add(dtFinal.Copy)
                End If

                'Add data table containing status data to the data set
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    CardDataSet.Tables.Add(dtStatus.Copy)
                End If

            Else 'Match GUID
                'Wrong GUID
                RetMsg = "Invalid GUID"
                'Add data table containing status data to the data set
                rowS = dtStatus.NewRow()
                rowS.Item("StatusCode") = StatusCodes.INVALID_GUID
                rowS.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(rowS)
                dtStatus.AcceptChanges()
                CardDataSet.Tables.Add(dtStatus.Copy)
            End If

        Catch ex As Exception
            'Add data table containing status data to the data set
            rowS = dtStatus.NewRow()
            rowS.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            rowS.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(rowS)
            dtStatus.AcceptChanges()
            CardDataSet.Tables.Add(dtStatus.Copy())
        Finally
            'Close the connection to the database
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Return CardDataSet
    End Function

    Private Function IsValidAdminUserID(ByVal AdminUser As String, ByRef ErrorMessage As String) As Boolean
        Dim AdminUserID As Integer = 0
        Dim dt As DataTable
        Dim Valid As Boolean = False

        If String.IsNullOrEmpty(AdminUser) OrElse AdminUser.Trim() = "" Then
            ErrorMessage = "Admin User ID is not provided."
        ElseIf Not Integer.TryParse(AdminUser, AdminUserID) Then
            ErrorMessage = "Admin User ID is not a number."
        ElseIf AdminUserID <= 0 Then
            ErrorMessage = "Admin User ID is less than or equal to 0."
        Else
            MyCommon.QueryStr = "select AdminUserID from AdminUsers with (NoLock) where AdminUserID= @AdminUserID"
            MyCommon.DBParameters.Add("@AdminUserID", SqlDbType.Int).Value = AdminUserID
            dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
            Valid = (dt.Rows.Count > 0)
        End If

        Return Valid
    End Function

    'Check for valid Customer Note Type ID
    Private Function IsValidCustNoteTypeID(ByVal CustNoteType As Integer) As Boolean
        Dim Valid As Boolean = False
        MyCommon.QueryStr = "dbo.pt_CheckValidCustNoteTypeId"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustNoteTypeId", SqlDbType.Int).Value = CustNoteType
        MyCommon.LXSsp.Parameters.Add("@Result", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        If (MyCommon.LXSsp.Parameters("@Result").Value > 0) Then
            Valid = True
        End If
        MyCommon.Close_LXSsp()

        Return Valid
    End Function
<WebMethod()> _
    Public Function GetStoredValueHistory(ByVal GUID As String, ByVal ProgramID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, _
                                        ByVal StartDate As String, ByVal EndDate As String) As DataSet
        Dim iCardTypeID As Integer = -1
        Dim sStartDate As Date = "01-01-1900"
        Dim sEndDate As Date = "01-01-1900"
        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        If IsValidDate(StartDate) Then sStartDate = Convert.ToDateTime(StartDate)
        If IsValidDate(EndDate) Then sEndDate = Convert.ToDateTime(EndDate)
        Return _GetCustomerStoredValueHistory(GUID, ProgramID, ExtCardID, iCardTypeID, sStartDate, sEndDate)
    End Function
    

    Private Function _GetCustomerStoredValueHistory(ByVal GUID As String, ByVal ProgramID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, _
                                                ByVal StartDate As Date, ByVal EndDate As Date) As DataSet
        Dim ResultDataSet As New System.Data.DataSet("SVHistory")
        Dim dtStatus As DataTable
        Dim dtSVHistory As DataTable = Nothing
        Dim row, dr As DataRow
        Dim dt As DataTable
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim CustomerPK As Long = 0
       
        Dim Program_Id As Integer = Convert.ToInt32(ProgramID)
                      
            
        Dim LocId As Long
        Dim ProgramName As String = "null"
         
        Dim AdjAmount As Integer = 0
		Dim ActionCode As Integer=-1
        Dim AdjAction As String = String.Empty
        
        Dim extlocationcode As String = String.Empty
        Dim extlocationname As String = String.Empty
       
   
        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
                             

            If IsValidGUID(GUID, "GetStoredValueHistory") Then
                ' Lookup the customer
                If IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then
                    'Ignore
                End If
                RetCode = RetCode
                RetMsg = RetMsg

                If (StartDate = "01-01-1900") Then
                    'Bad Start Date
                    RetCode = StatusCodes.INVALID_STARTDATE
                    RetMsg = "Failure: Invalid StartDate"
                ElseIf (EndDate = "01-01-1900") Then
                    'Bad End Date
                    RetCode = StatusCodes.INVALID_ENDDATE
                    RetMsg = "Failure: Invalid EndDate"
                End If
                If RetCode = StatusCodes.SUCCESS Then

                    ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
                    CustomerPK = LookupCustomerPK(ExtCardID, CardTypeID, RetCode, RetMsg)
                    If CustomerPK > 0 AndAlso RetCode = StatusCodes.SUCCESS Then
              
                        'Create a new datatable to hold the results we'll be assembling
                        dtSVHistory = New DataTable("StoredValueHistory")
                        dtSVHistory.Columns.Add("ProgramID", System.Type.GetType("System.Int32"))
                        dtSVHistory.Columns.Add("Amount", System.Type.GetType("System.Int64"))
						dtSVHistory.Columns.Add("Action", System.Type.GetType("System.String"))
                        dtSVHistory.Columns.Add("EarnRedeemDate", System.Type.GetType("System.String"))
                        dtSVHistory.Columns.Add("ExpirationDate", System.Type.GetType("System.String"))
                        dtSVHistory.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
                        dtSVHistory.Columns.Add("LocationName", System.Type.GetType("System.String"))
                        

                        MyCommon.QueryStr = "Select DISTINCT SVProgramID, QtyEarned, QtyUsed, EarnedLocationID, EarnedDate, " & _
                        "ExpireDate, StatusFlag from SVHistory WITH (NOLOCK) where customerpk = @CustomerPK and SVProgramID=@Program_Id" & _
                        " and (LastUpdate >= @StartDate) AND (LastUpdate <= @EndDate)"
                        
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.Int).Value = CustomerPK
                        MyCommon.DBParameters.Add("@Program_Id", SqlDbType.Int).Value = Program_Id
                        MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
                        MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
                        
                        dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                        If dt.Rows.Count > 0 Then
                            For Each dr In dt.Rows
                                Programid = MyCommon.NZ(dr.Item("SVProgramID"),0)
                                LocId = MyCommon.NZ(dr.Item("EarnedLocationID"),0)
								ActionCode=MyCommon.NZ(dr.Item("StatusFlag"),0)
                                '  If GetExtLocationCode(MyCommon, LocId).Rows.Count > 0 Then
                                
                                
                                MyCommon.QueryStr = "select Name from StoredValuePrograms WITH (NOLOCK) where SVProgramID = @Programid"
                                MyCommon.DBParameters.Add("@Programid", SqlDbType.Int).Value = Program_Id
                                Dim dtProgramName As New DataTable
                                dtProgramName = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                If dtProgramName.Rows.Count > 0 Then
                                    ProgramName = MyCommon.NZ(dtProgramName.Rows(0).Item("Name"), "")
                                End If

                                Dim dtExtLocation As New DataTable
                                MyCommon.QueryStr = "SELECT ExtLocationCode, LocationName FROM Locations WHERE LocationID= @LocId "
                                MyCommon.DBParameters.Add("@LocId", SqlDbType.BigInt).Value = LocId
                                dtExtLocation = MyCommon.ExecuteQuery(DataBases.LogixRT)
                                   
                                If (dtExtLocation.Rows.Count = 1) Then
                                    extlocationcode = MyCommon.NZ(dtExtLocation.Rows(0).Item("ExtLocationCode"), "")
                                    extlocationname = MyCommon.NZ(dtExtLocation.Rows(0).Item("LocationName"), "")
                                End If
                                                              
                                'AdjAmount
                                
                                Dim QtyEarned as Integer =dr.Item("QtyEarned")
                                Dim QtyUsed as Integer = dr.Item("QtyUsed")
                                Dim dtExpire As Date = MyCommon.NZ(dr.Item("ExpireDate"),"01-01-1900" )
                                                   
								'If value available in SVHistory.QtyEarned column, then it will be treated as Earn
								If (QtyEarned <> 0) Then
									AdjAmount = QtyEarned
									
								End If
												   
								' If value available in SVHistory.QtyUsed column, then it will be treated as Redeem
								If (QtyUsed <> 0) Then
									AdjAmount = QtyUsed
									
								End If
									
                                ' If SV program is already expired (i.e. it is within the input StartDate and EndDate range but expired according to the SV program creation details), then it will be treated as Expired. However value of QtyEarned/QtyUsed will be sent
                                
								  Select Case ActionCode
                                    Case 1
                                        AdjAction = "Earn"
                                    Case 2
                                        AdjAction = "Revoke"
                                    Case 3
                                        AdjAction = "Expire"
                                    Case 4
                                        AdjAction = "Redeem"
                                  End Select
                                    
                                row = dtSVHistory.NewRow()
                                row.Item("ProgramID") = Program_Id
                                row.Item("Amount") = AdjAmount
                                row.Item("Action") = AdjAction
                                row.Item("EarnRedeemDate") = MyCommon.NZ(dr.Item("EarnedDate"), "")
                                row.Item("ExpirationDate") = MyCommon.NZ(dr.Item("ExpireDate"), "")
                                row.Item("ExtLocationCode") = extlocationcode
                                row.Item("LocationName") = extlocationname
                                                              
                                dtSVHistory.Rows.Add(row)
                                
                                
                                ' End If
                            Next
                        End If

                        If dtSVHistory.Rows.Count > 0 Then
                            dtSVHistory.AcceptChanges()
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success."
                            dtStatus.Rows.Add(row)
                        Else
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Records Not Found"
                            dtStatus.Rows.Add(row)
                        End If
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = RetCode
                        row.Item("Description") = "CardID: " & ExtCardID & " with CardTypeID: " & CardTypeID & " not found."
                        dtStatus.Rows.Add(row)
                    End If
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If
                If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                    dtStatus.AcceptChanges()
                    ResultDataSet.Tables.Add(dtStatus.Copy())
                End If
                If dtSVHistory IsNot Nothing Then ResultDataSet.Tables.Add(dtSVHistory)

            Else
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultDataSet.Tables.Add(dtStatus.Copy())
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        
        End Try

        Return ResultDataSet
    End Function
End Class