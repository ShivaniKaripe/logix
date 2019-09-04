<%@ WebService Language="VB" Class="Service" %>
' version:7.3.1.138972.Official Build (SUSDAY10202) 

Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Serialization
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.IO
Imports Copient.CommonInc
Imports Copient.IdNotFoundException


<WebService(Namespace:="http://www.copienttech.com/FuelPartnerInterface/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService

    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib
  
  Public Enum StatusCodes As Integer
    SUCCESS = 0
    INVALID_GUID = 1
    INVALID_CUSTOMERID = 2
    INVALID_CUSTOMERTYPEID = 3
    INVALID_CUSTOMERGROUPID = 4
    INVALID_MODE = 5
    INVALID_STORED_VALUE_PROGRAM = 6
    INVALID_STARTDATE = 7
    INVALID_ENDDATE = 8
    INVALID_DATETIME = 9
    INVALID_EXTLOCATIONID = 10
    NOTFOUND_CUSTOMER = 11
    NOTFOUND_HOUSEHOLD = 12
    NON_POSITIVE_SV_ADJUST_AMOUNT = 13
    ADJUSTMENT_FAILED = 14
    INVALID_VALUE = 15
    INVALID_SVPROGRAMID = 16
    NOTFOUND_CONVERSIONDATA = 17
    INVALID_PARTNERID = 18
    CUSTOMER_NOLOCK = 19
    CUSTOMER_ALREADYLOCKED = 20
    APPLICATION_EXCEPTION = 9999
  End Enum
  
  <WebMethod()> _
  Public Function GetRewardBalances(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String, _
                                    SVProgramID As Long, LockCustomerID As Boolean, ExtLocationCode As String, _
                                    PartnerID As Long) As DataSet
    Dim iCardTypeID As Integer = -1
    If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
    Return _GetRewardBalances(GUID, CardID, iCardTypeID, SVProgramID, LockCustomerID, _
                              ExtLocationCode, PartnerID)
  End Function

  Private Function _GetRewardBalances(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, _
                                    SVProgramID As Long, LockCustomerID As Boolean, ExtLocationCode As String, _
                                    PartnerID As Long) As DataSet
    Dim ResultDataSet As New System.Data.DataSet("TransactionTotal")
    Dim CustomerPK As Long = 0
    Dim dtStatus As DataTable
    Dim dtStoredValueRewardBalances As DataTable = Nothing
    Dim row As DataRow
    Dim dt As DataTable
    Dim RetCode As StatusCodes = StatusCodes.SUCCESS
    Dim RetMsg As String = ""
    Dim SVProgramName As String = ""
    Dim SVMonetaryValue As Decimal = 0
    Dim SVPointsValue As Integer = 0
    Dim StoredValue As New Copient.StoredValue
    Dim PointsBalance As Integer = 0
    Dim SVRewardBalance As Decimal = 0
    Dim LocationPK As Long = -9
    
    'Fill PartnerID if not provided
    MyCommon.NZ(PartnerID, -1)

    'Initialize the status table, which will report the success or failure of the operation
    dtStatus = New DataTable
    dtStatus.TableName = "Status"
    dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
    dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
      If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()

      'Perform initial validation
      If IsValidGUID(GUID, MyCommon) Then

        ' Lookup the given customer card and CardType
        MyCommon.QueryStr = "Select CustomerPK from cardids where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & _
                                   "' and CardTypeID=" & CardTypeID
        dt = MyCommon.LXS_Select
         If dt.Rows.Count <= 0 OrElse  ExtCardID Is Nothing OrElse ExtCardID.Trim = "" Then
          'Bad customer ID
          RetCode = StatusCodes.INVALID_CUSTOMERID
          RetMsg = "Failure: Entered ExtCardID and CardTypeID was not found or is invalid"
        Else
          CustomerPK = dt.Rows(0).Item("CustomerPK")
        End If

        'Make sure that conversion data exists for this SVProgram
        MyCommon.QueryStr = "Select SVProgramID, SVMonetaryValue, SVPointsValue  from StoredValuePointsConversion where SVProgramID=" & SVProgramID
        dt = MyCommon.LXS_Select
        If  dt.Rows.Count <= 0 Then
          'Conversion information not found
          RetCode = StatusCodes.NOTFOUND_CONVERSIONDATA
          RetMsg = "Failure: The entered Stored Value Program is not fuel partner enabled"
        Else
          SVMonetaryValue = dt.Rows(0).Item("SVMonetaryValue")
          SVPointsValue = dt.Rows(0).Item("SVPointsValue")
        End If

        ' Lookup the given ExtLocationID
        MyCommon.QueryStr = "Select LocationID,LocationName from Locations where ExtLocationCode='" & ExtLocationCode & "';"
        dt = MyCommon.LRT_Select
        'If dt.Rows.Count <= 0 OrElse  ExtLocationCode Is Nothing OrElse ExtLocationCode.Trim = "" Then
        '  'Provided location not found
        '  RetCode = StatusCodes.INVALID_EXTLOCATIONID
        '  RetMsg = "Failure: Entered ExtLocationID was not found or is invalid"
        'Else
        'End If
        If dt.Rows.Count > 0 Then
          LocationPK = MyCommon.NZ(dt.Rows(0).Item("LocationID"),0)
        End If

        ' Lookup provided Stored Value Program
        MyCommon.QueryStr = "Select SVProgramID, SVTypeID, Name from StoredValuePrograms" & " Where SVProgramID = " & SVProgramID & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count <= 0 OrElse dt.Rows(0).Item("SVTypeID") <> 1 Then
          'Provided fuel partner id not found
          RetCode = StatusCodes.INVALID_SVPROGRAMID
          RetMsg = "Failure: Entered Stored Value Program was not found or is not a Stored Value Points Program"
        Else
          SVProgramName = dt.Rows(0).Item("Name")
        End If

        ' Check if the customer is already locked
        If LockCustomerID Then
          MyCommon.QueryStr = "Select * from CustomerLock where CustomerPK=" & CustomerPK & ";"
          dt = MyCommon.LXS_Select
          If dt.Rows.Count > 0 Then
            Dim sLockedDate As String = ""
            Dim dtLocked As Date
            Dim lMinutes As Long
            Dim elapsed_time As TimeSpan
            Dim CustLockDelMin As Integer
            sLockedDate = dt.Rows(0).Item("LockedDate")
            dtLocked = Date.Parse(sLockedDate)
            elapsed_time = DateTime.Now.Subtract(dtLocked)
            lMinutes = elapsed_time.TotalMinutes
            CustLockDelMin = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(68))
            If (CustLockDelMin = 0) OrElse (lMinutes < CustLockDelMin) Then
              'Customer is already locked
              RetCode = StatusCodes.CUSTOMER_ALREADYLOCKED
              RetMsg = "Failure: This customer is already locked"
            Else
              ' lock has expired
            End If
          End If
        End If
        
        'Validation passed, process request
        If RetCode = StatusCodes.SUCCESS Then
          'Return the total available points for this customer for this stored value program
          'If LockCustomerID = true, Lock this customer
            'If this customer is already locked, return immediately with an error
          'find the current total possible value of discount for this customer
            'look up the conversion information for this stored value program
            'total available value found by dividing total points by svpointsvalue for this sv program id, 
              'rounding this answer down to the nearest integer and multplying the resulting value by the monetaryvalue field for this program id

          'Lock the customer record if necessary
          If LockCustomerID Then
            CustomerLockSet(CustomerPk, ExtLocationCode, "0", "0", 0)
          End If

          'Create a new datatable to hold the results we'll be assembling
          dtStoredValueRewardBalances = New DataTable("TotalRewardAvailable")
          dtStoredValueRewardBalances.Columns.Add("SVProgramID", System.Type.GetType("System.Int64"))
          dtStoredValueRewardBalances.Columns.Add("SVProgramName", System.Type.GetType("System.String"))
          dtStoredValueRewardBalances.Columns.Add("CustomerID", System.Type.GetType("System.String"))
          dtStoredValueRewardBalances.Columns.Add("SVRewardBalance", System.Type.GetType("System.Decimal"))

          'Find the total points available to this customer
          PointsBalance  = StoredValue.GetQuantityBalance(CustomerPK, SVProgramID)
          
          'Use the conversion data and points data found to compute the possible reward for this customer
          SVRewardBalance = SVMonetaryValue * (PointsBalance\SVPointsValue)

          row = dtStoredValueRewardBalances.NewRow()
          row.Item("SVProgramID") = SVProgramID
          row.Item("CustomerID") = ExtCardID
          row.Item("SVProgramName") = SVProgramName
          row.Item("SVRewardBalance") = SVRewardBalance
          dtStoredValueRewardBalances.Rows.Add(row)

          row = dtStatus.NewRow()
          row.Item("StatusCode") = StatusCodes.SUCCESS
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
          End If
          If dtStoredValueRewardBalances IsNot Nothing Then ResultDataSet.Tables.Add(dtStoredValueRewardBalances)
        Else
          'Validation did not pass
          row = dtStatus.NewRow()
          row.Item("StatusCode") = RetCode
          row.Item("Description") = RetMsg
          dtStatus.Rows.Add(row)
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
      If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
    End Try

    Return ResultDataSet
  End Function

  <WebMethod()> _
  Public Function UpdateRewardBalances(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String, _
                                    SVProgramID As Long, ExtLocationCode As String, _
                                    InquiryDate As String, TransactionID As String, _
                                    RewardAmountRedeemed As Decimal, UnitsPumped As Decimal, PartnerID As Integer) As DataSet
    Dim iCardTypeID As Integer = -1
    If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
    Return _UpdateRewardBalances(GUID, CardID, CardTypeID, _
                                 SVProgramID, ExtLocationCode, _
                                 InquiryDate, TransactionID, _
                                 RewardAmountRedeemed, UnitsPumped, PartnerID)
  End Function

  Private Function _UpdateRewardBalances(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, _
                                    SVProgramID As Long, ExtLocationCode As String, _
                                    InquiryDate As String, TransactionID As String, _
                                    RewardAmountRedeemed As Decimal, UnitsPumped As Decimal, PartnerID As Integer) As DataSet
    Dim ResultDataSet As New System.Data.DataSet("TransactionTotal")
    Dim CustomerPK As Long = 0
    Dim dtStatus As DataTable
    Dim dtOperationStatus As DataTable = Nothing
    Dim row As DataRow
    Dim dt As DataTable
    Dim RetCode As StatusCodes = StatusCodes.SUCCESS
    Dim RetMsg As String = ""
    Dim SVProgramName As String = ""
    Dim SVMonetaryValue As Decimal = 0
    Dim SVPointsValue As Integer = 0
    Dim StoredValue As New Copient.StoredValue
    Dim PointsBalance As Integer = 0
    Dim SVRewardBalance As Decimal = 0
    Dim ReturnMessage As String = ""
    Dim SVAdjust As String = ""
    Dim PointsUsed As Integer = 0 ' This should be the number of points used in the transaction
    Dim AdjustAmt As String = "" 'This is PointsUsed formatted for our StoredValue Library call, a negative sign needs to be added for the call
    Dim LocationPK As Long = -9

    'Fill PartnerID if not provided
    MyCommon.NZ(PartnerID, -1)

    'Initialize the status table, which will report the success or failure of the operation
    dtStatus = New DataTable
    dtStatus.TableName = "Status"
    dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
    dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
    dtStatus.Columns.Add("SVPointsValue", System.Type.GetType("System.Int32"))
    dtStatus.Columns.Add("SVMonetaryValue", System.Type.GetType("System.Decimal"))
    dtStatus.Columns.Add("RewardAmountRedeemed", System.Type.GetType("System.Decimal"))
    dtStatus.Columns.Add("AdjustAmt", System.Type.GetType("System.String"))
    dtStatus.Columns.Add("SVProgramID", System.Type.GetType("System.Int64"))
    dtStatus.Columns.Add("CustomerPK", System.Type.GetType("System.Int64"))
    dtStatus.Columns.Add("LocationPK", System.Type.GetType("System.Int64"))

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
      If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()

      'Perform initial validation
      If IsValidGUID(GUID, MyCommon) Then
        
        ' Lookup the given customer card and CardType
        MyCommon.QueryStr = "Select CustomerPK from cardids where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & _
                                   "' and CardTypeID=" & CardTypeID
        dt = MyCommon.LXS_Select
         If dt.Rows.Count <= 0 OrElse  ExtCardID Is Nothing OrElse ExtCardID.Trim = "" Then
          'Bad customer ID
          RetCode = StatusCodes.INVALID_CUSTOMERID
          RetMsg = "Failure: Entered ExtCardID and CardTypeID was not found or is invalid"
        Else
          CustomerPK = dt.Rows(0).Item("CustomerPK")
          ' Customer must be locked 
                    MyCommon.QueryStr = "Select * from CustomerLock where CustomerPK=" & CustomerPK & ";"
          dt = MyCommon.LXS_Select
          If dt.Rows.Count <= 0 Then
            'No lock found
            RetCode = StatusCodes.CUSTOMER_NOLOCK
            RetMsg = "Failure: No lock was found for this customer"
          Else
          End If
        End If

        'Make sure that conversion data exists for this SVProgram
        MyCommon.QueryStr = "Select SVProgramID, SVMonetaryValue, SVPointsValue  from StoredValuePointsConversion where SVProgramID=" & SVProgramID
        dt = MyCommon.LXS_Select
        If  dt.Rows.Count <= 0 Then
          'Conversion information not found
          RetCode = StatusCodes.NOTFOUND_CONVERSIONDATA
          RetMsg = "Failure: The entered Stored Value Program is not fuel partner enabled"
        Else
          SVMonetaryValue = dt.Rows(0).Item("SVMonetaryValue")
          SVPointsValue = dt.Rows(0).Item("SVPointsValue")
        End If

        ' Lookup the given ExtLocationID
        MyCommon.QueryStr = "Select LocationID,LocationName from Locations where ExtLocationCode='" & ExtLocationCode & "';"
        dt = MyCommon.LRT_Select
        'If dt.Rows.Count <= 0 OrElse  ExtLocationCode Is Nothing OrElse ExtLocationCode.Trim = "" Then
        '  'Provided location not found
        '  RetCode = StatusCodes.INVALID_EXTLOCATIONID
        '  RetMsg = "Failure: Entered ExtLocationID was not found or is invalid"
        'Else
        'End If
        If dt.Rows.Count > 0 Then
          LocationPK = MyCommon.NZ(dt.Rows(0).Item("LocationID"),0)
        End If

        ' Lookup provided Stored Value Program
        MyCommon.QueryStr = "Select SVProgramID, SVTypeID, Name from StoredValuePrograms" & " Where SVProgramID = " & SVProgramID & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count <= 0 OrElse dt.Rows(0).Item("SVTypeID") <> 1 Then
          'Provided fuel partner id not found
          RetCode = StatusCodes.INVALID_SVPROGRAMID
          RetMsg = "Failure: Entered Stored Value Program was not found or is not a Stored Value Points Program"
        Else
          SVProgramName = dt.Rows(0).Item("Name")
        End If

        'Validation passed, process request
        If RetCode = StatusCodes.SUCCESS Then
          'If this customerID is not locked, immediately return an error
          'Lock this customer
          'find the points value for the reward amount in this transaction
            'look up the conversion information for this stored value program
            'total points expended found by multiplying svpointsvalue by rewardvalue and then dividing this value by the monetary value for this sv program id
              '(reqdpts * rewardval)\monetaryval = expendedpts
          'call SV lib to adjustpts correctly
          'report success/failure

          'Release the customer Record
          CustomerLockRelease(0,ExtCardID, ExtLocationCode, "0")

          'Create a new datatable to hold the results we'll be assembling
          dtOperationStatus = New DataTable("RewardsTransaction")
          dtOperationStatus.Columns.Add("SVReturnMessage", System.Type.GetType("System.String"))

          ' No Rewards Redeemed
          If RewardAmountRedeemed <> 0 Then
            'Find the total points available to this customer
            PointsBalance  = StoredValue.GetQuantityBalance(CustomerPK, SVProgramID)

            'Find the total number of points expended during this transaction.
            PointsUsed=SVPointsValue*(RewardAmountRedeemed/SVMonetaryValue)  'We expect this to be an integer and that the reward amounts are correctly reported
            
            If PointsUsed <= PointsBalance Then
              'Cast for StoredValue Call
              PointsUsed = PointsUsed * -1
              AdjustAmt = System.Convert.ToString(PointsUsed)

              'Adjust this customer's points balance, use default AdminUserID
              ReturnMessage = StoredValue.AdjustStoredValue(1, SVProgramID, CustomerPK, _
                           AdjustAmt,,,,LocationPK)

              If  ReturnMessage <> "No Stored Value records are available to revoke." Then
                'Add a new entry into the StoredValueThirdPartyTransactions Table
                MyCommon.QueryStr = "insert into StoredValueThirdPartyTransactions with (RowLock) (CustomerPK,SVProgramID,RewardAmtRedeemed,UnitsPumped,TransactionID,LastUpdate,PartnerID,ExtLocationCode) " & _
                                                   "values (" & CustomerPK & "," & SVProgramID & "," & RewardAmountRedeemed & "," & UnitsPumped & ",'" & TransactionID & "','" & InquiryDate & "'," & PartnerID & ",'" & ExtLocationCode  & "');"
                MyCommon.LXS_Execute()
              End If

              row = dtOperationStatus.NewRow()
              row.Item("SVReturnMessage") = ReturnMessage
              dtOperationStatus.Rows.Add(row)
            Else
              row = dtOperationStatus.NewRow()
              row.Item("SVReturnMessage") = "Total Points used greater than points available, no action taken"
              dtOperationStatus.Rows.Add(row)
            End If
          Else
            'Add a new entry into the StoredValueThirdPartyTransactions Table
            MyCommon.QueryStr = "insert into StoredValueThirdPartyTransactions with (RowLock) (CustomerPK,SVProgramID,RewardAmtRedeemed,UnitsPumped,TransactionID,LastUpdate,PartnerID,ExtLocationCode) " & _
                                               "values (" & CustomerPK & "," & SVProgramID & "," & RewardAmountRedeemed & "," & UnitsPumped & ",'" & TransactionID & "','" & InquiryDate & "'," & PartnerID & ",'" & ExtLocationCode  & "');"
            MyCommon.LXS_Execute()
            row = dtOperationStatus.NewRow()
            row.Item("SVReturnMessage") = "No Rewards Redeemed"
            dtOperationStatus.Rows.Add(row)
          End If

          row = dtStatus.NewRow()
          row.Item("StatusCode") = StatusCodes.SUCCESS
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
            dtStatus.AcceptChanges()
            ResultDataSet.Tables.Add(dtStatus.Copy())
          End If
          If dtOperationStatus IsNot Nothing Then ResultDataSet.Tables.Add(dtOperationStatus)
        Else
          'Validation did not pass
          row = dtStatus.NewRow()
          row.Item("StatusCode") = RetCode
          row.Item("Description") = RetMsg
          dtStatus.Rows.Add(row)
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
      row.Item("SVPointsValue") = SVPointsValue
      row.Item("RewardAmountRedeemed") = RewardAmountRedeemed
      row.Item("SVMonetaryValue") = SVMonetaryValue
      row.Item("AdjustAmt") = AdjustAmt
      row.Item("SVProgramID") = SVProgramID
      row.Item("CustomerPK") = CustomerPK
      row.Item("LocationPK") = LocationPK
      dtStatus.Rows.Add(row)
      dtStatus.AcceptChanges()
      ResultDataSet.Tables.Add(dtStatus.Copy())
    Finally
      If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
    End Try

    Return ResultDataSet
  End Function

  Private Function IsValidGUID(ByVal GUID As String, ByRef MyCommon As Copient.CommonInc) As Boolean
    Dim IsValid As Boolean = False
    Dim ConnInc As New Copient.ConnectorInc
    
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      ' verfiy using GUIDs defined for the FuelPartnerInterface Connector (54)
      IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 54, GUID)
    Catch ex As Exception
      IsValid = False
    End Try
      
    Return IsValid
  End Function
  
  Private Function CustomerLockSet(ByVal lCustomerPk As Long, ByVal sStoreNum As String, ByVal sTerminal As String, ByVal sTrxNum As String, ByRef lUpdateCount As Long) As Integer
    Dim iLockedStatus As Integer = 0
    Dim sLockedDate As String
    Dim sPrepay As String
    Dim dt As DataTable
    Dim dtLocked As Date
    Dim lMinutes As Long
    Dim lLocationTemp As Long
    Dim sLocationId As String = "0"
    Dim CustLockDelMin As Integer
    Dim elapsed_time As TimeSpan
    Dim bCheckForLock As Boolean = True
    Dim iCheckForLockCount As Int16 = 0
    Dim bUpdateExistingLock As Boolean = False
    Dim sLocationIdLocked As String
    Dim sTerminalNumberLocked As String
    'sCurrentMethod = "CustomerLockSet"

    MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreNum & "';"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      sLocationId = dt.Rows(0).Item(0)
    Else If Integer.TryParse(sStoreNum, lLocationTemp)  Then
        sLocationID = sStoreNum
    Else
      'Throw New ApplicationException("Location with ExtLocationCode='" & sStoreNum & "' does not exist!")
      sLocationId = "-9"
    End If
    Do
      Try
                MyCommon.QueryStr = "select Prepay, LockedDate, LocationId, TerminalNumber from CustomerLock with (NoLock) where CustomerPK=" & lCustomerPk & ";"
        dt = MyCommon.LXS_Select()
        If dt.Rows.Count > 0 Then
          sTerminalNumberLocked = dt.Rows(0).Item("TerminalNumber")
          sLocationIdLocked = dt.Rows(0).Item("LocationId")
          If sTerminalNumberLocked.Trim = sTerminal.Trim And sLocationIdLocked.Trim = sLocationId.Trim Then
            ' same store and terminal so just "get" new lock
            bUpdateExistingLock = True
          Else
            sLockedDate = dt.Rows(0).Item("LockedDate")
            dtLocked = Date.Parse(sLockedDate)
            elapsed_time = DateTime.Now.Subtract(dtLocked)
            lMinutes = elapsed_time.TotalMinutes
            CustLockDelMin = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(68))
            If (CustLockDelMin = 0) OrElse (lMinutes < CustLockDelMin) Then
              ' customer is already locked
              iLockedStatus = 1
              sPrepay = dt.Rows(0).Item("Prepay")
              If sPrepay.ToLower = "true" Then
                ' and was locked at a prepay terminal
                iLockedStatus = 2
              End If
            Else
              ' lock has expired
              bUpdateExistingLock = True
            End If
          End If
          If bUpdateExistingLock Then
                        MyCommon.QueryStr = "update CustomerLock with (RowLock) set Prepay=0,LockedDate=getdate(), LocationId=" & _
                                sLocationId & ",TerminalNumber=" & sTerminal & ",TransactionNumber=" & sTrxNum & _
                                " where CustomerPK=" & lCustomerPk & ";"
            MyCommon.LXS_Execute()
            MyCommon.QueryStr = "update Customers with (RowLock) set UpdateCount=UpdateCount + 1 where CustomerPK=" & lCustomerPk & ";"
            MyCommon.LXS_Execute()
            MyCommon.QueryStr = "select UpdateCount from Customers with (NoLock) where CustomerPK=" & lCustomerPk & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
              lUpdateCount = MyCommon.NZ(dt.Rows(0).Item("UpdateCount"), 0)
            End If
          End If
          bCheckForLock = False
        Else
                    MyCommon.QueryStr = "insert into CustomerLock with (RowLock) (CustomerPK,LocationID,TerminalNumber,TransactionNumber,Prepay,LockedDate)" & _
                    " values (" & lCustomerPk & "," & sLocationId.Trim & "," & sTerminal.Trim & "," & sTrxNum.Trim & ",0,getdate());"
          MyCommon.LXS_Execute()
          bCheckForLock = False
        End If
      Catch exApp As ApplicationException
        Throw
      'Catch exSql As SqlException
      '  If exSql.Number = 2627 Then
      '    ' duplicate - a row was inserted since the above select was attempted, so try again!
      '    bCheckForLock = True
      '    iCheckForLockCount += 1
      '    If iCheckForLockCount > 1 Then
      '      If iCheckForLockCount > 3 Then
      '        Throw New ApplicationException("Exceeded Customer Lock contentention limit (" & iCheckForLockCount & ") (CustomerPK = " & lCustomerPk & ")!")
      '      Else
       '       WriteLog("Customer Lock contentention (" & iCheckForLockCount & ") (CustomerPK = " & lCustomerPk & ")!", MessageType.Warning)
        '    End If
        '  End If
       ' Else
       '   Throw
       ' End If
      Catch ex As Exception
        Throw
      End Try
    Loop While bCheckForLock
    Return iLockedStatus
  End Function

  Private Sub CustomerLockRelease(ByVal sPrepay As String, ByVal sCustomerId As String, ByVal sStoreNum As String, ByVal sTerminal As String)
    Dim dt As DataTable
    Dim lCustomerPk As Long
    Dim lLocationTemp As Long
    Dim lHHPk As Long
    Dim sLocationId As String = "0"
    Try
      MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreNum & "';"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        sLocationId = dt.Rows(0).Item(0)
      Else If Integer.TryParse(sStoreNum, lLocationTemp)  Then
        sLocationID = sStoreNum
      Else
        'Throw New ApplicationException("Location with ExtLocationCode='" & sStoreNum & "' does not exist!")
        sLocationId = "-9"
      End If
            
      sCustomerId = MyCommon.Pad_ExtCardID(sCustomerId,0)

      ' get the CustomerPK from the Card number
      MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustPK"
      MyCommon.Open_LXSsp()
      MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(sCustomerId)
      dt = MyCommon.LXSsp_select()
      MyCommon.Close_LXSsp()
      If dt.Rows.Count > 0 Then
        lCustomerPk = dt.Rows(0).Item("CustomerPK")
        lHHPk = dt.Rows(0).Item("HHPK")
        If lHHPk <> 0 Then
          lCustomerPk = lHHPk
        End If
        If sPrepay.ToLower = "true" Then
                    MyCommon.QueryStr = "update CustomerLock with (RowLock) set Prepay=1 where CustomerPK=" & lCustomerPk & _
                              " and LocationId=" & sLocationId.Trim & " and TerminalNumber=" & sTerminal.Trim & ";"
        Else
                    MyCommon.QueryStr = "delete from CustomerLock with (RowLock) where CustomerPK=" & lCustomerPk & _
                              " and LocationId=" & sLocationId.Trim & " and (Prepay=1 or TerminalNumber=" & sTerminal.Trim & ");"
        End If
        MyCommon.LXS_Execute()
      End If
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try

  End Sub

End Class