<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%
  ' *****************************************************************************
  ' * FILENAME: GetCustomerInfo.aspx
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
    Private MyCryptlib As New Copient.CryptLib
  Public Connector As New Copient.ConnectorInc
  Public MyAltID As New Copient.AlternateID
  Public CAM As New Copient.CAM
  Public TextData As String
  Public IPL As Boolean
  Public LogFile As String
  Dim StartTime As Object
  Dim TotalTime As Object
  Public LSVersionNum As Copient.ConnectorInc.LSVersionRec
  Public ResponseStringHeader As String = "GetCustomerInfo"

  ' -----------------------------------------------------------------------------------------------

  Function Parse_Bit(ByVal BooleanField As Boolean) As String
    If BooleanField Then
      Parse_Bit = "1"
    Else
      Parse_Bit = "0"
    End If
  End Function

  ' -----------------------------------------------------------------------------------------------

  Function Construct_Table(ByVal TableName As String, ByVal Operation As String, ByVal DelimChar As Integer, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal dst As DataTable) As String

    Dim TempResults As String
    Dim NumRecs As Long
    Dim row As DataRow
    Dim SQLCol As DataColumn
    Dim TempOut As String
    Dim Index As Integer
    Dim FieldList As String
    Dim LineOut As String

    TempOut = ""
    TempResults = ""
    NumRecs = 0
    FieldList = ""
    If dst.Rows.Count > 0 Then
      For Each SQLCol In dst.Columns
        If Not (FieldList = "") Then FieldList = FieldList & Chr(DelimChar)
        FieldList = FieldList & SQLCol.ColumnName
      Next
      For Each row In dst.Rows
        Index = 0
        LineOut = ""
        For Each SQLCol In dst.Columns
          If Not (LineOut = "") Then
            LineOut = LineOut & Chr(DelimChar)
          End If
          If SQLCol.DataType.Name = "Boolean" Then 'if it is a binary field
            LineOut = LineOut & Parse_Bit(Common.NZ(row(Index), 0))
          ElseIf SQLCol.DataType.Name = "Int32" Or SQLCol.DataType.Name = "Int64" Then 'if it is an Int or BigInt Field
            LineOut = LineOut & Common.NZ(row(Index), 0)
          ElseIf SQLCol.DataType.Name = "DateTime" Then 'time, so format with zone
            LineOut = LineOut & Format(Common.NZ(row(Index), Now()), "yyyy-MM-dd hh:mm:sszzz")
          Else 'else treat it as a string
                        If UCase(TableName) = "USERS" And UCase(SQLCol.ColumnName) = "ClientUserID1" Then
                            LineOut = LineOut & Common.NZ(MyCryptlib.SQL_StringDecrypt(row(Index).ToString()), 0)
                        
                        ElseIf UCase(TableName) = "CARDIDS" And UCase(SQLCol.ColumnName) = "ExtCardID" Then
                            LineOut = LineOut & Common.NZ(MyCryptlib.SQL_StringDecrypt(row(Index).ToString()), 0)
                            
                        ElseIf UCase(TableName) = "AltIDCacheCardIDs" And UCase(SQLCol.ColumnName) = "ExtCardID" Then
                            LineOut = LineOut & Common.NZ(MyCryptlib.SQL_StringDecrypt(row(Index).ToString()), 0)
                            
                        ElseIf UCase(TableName) = "AltIDCacheUsers" And UCase(SQLCol.ColumnName) = "ClientUserID1" Then
                            LineOut = LineOut & MyCryptlib.SQL_StringDecrypt(row(Index).ToString())
                        Else
                            LineOut = LineOut & Common.NZ(row(Index), "")
                        End If
          End If
          Index = Index + 1
        Next
        TempResults = TempResults & LineOut & vbCrLf
        NumRecs = NumRecs + 1
      Next
      TempOut = TempOut & "1:" & TableName & vbCrLf
      TempOut = TempOut & "2:" & Operation & vbCrLf
      TempOut = TempOut & "3:" & FieldList & vbCrLf
      TempOut = TempOut & "4:" & NumRecs & vbCrLf
      'Common.Write_Log(LogFile, TempOut)
      TempOut = TempOut & TempResults
    End If

    Construct_Table = TempOut

  End Function

  ' -----------------------------------------------------------------------------------------------

  'LockStatus = 0 - Lock granted
  'LockStatus = 1 - Lock already exists for the specified locking group
  'LockStatus = 2 - Customer not found - lock can't be granted
  'LockStatus = 3 - Customer record is not locked

  Function HandleCustomerLocking(ByVal sLockOption As String, ByVal CustomerPK As Long, ByVal LockingGroupId As Long, ByVal LocationID As Long, ByRef LockedCustomerPK As Long, ByRef LockID As Long) As Integer

    Dim dst As DataTable
    Dim iLockStatus As Integer
    Dim lCuPk As Long
    Dim lHhPk As Long = 0
    Dim sHouseholdText As String
    Dim LockLocationID As Long
    Dim ExistingLockExpireDate As DateTime   'When an existing lock (if there is one) expires
    Dim NewLockExpireDate As DateTime        'When a new lock (if we create one) should expire
    Dim LockExpireMinutes As Long
    Dim LocationLockRestart As Integer       'From UE_SystemOption 133: 0=A store can not create a new lock for a terminal group if one already exists  1=A store CAN create a new lock (and delete the current one) for a terminal group if one already exists

    LockID = 0
    Long.TryParse(Common.Fetch_SystemOption(68), LockExpireMinutes)
    If LockExpireMinutes < 0 Then LockExpireMinutes = 0
    NewLockExpireDate = DateAdd(DateInterval.Minute, LockExpireMinutes, DateTime.Now)

    LocationLockRestart = 0
    Integer.TryParse(Common.Fetch_UE_SystemOption(133), LocationLockRestart)

    Common.QueryStr = "select CustomerPK, isnull(HHPK, 0) as HHPK " & _
                      "from Customers with (NoLock) where Customers.CustomerPK=" & CustomerPK & ";"
    dst = Common.LXS_Select
    If dst.Rows.Count > 0 Then
      lHhPk = dst.Rows(0).Item("HHPK")
      If lHhPk = 0 Then
        lCuPk = CustomerPK
        sHouseholdText = ""
      Else
        lCuPk = lHhPk
        sHouseholdText = " (HouseholdPK: " & lHhPk & ")"
      End If
      LockedCustomerPK = lCuPk
      Common.Write_Log(LogFile, "Using CustomerPK=" & LockedCustomerPK & " for account locking")
      dst.Dispose()
      dst = Nothing
      ExistingLockExpireDate = CDate("1/1/1980")
      Common.QueryStr = "dbo.pa_UE_CustomerLockFetch"
      Common.Open_LXSsp()
      Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCuPk
      Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = LockingGroupId
      dst = Common.LXSsp_select
      Common.Close_LXSsp()
      If dst.Rows.Count > 0 Then
        LockID = dst.Rows(0).Item("LockID")
        LockLocationID = dst.Rows(0).Item("LocationID")
        ExistingLockExpireDate = Common.NZ(dst.Rows(0).Item("UE_LockExpireDate"), "1/1/2050")
      End If
      dst = Nothing

      'if (record exists and not expired) and ((from a different location) or (from the same location and RestartLock=0))
						If (LockID <> 0 And ExistingLockExpireDate > DateTime.Now) And ((LockLocationID <> LocationID) Or (LockLocationID = LocationID And LocationLockRestart = 0)) Then
								'There is a lock record in the table for this customer/terminalgroup, and the lock is NOT expired
								iLockStatus = 1
								Common.Write_Log(LogFile, "Lock for CustomerPK: " & CustomerPK & sHouseholdText & " and TerminalLockingGroupID: " & LockingGroupId & " already exists. (Lock does not expire until: " & ExistingLockExpireDate & ")")
      End If
								'if there is a lock record, it's not expired, but the locking location is the same as the requested location and LocationLockRestarts are allowed
								'-or- a record exists and it is expired (this is necessary to keep from violating the unique constraint on CPE_CustomerLocks
      If (LockID <> 0 And ExistingLockExpireDate > DateTime.Now And LockLocationID = LocationID And LocationLockRestart = 1) Or (LockID <> 0 And ExistingLockExpireDate < DateTime.Now) Then
								'delete the existing lock record because we are about to create a new one (the local server is over-riding the existing lock
								Common.QueryStr = "delete from CustomerLock with (RowLock) where LockID=" & LockID & ";"
								Common.LXS_Execute()
								Common.Write_Log(LogFile, "Deleted lockID (" & LockID & ") to make way for a new lock for CustomerPK: " & CustomerPK & sHouseholdText & " and TerminalLockingGroupID: " & LockingGroupId)
      End If

								'if there is no lock record
								' -or- there is a lock record, but it is expired
								' -or- there is a lock record, it's not expired, but the locking location is the same as the requested location and LocationLockRestarts are allowed
      If (LockID = 0) Or (LockID <> 0 And ExistingLockExpireDate < DateTime.Now) Or (LockID <> 0 And ExistingLockExpireDate > DateTime.Now And LockLocationID = LocationID And LocationLockRestart = 1) Then
								'create a new lock
								Common.QueryStr = "dbo.pa_UE_CustomerLockInsert"
								Common.Open_LXSsp()
								Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCuPk
								Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = LockingGroupId
								Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
								Common.LXSsp.Parameters.Add("@UE_LockExpireDate", SqlDbType.DateTime).Value = NewLockExpireDate
								Common.LXSsp.Parameters.Add("@TerminalNumber", SqlDbType.Int).Value = 0
								Common.LXSsp.Parameters.Add("@TransactionNumber", SqlDbType.NVarChar, 128).Value = "0"
								Common.LXSsp.Parameters.Add("@LockID", SqlDbType.BigInt).Direction = ParameterDirection.Output
								Common.LXSsp.ExecuteNonQuery()
								LockID = Common.LXSsp.Parameters("@LockID").Value
								Common.Close_LXSsp()
								iLockStatus = 0
								Common.Write_Log(LogFile, "Inserted lock for CustomerPK: " & CustomerPK & sHouseholdText & " and TerminalLockingGroupID: " & LockingGroupId)
						End If
				Else
						iLockStatus = 2
						Common.Write_Log(LogFile, "Customer (CustomerPK: " & CustomerPK & ")" & " does not exist!")
				End If

				Return iLockStatus

  End Function

  ' -----------------------------------------------------------------------------------------------

  Function Fetch_CustomerInfo_Data(ByVal CustomerPK As Long, ByVal HHPK As Long, ByVal DelimChar As String, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal NewUser As Integer, ByVal ExtCardID As String) As String

    Dim TempStr As String
    Dim dst As DataTable
    Dim row As DataRow
    Dim AltIDColumn As String = ""
    Dim VerifierColumn As String = ""

    'OperationType decoder ring:
    '  1 - text INSERT / UPDATE
    '  2 - record DELETE
    '  3 - image (binary) INSERT
    '  4 - new central key
    '  5 - mass insert
    '  6 - forced insert (same as #1 except we aren't looking to see if the record already exists)
    '  7 - INSERT/UPDATE with 2 fields
    '      (this is the same as #1 except we are looking at the first 2 fields to see if the record already exists)
    '  8 - INSERT/UPDATE with 3 fields
    '  10 - Incentive delete, without deleting associated customer data for the incentive
    '  11 - UPDATE with 2 fields (only executed if record already exists, will not insert)
    '  12 - DELETE with 2 fields

    TempStr = ""
    AltIDColumn = Common.Fetch_SystemOption(60)
    VerifierColumn = Common.Fetch_SystemOption(61)
    If Not (AltIDColumn = "") Then
      AltIDColumn = ", isnull(" & AltIDColumn & ", '') as AlternateID "
    Else
      AltIDColumn = ", '' as AlternateID"
    End If
    If Not (VerifierColumn = "") Then
      VerifierColumn = ", isnull(" & VerifierColumn & ", '') as Verifier "
    Else
      VerifierColumn = ", '' as Verifier"
    End If

    'Users
    'send the data from the Users table
    Common.QueryStr = "select Customers.CustomerPK as UserID, isnull(Customers.InitialCardID,'') as ClientUserID1, isnull(Customers.HHPK, 0) as HHPrimaryID, isnull(Customers.FirstName, '') as FirstName, isnull(Customers.LastName, '') as LastName, " & _
                      "  case CustomerTypeID when 2 then 0 else CustomerTypeID end as HHRec, CustomerTypeID, Customers.Employee, Customers.CurrYearSTD, " & _
                      "  Customers.LastYearSTD, isnull(Customers.CustomerStatusID, 0) as CustomerStatusID, AltIDOptOut " & AltIDColumn & VerifierColumn & ", " & _
                      "  isnull(Customers.EmployeeID,'') as EmployeeID, isnull(CustomerExt.AirmileMemberID,'') as AirmileMemberID, isnull(Customers.Prefix,'') as Prefix, " & _
                      "  isnull(Customers.Suffix,'') as Suffix,  Customers.CreatedDate" & _
                      " from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                      " where Customers.CustomerPK=" & CustomerPK & If(HHPK = 0, "", " or Customers.CustomerPK=" & HHPK) & ";"
    dst = Common.LXS_Select
    TempStr = TempStr & Construct_Table("Users", 1, DelimChar, LocalServerID, LocationID, dst)

    'CardIDs
    'send the data from the CardIDs table
    Common.QueryStr = "select CardIDs.CardPK, CardIDs.CustomerPK as UserID, CardIDs.ExtCardID, CardIDs.CardStatusID, CardIDs.CardTypeID " & _
                      " from CardIDs with (NoLock) " & _
                      " where CardIDs.CustomerPK=" & CustomerPK & If(HHPK = 0, "", " or CardIDs.CustomerPK=" & HHPK) & ";"
    dst = Common.LXS_Select
    TempStr = TempStr & Construct_Table("CardIDs", 1, DelimChar, LocalServerID, LocationID, dst)

    'CustomerAttributes
    'send the data from the CustomerAttributes table
    Common.QueryStr = "select CustomerPK, AttributeTypeID, AttributeValueID " & _
                      " from CustomerAttributes with (NoLock) " & _
                      " where CustomerAttributes.CustomerPK=" & CustomerPK & If(HHPK = 0, "", " or CustomerAttributes.CustomerPK=" & HHPK) & ";"
    dst = Common.LXS_Select
    TempStr = TempStr & Construct_Table("CustomerAttributes", 1, DelimChar, LocalServerID, LocationID, dst)

    'GroupMembership
    'send the data from the GroupMembership table
    Common.QueryStr = "dbo.pa_UE_GCI_GM"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
    Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
    dst = Common.LXSsp_select
    Common.Close_LXSsp()
    TempStr = TempStr & Construct_Table("GroupMembership", 7, DelimChar, LocalServerID, LocationID, dst)

    'RewardAccumulation
    'send the data from the RewardAccumulation table
    Common.QueryStr = "dbo.pa_UE_GCI_RA"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
    Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
    dst = Common.LXSsp_select
    Common.Close_LXSsp()
    TempStr = TempStr & Construct_Table("RewardAccumulation", 7, DelimChar, LocalServerID, LocationID, dst)

    'RewardDistribution
    'send the data from the RewardDistribution table
    Common.QueryStr = "dbo.pa_UE_GCI_RD"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
    Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
    dst = Common.LXSsp_select
    Common.Close_LXSsp()
    TempStr = TempStr & Construct_Table("RewardDistribution", 1, DelimChar, LocalServerID, LocationID, dst)

    'UserResponses
    Common.QueryStr = "dbo.pa_UE_GCI_CR"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
    Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
    dst = Common.LXSsp_select
    Common.Close_LXSsp()
    TempStr = TempStr & Construct_Table("UserResponses", 1, DelimChar, LocalServerID, LocationID, dst)

    'Points
    Common.QueryStr = "dbo.pa_UE_GCI_Points"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
    Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
    dst = Common.LXSsp_select
    Common.Close_LXSsp()
    TempStr = TempStr & Construct_Table("Points", 6, DelimChar, LocalServerID, LocationID, dst)

    If Common.Fetch_UE_SystemOption(134) = "1" Then  'see if Preference Data Distribution is enabled
      'CustomerPreferences
      Common.QueryStr = "dbo.pa_CPE_IN_CustomerPrefs"
      Common.Open_LXSsp()
      Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
      dst = Common.LXSsp_select
      Common.Close_LXSsp()
      TempStr = TempStr & Construct_Table("CustomerPreferences", 7, DelimChar, LocalServerID, LocationID, dst)

      'CustomerPreferencesMV
      Common.QueryStr = "dbo.pa_CPE_IN_CustomerPrefsMV"
      Common.Open_LXSsp()
      Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
      dst = Common.LXSsp_select
      Common.Close_LXSsp()
      TempStr = TempStr & Construct_Table("CustomerPreferencesMV", 8, DelimChar, LocalServerID, LocationID, dst)
    End If

    'StoredValue
    Common.QueryStr = "dbo.pa_UE_GCI_StoredValue"
    Common.Open_LXSsp()
    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
    Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
    dst = Common.LXSsp_select
    Common.Close_LXSsp()
    TempStr = TempStr & Construct_Table("StoredValue", 6, DelimChar, LocalServerID, LocationID, dst)

	' If Pending Points enabled
    If Common.Fetch_SystemOption(251) = "1" Then
       'Points Pending
       Common.QueryStr = "dbo.pa_UE_GCI_PointsPending"
       Common.Open_LXSsp()
       Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
       Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
       dst = Common.LXSsp_select
       Common.Close_LXSsp()
       TempStr = TempStr & Construct_Table("PointsPending", 6, DelimChar, LocalServerID, LocationID, dst)

       'RewardDistribution Pending
       Common.QueryStr = "dbo.pa_UE_GCI_RewardDistributionPending"
       Common.Open_LXSsp()
       Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
       Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
       dst = Common.LXSsp_select
       Common.Close_LXSsp()
       TempStr = TempStr & Construct_Table("RewardDistributionPending", 1, DelimChar, LocalServerID, LocationID, dst)
       
       'RewardLimitVariables Pending
       Common.QueryStr = "dbo.pa_UE_GCI_RewardLimitVariablesPending"
       Common.Open_LXSsp()
       Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
       Common.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = HHPK
       dst = Common.LXSsp_select
       Common.Close_LXSsp()
       TempStr = TempStr & Construct_Table("RewardLimitVariablesPending", 1, DelimChar, LocalServerID, LocationID, dst)
    End If

    Fetch_CustomerInfo_Data = TempStr

  End Function

  ' -----------------------------------------------------------------------------------------------

  Class Customer

    Public m_cust_extCardID As String = ""
    Public m_cardtype As Integer = 0
    Public m_custpk As Long = 0
    Public m_custtype As Integer = 0
    Public m_hhpk As Long = 0
    Public m_hh_extCardID As String = ""
    Public m_should_create As Boolean = False

        Private m_Common As Copient.CommonInc ' needs to be open when it's passed to us
        Private MyCryptlib As New Copient.CryptLib

    Enum CardTypes
      CUSTOMER = 0
      HOUSEHOLD = 1
      CAM = 2
      FASTFORWARD = 3
    End Enum

    Enum CustomerTypes
      CUSTOMER = 0
      HOUSEHOLD = 1
      CAM = 2
    End Enum

    Enum CardStatuses
      ACTIVE = 1
      INACTIVE = 2
      CANCELED = 3
      EXPIRED = 4
      LOST_STOLEN = 5
      DEFAULT_CARD = 6
    End Enum

    Private Sub addMissingCardToCardIDsTable(ByVal custpk As Long, ByVal extCardID As String, ByVal cardtype As Integer)
      extCardID = m_Common.Pad_ExtCardID(extCardID, cardtype)

      m_Common.QueryStr = _
          "INSERT INTO cardids ( customerpk, extcardid, cardstatusid, cardtypeid ) " & _
          "  VALUES ( " & custpk & ", '" & MyCryptlib.SQL_StringEncrypt(extCardID) & "', 1, " & cardtype & " );"
      m_Common.LXS_Execute()
    End Sub ' addMissingCardToCardIDsTable

    Sub refresh()

      Dim dst As DataTable
      Dim CustomerPK As Long
      m_cust_extCardID = m_Common.Pad_ExtCardID(m_cust_extCardID, m_cardtype)

      CustomerPK = 0
      ' get the CustomerPK and HouseholdPK based on the ExtCardID
      m_Common.QueryStr = "select isnull(CustomerPK, 0) as CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCryptlib.SQL_StringEncrypt(m_cust_extCardID) & "' and CardTypeID=" & m_cardtype & ";"
      dst = m_Common.LXS_Select
      If dst.Rows.Count > 0 Then
        CustomerPK = dst.Rows(0).Item("CustomerPK")
      End If
      If CustomerPK = 0 AndAlso m_should_create Then  ' If it wasn't found from the looking up the card in the CardIDs table, try looking it up from the InitialCardID field
                m_Common.QueryStr = "SELECT isnull(CustomerPK, 0) as CustomerPK from Customers with (NoLock) where InitialCardID = '" & MyCryptlib.SQL_StringEncrypt(m_cust_extCardID) & "' and InitialCardTypeID = " & m_cardtype & ";"
        dst = m_Common.LXS_Select
        If dst.Rows.Count > 0 Then
          CustomerPK = dst.Rows(0).Item("CustomerPK")
          ' the customer exists with this card, but the card isn't in the CardIDs table
          addMissingCardToCardIDsTable(CustomerPK, m_cust_extCardID, m_cardtype)
        End If
      End If

      m_custpk = CustomerPK
      m_custtype = 0
      m_hhpk = 0

      If m_custpk > 0 Then
        'Now that we know the customerPK, fetch the CustomerType and HHPK
        m_Common.QueryStr = "SELECT ISNULL(Customers.HHPK, 0) AS HHPK, isnull(CustomerTypeID, 0) as CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
        dst = m_Common.LXS_Select
        If dst.Rows.Count > 0 Then
          m_custtype = dst.Rows(0).Item("CustomerTypeID")
          m_hhpk = dst.Rows(0).Item("HHPK")
        End If

        ' now get the associated household card
        If m_hhpk <> 0 AndAlso m_custtype <> CustomerTypes.HOUSEHOLD Then
          m_hhpk = m_Common.Pad_ExtCardID(m_hhpk, CardTypes.HOUSEHOLD)

          m_Common.QueryStr = "SELECT ExtCardID from CardIDs with (NoLock) where CustomerPK=" & m_hhpk & " and CardTypeID=" & CardTypes.HOUSEHOLD & ";"
          dst = m_Common.LXS_Select
          If dst.Rows.Count > 0 Then
            m_hh_extCardID = MyCryptlib.SQL_StringDecrypt(dst.Rows(0).Item("ExtCardID").ToString())
          End If
        End If
      End If  'm_custpk > 0

    End Sub ' refresh()

    Sub New(ByVal custid As String, ByVal cardtype As Integer, ByVal shouldcreate As Boolean, ByRef Common As Copient.CommonInc)

      custid = Common.Pad_ExtCardID(custid, cardtype)

      m_cust_extCardID = custid
      m_cardtype = cardtype
      m_should_create = shouldcreate
      m_Common = Common
      refresh()
    End Sub ' new

    Function needsCustomerRecord() As Boolean
      Return (m_custpk = 0)
    End Function

    Function needsHouseHold() As Boolean
      Return (m_hhpk = 0)
    End Function

    Function needsHouseHoldCard() As Boolean
      Return (m_hh_extCardID = "")
    End Function

    Function isNotComplete() As Boolean
      Return needsCustomerRecord() OrElse needsHouseHold() OrElse needsHouseHoldCard()
    End Function

    Public Overrides Function toString() As String
      Return "m_cust_extCardID: " & m_cust_extCardID & _
           "; m_cardtype: " & m_cardtype & _
           "; m_custpk: " & m_custpk & _
           "; m_custtype: " & m_custtype & _
           "; m_hhpk: " & m_hhpk & _
           "; m_hh_extCardID: " & m_hh_extCardID
    End Function

  End Class ' Customer ------------------------------------------------------------------------

  Function createCustomersInMasterDataLibrary() As Boolean
    Static doCreate As Boolean = (Common.Fetch_UE_SystemOption(113) = "1")
    Return doCreate
  End Function

  ' -----------------------------------------------------------------------------------------------

  Sub SetCustomerHouseHold(ByVal CustomerPK As Integer, ByVal hhpk As Integer)
    Common.QueryStr = "UPDATE Customers SET hhpk = " & hhpk & " WHERE customerpk = " & CustomerPK & ";"
    Common.LXS_Execute()
  End Sub

  ' -----------------------------------------------------------------------------------------------

  Function CreateCustomer(ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal LocationID As Long, ByVal BannerID As Integer, Optional ByVal hhpk As Long = 0) As Long
    'run the Stored Procedure to insert a record into Customers (returns the new PrimaryKey)

    ExtCardID = Common.Pad_ExtCardID(ExtCardID, CardTypeID)

    Common.QueryStr = "dbo.pa_CPE_IN_CreateCustomer"
    Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@InitialCardID", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID, True)
    Common.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CardTypeID
    Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
    Common.LXSsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
    Common.LXSsp.Parameters.Add("@InitialCardTypeID", SqlDbType.Int).Value = CardTypeID
        Common.LXSsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        Common.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(ExtCardID)
    Common.LXSsp.ExecuteNonQuery()
    Dim CustomerPK As Long = Common.LXSsp.Parameters("@PKID").Value
    Common.Close_LXSsp()

    If hhpk <> 0 Then
      SetCustomerHouseHold(CustomerPK, hhpk)
    End If

    Return CustomerPK

  End Function

  ' -----------------------------------------------------------------------------------------------

  Function AddCardToHouseHold(ByVal HHExtCardID As String, ByVal HHPK As Long) As Long
    '[dbo].[pt_NewCardIDs_Insert] @ExtCardID nvarchar(26), @CardTypeID int, @CustomerPK bigint, @CardStatusID int, @CardPK bigint OUTPUT, @Created bit = 0 OUTPUT
    HHExtCardID = Common.Pad_ExtCardID(HHExtCardID, Customer.CardTypes.HOUSEHOLD)
    Common.QueryStr = "dbo.pt_NewCardIDs_Insert"
    Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(HHExtCardID, True)
    Common.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = Customer.CardTypes.HOUSEHOLD
    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = HHPK
    Common.LXSsp.Parameters.Add("@CardStatusID", SqlDbType.BigInt).Value = Customer.CardStatuses.ACTIVE
        Common.LXSsp.Parameters.Add("@CardPK", SqlDbType.BigInt).Direction = ParameterDirection.Output
        Common.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(HHExtCardID)
    Common.LXSsp.ExecuteNonQuery()
    Dim newcardpk As Long = Common.LXSsp.Parameters("@CardPK").Value
    Common.Close_LXSsp()
    Return newcardpk
  End Function

  ' -----------------------------------------------------------------------------------------------

  Function DoCustomerCreation(ByVal theCust As Customer, ByVal ExtCardID As String, ByVal LocationID As String, ByVal BannerID As String, ByRef newHHPK As Long) As Long

    Const INTERFACE_OPTION_MDM_USERNAME As Integer = 28
    Const INTERFACE_OPTION_MDM_PASSWORD As Integer = 29
    Const INTERFACE_OPTION_MDM_URL As Integer = 30

    Try
      ' get the username, password, and URL from the interface options
      Dim mdm_user As String = Common.Fetch_InterfaceOption(INTERFACE_OPTION_MDM_USERNAME)
      Dim mdm_pw As String = Common.Fetch_InterfaceOption(INTERFACE_OPTION_MDM_PASSWORD)
      Dim mdm_url As String = Common.Fetch_InterfaceOption(INTERFACE_OPTION_MDM_URL)

      Dim mdm As New Copient.MasterDataLibrary(mdm_user, mdm_pw, mdm_url)
      Dim newHouseHoldExtCardID As String = ""
      Dim CustomerPK As Long = theCust.m_custpk

      If mdm.createCustomer(ExtCardID, newHouseHoldExtCardID) Then
        'NewUser = 1
        If theCust.needsCustomerRecord() Then
          CustomerPK = CreateCustomer(ExtCardID, Customer.CardTypes.CUSTOMER, LocationID, BannerID)
        End If

        If newHouseHoldExtCardID <> "" Then
          If theCust.needsHouseHold() Then ' create the household and add the existing customer to the household
            newHHPK = CreateCustomer(newHouseHoldExtCardID, Customer.CardTypes.HOUSEHOLD, LocationID, BannerID)
            SetCustomerHouseHold(CustomerPK, newHHPK)
          ElseIf theCust.needsHouseHoldCard() Then ' just add the new house hold card id to the household that already exists
            AddCardToHouseHold(newHouseHoldExtCardID, theCust.m_hhpk)
          End If
        End If ' newHouseHoldExtCardID <> ""

        Return CustomerPK
      End If ' mdm.createCustomer()

    Catch e As ApplicationException
      Common.Write_Log(LogFile, "Failed to create user in master data: " & e.Message)
    End Try

    Send_Response_Header("Failed to create user in master data", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Common.Write_Log(LogFile, "Failed to create user in master data")
    Return 0
  End Function

  ' -----------------------------------------------------------------------------------------------

  Sub Handle_Lookup_Request(ByVal LocalServer As Copient.ConnectorInc.UELSRec, ByVal Mode As String)

    Dim ex As System.Data.SqlClient.SqlException
    Dim DupKey As Boolean
    Dim ExtCardID As String
    Dim CustomerPK As Long
    Dim HHPK As Long
    Dim OutStr As String
    Dim DelimChar As Integer
    Dim dst As DataTable
    Dim CardTypeID As Integer
    Dim NewUser As Integer = 0
    Dim LockingGroupId As Long
    Dim sLockOption As String
    Dim iLockStatus As Integer
    Dim CAMErrorMsg As String
    Dim LockedCustomerPK As Long
    Dim ShouldCreateCard As Boolean = False
    Dim CreateCust As String
    Dim NoLock As String
    Dim LockID As Long 'LockID from CPE_CustomerLocks (if the customer account is already locked, or a lock is created during this response)
    Dim isCardValid As Boolean
    Dim CardResponse As CardValidationResponse

    DelimChar = 30

    On Error GoTo QueryError

    NoLock = GetCgiValue("nolock")  'is used to prevent locking a customer record if locking is enabled

    OutStr = ""
    OutStr = OutStr & "LocationID=" & LocalServer.LocationID & vbCrLf

    'Update the LocalServers.InfoNowLastHeard date
    Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServer.LocalServerID
    Common.LRTsp.ExecuteNonQuery()
    Common.Close_LRTsp()

    'grab the customer's identifier
    ExtCardID = Trim(Request.QueryString("id"))
    If ExtCardID = "" Then
      Send_Response_Header("Missing customer ID (id)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "Missing customer ID (id)")
      Exit Sub
    End If

    'see if we are supposed to create a new customer record if the customer is not found
    CreateCust = Request.QueryString("CreateCust")
    If CreateCust = "" Then
      ShouldCreateCard = True
    Else
      ShouldCreateCard = (CreateCust = "1")
    End If

    CardTypeID = Common.Extract_Val(Request.QueryString("ctype"))
    'Verify the customer\card id complies with CardTypeID supplied if any, default CardTypeID is 0
    isCardValid = Common.AllowToProcessCustomerCard(ExtCardID, CardTypeID, CardResponse)
    If Not isCardValid Then
      If CardResponse = CardValidationResponse.INVALIDCARDTYPEFORMAT Or CardResponse = CardValidationResponse.CARDTYPENOTFOUND Then
        Send_Response_Header("Card type ID(ctype): " & CardTypeID & " is an invalid card type.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Card type ID(ctype): " & CardTypeID & " is an invalid card type.")
        Exit Sub
      ElseIf CardResponse = CardValidationResponse.CARDIDNOTNUMERIC Then
        Send_Response_Header("Customer Id (Id): " & ExtCardID & " must be numeric", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                Common.Write_Log(LogFile, "Customer Id (Id): " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & " of type id " & CardTypeID & " must be numeric")
        Exit Sub
      ElseIf CardResponse = CardValidationResponse.INVALIDCARDFORMAT Then
        Send_Response_Header("Customer Id (Id): " & ExtCardID & " is invalid", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                Common.Write_Log(LogFile, "Customer Id (Id): " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & " of type id " & CardTypeID & " is invalid")
        Exit Sub
      ElseIf CardResponse = CardValidationResponse.ERROR_APPLICATION Then
        Send_Response_Header("Application Error encountered", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Application Error encountered")
        Exit Sub
      End If
    End If

    ExtCardID = Common.Pad_ExtCardID(ExtCardID, CardTypeID)
    If ExtCardID = "" Then
      Send_Response_Header("Missing customer ID", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Common.Write_Log(LogFile, "Missing customer ID!")
    Else
      'look up the Card Number we got
            Common.Write_Log(LogFile, "Received customer ID (id): " & Copient.MaskHelper.MaskCard(ExtCardID, CardTypeID) & "   CardTypeID (ctype)=" & CardTypeID)

      If CardTypeID = 2 Then 'make sure this is a valid CAM card
        CAMErrorMsg = ""
        If Not (CAM.VerifyCardNumber(ExtCardID, CAMErrorMsg)) Then
          Send_Response_Header("Invalid CAM Card Number - " & CAMErrorMsg, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Common.Write_Log(LogFile, "Invalid CAM Card Number - " & CAMErrorMsg)
          Exit Sub
        End If
      End If

      Dim theCust As New Customer(ExtCardID, CardTypeID, ShouldCreateCard, Common)
      CustomerPK = theCust.m_custpk
      HHPK = theCust.m_hhpk

      If createCustomersInMasterDataLibrary() Then
        If theCust.isNotComplete() Then ' the customer is missing some vital piece
          If ShouldCreateCard Then
            HHPK = 0
            CustomerPK = DoCustomerCreation(theCust, ExtCardID, LocalServer.LocationID, LocalServer.BannerID, HHPK)
            If CustomerPK <> 0 Then NewUser = 1
          Else
            Send_Response_Header("User Not Found Or Not Complete", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "User not found or not complete - create request not sent from store")
            Exit Sub
          End If ' should create card
        Else
          Common.Write_Log(LogFile, "Found existing CustomerPK: " & CustomerPK)
        End If
      Else
        If theCust.needsCustomerRecord() Then  'The customer was not found in the Logix database and we are NOT using the master data library
          If ShouldCreateCard Then
            CustomerPK = CreateCustomer(ExtCardID, CardTypeID, LocalServer.LocationID, LocalServer.BannerID)
            Common.Write_Log(LogFile, "User not found - created UserID: " & CustomerPK)
            NewUser = 1
          Else 'If CustomerPK = 0 AndAlso Not ShouldCreateCard Then
            Send_Response_Header("User Not Found", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "User not found - create request not sent from store")
            Exit Sub
          End If
        Else
          Common.Write_Log(LogFile, "Found existing CustomerPK: " & CustomerPK)
        End If
      End If

      If LocalServer.LocationAssociationType = Copient.ConnectorInc.AssociationType.ActiveServer Then
        'make sure this customer is associated with this location
        Common.QueryStr = "dbo.pa_CPE_IN_UpdateCustomerLocation"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocalServer.LocationID
        Common.LXSsp.ExecuteNonQuery()
        Common.Close_LXSsp()

        If Not HHPK = 0 Then 'if the customer is a member of a household ...
          'see if the household is already associated with this location
          Dim IsLocationTypeServer As Boolean = False
          Common.QueryStr = "select LocationTypeID from Locations where LocationId = " & LocalServer.LocationID & " and LocationTypeID = 2 and EngineID = 9"
          dst = Common.LRT_Select
          If dst.Rows.Count > 0
            IsLocationTypeServer = True
          End If          
          
          Common.QueryStr = "select 1 from CustomerLocations with (NoLock) where CustomerPK=" & HHPK & " and LocationID=" & LocalServer.LocationID & ";"
          dst = Common.LXS_Select
                    
          If dst.Rows.Count > 0 AndAlso Not IsLocationTypeServer  'household is already associated with this location - no need to update CustomerLocations or send data for the household
            HHPK = 0
          Else
            Common.QueryStr = "dbo.pa_CPE_IN_UpdateCustomerLocation"
            Common.Open_LXSsp()
            Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = HHPK
            Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocalServer.LocationID
            Common.LXSsp.ExecuteNonQuery()
            Common.Close_LXSsp()
          End If
        End If
      End If

      OutStr = OutStr & Fetch_CustomerInfo_Data(CustomerPK, HHPK, DelimChar, LocalServer.LocalServerID, LocalServer.LocationID, NewUser, ExtCardID)

      'Customer Locking
      sLockOption = Common.Fetch_UE_SystemOption(86)
      If sLockOption = "1" Or sLockOption = "2" Then  'lock based on CustID only -or- lock based on CustID & terminal lock group

        If sLockOption = "1" Then
          LockingGroupId = 0
        Else
          LockingGroupId = Common.NZ((Request.QueryString("lockinggroupid")), 0)
        End If
        LockedCustomerPK = CustomerPK
        iLockStatus = -1
        If Not (NoLock = "1") Then
          iLockStatus = HandleCustomerLocking(sLockOption, CustomerPK, LockingGroupId, LocalServer.LocationID, LockedCustomerPK, LockID)
        End If  'NoLock = 1
        Common.QueryStr = "dbo.pa_UE_CustomerLockStatus"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = LockedCustomerPK
        Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = LockingGroupId
        Common.LXSsp.Parameters.Add("@LockStatus", SqlDbType.Int).Value = iLockStatus
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        'LockStatus = 0 - Lock granted
        'LockStatus = 1 - Lock already exists for the specified locking group
        'LockStatus = 2 - Customer not found - lock can't be granted
        'LockStatus = 3 - Customer record is not locked
        OutStr = OutStr & Construct_Table("UsersLockStatus", 12, DelimChar, LocalServer.LocalServerID, LocalServer.LocationID, dst)
      End If 'sLockOption=1 or 2

      Common.Write_Log(LogFile, "Returned the following data for this customer:")

      Send_Response_Header(ResponseStringHeader, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)

      'send the total execution time to the local server so it can be logged there
      TotalTime = DateAndTime.Timer - StartTime
      OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
      OutStr = Len(OutStr) & vbCrLf & OutStr
      Send(OutStr)
      Common.Write_Log(LogFile, OutStr)

    End If

AllDone:
    Exit Sub

QueryError:
    DupKey = False
    If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
      ex = Err.GetException
      If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
        DupKey = True
      End If
    End If
    If DupKey Then
      Response.Write("Duplicate CardIDs.ExtCardID violation - two simultaneous GetCustomerInfo's for the same card must have been received. " & ex.ToString)
      Common.Write_Log(LogFile, "Duplicate CardIDs.ExtCardID violation - two simultaneous GetCustomerInfo's for the same card must have been received." & " serial=" & LocalServer.LocalServerID & "Mac IPAddress=" & LocalServer.MacAddress & " server=" & Environment.MachineName)
    Else
      Response.Write(Common.Error_Processor(, "Serial= " & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IP=" & LocalServer.IPAddress & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocalServer.LocationID, , Common.InstallationName))
      Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
    End If
    Response.End()

  End Sub

  ' -----------------------------------------------------------------------------------------------

  Sub Handle_AltID_Request(ByVal LocalServer As Copient.ConnectorInc.UELSRec, ByVal Mode As String)

    Dim ex As System.Data.SqlClient.SqlException
    Dim DupKey As Boolean
    Dim AltID As String
    Dim OutStr As String
    Dim DelimChar As Integer
    Dim dst As DataTable
    Dim AltIDColumn As String
    Dim VerifierColumn As String
    Dim AltIDUniqueness As Integer
    Dim OrigAltIdColumn As String
    Dim OrigVerifierColumn As String
    Dim VerifierID As String
    Dim ProcessOK As Boolean
    Dim CustList As String
    Dim row As DataRow

    DelimChar = 30

    On Error GoTo QueryError

    AltIDColumn = Trim(Common.Fetch_SystemOption(60))
    OrigAltIdColumn = AltIDColumn
    VerifierColumn = Trim(Common.Fetch_SystemOption(61))
    OrigVerifierColumn = VerifierColumn
    If Not (AltIDColumn = "") Then
      AltIDColumn = "isnull(" & AltIDColumn & ", '') as AlternateID"
    End If
    If Not (VerifierColumn = "") Then
      VerifierColumn = "isnull(" & VerifierColumn & ", '') as Verifier"
    Else
      VerifierColumn = "'' as Verifier"
    End If

    AltIDUniqueness = Common.Fetch_SystemOption(81)

    OutStr = ""
    OutStr = OutStr & "LocationID=" & LocalServer.LocationID & vbCrLf

    Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServer.LocalServerID
    Common.LRTsp.ExecuteNonQuery()
    Common.Close_LRTsp()

    AltID = Trim(Request.QueryString("altid"))
    VerifierID = Trim(Request.QueryString("verifier"))  'need to modify code to save this
    Common.Write_Log(LogFile, "altid=" & AltID & "  verifier=" & VerifierID & " Serial=" & LocalServer.LocalServerID & " IP=" & LocalServer.IPAddress & " MacAddress=" & LocalServer.MacAddress)

    'look for the things we NEED to HAVE in order to continue
    ProcessOK = True
    If AltID = "" Then
      OutStr = OutStr & "R:04 Missing AltID" & vbCrLf
      ProcessOK = False
    End If
    'If Not (OrigVerifierColumn = "") And VerifierID = "" Then  'a verifier is required and we didn't get one from the local server
    '  OutStr = OutStr & "R:04 Missing VerifierID (verifier)" & vbCrLf
    '  ProcessOK = False
    'End If
    If OrigAltIdColumn = "" Then
      OutStr = OutStr & "R:05 No Customer Alternate ID has been specified in SystemOptions" & vbCrLf
      ProcessOK = False
    End If

    If ProcessOK Then

      'look up the Card Number we got
            Common.Write_Log(LogFile, "Received AltID: " & AltID)
            
            If AltIDUniqueness <> 2 Then
                'VerifierCondition = ""
                'If AltIDUniqueness = 3 Then  'unique by AltID and Verifier
                '  VerifierCondition = " and " & OrigVerifierColumn & "='" & VerifierID & "' "
                'End If
                Common.QueryStr = "select Customers.CustomerPK, " & AltIDColumn & ", " & VerifierColumn & ", isnull(Customers.InitialCardID, '') as ClientUserID1, isnull(Customers.CustomerTypeID, 0) as HHRec, isnull(Customers.Employee, 0) as Employee, 0 as RecordStatus, isnull(Customers.FirstName, '') as FirstName, isnull(Customers.LastName, '') as LastName, isnull(Customers.EmployeeID, '') as EmployeeID " & _
                                  "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                                  "where isnull(Customers.CustomerStatusID, 1)=1 and " & OrigAltIdColumn & "='" & AltID & "';"
            Else
                Common.QueryStr = "select Customers.CustomerPK, " & AltIDColumn & ", " & VerifierColumn & ", isnull(Customers.InitialCardID, '') as ClientUserID1, isnull(Customers.CustomerTypeID, 0) as HHRec, isnull(Customers.Employee, 0) as Employee, 0 as RecordStatus, isnull(Customers.FirstName, '') as FirstName, isnull(Customers.LastName, '') as LastName, isnull(Customers.EmployeeID, '') as EmployeeID " & _
                                  "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                                  "where isnull(Customers.CustomerStatusID, 1)=1 and " & OrigAltIdColumn & "='" & AltID & "' and Customers.BannerID=" & LocalServer.BannerID & ";"
            End If
            dst = Common.LXS_Select
            OutStr = OutStr & Construct_Table("AltIDCacheUsers", 1, 30, LocalServer.LocalServerID, LocalServer.LocationID, dst)

            If dst.Rows.Count > 0 Then
                CustList = "(-77"
                For Each row In dst.Rows
                    CustList = CustList & "," & row.Item("CustomerPK")
                Next
                CustList = CustList & ")"

                Common.QueryStr = "select CardIDs.CardPK, CardIDs.CustomerPK, CardIDs.ExtCardID, CardIDs.CardStatusID, CardIDs.CardTypeID " & _
                                  " from CardIDs with (NoLock) " & _
                                  " where CardIDs.CustomerPK IN " & CustList & ";"
                dst = Common.LXS_Select
                OutStr = OutStr & Construct_Table("AltIDCacheCardIDs", 1, 30, LocalServer.LocalServerID, LocalServer.LocationID, dst)
            End If
        End If

        Common.Write_Log(LogFile, "Returned the following data for this AltID lookup:")

        Send_Response_Header(ResponseStringHeader, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        'send the total execution time to the local server so it can be logged there
        TotalTime = DateAndTime.Timer - StartTime
        OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf & " Serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress
        OutStr = Len(OutStr) & vbCrLf & OutStr
        Send(OutStr)
        Common.Write_Log(LogFile, OutStr)

AllDone:
        Exit Sub

QueryError:
        DupKey = False
        If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
            ex = Err.GetException
            If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
                DupKey = True
            End If
        End If
        If DupKey Then
            Resume Next
        Else
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IPAddress=" & LocalServer.IPAddress & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocalServer.LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()

    End Sub

    ' -----------------------------------------------------------------------------------------------

    Sub Handle_AltIDEnrollment_Request(ByVal LocalServer As Copient.ConnectorInc.UELSRec, ByVal Mode As String)

        Dim ex As System.Data.SqlClient.SqlException
        Dim DupKey As Boolean
        Dim AltID As String
        Dim OutStr As String
        Dim DelimChar As Integer
        Dim dst As DataTable
        Dim AltIDColumn As String
        Dim VerifierColumn As String
        Dim VerifierCondition As String
        Dim AltIDUniqueness As Integer
        Dim OrigAltIdColumn As String
        Dim OrigVerifierColumn As String
        Dim AutoGenOn As Boolean = False
        Dim ExtCardID As String = ""
        Dim CustomerPK As Long
        Dim NewCustomerPK As Long
        Dim ShouldCreateCard As Boolean = False
        Dim CreateCust As String
        Dim VerifierID As String
        Dim AltTable() As String
        Dim VerifierTable() As String
        Dim NumRecs As Long
        Dim CustomerData As String = "" 'the standard GetCustomerInfo data returned when creating a NEW customer record
        Dim ProcessOK As Boolean
        Dim SendAllCustData As Boolean

        DelimChar = 30

        On Error GoTo QueryError

        SendAllCustData = 0
        AltIDColumn = Trim(Common.Fetch_SystemOption(60))
        OrigAltIdColumn = AltIDColumn
        VerifierColumn = Trim(Common.Fetch_SystemOption(61))
        OrigVerifierColumn = VerifierColumn
        If Not (AltIDColumn = "") Then
            AltIDColumn = "isnull(" & AltIDColumn & ", '') as AlternateID"
        End If
        If Not (VerifierColumn = "") Then
            VerifierColumn = "isnull(" & VerifierColumn & ", '') as Verifier"
        Else
            VerifierColumn = "'' as Verifier"
        End If

        AltIDUniqueness = Common.Fetch_SystemOption(81)

        OutStr = ""
        OutStr = OutStr & "LocationID=" & LocalServer.LocationID & vbCrLf

        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServer.LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()

        AltID = Trim(Request.QueryString("altid"))
        VerifierID = Trim(Request.QueryString("verifier"))  'need to modify code to save this
        CustomerPK = Common.Extract_Val(Trim(Request.QueryString("cardid"))) 'this is the CustomerPK
        Common.Write_Log(LogFile, "altid=" & AltID & "  verifier=" & VerifierID & "  cardid=" & CustomerPK & "  createcust=" & Request.QueryString("createcust"))

        'look for the things we NEED to HAVE in order to continue
        ProcessOK = True
        If AltID = "" Then
            OutStr = OutStr & "R:04 Missing AltID" & vbCrLf
            ProcessOK = False
        End If
        If Not (OrigVerifierColumn = "") AndAlso VerifierID = "" Then  'a verifier is required and we didn't get one from the local server
            OutStr = OutStr & "R:04 Missing VerifierID (verifier)" & vbCrLf
            ProcessOK = False
        End If
        If OrigAltIdColumn = "" Then
            OutStr = OutStr & "R:05 No Customer Alternate ID has been specified in SystemOptions" & vbCrLf
            ProcessOK = False
        End If

        If ProcessOK Then
            AutoGenOn = False
            If Common.Fetch_UE_SystemOption(87) = "1" Then AutoGenOn = True
            If AutoGenOn Then
                CreateCust = Request.QueryString("createcust")
                If CreateCust = "" Then
                    ShouldCreateCard = False
                Else
                    ShouldCreateCard = (CreateCust = 1)
                End If
            Else
                ShouldCreateCard = False
            End If

            If CustomerPK = 0 And Not ShouldCreateCard Then
                OutStr = OutStr & "R:06 Missing CustomerPK (cardid) in AltID Enrollment.  No create card request sent from the store." & vbCrLf
                ProcessOK = False
            End If

            If Not (CustomerPK = 0) Then
                'check to see if this Customers record really exists
                Common.QueryStr = "select count(*) as NumRecs from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
                dst = Common.LXS_Select
                NumRecs = dst.Rows(0).Item("NumRecs")
                If NumRecs = 0 Then  'the CustomerPK sent from the local server is not valid!
                    OutStr = OutStr & "R:06 The specified CustomerPK (" & CustomerPK & ") is invalid." & vbCrLf
                    ProcessOK = False
                    Exit Sub
                End If
            End If
        End If
        If ProcessOK Then
            If ShouldCreateCard AndAlso CustomerPK = 0 Then
                Common.Write_Log(LogFile, "Received in non-carded Alt ID enroll mode")
            Else
                Common.Write_Log(LogFile, "Received CardID (CustomerPK): " & CustomerPK & " in enroll mode")
            End If
            Common.Write_Log(LogFile, "Received AltID: " & AltID & " in enroll mode")

            If Not (AltIDUniqueness = 0) Then  'AltID's should be unique
                If Not (AltIDUniqueness = 2) Then 'AltID's are not unique by BannerID
                    ' see if the AltID is already in use by someone other than the current CustomerPK
                    VerifierCondition = ""
                    If AltIDUniqueness = 3 Then  'unique by AltID and Verifier
                        VerifierCondition = " and " & OrigVerifierColumn & "='" & VerifierID & "' "
                    End If
                    Common.QueryStr = "select count(*) as NumRecs " & _
                                      "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                                      "where not(Customers.customerpk=" & CustomerPK & ") and " & OrigAltIdColumn & "='" & AltID & "' " & VerifierCondition & ";"
                Else  'AltID's ARE unique by BannerID
                    ' see if the AltID is already in use by someone other than the current CustomerPK within this banner
                    Common.QueryStr = "select count(*) as NumRecs " & _
                                      "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                                      "where not(Customers.customerpk=" & CustomerPK & ") and " & OrigAltIdColumn & "='" & AltID & "' and Customers.BannerID=" & LocalServer.BannerID & ";"
                End If
                dst = Common.LXS_Select
                NumRecs = dst.Rows(0).Item("NumRecs")
            Else
                NumRecs = 0
            End If

            If NumRecs > 0 Then
                OutStr = OutStr & "R:03 Alternate ID is already in use" & vbCrLf
            Else
                ' New Alt ID is not in use so we can update the customer record
                AltTable = OrigAltIdColumn.Split(".")
                VerifierTable = OrigVerifierColumn.Split(".")

                'Send("update " & AltTable(0) & " set " & BareAltIdColumn & "='" & AltID & "' where CustomerPK=" & CardID)

                Common.Write_Log(LogFile, "CustomerPK=" & CustomerPK & "   ShouldCreateCard=" & ShouldCreateCard)
                ' if the auto-generate card number option is on, then create a new card number for this non-card AltID enroll request
                If CustomerPK = 0 AndAlso ShouldCreateCard Then
                    NewCustomerPK = MyAltID.GetAutoGeneratedCustomerPK(ExtCardID, LocalServer.BannerID) 'this is in the CustomerInquiry DLL
                    If NewCustomerPK > 0 Then
                        CustomerPK = NewCustomerPK
                        Common.Write_Log(LogFile, "Auto-generated customer " & Copient.MaskHelper.MaskCard(ExtCardID, Copient.commonShared.CardTypes.CUSTOMER) & " (" & CustomerPK & ") during AltID enrollment.")

                        If LocalServer.LocationAssociationType = Copient.ConnectorInc.AssociationType.ActiveServer Then
                            'make sure this customer is associated with this location
                            Common.QueryStr = "dbo.pa_CPE_IN_UpdateCustomerLocation"
                            Common.Open_LXSsp()
                            Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                            Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocalServer.LocationID
                            Common.LXSsp.ExecuteNonQuery()
                            Common.Close_LXSsp()
                            Common.Write_Log(LogFile, "Associated new customer (" & CustomerPK & ") with this location (" & LocalServer.LocationID & ").")
                        End If

                        'send the normal load of GetCustomerInfo data for the newly created customer record
                        SendAllCustData = True
                    Else
                        CustomerPK = 0
                        Send_Response_Header(MyAltID.ErrorMessage, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                        Common.Write_Log(LogFile, MyAltID.ErrorMessage & " Serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IP=" & LocalServer.IPAddress & " server=" & Environment.MachineName)
                    End If
                End If

                If Not (CustomerPK = 0) Then
                    Common.QueryStr = "update " & AltTable(0) & " with (rowlock) set " & OrigAltIdColumn & "='" & AltID & "' where CustomerPK=" & CustomerPK
                    Common.LXS_Execute()

                    If (Common.RowsAffected > 0) Then
                        ' the update worked
                        If Not (OrigVerifierColumn = "") Then
                            Common.QueryStr = "Update " & VerifierTable(0) & " with (rowlock) set " & OrigVerifierColumn & "='" & VerifierID & "' where CustomerPK= " & CustomerPK
                            Common.LXS_Execute()
                        End If
                        OutStr = OutStr & "R:01 Alternate ID Stored" & vbCrLf
                    Else

                        ' if we got here we need to attempt an insert instead of an update
                        Common.QueryStr = "insert into " & AltTable(0) & " with (rowlock) ( customerpk," & OrigAltIdColumn & ") values (" & CustomerPK & ",'" & AltID & "')"
                        Common.LXS_Execute()

                        If (Common.RowsAffected > 0) Then
                            ' the insert worked
                            If Not (OrigVerifierColumn = "") Then
                                Common.QueryStr = "Update " & VerifierTable(0) & " with (rowlock) set " & OrigVerifierColumn & "='" & VerifierID & "' where CustomerPK= " & CustomerPK
                                Common.LXS_Execute()
                            End If
                            OutStr = OutStr & "R:01 Alternate ID Stored" & vbCrLf
                        Else

                            ' the update didn't work
                            OutStr = OutStr & "R:02 Alternate not updated"
                        End If

                    End If  'Common.RowsAffected>0

                End If  'not(CustomerPK=0)
                'now that the AltID and Verifier have been set/updated, pull out the customer data if appropriate
                'send the normal load of GetCustomerInfo data for the newly created customer record
                CustomerData = Fetch_CustomerInfo_Data(CustomerPK, 0, DelimChar, LocalServer.LocalServerID, LocalServer.LocationID, 1, ExtCardID)

            End If  'NumRecs>0

        End If 'ProcessOK

        Common.Write_Log(LogFile, "Returned the following data for this AltID lookup:")

        OutStr = OutStr & CustomerData
        Send_Response_Header(ResponseStringHeader, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        'send the total execution time to the local server so it can be logged there
        TotalTime = DateAndTime.Timer - StartTime
        OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
        OutStr = Len(OutStr) & vbCrLf & OutStr
        Send(OutStr)
        Common.Write_Log(LogFile, OutStr)

AllDone:
        Exit Sub

QueryError:
        DupKey = False
        If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
            ex = Err.GetException
            If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
                DupKey = True
            End If
        End If
        If DupKey Then
            Resume Next
        Else
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IP=" & LocalServer.IPAddress & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocalServer.LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()

    End Sub

  ' -----------------------------------------------------------------------------------------------

  Sub Handle_CustomerLocking_Request(ByVal LocalServer As Copient.ConnectorInc.UELSRec, ByVal Mode As String)

    Dim ex As System.Data.SqlClient.SqlException
    Dim DupKey As Boolean
    Dim OutStr As String
    Dim DelimChar As Integer
    Dim sLockOption As Integer
    Dim sCustomerPK As String
    Dim sLockingGroupId As String
    Dim sLockingMode As String
    Dim iLockStatus As Integer
    Dim sTerminal As String
    Dim sTrxNum As String
    Dim LockedCustomerPK As Long
    Dim LockID As Long
    Dim NewLockExpireDate As DateTime        'When an lock is removed, there may be a delay period - this is the date/time when that delay time is over
    Dim LockExpireMinutes As Long
    Dim iRowCount As Integer

    LockID = 0
    Long.TryParse(Common.Fetch_UE_SystemOption(132), LockExpireMinutes)
    If LockExpireMinutes < 0 Then LockExpireMinutes = 0
    NewLockExpireDate = DateAdd(DateInterval.Minute, LockExpireMinutes, DateTime.Now)

    DelimChar = 30

    On Error GoTo QueryError

    OutStr = ""
    OutStr = OutStr & "LocationID=" & LocalServer.LocationID & vbCrLf

    Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
    Common.Open_LRTsp()
    Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServer.LocalServerID
    Common.LRTsp.ExecuteNonQuery()
    Common.Close_LRTsp()

    sCustomerPK = Trim(Request.QueryString("customerpk"))
    sLockingGroupId = Trim(Request.QueryString("lockinggroupid"))
    sLockingMode = Trim(Request.QueryString("lockmode")).ToUpper
    If (sLockingMode = "") Then
      sLockingMode = "LOCK"
    End If
    sTerminal = Trim(Request.QueryString("terminalid"))
    If sTerminal.Length = 0 Then
      sTerminal = "0"
    End If
    sTrxNum = Trim(Request.QueryString("terminalid"))
    If sTrxNum.Length = 0 Then
      sTrxNum = "0"
    End If

    sLockOption = Common.Fetch_UE_SystemOption(86)

    If sLockOption = "0" Then
      iLockStatus = 3
      Common.Write_Log(LogFile, "Customer locking is turned off (LockMode: " & sLockingMode & ", CustomerPK: " & sCustomerPK & ", TerminalLockingGroupID: " & sLockingGroupId & ")")

      OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
      TotalTime = DateAndTime.Timer - StartTime
      OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
      OutStr = Len(OutStr) & vbCrLf & OutStr

      Send_Response_Header(ResponseStringHeader, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
      Send(OutStr)
      Common.Write_Log(LogFile, OutStr)
    Else
      'locking is enabled
      If (sLockingMode = "LOCK") Then
        If sCustomerPK = "" Then
          Send_Response_Header("Missing CustomerPK (CustomerPK)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Common.Write_Log(LogFile, "Missing CustomerPK (CustomerPK)!")
        ElseIf sLockingGroupId = "" AndAlso sLockOption = "2" Then
          Send_Response_Header("Missing TerminalLockingGroupID (lockingGroupID)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Common.Write_Log(LogFile, "Missing TerminalLockingGroupID (lockingGroupID)")
        Else
          If sLockOption = "2" Then
            Common.Write_Log(LogFile, "Received LockMode: Lock, CustomerPK: " & sCustomerPK & ", TerminalLockingGroupID: " & sLockingGroupId)
          Else
            Common.Write_Log(LogFile, "Received LockMode: Lock, CustomerPK: " & sCustomerPK)
            sLockingGroupId = "0"
          End If

          iLockStatus = HandleCustomerLocking(sLockOption, Long.Parse(sCustomerPK), Long.Parse(sLockingGroupId), LocalServer.LocationID, LockedCustomerPK, LockID)

          OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
          OutStr = OutStr & "LockID=" & LockID & vbCrLf
          TotalTime = DateAndTime.Timer - StartTime
          OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
          OutStr = Len(OutStr) & vbCrLf & OutStr

          Send_Response_Header(ResponseStringHeader, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Send(OutStr)
          Common.Write_Log(LogFile, OutStr)
        End If
      ElseIf (sLockingMode = "UNLOCK") Then
        If sCustomerPK = "" Then
          Send_Response_Header("Missing CustomerPK (CustomerPK)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Common.Write_Log(LogFile, "Missing CustomerPK (CustomerPK)!" & " IP=" & LocalServer.IPAddress & " MacAddress=" & LocalServer.MacAddress)
        ElseIf sLockingGroupId = "" AndAlso sLockOption = "2" Then
          'locking is based on CustomerPK and terminal lock group
          Send_Response_Header("Missing TerminalLockingGroupID (lockingGroupID)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Common.Write_Log(LogFile, "Missing TerminalLockingGroupID (lockingGroupID)" & " Serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IP=" & LocalServer.IPAddress & " server=" & Environment.MachineName)
        Else
          If sLockOption = "2" Then  'locking is based on CustomerPK and terminal lock group
            Common.Write_Log(LogFile, "Received LockMode: Unlock, CustomerPK: " & sCustomerPK & ", TerminalLockingGroupID: " & sLockingGroupId)
          Else  'locking is based on CustomerPK only
            Common.Write_Log(LogFile, "Received LockMode: Unlock, CustomerPK: " & sCustomerPK)
            sLockingGroupId = "0"
          End If
          Common.QueryStr = "dbo.pa_UE_CustomerLockSetUnlockDelay"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = sCustomerPK
          Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = sLockingGroupId
          Common.LXSsp.Parameters.Add("@UE_LockExpireDate", SqlDbType.DateTime).Value = NewLockExpireDate  'Apply the unlock delay time to the UE_LockExpireDate field of the record
          Common.LXSsp.Parameters.Add("@Count", SqlDbType.Int).Direction = ParameterDirection.Output
          Common.LXSsp.ExecuteNonQuery()
          iRowCount = Common.LXSsp.Parameters("@Count").Value
          Common.Close_LXSsp()
          If iRowCount > 0 Then
            iLockStatus = 0
            Common.Write_Log(LogFile, "Deleted lock for CustomerPK: " & sCustomerPK & " and TerminalLockingGroupID: " & sLockingGroupId)
          Else
            iLockStatus = 1
            Common.Write_Log(LogFile, "Lock for CustomerPK: " & sCustomerPK & " and TerminalLockingGroupID: " & sLockingGroupId & " updated with an lock expire time of " & NewLockExpireDate)
          End If
          OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
          TotalTime = DateAndTime.Timer - StartTime
          OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
          OutStr = Len(OutStr) & vbCrLf & OutStr

          Send_Response_Header(ResponseStringHeader, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Send(OutStr)
          Common.Write_Log(LogFile, OutStr)
        End If

      ElseIf (sLockingMode = "FORCELOCK") Then

        If sCustomerPK = "" Then
          Send_Response_Header("Missing CustomerPK (CustomerPK)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Common.Write_Log(LogFile, "Missing CustomerPK (CustomerPK)!" & " serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IP=" & LocalServer.IPAddress & " server=" & Environment.MachineName)
        ElseIf sLockingGroupId = "" AndAlso sLockOption = "2" Then
          Send_Response_Header("Missing TerminalLockingGroupID (lockingGroupID)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Common.Write_Log(LogFile, "Missing TerminalLockingGroupID (lockingGroupID)" & " Serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IP=" & LocalServer.IPAddress & " server=" & Environment.MachineName)
        Else
          If sLockOption = "2" Then
            Common.Write_Log(LogFile, "Received LockMode: ForceLock, CustomerPK: " & sCustomerPK & ", TerminalLockingGroupID: " & sLockingGroupId)
          Else
            Common.Write_Log(LogFile, "Received LockMode: ForceLock, CustomerPK: " & sCustomerPK)
            sLockingGroupId = "0"
          End If
          Common.QueryStr = "dbo.pa_CPE_CustomerLockDelete"
          Common.Open_LXSsp()
          Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = sCustomerPK
          Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = sLockingGroupId
          Common.LXSsp.Parameters.Add("@Count", SqlDbType.Int).Direction = ParameterDirection.Output
          Common.LXSsp.ExecuteNonQuery()
          iRowCount = Common.LXSsp.Parameters("@Count").Value
          Common.Close_LXSsp()
          If iRowCount > 0 Then
            iLockStatus = 0
            Common.Write_Log(LogFile, "Deleted lock for CustomerPK: " & sCustomerPK & " and TerminalLockingGroupID: " & sLockingGroupId)
          Else
            iLockStatus = 0
            Common.Write_Log(LogFile, "Lock for CustomerPK: " & sCustomerPK & " and TerminalLockingGroupID: " & sLockingGroupId & " does not exist.")
          End If

          iLockStatus = HandleCustomerLocking(sLockOption, Long.Parse(sCustomerPK), Long.Parse(sLockingGroupId), LocalServer.LocationID, LockedCustomerPK, LockID)

          OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
          OutStr = OutStr & "LockID=" & LockID & vbCrLf
          TotalTime = DateAndTime.Timer - StartTime
          OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
          OutStr = Len(OutStr) & vbCrLf & OutStr

          Send_Response_Header(ResponseStringHeader, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
          Send(OutStr)
          Common.Write_Log(LogFile, OutStr)
        End If

      End If
    End If  'locking is enabled

AllDone:
    Exit Sub

QueryError:
    DupKey = False
    If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
      ex = Err.GetException
      If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
        DupKey = True
      End If
    End If
    If DupKey Then
      Resume Next
    Else
      Response.Write(Common.Error_Processor(, "Serial=" & LocalServer.LocalServerID & " MacAddress=" & LocalServer.MacAddress & " IPAddress=" & LocalServer.IPAddress & " Process Info: Server Name:" & Environment.MachineName & " Invoking LocationID=" & LocalServer.LocationID, , Common.InstallationName))
      Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
    End If
    Response.End()

  End Sub
</script>
<%
  '-----------------------------------------------------------------------------------------------
  'Main Code - Execution starts here

  Dim LocalServerID As Integer
  Dim LocalServer As Copient.ConnectorInc.UELSRec
  Dim Mode As String
  Dim MacAddress As String
  Dim LocalServerIP As String

  Common.AppName = "GetCustomerInfo.aspx"
  Response.Expires = 0
  On Error GoTo ErrorTrap
  StartTime = DateAndTime.Timer

  MacAddress = Trim(Request.QueryString("mac"))

  If MacAddress = "" Or MacAddress = "0" Then
    'MacAddress = Trim(Request.UserHostAddress)
    MacAddress = "Not Specified"
  End If
  LocalServerID = CType(Common.Extract_Val(Request.QueryString("serial")), Integer)
  LocalServerIP = CType(Common.Extract_Val(Request.QueryString("IP")), String)
  If LocalServerIP = "" Or LocalServerIP = "0" Then
    LocalServerIP = Trim(Request.UserHostAddress) & " :IP from requesting browser. "
  End If

  LSVersionNum = Connector.Convert_LSVersion(Common, Trim(Request.QueryString("lsversion")), Trim(Request.QueryString("lsbuild")))
  If(Convert.ToDecimal(Trim(Request.QueryString("lsversion"))) < 6) Then
    ResponseStringHeader = "PhoneHome"
  End If

  Mode = UCase(Request.QueryString("mode"))
  LogFile = "UE-GetCustomerInfoLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

  Common.Open_LogixRT()
  Common.Open_LogixXS()
  Common.Load_System_Info()
  Connector.Load_System_Info(Common)

  LocalServer = Connector.Get_UELS_Info(Common, LocalServerID, Copient.ConnectorInc.LastHeardType.PhoneHome, LocalServerIP, MacAddress)
  Common.Write_Log(LogFile, "----------------------------------------------------------------")
  Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "  Primary/Secondary: " & If(LocalServer.LocationAssociationType = Copient.ConnectorInc.AssociationType.SecondaryServer, "Secondary", "Primary") & "  LocationID: " & LocalServer.LocationID & "  Mode: " & Mode & "  LSVersion=" & LSVersionNum.LSMajorVersion & "." & LSVersionNum.LSMinorVersion & "b" & LSVersionNum.BuildMajorVersion & "." & LSVersionNum.BuildMinorVersion & "  Process running on server:" & Environment.MachineName & " with MacAddress=" & MacAddress & " IP=" & LocalServerIP)

  If LocalServer.LocationID = 0 Then
    Send_Response_Header("Invalid Serial - does not resolve to a location", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
    Common.Write_Log(LogFile, "Invalid Serial - does not resolve to a location" & " Serial=" & LocalServerID & " MacAddress=" & MacAddress)
  ElseIf Not (Connector.IsValidLocationForConnectorEngine(Common, LocalServer.LocationID, 9)) Then
    'the location calling TransDownload is not associated with the UE promoengine
    Common.Write_Log(LogFile, "This location is associated with a promotion engine other than UE.  Can not proceed.", True)
    Send_Response_Header("This location is associated with a promotion engine other than UE.  Can not proceed.", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
  Else
    Select Case Mode
      Case "LOOKUP"
        Handle_Lookup_Request(LocalServer, Mode)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "Lookup Run Time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case "ALTID"
        Handle_AltID_Request(LocalServer, Mode)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "AltID Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case "ALTIDENROLL"
        Handle_AltIDEnrollment_Request(LocalServer, Mode)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "AltIDEnrollment Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case "CUSTOMERLOCKING"
        Handle_CustomerLocking_Request(LocalServer, Mode)
        TotalTime = DateAndTime.Timer - StartTime
        Common.Write_Log(LogFile, "CustomerLocking Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
      Case Else
        Send_Response_Header("Invalid Request - bad mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Received invalid request!" & " Serial=" & LocalServerID & " MacAddress=" & MacAddress)
        Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
        Common.Write_Log(LogFile, Get_Raw_Form(Request.InputStream))
    End Select
  End If 'LocalServer.LocationID="0"

  If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
  If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()

%>
<%
  Response.End()
ErrorTrap:
  Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocalServer.LocationID, , Common.InstallationName))
  Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
  Common = Nothing
%>