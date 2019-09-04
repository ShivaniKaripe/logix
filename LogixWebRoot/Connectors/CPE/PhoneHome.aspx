<%@ Page CodeFile="..\ConnectorCB.vb" Debug="true" Inherits="ConnectorCB" Language="vb" %>

<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="Copient.commonShared" %>
<%' version:5.99.1.68274.Official Build (BERYLLIUM) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient" %>
<%
    ' *****************************************************************************
    ' * FILENAME: PhoneHome.aspx 
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
    Public LocalServerID As String
    Public MacAddress As String
    Public LocalServerIP As String
    Dim StartTime As Object
    Dim TotalTime As Object
    Public LSVerMajor As Integer
    Public LSVerMinor As Integer
    Public LSBuildMajor As Integer
    Public LSBuildMinor As Integer
    Public SupressLogging As Boolean = True
    Public LanguageID As Integer = 1
    Public AdminUserID As Integer = 1
    Private cardValidationResp As CardValidationResponse
    
    '------------------------------------------------------------------------------------------------
  

    Sub Fetch_Serial()
    
        Dim dst As DataTable
        Dim ClientLocationCode As String
        Dim LocationID As Long
        Dim ImageFetchURL As String
        Dim IncentiveFetchURL As String
        Dim PhoneHomeIP As String
        Dim OfflineFTPUser As String
        Dim OfflineFTPPass As String
        Dim OfflineFTPPath As String
        Dim OfflineFTPIP As String
        Dim Failover As Integer
        Dim LSFailover As Integer
        Dim MustIPL As Integer
        Dim OldLSID As Long
        Dim LastIP As String
        'Dim LocalServerIP As String
        Dim EngineID As Integer
        Dim Mode As String
    
        LocalServerID = 0
        LocationID = 0
    
        Failover = Common.Extract_Val(Request.QueryString("failover"))
        If Not (Failover = 1) Then Failover = 0
        Mode = UCase(Trim(Request.QueryString("mode")))
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If
    
        LastIP = "NotSpecified"
        ClientLocationCode = Common.Parse_Quotes(Trim(Request.QueryString("clc")))

        MacAddress = Trim(Request.QueryString("mac"))
        If MacAddress = "" Then
            Send_Response_Header("Missing MAC Address", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "Missing MAC Address!  ClientLocationCode:" & ClientLocationCode & " from Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
            Exit Sub
        End If

        LocationID = 0
        If Not (ClientLocationCode = "") Then
            Common.QueryStr = "select LocationID, EngineID from Locations with (NoLock) where ExtLocationCode='" & ClientLocationCode & "' and Deleted=0;"
            dst = Common.LRT_Select
            If dst.Rows.Count > 0 Then
                EngineID = Common.NZ(dst.Rows(0).Item("EngineID"), 0)
                If Not (EngineID = 2) Then
                    Send_Response_Header("This location code is not associated with the CPE engine", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, "Invalid ClientLocationCode - clc=" & ClientLocationCode & " - not associated with the CPE engine.   EngineID=" & EngineID & " Serial=" & LocalServerID & " IPAddress=" & Trim(Request.UserHostAddress) & " server=" & Environment.MachineName)
                    Exit Sub
                Else
                    LocationID = Common.NZ(dst.Rows(0).Item("LocationID"), 0)
                End If
            End If
            If LocationID = 0 Then
                Send_Response_Header("Invalid Location Code", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                Common.Write_Log(LogFile, "Invalid ClientLocationCode - clc:" & ClientLocationCode & " Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
                Exit Sub
            End If
        End If
        If ((Not Mode.Equals("SRL")) Or (Not SupressLogging)) Then Common.Write_Log(LogFile, "Found LocationID " & LocationID & " for CLC:'" & ClientLocationCode & "'" & " Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
        'get the image and incentive fetch URLs
        IncentiveFetchURL = Common.Fetch_CPE_SystemOption(43)
        ImageFetchURL = Common.Fetch_CPE_SystemOption(39)
        PhoneHomeIP = Trim(Common.Fetch_CPE_SystemOption(98))
        OfflineFTPUser = Common.Fetch_CPE_SystemOption(75)
        OfflineFTPPass = Common.Fetch_CPE_SystemOption(76)
        OfflineFTPPath = Common.Fetch_CPE_SystemOption(77)
        OfflineFTPIP = Common.Fetch_CPE_SystemOption(78)
    
        LocalServerID = 0
    
        'see if we have a LocalServers record for this MacAddress
        Common.QueryStr = "select LocalServerID, FailoverServer, LastIP from LocalServers with (NoLock) where MacAddress='" & MacAddress & "';"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            LocalServerID = Common.NZ(dst.Rows(0).Item("LocalServerID"), 0)
            LSFailover = Common.NZ(dst.Rows(0).Item("FailoverServer"), 0)
            LSFailover = Math.Abs(LSFailover)
            LastIP = Trim(Common.NZ(dst.Rows(0).Item("LastIP"), "NotSpecified"))
        End If
        If (LocalServerID = 0) And (LocationID <> 0) Then 'check and see if there is a record for this LocalServerID that has an empty MacAddress and use that record if it exists
            'see if we have a LocalServers record for this location
            Common.QueryStr = "select LocalServerID, FailoverServer from LocalServers with (NoLock) where LocationID=" & LocationID & " and MacAddress is NULL;"
            If ((Not Mode.Equals("SRL")) Or (Not SupressLogging)) Then Common.Write_Log(LogFile, Common.QueryStr & " Serial= " & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
            dst = Common.LRT_Select
            If dst.Rows.Count > 0 Then
                LocalServerID = Common.NZ(dst.Rows(0).Item("LocalServerID"), 0)
                LSFailover = Math.Abs(Common.NZ(dst.Rows(0).Item("FailoverServer"), 0))
            End If
            If ((Not Mode.Equals("SRL")) Or (Not SupressLogging)) Then Common.Write_Log(LogFile, "Serial: " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
            If Not (LocalServerID = 0) Then 'update the localservers record with the MacAddress
                Common.QueryStr = "update LocalServers with (RowLock) set MacAddress='" & MacAddress & "' where LocalServerID=" & LocalServerID & ";"
                Common.LRT_Execute()
            End If
        End If
        If LocalServerID = 0 Then 'create a localservers record for this MacAddress
            Common.QueryStr = "pa_CPE_PH_CreateLS"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@ImageFetchURL", SqlDbType.VarChar, 255).Value = ImageFetchURL
            Common.LRTsp.Parameters.Add("@IncentiveFetchURL", SqlDbType.VarChar, 255).Value = IncentiveFetchURL
            Common.LRTsp.Parameters.Add("@PhoneHomeIPOverride", SqlDbType.VarChar, 255).Value = PhoneHomeIP
            Common.LRTsp.Parameters.Add("@OfflineFTPUser", SqlDbType.VarChar, 255).Value = OfflineFTPUser
            Common.LRTsp.Parameters.Add("@OfflineFTPPass", SqlDbType.VarChar, 255).Value = OfflineFTPPass
            Common.LRTsp.Parameters.Add("@OfflineFTPPath", SqlDbType.VarChar, 255).Value = OfflineFTPPath
            Common.LRTsp.Parameters.Add("@OfflineFTPIP", SqlDbType.VarChar, 255).Value = OfflineFTPIP
            Common.LRTsp.Parameters.Add("@MacAddress", SqlDbType.VarChar, 50).Value = MacAddress
            Common.LRTsp.Parameters.Add("@Failover", SqlDbType.Bit).Value = Failover
            Common.LRTsp.Parameters.Add("@LastIP", SqlDbType.VarChar, 15).Value = Left(LocalServerIP, 15)
            Common.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
            Common.LRTsp.ExecuteNonQuery()
            LocalServerID = Common.LRTsp.Parameters("@PKID").Value
            Common.Close_LRTsp()
            If ((Not Mode.Equals("SRL")) Or (Not SupressLogging)) Then Common.Write_Log(LogFile, "Created LocalServers record (" & LocalServerID & ") for MacAddress '" & MacAddress & "'")
        Else
            If ((Not Mode.Equals("SRL")) Or (Not SupressLogging)) Then Common.Write_Log(LogFile, "Found Serial= " & LocalServerID & " MacAddress:" & MacAddress & " IP:" & LocalServerIP & " server:" & Environment.MachineName)
            If Not ((Failover = LSFailover) And (LastIP = LocalServerIP)) Then
                Common.QueryStr = "Update LocalServers with (RowLock) set FailoverServer=" & Failover & ", LastIP='" & LocalServerIP & "' where LocalServerID=" & LocalServerID & ";"
                Common.LRT_Execute()
            End If
        End If
    
        If Not (LocationID = "0") Then
            MustIPL = 0
            OldLSID = 0
            Common.QueryStr = "select LocalServerID, isnull(MustIPL, 0) as MustIPL from LocalServers with (NoLock) where LocationID=" & LocationID & ";"
            dst = Common.LRT_Select
            If dst.Rows.Count > 0 Then
                OldLSID = Common.NZ(dst.Rows(0).Item("LocalServerID"), 0)
                If OldLSID <> LocalServerID Then
                    MustIPL = 1
                End If
                If (MustIPL = 0) Then
                    If ((Not Mode.Equals("SRL")) Or (Not SupressLogging)) Then Common.Write_Log(LogFile, "MustIPL from DB=" & dst.Rows(0).Item("MustIPL") & " LocalServerID=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP)
                    If dst.Rows(0).Item("MustIPL") = True Then MustIPL = 1
                End If
            Else
                MustIPL = 1
            End If
            If ((Not Mode.Equals("SRL")) Or (Not SupressLogging) Or (MustIPL = 1)) Then Common.Write_Log(LogFile, "MustIPL=" & MustIPL & " LocalServerID=" & LocalServerID & " MacAddress=" & MacAddress)
            If MustIPL Then
                Common.Write_Log(LogFile, "Changing the server handling LocationID " & LocationID & " (" & ClientLocationCode & ") to LocalServerID=" & LocalServerID & ";")
                Common.QueryStr = "update CPE_IncentiveDLBuffer with (RowLock) set WaitingACK=2 where WaitingACK>=0 and WaitingACK<2 and LocalServerID=" & OldLSID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "Update LocalServers with (RowLock) set LastLocationID=LocationID where LocationID=" & LocationID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "Update LocalServers with (RowLock) set LocationID=0, MustIPL=1 where LocationID=" & LocationID & ";"
                Common.LRT_Execute()
                Common.QueryStr = "Update LocalServers with (RowLock) set LocationID=" & LocationID & ", LastIP='" & Left(LocalServerIP, 15) & "', MustIPL=1 where MacAddress='" & MacAddress & "';"
                Common.LRT_Execute()
            Else
                If ((Not Mode.Equals("SRL")) Or (Not SupressLogging)) Then Common.Write_Log(LogFile, "Server re-requested serial number - Location and MacAddress have not changed")
            End If
        End If
    
        If ((Not Mode.Equals("SRL")) Or (Not SupressLogging) Or (MustIPL = 1)) Then Common.Write_Log(LogFile, "Fetched serial number '" & LocalServerID & "' for Local Server at ClientLocationCode '" & ClientLocationCode & "'" & " Serial= " & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
        If MustIPL Then
            Send_Response_Header("MustIPL", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Else
            Send_Response_Header("Serial", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        End If
        Send(LocalServerID)
    
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
                        If UCase(TableName) = "USERS" And UCase(SQLCol.ColumnName) = "CLIENTUSERID1" Then
                            LineOut = LineOut & Common.NZ(MyCryptlib.SQL_StringDecrypt(row(Index).ToString()), 0)
                        
                        ElseIf UCase(TableName) = "CARDIDS" And UCase(SQLCol.ColumnName) = "EXTCARDID" Then
                            LineOut = LineOut & Common.NZ(MyCryptlib.SQL_StringDecrypt(row(Index).ToString()), 0)
                            
                        ElseIf UCase(TableName) = "AltIDCacheCardIDs" And UCase(SQLCol.ColumnName) = "EXTCARDID" Then
                            LineOut = LineOut & Common.NZ(MyCryptlib.SQL_StringDecrypt(row(Index).ToString()), 0)
                            
                        ElseIf UCase(TableName) = "AltIDCacheUsers" And UCase(SQLCol.ColumnName) = "CLIENTUSERID1" Then
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
  
    Function HandleCustomerLocking(ByVal sLockOption As String, ByVal CustomerPK As Long, ByVal LockingGroupId As Long, ByVal LocationID As Long, ByRef LockedCustomerPK As Long) As Integer
    
        Dim dst As DataTable
        Dim sLockedDate As String
        Dim dtLocked As Date
        Dim iLockStatus As Integer
        Dim lMinutes As Long
        Dim CustLockDelMin As Integer
        Dim elapsed_time As TimeSpan
        Dim lCuPk As Long
        Dim lHhPk As Long = 0
        Dim sHouseholdText As String
        Dim LockExpireMinutes As Long
        Dim NewLockExpireDate As DateTime
        Long.TryParse(Common.Fetch_SystemOption(68), LockExpireMinutes)
        If LockExpireMinutes < 0 Then LockExpireMinutes = 0
        NewLockExpireDate = DateAdd(DateInterval.Minute, LockExpireMinutes, DateTime.Now)


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
            Common.QueryStr = "dbo.pa_CPE_CustomerLockFetch"
            Common.Open_LXSsp()
            Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCuPk
            Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = LockingGroupId
            dst = Common.LXSsp_select
            Common.Close_LXSsp()
            If dst.Rows.Count > 0 Then
                sLockedDate = dst.Rows(0).Item("LockedDate")
                dtLocked = Date.Parse(sLockedDate)
                elapsed_time = DateTime.Now.Subtract(dtLocked)
                lMinutes = elapsed_time.TotalMinutes
                CustLockDelMin = Common.Extract_Val(Common.Fetch_SystemOption(68))
                If (CustLockDelMin = 0) OrElse (lMinutes < CustLockDelMin) Then
                    iLockStatus = 1
                    Common.Write_Log(LogFile, "Lock for CustomerPK: " & CustomerPK & sHouseholdText & " and TerminalLockingGroupID: " & LockingGroupId & " already exists. (Previously locked on: " & sLockedDate & ")")
                Else
                    Common.QueryStr = "dbo.pa_CPE_CustomerLockUpdate"
                    Common.Open_LXSsp()
                    Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCuPk
                    Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = LockingGroupId
                    Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                    Common.LXSsp.Parameters.Add("@TerminalNumber", SqlDbType.Int).Value = 0
                    Common.LXSsp.Parameters.Add("@TransactionNumber", SqlDbType.NVarChar, 128).Value = "0"
                    Common.LXSsp.Parameters.Add("@LockedBy", SqlDbType.BigInt).Value = CustomerPK
                    Common.LXSsp.Parameters.Add("@LockExpireDate", SqlDbType.DateTime).Value = NewLockExpireDate
                    Common.LXSsp.ExecuteNonQuery()
                    Common.Close_LXSsp()
                    iLockStatus = 0
                    Common.Write_Log(LogFile, "Updated expired lock for CustomerPK: " & CustomerPK & sHouseholdText & " and TerminalLockingGroupID: " & LockingGroupId & " (Previously locked on: " & sLockedDate & ")")
                End If
            Else
                Common.QueryStr = "dbo.pa_CPE_CustomerLockInsert"
                Common.Open_LXSsp()
                Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCuPk
                Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = LockingGroupId
                Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                Common.LXSsp.Parameters.Add("@TerminalNumber", SqlDbType.Int).Value = 0
                Common.LXSsp.Parameters.Add("@TransactionNumber", SqlDbType.NVarChar, 128).Value = "0"
                Common.LXSsp.Parameters.Add("@LockedBy", SqlDbType.BigInt).Value = CustomerPK
                Common.LXSsp.Parameters.Add("@LockExpireDate", SqlDbType.DateTime).Value = NewLockExpireDate
                Common.LXSsp.ExecuteNonQuery()
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
  
    Function Fetch_PhoneHome_Data(ByVal CustomerPK As Long, ByVal DelimChar As String, ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal NewUser As Integer, ByVal ExtCardID As String) As String
    
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
    
        'if the customer record was just created on central, then send the
        'special NewUser table so that the promotion engine executes offers 
        'that are targeted to new cardholders
        If NewUser = 1 Then
            dst = New DataTable
            dst.Columns.Add("Value", System.Type.GetType("System.Int32"))
            row = dst.NewRow
            row.Item("Value") = 1
            dst.Rows.Add(row)
            Common.Write_Log(LogFile, "sending table")
            TempStr = TempStr & Construct_Table("NewUser", 11, DelimChar, LocalServerID, LocationID, dst)
            dst = Nothing
            Common.QueryStr = "update Customers set NewCustomer=0 where CustomerPK=" & CustomerPK & ";"
            Common.LXS_Execute()
        End If
    
        'Users
        'send the data from the Users table
        Common.QueryStr = "select Customers.CustomerPK as UserID, InitialCardIDOriginal as ClientUserID1, isnull(Customers.HHPK, 0) as HHPrimaryID, isnull(Customers.FirstName, '') as FirstName, isnull(Customers.LastName, '') as LastName, " & _
                      "  case CustomerTypeID when 2 then 0 else CustomerTypeID end as HHRec, CustomerTypeID, Customers.Employee, Customers.CurrYearSTD, " & _
                      "  Customers.LastYearSTD, isnull(Customers.CustomerStatusID, 0) as CustomerStatusID, AltIDOptOut " & AltIDColumn & VerifierColumn & ", " & _
                      "  isnull(Customers.EmployeeID,'') as EmployeeID, isnull(CustomerExt.AirmileMemberID,'') as AirmileMemberID, isnull(Customers.Prefix,'') as Prefix, " & _
                      "  isnull(Customers.Suffix,'') as Suffix, isnull(Customers.RestrictedRedemption,0) as RestrictedRedemption" & _
                      " from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                      " where Customers.CustomerPK=" & CustomerPK & ";"
        dst = Common.LXS_Select
        TempStr = TempStr & Construct_Table("Users", 1, DelimChar, LocalServerID, LocationID, dst)
    
        'CardIDs
        'send the data from the CardIDs table
        Common.QueryStr = "select CardIDs.CardPK, CardIDs.CustomerPK as UserID, CardIDs.ExtCardIDOriginal as ExtCardID, CardIDs.CardStatusID, CardIDs.CardTypeID " & _
                          " from CardIDs with (NoLock) " & _
                          " where CardIDs.CustomerPK=" & CustomerPK & ";"
        dst = Common.LXS_Select
        TempStr = TempStr & Construct_Table("CardIDs", 1, DelimChar, LocalServerID, LocationID, dst)

        'CustomerAttributes
        'send the data from the CustomerAttributes table
        Common.QueryStr = "select CustomerPK, AttributeTypeID, AttributeValueID " & _
                          " from CustomerAttributes with (NoLock) " & _
                          " where CustomerAttributes.CustomerPK=" & CustomerPK & " and CustomerAttributes.Deleted!=1;"
        dst = Common.LXS_Select
        TempStr = TempStr & Construct_Table("CustomerAttributes", 1, DelimChar, LocalServerID, LocationID, dst)
    
        'GroupMembership
        'send the data from the GroupMembership table
        Common.QueryStr = "dbo.pa_CPE_IN_GMActive"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        TempStr = TempStr & Construct_Table("GroupMembership", 7, DelimChar, LocalServerID, LocationID, dst)
    
        'RewardAccumulation
        'send the data from the RewardAccumulation table
        Common.QueryStr = "dbo.pa_CPE_IN_RAActive"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        TempStr = TempStr & Construct_Table("RewardAccumulation", 7, DelimChar, LocalServerID, LocationID, dst)
    
        'RewardDistribution
        'send the data from the RewardDistribution table
        Common.QueryStr = "dbo.pa_CPE_IN_RDActive"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        TempStr = TempStr & Construct_Table("RewardDistribution", 7, DelimChar, LocalServerID, LocationID, dst)
    
        'UserResponses
        Common.QueryStr = "dbo.pa_CPE_IN_CRActive"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        TempStr = TempStr & Construct_Table("UserResponses", 1, DelimChar, LocalServerID, LocationID, dst)
    
        'Points
        Common.QueryStr = "dbo.pa_CPE_IN_PointsActive"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        TempStr = TempStr & Construct_Table("Points", 6, DelimChar, LocalServerID, LocationID, dst)
    
        If Common.Fetch_CPE_SystemOption(128) = "1" Then  'see if Preference Data Distribution is enabled
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
        Common.QueryStr = "dbo.pa_CPE_IN_StoredValue"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        TempStr = TempStr & Construct_Table("StoredValue", 6, DelimChar, LocalServerID, LocationID, dst)
    
        'StoredFranking
        Common.QueryStr = "dbo.pa_CPE_IN_StoredFranking"
        Common.Open_LXSsp()
        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NChar, 26).Value = CustomerPK
        dst = Common.LXSsp_select
        Common.Close_LXSsp()
        TempStr = TempStr & Construct_Table("StoredFranking", 6, DelimChar, LocalServerID, LocationID, dst)
    
        Fetch_PhoneHome_Data = TempStr
    
    End Function
  
    ' -----------------------------------------------------------------------------------------------
    ' 
    Class Customer
        
        Public m_cust_extCardID As String = ""
        Public m_cardtype As Integer = 0
        Public m_custpk As Long = 0
        Public m_cust_extCardPK As Long = 0
        Public m_custtype As Integer = 0
        Public m_hhpk As Long = 0
        Public m_hh_extCardID As String = ""
        
        Private m_Common As Copient.CommonInc ' needs to be open when it's passed to us
        
        Private MyCryptlib As New Copient.CryptLib
        
        Enum CardTypes
            CUSTOMER = 0
            HOUSEHOLD = 1
            CAM = 2
            ALTERNATEID = 3
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
        
        Private Function fetchCustomerInfo(ByVal extCardID As String, ByVal cardtype As Integer) As DataTable
            extCardID = m_Common.Pad_ExtCardID(extCardID, cardtype)
            ' get the customer pk and household pk based on the ext card id
            m_Common.QueryStr = _
                "SELECT Customers.CustomerPK AS UserID, ISNULL(Customers.HHPK, 0) AS HHPrimaryID, customers.customertypeid AS customertypeid, cardids.cardpk as cardpk " & _
                " FROM customers WITH (NoLock) JOIN cardids WITH (NoLock) ON customers.customerpk = cardids.customerpk " & _
                " WHERE cardids.extcardid = '" & MyCryptlib.SQL_StringEncrypt(extCardID, True) & "' AND cardids.cardtypeid = " & cardtype & " ;"
            Dim dst As DataTable = m_Common.LXS_Select

            If dst.Rows.Count = 0 Then ' If it wasn't found from the looking up the card in the cardids table, try looking it up from the initialcardid field
                m_Common.QueryStr = _
                    "SELECT Customers.CustomerPK AS UserID, ISNULL(Customers.HHPK, 0) AS HHPrimaryID, customers.customertypeid AS customertypeid, 0 as cardpk " & _
                    " FROM customers WITH (NoLock)" & _
                    " WHERE customers.InitialCardID = '" & MyCryptlib.SQL_StringEncrypt(extCardID, True) & "' AND customers.initialcardtypeid = " & cardtype & " ;"
                dst = m_Common.LXS_Select
            End If
            
            Return dst
        End Function ' fetchCustomerInfo
        
        
        
        Private Sub addMissingCardToCardIDsTable(ByVal custpk As Long, ByVal extCardID As String, ByVal cardtype As Integer)
            Dim cardValidationResp As CardValidationResponse
            If (m_Common.AllowToProcessCustomerCard(extCardID, cardtype, cardValidationResp) <> True) Then
                Throw New ArgumentException(m_Common.CardValidationResponseMessage(extCardID, cardtype, cardValidationResp))
            End If
            extCardID = m_Common.Pad_ExtCardID(extCardID, cardtype)
            m_Common.QueryStr = _
                "INSERT INTO cardids ( customerpk, extcardid, cardstatusid, cardtypeid, extcardidoriginal ) " & _
                "  VALUES ( " & custpk & ", '" & MyCryptlib.SQL_StringEncrypt(extCardID, True) & "', 1, " & cardtype & " , '" & MyCryptlib.SQL_StringEncrypt(extCardID) & "');"
            m_Common.LXS_Execute()
        End Sub ' addMissingCardToCardIDsTable
        
        
        Sub refresh()

            Dim dst As DataTable = fetchCustomerInfo(m_cust_extCardID, m_cardtype)
            If dst.Rows.Count > 0 Then
                m_custpk = dst.Rows(0).Item("UserID")
                m_hhpk = dst.Rows(0).Item("HHPrimaryID")
                m_custtype = dst.Rows(0).Item("customertypeid")
                m_cust_extCardPK = dst.Rows(0).Item("cardpk")
                
                If m_cust_extCardPK = 0 Then ' the customer exists with this card, but the card isn't in the cardIDs table
                    addMissingCardToCardIDsTable(m_custpk, m_cust_extCardID, m_cardtype)
                End If
            End If
            
            ' now get the associated household card
            If m_hhpk <> 0 AndAlso m_custtype <> CustomerTypes.HOUSEHOLD Then
                m_Common.QueryStr = _
                    "SELECT CardIDs.ExtCardIDOriginal as ExtCardID FROM CardIDs WITH (NoLock) " & _
                    " WHERE CardIDs.CustomerPK = " & m_hhpk & " AND cardids.cardtypeid = " & CardTypes.HOUSEHOLD & ";"
                dst = m_Common.LXS_Select
                If dst.Rows.Count > 0 Then
                    m_hh_extCardID = MyCryptlib.SQL_StringDecrypt(dst.Rows(0).Item("ExtCardID").ToString())
                End If
            End If
        End Sub ' refresh()


        Sub New(ByVal custid As String, ByVal cardtype As Integer, ByRef Common As Copient.CommonInc)
            m_cust_extCardID = custid
            m_cardtype = cardtype
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
        Static doCreate As Boolean = (Common.Fetch_CPE_SystemOption(113) = "1")
        Return doCreate
    End Function
    
    ' -----------------------------------------------------------------------------------------------    

    Sub SetCustomerHouseHold(ByVal CustomerPK As Integer, ByVal hhpk As Integer)
        Common.QueryStr = "UPDATE Customers SET hhpk = " & hhpk & " WHERE customerpk = " & CustomerPK & ";"
        Common.LXS_Execute()
    End Sub
  
    ' ----------------------------------------------------------------------------------------------- 
  
    Function IsNewCustomer(ByVal CustomerPK As Integer) As Boolean
        Dim dt As DataTable
        Dim IsNew As Boolean = False
    
        Common.QueryStr = "SELECT NewCustomer FROM Customers WITH (NoLock) WHERE CustomerPK=" & CustomerPK & ";"
        dt = Common.LXS_Select()
        If dt.Rows.Count > 0 Then
            IsNew = IIf(Common.NZ(dt.Rows(0).Item("NewCustomer"), False), True, False)
        End If
    
        Return IsNew
    End Function
    
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
        'dbo].[pt_NewCardIDs_Insert] @ExtCardID nvarchar(26), @CardTypeID int, @CustomerPK bigint, @CardStatusID int, @CardPK bigint OUTPUT, @Created bit = 0 OUTPUT
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
    Function DoCustomerCreation(ByVal theCust As Customer, ByVal ExtCardID As String, ByVal LocationID As String, ByVal BannerID As String) As Long

        Const INTERFACE_OPTION_MDM_USERNAME As Integer = 28
        Const INTERFACE_OPTION_MDM_PASSWORD As Integer = 29
        Const INTERFACE_OPTION_MDM_URL As Integer = 30
        
        Try
            ' get the username, password, and url from the interface options
            Dim mdm_user As String = Common.Fetch_InterfaceOption(INTERFACE_OPTION_MDM_USERNAME)
            Dim mdm_pw As String = Common.Fetch_InterfaceOption(INTERFACE_OPTION_MDM_PASSWORD)
            Dim mdm_url As String = Common.Fetch_InterfaceOption(INTERFACE_OPTION_MDM_URL)
            
            Dim mdm As New Copient.MasterDataLibrary(mdm_user, mdm_pw, mdm_url)
            Dim newHouseHoldExtCardID As String = ""
            Dim CustomerPK As Long = theCust.m_custpk
                       
            MacAddress = Trim(Request.QueryString("mac"))

            LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
            LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
            
            If mdm.createCustomer(ExtCardID, newHouseHoldExtCardID) Then
                'NewUser = 1
                If theCust.needsCustomerRecord() Then
                    CustomerPK = CreateCustomer(ExtCardID, Customer.CardTypes.CUSTOMER, LocationID, BannerID)
                End If
                
                If newHouseHoldExtCardID <> "" Then
                    If theCust.needsHouseHold() Then ' create the household and add the existing customer to the household
                        Dim newhhpk As Long = CreateCustomer(newHouseHoldExtCardID, Customer.CardTypes.HOUSEHOLD, LocationID, BannerID)
                        SetCustomerHouseHold(CustomerPK, newhhpk)
                    ElseIf theCust.needsHouseHoldCard() Then ' just add the new house hold card id to the household that already exists
                        AddCardToHouseHold(newHouseHoldExtCardID, theCust.m_hhpk)
                    End If
                End If ' newHouseHoldExtCardID <> ""

                Return CustomerPK
            End If ' mdm.createCustomer()
        
        Catch e As ApplicationException
            Common.Write_Log(LogFile, "Failed to create user in master data: " & e.Message & " serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
        End Try
        
        Send_Response_Header("Failed to create user in master data", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Common.Write_Log(LogFile, "Failed to create user in master data" & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
        Return 0
    End Function
    
    ' -----------------------------------------------------------------------------------------------       
    Sub Handle_User_Request(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal BannerID As Integer, ByVal Mode As String)
    
        Dim ex As System.Data.SqlClient.SqlException
        Dim DupKey As Boolean
        Dim ErrDesc As String
        Dim ExtCardID As String
        Dim CustomerPK As Long
        Dim OutStr As String
        Dim DelimChar As Integer
        Dim NumRecs As Long
        Dim LastUserLocationID As Long
        Dim UniqueError As Boolean
        Dim dst As DataTable
        Dim UpdateTime As String
        Dim tempstr As String
        Dim CardTypeID As Integer
        Dim TempCType As String
        Dim NewUser As Integer = 0
        Dim LockingGroupId As Long
        Dim sLockOption As String
        Dim iLockStatus As Integer
        Dim CAMErrorMsg As String
        Dim LockedCustomerPK As Long
        Dim ShouldCreateCard As Boolean = False
        Dim CreateCust As String
    
        DelimChar = 30
    
        Try
            MacAddress = Trim(Request.QueryString("mac"))
    
            LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
            LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
            If LocalServerIP = "" Or LocalServerIP = "0" Then
                Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
                LocalServerIP = Trim(Request.UserHostAddress)
            End If
    
            OutStr = ""
            OutStr = OutStr & "LocationID=" & LocationID & vbCrLf
    
            Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
            Common.Open_LRTsp()
            Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
            Common.LRTsp.ExecuteNonQuery()
            Common.Close_LRTsp()
    
            If Common.Fetch_CPE_SystemOption(91) = "1" Then
                Common.Write_Log(LogFile, "Operate at enterprise is enabled (CPE_SystemOption 91), so PhoneHome customer lookups are not allowed." & " serial=" & LocalServerID & "Mac Address=" & Trim(Request.UserHostAddress) & " server=" & Environment.MachineName)
                Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                Send("PhoneHome customer lookups not allowed while operate at enterprise is enabled (CPE_SystemOption 91)")
                Exit Sub
            End If
    
            ExtCardID = Trim(Request.QueryString("id"))
   
            'look for the hhrec parameter, and then override it with ctype (CustomerTypeID) if it exists
            CardTypeID = Common.Extract_Val(Request.QueryString("hhrec"))
            TempCType = Trim(Request.QueryString("ctype"))
            If Not (TempCType = "") Then
                CardTypeID = Common.Extract_Val(TempCType)
            End If
    
            Dim CreateCardType As Boolean = False
            If (CardTypeID = 0 OrElse CardTypeID = 1) Then CreateCardType = True
            If (Common.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CAM) AndAlso CardTypeID = 2) Then CreateCardType = True
        
            CreateCust = Request.QueryString("CreateCust")
            If CreateCust = "" Then
                ShouldCreateCard = True
            Else
                ShouldCreateCard = (CreateCust = 1)
            End If
    
            If ExtCardID = "" Then
                Send_Response_Header("Missing Card Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                Common.Write_Log(LogFile, "Missing Card Number!" & " serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
            Else
                'look up the Card Number we got
                Common.Write_Log(LogFile, "Received Card Number (id): " & ExtCardID & "   CardTypeID (ctype)=" & CardTypeID & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                If Not Common.AllowToProcessCustomerCard(ExtCardID, CardTypeID, cardValidationResp) Then
                    Send_Response_Header(Common.CardValidationResponseMessage(ExtCardID, CardTypeID, cardValidationResp), Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, Common.CardValidationResponseMessage(ExtCardID, CardTypeID, cardValidationResp))
                    Return
                End If
                If CardTypeID = 2 Then 'make sure this is a valid CAM card
                    CAMErrorMsg = ""
                    If Not (CAM.VerifyCardNumber(ExtCardID, CAMErrorMsg)) Then
                        Send_Response_Header("Invalid CAM Card Number - " & CAMErrorMsg, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                        Common.Write_Log(LogFile, "Invalid CAM Card Number - " & CAMErrorMsg)
                        Exit Sub
                    End If
                End If
      
                Dim theCust As Customer
                Try
                    theCust = New Customer(ExtCardID, CardTypeID, Common)
                Catch argExcp As ArgumentException
                    Send_Response_Header(argExcp.Message, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, argExcp.Message)
                    Return
                End Try
                CustomerPK = theCust.m_custpk
            
                If createCustomersInMasterDataLibrary() Then
                    If theCust.isNotComplete() Then ' the customer is missing some vital piece
                        If ShouldCreateCard AndAlso CreateCardType Then
                            CustomerPK = DoCustomerCreation(theCust, ExtCardID, LocationID, BannerID)
                            If CustomerPK <> 0 Then NewUser = 1
                        Else
                            Send_Response_Header("User Not Found Or Not Complete", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                            If (ShouldCreateCard = False) Then
                                Common.Write_Log(LogFile, "User not found or not complete - create request not sent from store" & " serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                            ElseIf (ShouldCreateCard = True AndAlso CreateCardType = False) Then
                                Common.Write_Log(LogFile, "User not found or not complete - cannot create a card for cardtype " & CardTypeID & " serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                            End If ' should create card
                        End If
                    Else
                        Common.Write_Log(LogFile, "Found existing CustomerPK: " & CustomerPK)
                        If IsNewCustomer(CustomerPK) Then
                            NewUser = 1
                        End If
                    End If
                Else
                    If theCust.needsCustomerRecord() Then
                        If ShouldCreateCard AndAlso CreateCardType Then
                            CustomerPK = CreateCustomer(ExtCardID, CardTypeID, LocationID, BannerID)
                            Common.Write_Log(LogFile, "User not found - created UserID: " & CustomerPK)
                            NewUser = 1
                        Else 'If CustomerPK = 0 AndAlso Not ShouldCreateCard Then
                            Send_Response_Header("User Not Found", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                            If (ShouldCreateCard = False) Then
                                Common.Write_Log(LogFile, "User not found - create request not sent from store")
                            ElseIf (CreateCardType = False) Then
                                Common.Write_Log(LogFile, "User not found or not complete - cannot create a card for cardtype " & CardTypeID & " serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                            End If
                        End If
                    Else
                        Common.Write_Log(LogFile, "Found existing CustomerPK: " & CustomerPK)
                        If IsNewCustomer(CustomerPK) Then
                            NewUser = 1
                        End If
                    End If
                End If
      
                'make sure this customer is associated with this location
                Common.QueryStr = "dbo.pa_CPE_IN_UpdateCustomerLocation"
                Common.Open_LXSsp()
                Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                Common.LXSsp.ExecuteNonQuery()
                Common.Close_LXSsp()
      
                OutStr = OutStr & Fetch_PhoneHome_Data(CustomerPK, DelimChar, LocalServerID, LocationID, NewUser, ExtCardID)

                If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 6) Or (LSVerMajor = 5 And LSVerMinor = 6 And LSBuildMajor >= 2) Then
                    'Customer Locking
                    sLockOption = Common.Fetch_CPE_SystemOption(86)
                    If sLockOption = "1" Or sLockOption = "2" Then
                        If sLockOption = "1" Then
                            LockingGroupId = 0
                        Else
                            LockingGroupId = Common.NZ((Request.QueryString("lockinggroupid")), 0)
                        End If
                        iLockStatus = HandleCustomerLocking(sLockOption, CustomerPK, LockingGroupId, LocationID, LockedCustomerPK)
                        Common.QueryStr = "dbo.pa_CPE_CustomerLockStatus"
                        Common.Open_LXSsp()
                        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = LockedCustomerPK
                        Common.LXSsp.Parameters.Add("@LockingGroupID", SqlDbType.BigInt).Value = LockingGroupId
                        Common.LXSsp.Parameters.Add("@LockStatus", SqlDbType.Int).Value = iLockStatus
                        dst = Common.LXSsp_select
                        Common.Close_LXSsp()
                        OutStr = OutStr & Construct_Table("UsersLockStatus", 12, DelimChar, LocalServerID, LocationID, dst)
                    End If
                End If
  
                Common.Write_Log(LogFile, "Returned the following data for this customer:")
      
                Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)

                If (LSVerMajor > 5) Or (LSVerMajor = 5 And LSVerMinor > 5) Or (LSVerMajor = 5 And LSVerMinor = 5 And LSBuildMajor >= 2) Then
                    'send the total execution time to the local server so it can be logged there
                    TotalTime = DateAndTime.Timer - StartTime
                    OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
                End If
                OutStr = Len(OutStr) & vbCrLf & OutStr
                Send(OutStr)
                Common.Write_Log(LogFile, OutStr)
            End If

        Catch excp As Exception
            DupKey = False
            If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
                ex = Err.GetException
                If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
                    DupKey = True
                End If
            End If
            If DupKey Then
                Response.Write("Duplicate CardIDs.ExtCardID violation - two simultaneous PhoneHome's for the same card must have been received. " & ex.ToString)
                Common.Write_Log(LogFile, "Duplicate CardIDs.ExtCardID violation - two simultaneous PhoneHome's for the same card must have been received." & " serial=" & LocalServerID & "Mac IPAddress=" & MacAddress & " server=" & Environment.MachineName)
            Else
                Response.Write(Common.Error_Processor(, "Serial= " & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
                Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
            End If
            Response.End()
        End Try
    End Sub
  
    ' -----------------------------------------------------------------------------------------------
  
    Sub Handle_Primary_Request(ByVal LocalServerID As Long, ByVal LocationID As Long)
    
        Dim ex As System.Data.SqlClient.SqlException
        Dim DupKey As Boolean
        Dim ErrDesc As String
        Dim ExtCardID As String = ""
        Dim CustomerPK As Long
        Dim OutStr As String
        Dim DelimChar As Integer
        Dim NumRecs As Long
        Dim RunDate As String
        Dim RunTime As String
        Dim LastUserLocationID As Long
        Dim UniqueError As Boolean
        Dim dst As DataTable
        Dim tempstr As String
        Dim LockingGroupId As Long
        
        DelimChar = 30
    
        On Error GoTo QueryError
        MacAddress = Trim(Request.QueryString("mac"))
    
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If
        
        OutStr = ""
        OutStr = OutStr & "LocationID=" & LocationID & vbCrLf
    
        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Common.Write_Log(LogFile, "Operate at enterprise is enabled (CPE_SystemOption 91), so PhoneHome customer lookups are not allowed." & "LocalServerID=" & LocalServerID & "Mac IPAddress=" & MacAddress)
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send("PhoneHome customer lookups not allowed while operate at enterprise is enabled (CPE_SystemOption 91)")
            Exit Sub
        End If
    
        CustomerPK = Common.Extract_Val(Request.QueryString("id"))
        If CustomerPK = 0 Then
            Send_Response_Header("Missing UserID", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Common.Write_Log(LogFile, "Missing UserID (customerpk of household)!" & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
        Else
            Common.Write_Log(LogFile, "Received UserID: " & CustomerPK)
            'make sure this customer is associated with this location
            Common.QueryStr = "dbo.pa_CPE_IN_UpdateCustomerLocation"
            Common.Open_LXSsp()
            Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
            Common.LXSsp.ExecuteNonQuery()
            Common.Close_LXSsp()
            Common.Write_Log(LogFile, vbCrLf & "Returned the following data for this customer:" & vbCrLf)
      
            LockingGroupId = Common.NZ((Request.QueryString("lockinggroupid")), 0)
            OutStr = OutStr & Fetch_PhoneHome_Data(CustomerPK, DelimChar, LocalServerID, LocationID, 0, ExtCardID)
      
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
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
            Resume Next
        Else
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()
    
    End Sub
  


    ' -----------------------------------------------------------------------------------------------
  
    Sub Handle_AltID_Request(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal BannerID As Integer, ByVal Mode As String)
    
        Dim ex As System.Data.SqlClient.SqlException
        Dim DupKey As Boolean
        Dim ErrDesc As String
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
        Dim ExtCardID As String = ""
        Dim VerifierID As String
        Dim AltTable() As String
        Dim VerifierTable() As String
        Dim NumRecs As Long
        Dim CustomerData As String = "" 'the standard PhoneHome data returned when creating a NEW customer record
        Dim ProcessOK As Boolean
        Dim CustList As String
        Dim row As DataRow
    
        DelimChar = 30
    
        On Error GoTo QueryError
        
        MacAddress = Trim(Request.QueryString("mac"))
    
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If
                
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
        OutStr = OutStr & "LocationID=" & LocationID & vbCrLf
    
        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Common.Write_Log(LogFile, "Operate at enterprise is enabled (CPE_SystemOption 91), so AltID lookups are not allowed.")
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send("AltID lookups not allowed while operate at enterprise is enabled (CPE_SystemOption 91)")
            Exit Sub
        End If
    
        AltID = Trim(Request.QueryString("altid"))
        VerifierID = Trim(Request.QueryString("verifier"))  'need to modifiy code to save this
        Common.Write_Log(LogFile, "altid=" & AltID & "  verifier=" & VerifierID & " Serial=" & LocalServerID & " IP=" & LocalServerIP & " MacAddress=" & MacAddress)

    
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
                Common.QueryStr = "select Customers.CustomerPK, " & AltIDColumn & ", " & VerifierColumn & ", isnull(Customers.InitialCardIDOriginal, '') as ClientUserID1, isnull(Customers.CustomerTypeID, 0) as HHRec, isnull(Customers.Employee, 0) as Employee, 0 as RecordStatus, isnull(Customers.FirstName, '') as FirstName, isnull(Customers.LastName, '') as LastName, isnull(Customers.EmployeeID, '') as EmployeeID " & _
                          "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                          "where isnull(Customers.CustomerStatusID, 1)=1 and " & OrigAltIdColumn & "='" & AltID & "';"
            Else
                Common.QueryStr = "select Customers.CustomerPK, " & AltIDColumn & ", " & VerifierColumn & ", isnull(Customers.InitialCardIDOriginal, '') as ClientUserID1, isnull(Customers.CustomerTypeID, 0) as HHRec, isnull(Customers.Employee, 0) as Employee, 0 as RecordStatus, isnull(Customers.FirstName, '') as FirstName, isnull(Customers.LastName, '') as LastName, isnull(Customers.EmployeeID, '') as EmployeeID " & _
                                  "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                                  "where isnull(Customers.CustomerStatusID, 1)=1 and " & OrigAltIdColumn & "='" & AltID & "' and Customers.BannerID=" & BannerID & ";"
            End If
            dst = Common.LXS_Select
            OutStr = OutStr & Construct_Table("AltIDCacheUsers", 1, 30, LocalServerID, LocationID, dst)

            If dst.Rows.Count > 0 Then
                CustList = ""
                CustList = "(-77"
                For Each row In dst.Rows
                    CustList = CustList & "," & row.Item("CustomerPK")
                Next
                CustList = CustList & ")"
      
                Common.QueryStr = "select CardIDs.CardPK, CardIDs.CustomerPK, CardIDs.ExtCardIDOriginal as ExtCardID, CardIDs.CardStatusID, CardIDs.CardTypeID " & _
                                  " from CardIDs with (NoLock) " & _
                                  " where CardIDs.CustomerPK IN " & CustList & ";"
                dst = Common.LXS_Select
                OutStr = OutStr & Construct_Table("AltIDCacheCardIDs", 1, 30, LocalServerID, LocationID, dst)
            End If
        End If
      
        Common.Write_Log(LogFile, "Returned the following data for this AltID lookup:")
      
        Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        'send the total execution time to the local server so it can be logged there
        TotalTime = DateAndTime.Timer - StartTime
        OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf & " Serial=" & LocalServerID & " MacAddress=" & MacAddress
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
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IPAddress=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()
    
    End Sub


    ' -----------------------------------------------------------------------------------------------
  
  
    Sub Handle_AltIDEnrollment_Request(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal BannerID As Integer, ByVal Mode As String)
    
        Dim ex As System.Data.SqlClient.SqlException
        Dim DupKey As Boolean
        Dim ErrDesc As String
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
        Dim CustomerData As String = "" 'the standard PhoneHome data returned when creating a NEW customer record
        Dim ProcessOK As Boolean
        Dim SendAllCustData As Boolean
    
        DelimChar = 30
    
        On Error GoTo QueryError
        
        MacAddress = Trim(Request.QueryString("mac"))
    
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If
        
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
        OutStr = OutStr & "LocationID=" & LocationID & vbCrLf
    
        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Common.Write_Log(LogFile, "Operate at enterprise is enabled (CPE_SystemOption 91), so AltID enrollment not allowed." & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send("AltID enrollment is not allowed while operate at enterprise is enabled (CPE_SystemOption 91)")
            Exit Sub
        End If
    
    
        AltID = Trim(Request.QueryString("altid"))
        VerifierID = Trim(Request.QueryString("verifier"))  'need to modifiy code to save this
        CustomerPK = Common.Extract_Val(Trim(Request.QueryString("cardid"))) 'this is the CustomerPK
        Common.Write_Log(LogFile, "altid=" & AltID & "  verifier=" & VerifierID & "  cardid=" & CustomerPK & "  createcust=" & Request.QueryString("createcust"))


        'look for the things we NEED to HAVE in order to continue
        ProcessOK = True
        If AltID = "" Then
            OutStr = OutStr & "R:04 Missing AltID" & vbCrLf
            ProcessOK = False
        End If
        If Not (OrigVerifierColumn = "") And VerifierID = "" Then  'a verifier is required and we didn't get one from the local server
            OutStr = OutStr & "R:04 Missing VerifierID (verifier)" & vbCrLf
            ProcessOK = False
        End If
        If OrigAltIdColumn = "" Then
            OutStr = OutStr & "R:05 No Customer Alternate ID has been specified in SystemOptions" & vbCrLf
            ProcessOK = False
        End If
    
    
        If ProcessOK Then
            AutoGenOn = False
            If Common.Fetch_CPE_SystemOption(87) = "1" Then AutoGenOn = True
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
                    ' see if the altid is already in use by someone other than the current customerpk
                    VerifierCondition = ""
                    If AltIDUniqueness = 3 Then  'unique by AltID and Verifier
                        VerifierCondition = " and " & OrigVerifierColumn & "='" & VerifierID & "' "
                    End If
                    Common.QueryStr = "select count(*) as NumRecs " & _
                          "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                          "where not(Customers.customerpk=" & CustomerPK & ") and " & OrigAltIdColumn & "='" & AltID & "' " & VerifierCondition & ";"
                Else  'AltID's ARE unique by BannerID
                    ' see if the altid is already in use by someone other than the current customerpk within this banner
                    Common.QueryStr = "select count(*) as NumRecs " & _
                          "from Customers with (NoLock) Left Join CustomerExt with (NoLock) on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                          "where not(Customers.customerpk=" & CustomerPK & ") and " & OrigAltIdColumn & "='" & AltID & "' and Customers.BannerID=" & BannerID & ";"
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
                    NewCustomerPK = MyAltID.GetAutoGeneratedCustomerPK(ExtCardID, BannerID) 'this is in the CustomerInquiry DLL
                    If NewCustomerPK > 0 Then
                        CustomerPK = NewCustomerPK
                        Common.Write_Log(LogFile, "Auto-generated customer " & ExtCardID & " (" & CustomerPK & ") during AltID enrollment.")
                        'make sure this customer is associated with this location
                        Common.QueryStr = "dbo.pa_CPE_IN_UpdateCustomerLocation"
                        Common.Open_LXSsp()
                        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                        Common.LXSsp.ExecuteNonQuery()
                        Common.Close_LXSsp()
                        Common.Write_Log(LogFile, "Associated new customer (" & CustomerPK & ") with this location (" & LocationID & ").")
                        'send the normal load of PhoneHome data for the newly created customer record
                        SendAllCustData = 1
                    Else
                        CustomerPK = 0
                        Send_Response_Header(MyAltID.ErrorMessage, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                        Common.Write_Log(LogFile, MyAltID.ErrorMessage & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
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
            
                            ' the update didnt work
                            OutStr = OutStr & "R:02 Alternate not updated"
                        End If
              
                    End If  'Common.RowsAffected>0
              
                End If  'not(CustomerPK=0)
                'now that the AltID and Verifier have been set/updated, pull out the customer data if appropriate 
                'send the normal load of PhoneHome data for the newly created customer record
                CustomerData = Fetch_PhoneHome_Data(CustomerPK, DelimChar, LocalServerID, LocationID, 1, ExtCardID)
            
            End If  'NumRecs>0
      
        End If 'ProcessOK

    
        Common.Write_Log(LogFile, "Returned the following data for this AltID lookup:")
      
        OutStr = OutStr & CustomerData
        Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
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
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()
    
    End Sub
  
  
    ' -----------------------------------------------------------------------------------------------
    '' ''R:01 Customer does not have AlternateID ###.
    '' ''R:01 AlternateID ### edited to ###.
    '' ''R:02 AlternateID ### could not be updated.
    '' ''R:04 Missing AltID.
    '' ''R:04 Missing NewAltID.
    '' ''R:06 Missing CustomerPK (CardID).
    '' ''R:06 The specified CustomerPK (###) is invalid.
    '' ''R:06 The requested new Alternate ID ### is already in use.
    ' ------------------------------------------------------------------------------------------------

    Sub Handle_AltIDCardEnroll_Request(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal BannerID As Integer, ByVal Mode As String)
    
        'Sample request: PhoneHome.aspx?serial=1&mode=altidcard_enroll&altid=336699&createcust=0&cardid=206827&lsversion=5.15&lsbuild=0&IP=192.168.3.18
        Dim ex As System.Data.SqlClient.SqlException
        Dim MyLookup As New Copient.CustomerLookup
        Dim dst As DataTable
        Dim DelimChar As Integer = 30
        Dim AltID As String = ""
        Dim CustomerPK As Long = 0
        Dim NewCustomerPK As Long = 0
        Dim NewCardPK As Long = 0
        Dim CreateCust As String = ""
        Dim ShouldCreateCustomer As Boolean = False
        Dim SendAllCustData As Boolean = False
        Dim ProcessOK As Boolean = True
        Dim OutStr As String = ""
    
        On Error GoTo QueryError
    
        MacAddress = Trim(Request.QueryString("mac"))
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If
        OutStr = "LocationID=" & LocationID & vbCrLf
    
        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Common.Write_Log(LogFile, "Operate-at-enterprise (CPE system option 91) is enabled, so AltID enrollment is not allowed. Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Server=" & Environment.MachineName)
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send("AltID enrollment is not allowed while operate-at-enterprise (CPE system option 91) is enabled.")
            Exit Sub
        End If
        AltID = Trim(Request.QueryString("altid"))
        ' AltID = PadAltID(AltID) 
        CustomerPK = Common.Extract_Val(Trim(Request.QueryString("cardid"))) 'this is the CustomerPK
        CreateCust = Request.QueryString("createcust")
        ShouldCreateCustomer = IIf(CreateCust = "1", True, False)
        Common.Write_Log(LogFile, "altid=" & AltID & "  cardid=" & CustomerPK & "  createcust=" & CreateCust)
        'Verify the presence of necessary parameters
        If AltID = "" Then
            OutStr = OutStr & "R:04 Missing AltID." & vbCrLf
            ProcessOK = False
        End If
        If (ProcessOK = True And Common.AllowToProcessCustomerCard(AltID, Customer.CardTypes.ALTERNATEID, cardValidationResp) <> True) Then
            OutStr = OutStr & "R:06 AlternateID " & AltID & " should be numeric" & vbCrLf
            ProcessOK = False
            Common.Write_Log(LogFile, Common.CardValidationResponseMessage(AltID, Customer.CardTypes.ALTERNATEID, cardValidationResp))
        End If
   
        If ProcessOK Then
            If CustomerPK = 0 And ShouldCreateCustomer = False Then
                OutStr = OutStr & "R:06 Missing CustomerPK (CardID) in AltID enrollment. No create card request sent from the store." & vbCrLf
                ProcessOK = False
            End If
            If Not (CustomerPK = 0) Then
                If Not MyLookup.DoesCustomerExist(CustomerPK) Then
                    OutStr = OutStr & "R:06 The specified CustomerPK (" & CustomerPK & ") is invalid." & vbCrLf
                    ProcessOK = False
                End If
            End If
        End If
    
        If ProcessOK Then
            AltID = Common.Pad_ExtCardID(AltID, Customer.CardTypes.ALTERNATEID)
            Common.QueryStr = "select * from CardIDs with (NoLock) where CardTypeID=3 and ExtCardID= @AltID and CustomerPK= @CustomerPK"
            Common.DBParameters.Add("@AltID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(AltID)
            Common.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            dst = Common.ExecuteQuery(DataBases.LogixXS)
            If dst.Rows.Count > 0 Then
                'Specified customer already has the specified AltID
                OutStr = OutStr & "R:01 Customer already has AlternateID " & AltID & "." & vbCrLf
                ProcessOK = False
            Else
                'Specified customer does not have the specified AltID, so proceed
                Common.Write_Log(LogFile, "Received AltID " & AltID & " in enrollment mode.")
                If ShouldCreateCustomer AndAlso CustomerPK = 0 Then
                    Common.Write_Log(LogFile, "Customer creation requested.")
                Else
                    Common.Write_Log(LogFile, "Received CardID (CustomerPK) " & CustomerPK & ".")
                End If
                Common.QueryStr = "select OnePerCustomer from CardTypes with (NoLock) where CardTypeID=" & Customer.CardTypes.ALTERNATEID & " and OnePerCustomer=1;"
                dst = Common.LXS_Select
                If dst.Rows.Count > 0 Then
                    'See if the customer already has an AltID
                    Common.QueryStr = "select * from CardIDs with (NoLock) where CardTypeID=3 and CustomerPK=" & CustomerPK & ";"
                    dst = Common.LXS_Select
                    If dst.Rows.Count > 0 Then
                        'Specified customer already has an AltID
                        OutStr = OutStr & "R:01 Customer already has an AlternateID " & AltID & " has not been added, use edit function instead." & vbCrLf
                        ProcessOK = False
                    End If
                End If
                'See if the AltID is already in use by another customer
      
                Common.QueryStr = "select CustomerPK from CardIDs with (NoLock) where CardTypeID= @CardTypeID and ExtCardID= @AltID and CustomerPK <> @CustomerPK "
                Common.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = Customer.CardTypes.ALTERNATEID
                Common.DBParameters.Add("@AltID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(AltID.ConvertBlankIfNothing(), True)
                Common.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                dst = Common.ExecuteQuery(DataBases.LogixXS)
                If dst.Rows.Count > 0 Then
                    OutStr = OutStr & "R:03 AlternateID " & AltID & " is already in use by another customer (" & Common.NZ(dst.Rows(0).Item("CustomerPK"), 0) & ")." & vbCrLf
                    ProcessOK = False
                ElseIf (ProcessOK) Then
                    'New AltID is unused, so proceed
                    Common.Write_Log(LogFile, "CustomerPK=" & CustomerPK & ", ShouldCreateCard=" & ShouldCreateCustomer)
                    'If ShouldCreateCustomer was requested, then create a new customer record
                    If CustomerPK = 0 AndAlso ShouldCreateCustomer Then
                        Common.QueryStr = "dbo.pt_NewCustomer_Insert"
                        Common.Open_LXSsp()
                        Common.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(AltID, True)
                        Common.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = Customer.CardTypes.ALTERNATEID
                        Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Direction = ParameterDirection.Output
                        Common.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(AltID)
                        Common.LXSsp.ExecuteNonQuery()
                        NewCustomerPK = Common.LXSsp.Parameters("@CustomerPK").Value
                        Common.Close_LXSsp()
                        If NewCustomerPK > 0 Then
                            CustomerPK = NewCustomerPK
                            'Common.QueryStr = "select top 1 CardPK from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " and CardTypeID=" & Customer.CardTypes.ALTERNATEID & " and ExtCardID='" & AltID & "';"
                            'dst = Common.LXS_Select
                            Common.QueryStr = "select top 1 CardPK from CardIDs with (NoLock) where CustomerPK= @CustomerPK and CardTypeID= @CardTypeID and ExtCardID= @AltID"
                            Common.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                            Common.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = Customer.CardTypes.ALTERNATEID
                            Common.DBParameters.Add("@AltID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(AltID, True)
                            dst = Common.ExecuteQuery(DataBases.LogixXS)
                            If dst.Rows.Count > 0 Then NewCardPK = Common.NZ(dst.Rows(0).Item("CardPK"), 0)
                            OutStr = OutStr & "R:01 CardPK:" & NewCardPK & " AlternateID " & AltID & " created." & vbCrLf
                            Common.Write_Log(LogFile, "Auto-generated CustomerPK " & CustomerPK & " during AltID enrollment.")
                            Common.Activity_Log2(25, 10, CustomerPK, AdminUserID, Copient.PhraseLib.Lookup("history.customer-add-customer", LanguageID))
                            'Assign customer to banner
                            If BannerID > -1 Then
                                MyLookup.SetBannerForCustomer(CustomerPK, BannerID)
                            End If
                            'Associate customer with location
                            Common.QueryStr = "dbo.pa_CPE_IN_UpdateCustomerLocation"
                            Common.Open_LXSsp()
                            Common.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                            Common.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                            Common.LXSsp.ExecuteNonQuery()
                            Common.Close_LXSsp()
                            Common.Write_Log(LogFile, "Associated CustomerPK " & CustomerPK & " with this location (" & LocationID & ").")
                            'Send the normal load of PhoneHome data for the newly-created customer record
                            SendAllCustData = True
                        Else
                            Send_Response_Header(MyAltID.ErrorMessage, Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                            OutStr = OutStr & "R:02 AlternateID " & AltID & " could not be created."
                            Common.Write_Log(LogFile, MyAltID.ErrorMessage & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                            ProcessOK = False
                        End If
                    ElseIf CustomerPK > 0 Then
                        Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
                        If MyLookup.AddCardToCustomer(CustomerPK, AltID, Customer.CardTypes.ALTERNATEID, 1, ReturnCode) Then
                            'Common.QueryStr = "select top 1 CardPK from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & " and CardTypeID=" & Customer.CardTypes.ALTERNATEID & " and ExtCardID='" & AltID & "';"
                            'dst = Common.LXS_Select
                            Common.QueryStr = "select top 1 CardPK from CardIDs with (NoLock) where CustomerPK= @CustomerPK and CardTypeID= @CardTypeID and ExtCardID= @AltID"
                            Common.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                            Common.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = Customer.CardTypes.ALTERNATEID
                            Common.DBParameters.Add("@AltID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(AltID, True)
                            dst = Common.ExecuteQuery(DataBases.LogixXS)
                            If dst.Rows.Count > 0 Then NewCardPK = Common.NZ(dst.Rows(0).Item("CardPK"), 0)
                            OutStr = OutStr & "R:01 CardPK:" & NewCardPK & " AlternateID " & AltID & " created." & vbCrLf
                        Else
                            OutStr = OutStr & "R:02 AlternateID " & AltID & " could not be created."
                            ProcessOK = False
                        End If
                    End If
                End If
            End If
        End If
    
        Common.Write_Log(LogFile, "Returned the following data for this AltID lookup:")
    
        Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        'Send the total execution time to the local server so it can be logged there
        TotalTime = DateAndTime.Timer - StartTime
        OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
        OutStr = Len(OutStr) & vbCrLf & OutStr
        Send(OutStr)
        Common.Write_Log(LogFile, OutStr)
    
AllDone:
        Exit Sub
    
QueryError:
        Dim DupKey As Boolean = False
        If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
            ex = Err.GetException
            If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
                DupKey = True
            End If
        End If
        If DupKey Then
            Resume Next
        Else
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()
    
    End Sub
  
  
    ' -----------------------------------------------------------------------------------------------
  
  
    Sub Handle_AltIDCardEdit_Request(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal BannerID As Integer, ByVal Mode As String)
    
        'Sample request: PhoneHome.aspx?serial=1&mode=altidcard_edit&altid=336699&newaltid=334455&cardid=206827&lsversion=5.15&lsbuild=0&IP=192.168.3.18
        Dim ex As System.Data.SqlClient.SqlException
        Dim MyLookup As New Copient.CustomerLookup
        Dim dst As DataTable
        Dim DelimChar As Integer = 30
        Dim AltID As String = ""
        Dim NewAltID As String = ""
        Dim CustomerPK As Long = 0
        Dim CardPK As Long = 0
        Dim SendAllCustData As Boolean = False
        Dim ProcessOK As Boolean = True
        Dim OutStr As String = ""
    
        On Error GoTo QueryError
    
        MacAddress = Trim(Request.QueryString("mac"))
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If
        OutStr = "LocationID=" & LocationID & vbCrLf
    
        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Common.Write_Log(LogFile, "Operate-at-enterprise (CPE system option 91) is enabled, so AltID is not allowed. Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Server=" & Environment.MachineName)
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send("AltID is not allowed while operate-at-enterprise (CPE system option 91) is enabled.")
            Exit Sub
        End If
    
        AltID = Trim(Request.QueryString("altid"))
    
        AltID = Common.Pad_ExtCardID(AltID, Customer.CardTypes.ALTERNATEID)
        NewAltID = Trim(Request.QueryString("newaltid"))
    
        NewAltID = Common.Pad_ExtCardID(NewAltID, Customer.CardTypes.ALTERNATEID)
        CustomerPK = Common.Extract_Val(Trim(Request.QueryString("cardid"))) 'this is the CustomerPK
        Common.Write_Log(LogFile, "altid=" & AltID & "  newaltid=" & NewAltID & "  cardid=" & CustomerPK)
    
        'Verify the presence of necessary parameters
        If AltID = "" Then
            OutStr = OutStr & "R:04 Missing AltID." & vbCrLf
            ProcessOK = False
        End If
        If NewAltID = "" Then
            OutStr = OutStr & "R:04 Missing NewAltID." & vbCrLf
            ProcessOK = False
        End If
        If CustomerPK = 0 Then
            OutStr = OutStr & "R:04 Missing CustomerPK (CardID)." & vbCrLf
            ProcessOK = False
        Else
            If Not MyLookup.DoesCustomerExist(CustomerPK) Then
                OutStr = OutStr & "R:06 The specified CustomerPK (" & CustomerPK & ") is invalid." & vbCrLf
                ProcessOK = False
            End If
        End If
    
        'Verify that the customer possesses the specified AltID
        If ProcessOK Then
            Common.QueryStr = "select CardPK from CardIDs with (NoLock) where CardTypeID=" & Customer.CardTypes.ALTERNATEID & " and ExtCardID='" & MyCryptlib.SQL_StringEncrypt(AltID, True) & "' and CustomerPK=" & CustomerPK & ";"
            dst = Common.LXS_Select
            If dst.Rows.Count > 0 Then
                CardPK = Common.NZ(dst.Rows(0).Item("CardPK"), 0)
            Else
                OutStr = OutStr & "R:02 Customer does not have AlternateID " & AltID & "." & vbCrLf
                ProcessOK = False
            End If
        End If
        'Verify that no one possesses the specified NewAltID
        If ProcessOK Then
            Common.QueryStr = "select * from CardIDs with (NoLock) where CardTypeID=" & Customer.CardTypes.ALTERNATEID & " and ExtCardID='" & MyCryptlib.SQL_StringEncrypt(NewAltID, True) & "';"
            dst = Common.LXS_Select
            If dst.Rows.Count > 0 Then
                OutStr = OutStr & "R:06 The requested new AlternateID " & AltID & " is already in use." & vbCrLf
                ProcessOK = False
            End If
        End If
    
        If ProcessOK Then
            'Specified customer has the specified AltID, and the NewAltID is available, so proceed
            Common.Write_Log(LogFile, "Received AltID " & AltID & ", NewAltID " & NewAltID & ", and CardID (CustomerPK) " & CustomerPK & " in edit mode.")
            If (Common.AllowToProcessCustomerCard(NewAltID, Customer.CardTypes.ALTERNATEID, Nothing)) Then
                Common.QueryStr = "update CardIDs set ExtCardID='" & MyCryptlib.SQL_StringEncrypt(NewAltID, True) & "',  ExtCardIDOriginal='" & MyCryptlib.SQL_StringEncrypt(NewAltID) & "' where CardPK=" & CardPK & ";"
                Common.LXS_Execute()
                If (Common.RowsAffected > 0) Then
                    OutStr = OutStr & "R:01 AlternateID " & AltID & " edited to " & NewAltID & "." & vbCrLf
                    Common.Activity_Log2(25, 11, CustomerPK, AdminUserID, Copient.PhraseLib.Lookup("history.customer-edit-card", LanguageID) & _
                                                                          " " & AltID & " " & Copient.PhraseLib.Lookup("term.to", LanguageID) & _
                                                                          " " & NewAltID & " (" & Copient.PhraseLib.Lookup("term.alternateid", LanguageID) & ")")
                Else
                    OutStr = OutStr & "R:02 AlternateID " & AltID & " could not be updated."
                    ProcessOK = False
                End If
            Else
                OutStr = OutStr & "R:06 Requested new AlternateID " & NewAltID & " should be numeric." & vbCrLf
                Common.Write_Log(LogFile, "AlternateID:" & NewAltID & "should be numeric.")
                ProcessOK = False
            End If
        End If
    
        Common.Write_Log(LogFile, "Returned the following data for this AltID lookup:")
    
        Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        'Send the total execution time to the local server so it can be logged there
        TotalTime = DateAndTime.Timer - StartTime
        OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
        OutStr = Len(OutStr) & vbCrLf & OutStr
        Send(OutStr)
        Common.Write_Log(LogFile, OutStr)
    
AllDone:
        Exit Sub
    
QueryError:
        Dim DupKey As Boolean = False
        If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
            ex = Err.GetException
            If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
                DupKey = True
            End If
        End If
        If DupKey Then
            Resume Next
        Else
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()
    
    End Sub
  
  
    ' -----------------------------------------------------------------------------------------------

  
    Sub Handle_AltIDCardDelete_Request(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal BannerID As Integer, ByVal Mode As String)
    
        'Sample request: PhoneHome.aspx?serial=1&mode=altidcard_delete&altid=336699&cardid=206827&lsversion=5.15&lsbuild=0&IP=192.168.3.18
        Dim ex As System.Data.SqlClient.SqlException
        Dim MyLookup As New Copient.CustomerLookup
        Dim dst As DataTable
        Dim DelimChar As Integer = 30
        Dim AltID As String = ""
        Dim CustomerPK As Long = 0
        Dim CardPK As Long = 0
        Dim ProcessOK As Boolean = True
        Dim OutStr As String = ""
    
        On Error GoTo QueryError
    
        MacAddress = Trim(Request.QueryString("mac"))
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If
        OutStr = "LocationID=" & LocationID & vbCrLf
    
        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Common.LRTsp.ExecuteNonQuery()
        Common.Close_LRTsp()
    
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Common.Write_Log(LogFile, "Operate-at-enterprise (CPE system option 91) is enabled, so AltID is not allowed. Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Server=" & Environment.MachineName)
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send("AltID is not allowed while operate-at-enterprise (CPE system option 91) is enabled.")
            Exit Sub
        End If
    
        AltID = Trim(Request.QueryString("altid"))

        CustomerPK = Common.Extract_Val(Trim(Request.QueryString("cardid"))) 'this is the CustomerPK
        Common.Write_Log(LogFile, "altid=" & AltID & "  cardid=" & CustomerPK)
    
        'Verify the presence of necessary parameters
        If AltID = "" Then
            OutStr = OutStr & "R:04 Missing AltID." & vbCrLf
            ProcessOK = False
        End If
        If CustomerPK = 0 Then
            OutStr = OutStr & "R:04 Missing CustomerPK (CardID)." & vbCrLf
            ProcessOK = False
        Else
            CustomerPK = Common.Pad_ExtCardID(CustomerPK, Customer.CardTypes.CUSTOMER)
            If Not MyLookup.DoesCustomerExist(CustomerPK) Then
                OutStr = OutStr & "R:06 The specified CustomerPK (" & CustomerPK & ") is invalid." & vbCrLf
                ProcessOK = False
            End If
        End If
    
        AltID = Common.Pad_ExtCardID(AltID, Customer.CardTypes.ALTERNATEID)
       
        If ProcessOK Then
            Common.QueryStr = "select CardPK from CardIDs with (NoLock) " & _
                              "where CardTypeID=" & Customer.CardTypes.ALTERNATEID & " and ExtCardID='" & MyCryptlib.SQL_StringEncrypt(AltID, True) & "' and CustomerPK=" & CustomerPK & ";"
            dst = Common.LXS_Select
            If dst.Rows.Count = 0 Then
                'Specified customer does not have the specified AltID
                OutStr = OutStr & "R:01 Customer does not have AlternateID " & AltID & "." & vbCrLf
                ProcessOK = False
            Else
                'Specified customer has the specified AltID, so proceed
                CardPK = Common.NZ(dst.Rows(0).Item("CardPK"), 0)
                Common.Write_Log(LogFile, "Received AltID " & AltID & " and CardID (CustomerPK) " & CustomerPK & " in delete mode.")
                Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
                If MyLookup.RemoveCardFromCustomer(CustomerPK, CardPK, ReturnCode) Then
                    OutStr = OutStr & "R:01 AlternateID " & AltID & " deleted." & vbCrLf
                Else
                    OutStr = OutStr & "R:02 AlternateID " & AltID & " could not be deleted."
                    ProcessOK = False
                End If
            End If
        End If
    
        Common.Write_Log(LogFile, "Returned the following data for this AltID lookup:")
    
        Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        'Send the total execution time to the local server so it can be logged there
        TotalTime = DateAndTime.Timer - StartTime
        OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
        OutStr = Len(OutStr) & vbCrLf & OutStr
        Send(OutStr)
        Common.Write_Log(LogFile, OutStr)
    
AllDone:
        Exit Sub
    
QueryError:
        Dim DupKey As Boolean = False
        If Err.Source = ".Net SqlClient Data Provider" Then 'an SQL error occurred
            ex = Err.GetException
            If ex.Number = 2627 Or ex.Number = 2601 Then 'tried to insert a duplicate value into a column with a unique index - ignore it and go on
                DupKey = True
            End If
        End If
        If DupKey Then
            Resume Next
        Else
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " Process Info: Server Name=" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()
    
    End Sub
  
  
    ' -----------------------------------------------------------------------------------------------
  
  
    Sub Handle_CustomerLocking_Request(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal Mode As String)
    
        Dim ex As System.Data.SqlClient.SqlException
        Dim DupKey As Boolean
        Dim ErrDesc As String
        Dim OutStr As String
        Dim DelimChar As Integer
        Dim dst As DataTable
        Dim sLockOption As Integer
        Dim sCustomerPK As String
        Dim sLockingGroupId As String
        Dim sLockingMode As String
        Dim sLockedDate As String
        Dim dtLocked As Date
        Dim iLockStatus As Integer
        Dim lMinutes As Long
        Dim CustLockDelMin As Integer
        Dim elapsed_time As TimeSpan
        Dim sTerminal As String
        Dim sTrxNum As String
        Dim LockedCustomerPK As Long

        DelimChar = 30
    
        On Error GoTo QueryError
        
        MacAddress = Trim(Request.QueryString("mac"))
    
        LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
        LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
        If LocalServerIP = "" Or LocalServerIP = "0" Then
            Common.Write_Log(LogFile, "Could not get IP from query. Analyzing client request for IP ...")
            LocalServerIP = Trim(Request.UserHostAddress)
        End If

        OutStr = ""
        OutStr = OutStr & "LocationID=" & LocationID & vbCrLf
    
        Common.QueryStr = "dbo.pa_CPE_IN_LastHeardUpdate"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
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
    
        sLockOption = Common.Fetch_CPE_SystemOption(86)
        If Common.Fetch_CPE_SystemOption(91) = "1" Then
            Common.Write_Log(LogFile, "Operate at enterprise is enabled (CPE_SystemOption 91), so locking requests are not allowed.")
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send("Customer locking is not allowed while operate at enterprise is enabled (CPE_SystemOption 91)")
        ElseIf sLockOption = "0" Then
            iLockStatus = 3
            Common.Write_Log(LogFile, "Customer locking is turned off (LockMode: " & sLockingMode & ", CustomerPK: " & sCustomerPK & ", TerminalLockingGroupID: " & sLockingGroupId & ")")
      
            OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
            TotalTime = DateAndTime.Timer - StartTime
            OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
            OutStr = Len(OutStr) & vbCrLf & OutStr
      
            Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
            Send(OutStr)
            Common.Write_Log(LogFile, OutStr)
        Else
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
          
                    iLockStatus = HandleCustomerLocking(sLockOption, Long.Parse(sCustomerPK), Long.Parse(sLockingGroupId), LocationID, LockedCustomerPK)
        
                    OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
                    TotalTime = DateAndTime.Timer - StartTime
                    OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
                    OutStr = Len(OutStr) & vbCrLf & OutStr
      
                    Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Send(OutStr)
                    Common.Write_Log(LogFile, OutStr)
                End If
            ElseIf (sLockingMode = "UNLOCK") Then
                If sCustomerPK = "" Then
                    Send_Response_Header("Missing CustomerPK (CustomerPK)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, "Missing CustomerPK (CustomerPK)!" & " IP=" & LocalServerIP & " MacAddress=" & MacAddress)
                ElseIf sLockingGroupId = "" AndAlso sLockOption = "2" Then
                    Send_Response_Header("Missing TerminalLockingGroupID (lockingGroupID)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, "Missing TerminalLockingGroupID (lockingGroupID)" & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerID & " server=" & Environment.MachineName)
                Else
                    Dim iRowCount As Integer
                    If sLockOption = "2" Then
                        Common.Write_Log(LogFile, "Received LockMode: Unlock, CustomerPK: " & sCustomerPK & ", TerminalLockingGroupID: " & sLockingGroupId)
                    Else
                        Common.Write_Log(LogFile, "Received LockMode: Unlock, CustomerPK: " & sCustomerPK)
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
                    OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
                    TotalTime = DateAndTime.Timer - StartTime
                    OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
                    OutStr = Len(OutStr) & vbCrLf & OutStr
      
                    Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Send(OutStr)
                    Common.Write_Log(LogFile, OutStr)
                End If

            ElseIf (sLockingMode = "FORCELOCK") Then
            
                If sCustomerPK = "" Then
                    Send_Response_Header("Missing CustomerPK (CustomerPK)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, "Missing CustomerPK (CustomerPK)!" & " serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                ElseIf sLockingGroupId = "" AndAlso sLockOption = "2" Then
                    Send_Response_Header("Missing TerminalLockingGroupID (lockingGroupID)", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, "Missing TerminalLockingGroupID (lockingGroupID)" & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & Trim(Request.UserHostAddress) & " server=" & Environment.MachineName)
                Else
                    Dim iRowCount As Integer
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
                    '   OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
                    '   TotalTime = Timer - StartTime
                    '   OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
                    '   OutStr = Len(OutStr) & vbCrLf & OutStr
      
                    '   Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    '   Send(OutStr)
                    '   Common.Write_Log(LogFile, OutStr)
 
          
                    iLockStatus = HandleCustomerLocking(sLockOption, Long.Parse(sCustomerPK), Long.Parse(sLockingGroupId), LocationID, LockedCustomerPK)
        
                    OutStr = OutStr & "Status=" & iLockStatus & vbCrLf
                    TotalTime = DateAndTime.Timer - StartTime
                    OutStr = OutStr & "L:Central Execute Time: " & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)" & vbCrLf
                    OutStr = Len(OutStr) & vbCrLf & OutStr
      
                    Send_Response_Header("PhoneHome", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Send(OutStr)
                    Common.Write_Log(LogFile, OutStr)
                End If
              
            End If
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
            Resume Next
        Else
            Response.Write(Common.Error_Processor(, "Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IPAddress=" & LocalServerIP & " Process Info: Server Name:" & Environment.MachineName & " Invoking LocationID=" & LocationID, , Common.InstallationName))
            Common.Write_Log(LogFile, "***** An error occurred during processing! Please check the ErrorLog! *****")
        End If
        Response.End()
    
    End Sub



    ' -----------------------------------------------------------------------------------------------
  
    Sub Echo_System_Msg(ByVal LocalServerID As Long, ByVal LocationID As Long)
    
        Dim dst As DataTable
        Dim UpdateTime As String
        Dim OutBuffer As String
        Dim LocationName As String = ""
        Dim SubjectStr As String
        Dim ExtLocationCode As String = "0"
    
        UpdateTime = Microsoft.VisualBasic.DateAndTime.Today & " " & Microsoft.VisualBasic.DateAndTime.TimeOfDay
        Common.QueryStr = "select getdate() as CurrentTime;"
        dst = Common.LRT_Select()
        If dst.Rows.Count > 0 Then
            UpdateTime = Common.NZ(dst.Rows(0).Item("CurrentTime"), Microsoft.VisualBasic.DateAndTime.Today & " " & Microsoft.VisualBasic.DateAndTime.TimeOfDay)
        End If
    
        SubjectStr = Request.QueryString("subject")
        Common.Write_Log(LogFile, "Subject: " & SubjectStr)
    
        If (SubjectStr.Contains("incentiveFetch SQL error")) Then
            Update_IncentiveFetch_NAKs(LocalServerID, LocationID, SubjectStr)
        End If
    
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID.ToString & ";"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            LocationName = Common.NZ(dst.Rows(0).Item("LocationName"), "")
            ExtLocationCode = Common.NZ(dst.Rows(0).Item("ExtLocationCode"), "0")
        End If
    
        OutBuffer = "Local Server System Message" & vbCrLf
        OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString & vbCrLf
        OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
        OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "Report received: " & UpdateTime & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "Subject: " & SubjectStr
    
        Common.Send_Email(Common.Get_Error_Emails(4), Common.SystemEmailAddress, "System Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
        Send_Response_Header("SystemMsg", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("ACK")
    End Sub
  
    Sub Echo_Health_Msg(ByVal LocalServerID As Long, ByVal LocationID As Long)
    
        Dim dst As DataTable
        Dim UpdateTime As String
        Dim OutBuffer As String
        Dim ExtLocationCode As String = ""
        Dim LocationName As String = ""
        Dim SubjectStr As String
    
        UpdateTime = Microsoft.VisualBasic.DateAndTime.Today & " " & Microsoft.VisualBasic.DateAndTime.TimeOfDay
        Common.QueryStr = "select getdate() as CurrentTime;"
        dst = Common.LRT_Select()
        If dst.Rows.Count > 0 Then
            UpdateTime = Common.NZ(dst.Rows(0).Item("CurrentTime"), Microsoft.VisualBasic.DateAndTime.Today & " " & Microsoft.VisualBasic.DateAndTime.TimeOfDay)
        End If
    
        If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
        Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID.ToString() & ";"
        dst = Common.LRT_Select
        If dst.Rows.Count > 0 Then
            LocationName = Common.NZ(dst.Rows(0).Item("LocationName"), "")
            ExtLocationCode = Common.NZ(dst.Rows(0).Item("ExtLocationCode"), "0")
        End If

        SubjectStr = Request.QueryString("subject")
        Common.Write_Log(LogFile, "Subject: " & SubjectStr)
        
        OutBuffer = "Local Server Health Message" & vbCrLf
        OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString & vbCrLf
        OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
        OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "Report received: " & UpdateTime & vbCrLf
        OutBuffer = OutBuffer & vbCrLf & "Subject: " & SubjectStr
    
        Common.Send_Email(Common.Get_Error_Emails(7), Common.SystemEmailAddress, "Health Server Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
        Send_Response_Header("HealthMsg", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
        Send("ACK")
    End Sub
  
    Sub Update_IncentiveFetch_NAKs(ByVal LocalServerID As Long, ByVal LocationID As Long, ByVal SubjectStr As String)
        Dim ExtLocationCode As String = ""
        Dim LocationName As String = ""
    
        Dim MaxNAK As Integer = Common.Extract_Val(Common.Fetch_CPE_SystemOption(151))
        Common.QueryStr = "dbo.pt_CPE_UpdateIncentiveFetchNAKCounter"
        Common.Open_LRTsp()
        Common.LRTsp.Parameters.Add("@Max_NAK_Count", SqlDbType.Int).Value = MaxNAK
        Common.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
        Common.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = LocalServerID
        Dim dst As DataTable = Common.LRTsp_select
        Dim IF_Offline As String = Common.NZ(dst.Rows(0).Item("IncentiveFetchOffline"), "")
        Dim NAK_Count As String = Common.NZ(dst.Rows(0).Item("IncentiveFetchNAKCount"), "")
        Dim BatchFileName As String = Common.NZ(dst.Rows(0).Item("BatchFileName"), "")
        Common.Close_LRTsp()
    
        Common.Write_Log(LogFile, "Location " & LocationID.ToString() & " (LocalServer " & LocalServerID.ToString() & ") now has incentiveFetch error(s) " & _
                         NAK_Count & " time(s). Max NAK allowed is " & MaxNAK.ToString() & ". IncentiveFetchOffline is " & IF_Offline & _
                         ". First IncentiveFetch batch file name is " & BatchFileName & ".", True)
        
        If CInt(NAK_Count) >= MaxNAK Then
            If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
            Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID.ToString & ";"
            dst = Common.LRT_Select
            If dst.Rows.Count > 0 Then
                LocationName = Common.NZ(dst.Rows(0).Item("LocationName"), "")
                ExtLocationCode = Common.NZ(dst.Rows(0).Item("ExtLocationCode"), "0")
            End If
            
            Dim OutBuffer As String = "Local Server Incentive Fetch Max NAK Failure Message" & vbCrLf
            OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString() & vbCrLf
            OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
            OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID.ToString() & vbCrLf
            OutBuffer = OutBuffer & "Maximum NAK limit of " & MaxNAK.ToString() & " has been reached. " & vbCrLf
            If Len(BatchFileName) > 0 Then
                OutBuffer = OutBuffer & "First IncentiveFetch batch file name is " & BatchFileName & "." & vbCrLf
            End If
            OutBuffer = OutBuffer & vbCrLf & "Subject: " & SubjectStr
    
            Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Max NAK Message - " & Common.InstallationName & "  Store #" & ExtLocationCode, OutBuffer)
        End If
    End Sub
  
    Function PadAltID(ByVal AltID As String) As String
        AltID = Common.Pad_ExtCardID(AltID, Copient.commonShared.CardTypes.CUSTOMER)
        Return AltID
    End Function
  
</script>
<%
    '-----------------------------------------------------------------------------------------------
    'Main Code - Execution starts here
  
    Dim LocalServerID As Long
    Dim LocationID As Long
    Dim BannerID As Integer
    Dim LastHeard As String
    Dim ZipOutput As Boolean
    Dim DataFile As String
    Dim ZipFile As String
    Dim FileStamp As String
    Dim Mode As String
    Dim RawRequest As String
    Dim Index As Long
    Dim LSVerParts() As String
    Dim OutBuffer As String = ""
    Dim LocationName As String = ""
    Dim ExtLocationCode As String = ""
    Dim rst As New DataTable
  
    Common.AppName = "PhoneHome.aspx"
    Response.Expires = 0
    On Error GoTo ErrorTrap
    StartTime = DateAndTime.Timer
  
    LastHeard = "1/1/1980"
  
    MacAddress = Trim(Request.QueryString("mac"))
    If MacAddress = "" Or MacAddress = "0" Then
        MacAddress = "0"
    End If
    LocalServerID = Common.Extract_Val(Request.QueryString("serial"))
    LogFile = "CPE-PhoneHomeLog-" & Format(LocalServerID, "00000") & "." & Common.Leading_Zero_Fill(Year(Today), 4) & Common.Leading_Zero_Fill(Month(Today), 2) & Common.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
    LocalServerIP = Common.NZ(Request.QueryString("IP"), "").ToString.Trim
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
  
    Mode = UCase(Request.QueryString("mode"))
  
    Common.Open_LogixRT()
    Common.Load_System_Info()
    Connector.Load_System_Info(Common)
  
    If Mode = "SRL" Then
        'Common.Write_Log(LogFile, "----------------------------------------------------------------")
        'Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Mode: " & Mode & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "  Process running on server:" & Environment.MachineName & " Serial=" & LocalServerID & " MacAddress=" & MacAddress)
        Fetch_Serial()
        TotalTime = DateAndTime.Timer - StartTime
        'Common.Write_Log(LogFile, "Serial Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
    Else
        Common.Open_LogixXS()
        Connector.Get_LS_Info(Common, LocalServerID, LocationID, BannerID, LastHeard, 1, LocalServerIP)
        Common.Write_Log(LogFile, "----------------------------------------------------------------")
        Common.Write_Log(LogFile, "** " & Microsoft.VisualBasic.DateAndTime.Now & "  Serial: " & LocalServerID & "   LocationID: " & LocationID & "  Mode: " & Mode & "  LSVersion=" & LSVerMajor & "." & LSVerMinor & "b" & LSBuildMajor & "." & LSBuildMinor & "  Process running on server:" & Environment.MachineName & " with MacAddress=" & MacAddress & " IP=" & LocalServerIP)
        If Not (LocationID = 0) Then
            Select Case Mode
                Case "USER"
                    Handle_User_Request(LocalServerID, LocationID, BannerID, Mode)
                    'Case "ID2"
                    '    Handle_User_Request(LocalServerID, LocationID, Mode)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "User Run Time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "PRIMARY"
                    Handle_Primary_Request(LocalServerID, LocationID)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "Primary Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "ALTID"
                    Handle_AltID_Request(LocalServerID, LocationID, BannerID, Mode)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "AltID Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "ALTIDENROLL"
                    Handle_AltIDEnrollment_Request(LocalServerID, LocationID, BannerID, Mode)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "AltIDEnrollment Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "SYSTEM"
                    Echo_System_Msg(LocalServerID, LocationID)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "SystemMsg Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "CUSTOMERLOCKING"
                    Handle_CustomerLocking_Request(LocalServerID, LocationID, Mode)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "CustomerLocking Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "ALTIDCARD_ENROLL"
                    Handle_AltIDCardEnroll_Request(LocalServerID, LocationID, BannerID, Mode)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "AltIDCard_Enroll Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "ALTIDCARD_EDIT"
                    Handle_AltIDCardEdit_Request(LocalServerID, LocationID, BannerID, Mode)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "AltIDCard_Edit Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case "ALTIDCARD_DELETE"
                    Handle_AltIDCardDelete_Request(LocalServerID, LocationID, BannerID, Mode)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "AltIDCard_Delete Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case Else
                    Send_Response_Header("Invalid Request - bad mode", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    Common.Write_Log(LogFile, "Received invalid request!" & " Serial=" & LocalServerID & " MacAddress=" & MacAddress)
                    Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf & " Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                    RawRequest = Get_Raw_Form(Request.InputStream)
                    Common.Write_Log(LogFile, RawRequest)
                    
                    If Common.LRTadoConn.State = ConnectionState.Closed Then Common.Open_LogixRT()
                    Common.QueryStr = "select LocationName, ExtLocationCode from Locations with (NoLock) where LocationID=" & LocationID & ";"
                    rst = Common.LRT_Select
                    If rst.Rows.Count > 0 Then
                        LocationName = Common.NZ(rst.Rows(0).Item("LocationName"), "")
                        ExtLocationCode = Common.NZ(rst.Rows(0).Item("ExtLocationCode"), "0")
                    End If
    
                    OutBuffer = "Phone Home Received invalid request - bad mode" & vbCrLf
                    OutBuffer = OutBuffer & "LocationID: " & LocationID.ToString() & vbCrLf
                    OutBuffer = OutBuffer & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
                    OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
                    OutBuffer = OutBuffer & vbCrLf & "Subject: Phone Home  Received invalid request"
                    
                    Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Invalid Request Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID.ToString(), OutBuffer)
            End Select
        Else
            Select Case Mode
                Case "HEALTH"
                    Echo_Health_Msg(LocalServerID, LocationID)
                    TotalTime = DateAndTime.Timer - StartTime
                    Common.Write_Log(LogFile, "HealthMsg Run time=" & Int(TotalTime) & Format$(TotalTime - Fix(TotalTime), ".00") & "(sec)")
                Case Else
                    Common.Write_Log(LogFile, "Invalid Serial - associated LocationID not found-2")
                    Common.Write_Log(LogFile, "Received invalid request!" & "from Serial=" & LocalServerID & " MacAddress=" & MacAddress & " IP=" & LocalServerIP & " server=" & Environment.MachineName)
                    Common.Write_Log(LogFile, "Query String: " & Request.RawUrl & vbCrLf & "Form Data: " & vbCrLf)
                    RawRequest = Get_Raw_Form(Request.InputStream)
                    Common.Write_Log(LogFile, RawRequest)
                    Send_Response_Header("Invalid Serial Number", Connector.CSMajorVersion, Connector.CSMinorVersion, Connector.CSBuild, Connector.CSBuildRevision)
                    
                    OutBuffer = "Phone Home Received Invalid Serial from MacAddress:" & MacAddress & vbCrLf
                    OutBuffer = OutBuffer & "LocationID: " & LocationID & vbCrLf
                    OutBuffer = OutBuffer & "LocalServerSerial: " & LocalServerID & vbCrLf
                    OutBuffer = OutBuffer & "IP: " & LocalServerIP & vbCrLf
                    OutBuffer = OutBuffer & "Server: " & Environment.MachineName & vbCrLf
                    OutBuffer = OutBuffer & vbCrLf & "Subject: Phone Home  Invalid Serial from MacAddress"
                        
                    Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Invalid Serial Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID, OutBuffer)
            End Select
        End If 'locationid="0"
        If Not (Common.LXSadoConn.State = ConnectionState.Closed) Then Common.Close_LogixXS()
    End If
  
    If Not (Common.LRTadoConn.State = ConnectionState.Closed) Then Common.Close_LogixRT()
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
    
    Dim ErrorMsg As String = "Phone Home Error during Local Server Processing" & vbCrLf
    ErrorMsg = ErrorMsg & "LocationID: " & LocationID.ToString() & vbCrLf
    ErrorMsg = ErrorMsg & "Location: " & ExtLocationCode & " " & LocationName & vbCrLf
    ErrorMsg = ErrorMsg & "LocalServerSerial: " & LocalServerID.ToString & vbCrLf
    ErrorMsg = ErrorMsg & "MacAddress: " & MacAddress & vbCrLf
    ErrorMsg = ErrorMsg & "IP: " & LocalServerIP & vbCrLf
    ErrorMsg = ErrorMsg & vbCrLf & "Subject: Phone Home Error during Local Server Processing"
    
    Common.Send_Email(Common.Get_Error_Emails(5), Common.SystemEmailAddress, "Incentive Fetch Error in Processing Message - " & Common.InstallationName & "  LocalServerID " & LocalServerID.ToString(), ErrorMsg)
    
    Common = Nothing
%>